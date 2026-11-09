"""Point-in-time financial records. All amounts in normalized local millions."""
import csv, hashlib, io, json, math, sqlite3, uuid
from collections import defaultdict, Counter
from datetime import date, datetime, timedelta, timezone
from calendar import monthrange
from pathlib import Path
from time import perf_counter

FLOWS = "Revenue EBITDA EBIT NetIncome DA CapEx OperatingCashFlow".split()
BAL = "Debt Cash Minority Preferred OtherEV DilutedShares CurrentAssets CurrentLiabilities Assets Liabilities Equity EBT Tax".split()
FIELDS = FLOWS + BAL
META = "Issuer Kind Start End Published Currency Units Source".split()
SCALES = {"Actual":1e-6,"Thousands":.001,"Millions":1,"Billions":1000}
DAY = timedelta(days=1)
def dt(s): return date.fromisoformat(s)
def months(d,n):
    y,m=divmod(d.year*12+d.month-1+n,12)
    return date(y,m+1,min(d.day,monthrange(y,m+1)[1]))
def now(): return datetime.now(timezone.utc).isoformat()
def connect(path):
    db=sqlite3.connect(path)
    db.row_factory=sqlite3.Row
    db.executescript("""
    CREATE TABLE IF NOT EXISTS batches(id TEXT PRIMARY KEY, sha TEXT UNIQUE, file TEXT, stamp TEXT);
    CREATE TABLE IF NOT EXISTS records(id TEXT PRIMARY KEY,batch TEXT,rownum INTEGER,raw TEXT,norm TEXT,issue TEXT);
    CREATE TABLE IF NOT EXISTS reviews(seq INTEGER PRIMARY KEY,id TEXT,action TEXT,reviewer TEXT,note TEXT,stamp TEXT);
    """)
    return db
def normalize(row):
    for k in META:
        if not row[k].strip(): raise ValueError("Missing "+k)
    r={k:row[k].strip() for k in META}
    start,end,pub=map(dt,(r['Start'],r['End'],r['Published']))
    if start>end or pub<end: raise ValueError("Invalid period or publication dates")
    k=r["Kind"]
    if k not in ("FY","Q","YTD","STUB"): raise ValueError("Unknown period type")
    if k in ("Q","FY") and months(start,3 if k=="Q" else 12)-DAY!=end: raise ValueError("Nonstandard duration: supply STUB")
    if r['Units'] not in SCALES: raise ValueError("Unknown unit")
    if len(r['Currency'])!=3 or not r['Currency'].isalpha(): raise ValueError("Currency")
    r['Currency']=r['Currency'].upper()
    for field in FIELDS:
        value=row.get(field,'').strip()
        r[field]=float(value)*SCALES[r['Units']] if value else None
        if r[field] is not None and not math.isfinite(r[field]): raise ValueError("Nonfinite "+field)
    if r['Revenue'] is None or r['Revenue']<0: raise ValueError("Missing or negative revenue")
    for lhs,a,b,sign in [('Assets','Liabilities','Equity',1),('NetIncome','EBT','Tax',-1)]:
        if all(r[f] is not None for f in (lhs,a,b)) and abs(r[lhs]-r[a]-sign*r[b])>max(1e-6,abs(r[lhs])*1e-4): raise ValueError(lhs+" reconciliation failed")
    return r
def ingest(db,path,timings=None):
    t=perf_counter(); data=Path(path).read_bytes(); sha=hashlib.sha256(data).hexdigest()
    if db.execute("select 1 from batches where sha=?",(sha,)).fetchone(): return 0
    reader=csv.DictReader(io.StringIO(data.decode('utf-8-sig'),newline=''))
    if not set(META+['Revenue'])<=set(reader.fieldnames or []) or len(set(reader.fieldnames))!=len(reader.fieldnames): raise ValueError("Missing or repeated headers")
    batch=uuid.uuid4().hex; count=0; validation=0
    with db:
        db.execute("insert into batches values(?,?,?,?)",(batch,sha,Path(path).name,now()))
        for count,row in enumerate(reader,1):
            if None in row or None in row.values(): raise ValueError("Malformed row width")
            tick=perf_counter()
            try: norm=json.dumps(normalize(row)); issue=''
            except (ValueError,OverflowError) as e: norm=None; issue=str(e)
            validation+=perf_counter()-tick
            db.execute("insert into records values(?,?,?,?,?,?)",(sha+'-'+str(count+1),batch,count+1,json.dumps(row),norm,issue))
    if timings is not None: timings.update(import_io_sql_seconds=perf_counter()-t-validation,normalize_and_validate_seconds=validation)
    return count
def review(db,rid,decision,person,note):
    if decision not in ('Approve','Reject','Review') or not person.strip() or not note.strip(): raise ValueError("Decision, reviewer and reason required")
    rec=db.execute("select * from records where id=?",(rid,)).fetchone()
    if rec is None: raise ValueError("Unknown record")
    if decision=='Approve': normalize(json.loads(rec['raw']))
    last=db.execute("select action from reviews where id=? order by seq desc limit 1",(rid,)).fetchone()
    if decision!='Review' and (not last or last[0]!='Review'): raise ValueError("Review first")
    with db: db.execute("insert into reviews(id,action,reviewer,note,stamp) values(?,?,?,?,?)",(rid,decision,person,note,now()))
def snapshot(db,cut):
    groups=defaultdict(list); states={}; accepted=[]; logs={}
    for log in db.execute("select * from reviews order by seq"): logs[log['id']]=log
    for rec in db.execute("select * from records order by rowid"):
        rid=rec['id']; decision=logs[rid]['action'] if rid in logs else ''
        if decision=='Reject' and not rec['norm']: states[rid]=('Rejected','Reviewer decision'); continue
        if not rec['norm']: states[rid]=('Quarantined',rec['issue']); continue
        r=json.loads(rec['norm'])
        if dt(r['Published'])>cut: states[rid]=('Deferred','Published after valuation date'); continue
        key=tuple(r[x] for x in ('Issuer','Kind','Start','End'))
        groups[key].append((r,rid,decision))
    for versions in groups.values():
        publication=max(r['Published'] for r,_,_ in versions)
        candidates=[]
        for r,rid,decision in versions:
            if r['Published']<publication: states[rid]=('Superseded','Later published revision')
            else: candidates.append((r,rid,decision))
        selected=[v for v in candidates if v[2]=='Approve']
        remaining=[v for v in candidates if v[2]!='Reject']
        if not remaining:
            for _,rid,_ in candidates: states[rid]=('Rejected','Reviewer decision')
            continue
        if len(selected)==1: winner=selected[0]
        elif len({tuple(r[k] for k in ['Currency']+FIELDS) for r,_,_ in candidates})==1: winner=remaining[0]
        else:
            for _,rid,decision in candidates: states[rid]=('Rejected','Reviewer decision') if decision=='Reject' else ('Quarantined','Conflicting versions on same publication date')
            continue
        for r,rid,decision in candidates:
            if rid==winner[1]: states[rid]=('Accepted','Validated'); accepted.append(dict(r,id=rid))
            elif decision=='Reject': states[rid]=('Rejected','Reviewer decision')
            else: states[rid]=('Duplicate' if all(r[k]==winner[0][k] for k in ['Currency']+FIELDS) else 'Superseded','Selected '+winner[1])
    return accepted,states
def ltm(rows,end):
    """Exact common endpoint, no extrapolation across gaps or fiscal changes."""
    start=months(end+DAY,-12); lookup={(r['Kind'],dt(r['Start']),dt(r['End'])):r for r in rows}
    pieces=[]; method=''
    for off in range(0,12,3):
        a,b=months(start,off),months(start,off+3)-DAY
        if ('Q',a,b) in lookup: pieces.append((1,lookup['Q',a,b]))
    if len(pieces)==4: method='4 quarters'
    elif ('FY',start,end) in lookup: pieces=[(1,lookup['FY',start,end])]; method='Fiscal year'
    else:
        pieces=[]
        for (kind,a,b),cur in lookup.items():
            if kind not in ('YTD','STUB') or b!=end or not start<a<=end: continue
            old_start=months(a,-12)
            fy=lookup.get(('FY',old_start,a-DAY)); prev=lookup.get((kind,old_start,start-DAY))
            if fy and prev: pieces=[(1,fy),(1,cur),(-1,prev)];method='FY + YTD - prior YTD';break
    if not pieces or len({r['Currency'] for _,r in pieces})!=1: return None
    # Contradictory quarterly intervals cannot both contribute to a common-year window.
    quarters=sorted((dt(r['Start']),dt(r['End'])) for r in rows if r['Kind']=='Q' and dt(r['Start'])<=end and dt(r['End'])>=start)
    if any(b[0]<=a[1] for a,b in zip(quarters,quarters[1:])): return None
    result={k:sum(sign*r[k] for sign,r in pieces) if all(r[k] is not None for _,r in pieces) else None for k in FLOWS}
    closing=next(r for _,r in reversed(pieces) if r['End']==end.isoformat())
    result.update({k:closing[k] for k in BAL})
    return dict(result,method=method,start=str(start),end=str(end),currency=closing['Currency'],lineage=" ".join(('+' if sign==1 else '-')+r['id'] for sign,r in pieces))
