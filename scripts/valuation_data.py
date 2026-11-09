"""Prepare reviewable valuation and lineage tables for native Excel models."""
from periods import *
import argparse, os, tempfile
from collections import defaultdict

def csvrows(path):
    with path.open(encoding='utf-8-sig',newline='') as f: return list(csv.DictReader(f))
def output(folder,name,headers,rows):
    with (folder/(name+'.csv')).open('w',encoding='utf-8',newline='') as f:
        w=csv.writer(f);w.writerow(headers);w.writerows(rows)
def divide(a,b):
    return a/b if a is not None and b is not None and b>0 else None
def generate(db,folder,output_dir,cut,end,ccy):
    if end>cut: raise ValueError("Common period ends after valuation date")
    accepted,states=snapshot(db,cut); by=defaultdict(list)
    for r in accepted: by[r['Issuer']].append(r)
    quotes={}; rates={ccy:(1,str(cut),'Identity')}
    for row in csvrows(folder/'fx.csv'):
        day=dt(row['Date']);rate=float(row['Rate'])
        if not math.isfinite(rate) or rate<=0: raise ValueError("Invalid FX")
        if row['To']==ccy and day<=cut and row['From']!=ccy and row['Date']>=rates.get(row['From'],(0,'',''))[1]:
            rates[row['From']]=(rate,row['Date'],row['Source'])
    for row in csvrows(folder/'market.csv'):
        if not row['Source'].strip(): raise ValueError("Market source required")
        if dt(row['Date'])<=cut and row['Date']>=quotes.get(row['Issuer'],{}).get('Date',''):
            quotes[row['Issuer']]=row
    comps=[]; reconciliation=[]; control=[]; lineage=[]; history=[]; observations=[]
    for issuer,rows in sorted(by.items()):
        v=ltm(rows,end);prior=ltm(rows,months(end+DAY,-12)-DAY)
        if not v:
            reconciliation.append([issuer,'Unavailable',str(end),'Gap, overlap, currency change or incompatible duration']);continue
        for old in [prior,v]:
            if old: reconciliation.append([issuer,old['method'],old['end'],old['lineage'],old['start'],old['currency']]+[old[k] for k in FLOWS])
        m=quotes.get(issuer);fx=rates.get(v['currency'])
        if not m or not fx or m['Currency']!=v['currency']:continue
        def converted(k):return None if v[k] is None else v[k]*fx[0]
        price=float(m['Price'])*fx[0]
        if not math.isfinite(price) or price<=0:raise ValueError("Quote must be positive")
        shares=v['DilutedShares']; cap=price*shares if shares is not None and shares>0 else None
        ev_components=[converted(k) for k in ('Debt','Cash','Minority','Preferred','OtherEV')]
        bridge=None if None in ev_components else sum(ev_components)-2*ev_components[1]
        ev=cap+bridge if cap is not None and bridge is not None else None
        net=None if None in ev_components[:2] else ev_components[0]-ev_components[1]
        rev,ebitda,ebit,ni=[converted(k) for k in FLOWS[:4]]
        growth=divide(v['Revenue'],prior['Revenue']) if prior and prior['currency']==v['currency'] else None
        def ratios(f):
            def difference(a,b):return None if f[a] is None or f[b] is None else (f[a]-f[b])*fx[0]
            nd=difference('Debt','Cash')
            return [divide(f['EBITDA'],f['Revenue']),divide(nd, None if f['EBITDA'] is None else f['EBITDA']*fx[0]),divide(f['CurrentAssets'],f['CurrentLiabilities']),difference('CurrentAssets','CurrentLiabilities'),difference('OperatingCashFlow','CapEx')]
        previous=ratios(prior) if prior and prior['currency']==v['currency'] else [None]*5
        for field,old,new,unit in zip(['EBITDA margin','Net debt / EBITDA','Current ratio','Working capital','Operating cash flow less CapEx'],previous,ratios(v),['%','x','x',ccy+' m',ccy+' m']):
            if new is None:continue
            def fmt(x):return f'{100*x:.1f}%' if unit=='%' else f'{x:,.2f} {unit}'
            message=f'{field} was {fmt(new)}.' if old is None else f'{field} moved from {fmt(old)} to {fmt(new)}.'
            observations.append([issuer,field,old,new,unit,message,v['lineage'],prior['lineage'] if prior else ''])
        comps.append([issuer,m['Industry'],m['Geography'],m['BusinessModel'],m['Date'],v['end'],rev,ebitda,ebit,ni,shares,price,cap,*ev_components,net,ev]+[divide(ev,x) if ev is not None and ev>=0 else None for x in (rev,ebitda,ebit)]+[divide(cap,ni),divide(ebitda,rev),growth-1 if growth is not None else None,divide(net,ebitda),bridge,ccy,m['Source'],v['lineage'],fx[1],fx[2]])
        for key in FLOWS+['CurrentAssets','CurrentLiabilities','Debt','Cash']:
            history.append([issuer,key,prior[key]*fx[0] if prior and prior['currency']==v['currency'] and prior[key] is not None else None,converted(key),ccy,v['end'],v['lineage']])
    batches=list(db.execute("select * from batches"))
    source_totals=defaultdict(lambda:[0,0.,0.])
    for r in db.execute("select r.*,b.sha,b.file,b.stamp from records r join batches b on r.batch=b.id order by r.rowid"):
        raw=json.loads(r['raw']);norm=json.loads(r['norm']) if r['norm'] else {}
        state,reason=states[r['id']]
        if norm:
            group=source_totals[(r['file'],raw['Kind'],raw['Currency'],raw['Units'])]
            group[0]+=1;group[1]+=float(raw['Revenue']);group[2]+=norm['Revenue']
        lineage.append([r['id'],r['batch'],r['sha'],r['file'],'CSV',r['rownum'],r['stamp'],raw['Issuer'],raw['Kind'],raw['Start'],raw['End'],raw['Published'],raw['Currency'],raw['Units'],raw['Revenue'],norm.get('Revenue'),state,reason,raw['Source'],r['id'] if state=='Accepted' else '',sum(norm.get(k) is None for k in FIELDS),r['raw']])
    counts=Counter(s for s,_ in states.values())
    control=[['Imported files',len(batches),'INFO'],['Imported rows',len(states),'INFO']]+[[k,counts[k],'WARNING' if k=='Quarantined' and counts[k] else 'INFO'] for k in ('Accepted','Quarantined','Rejected','Duplicate','Superseded','Deferred')]
    control.append(['Disposition difference',len(states)-sum(counts.values()),'PASS' if len(states)==sum(counts.values()) else 'FAIL'])
    for (file,kind,currency,unit),(n,original,normalized) in source_totals.items():
        diff=original*SCALES[unit]-normalized
        control.append([f'{file} | {kind} | {currency} {unit}',n,'PASS' if abs(diff)<max(1e-6,abs(normalized)*1e-10) else 'FAIL',original,normalized,diff])
    output_dir.mkdir(exist_ok=True,parents=True)
    output(output_dir,'Trading_Data','Issuer Industry Geography Business_model Quote_date LTM_end Revenue EBITDA EBIT Net_income Diluted_shares Price Equity_value Debt Cash Minority Preferred Other_EV Net_debt EV EV_Revenue EV_EBITDA EV_EBIT PE EBITDA_margin Revenue_growth Net_debt_EBITDA EV_bridge Currency Quote_source Financial_record_IDs FX_date FX_source'.split(),comps)
    output(output_dir,'LTM_Reconciliation',['Issuer','Method','End','Signed source IDs / issue','Start','Currency']+FLOWS,reconciliation)
    output(output_dir,'Record_Lineage','Record_ID Batch_ID SHA256 File Sheet Row Imported Issuer Kind Start End Published Currency Units Source_revenue Revenue_m Status Decision Source Accepted_ID Missing_fields Original_values'.split(),lineage)
    output(output_dir,'Quality_Control',['Control','Count','Status','Source revenue','Normalized revenue (m)','Difference (m)'],control)
    output(output_dir,'Valuation_History',['Issuer','Metric','Prior LTM','Current LTM','Currency','End','Current IDs'],history)
    output(output_dir,'Observations',['Issuer','Metric','Prior','Current','Unit','Observation','Current IDs','Prior IDs'],observations)
    output(output_dir,'Review_History',['Sequence','Record','Action','Reviewer','Notes','UTC'],[tuple(x) for x in db.execute('select * from reviews order by seq')])
    return len(comps)

def main():
    base=Path(__file__).resolve().parents[1];p=argparse.ArgumentParser()
    p.add_argument('--folder',type=Path,default=base/'valuation');p.add_argument('--date',default='2024-03-31');p.add_argument('--end',default='2023-12-31');p.add_argument('--currency',default='USD')
    p.add_argument('--import-file',type=Path);p.add_argument('--review',nargs=4,metavar=('ID','ACTION','REVIEWER','NOTE'))
    args=p.parse_args();folder=args.folder; db=connect(folder/'valuation.sqlite')
    if args.review:review(db,*args.review)
    elif args.import_file:ingest(db,args.import_file)
    else:
        for file in sorted((folder/'filings').glob('*.csv')): ingest(db,file)
    generate(db,folder,folder/'tables',dt(args.date),dt(args.end),args.currency)
if __name__=='__main__':main()
