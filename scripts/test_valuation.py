"""Independent expected outcomes for point-in-time ingestion, LTM and review."""
from valuation_data import *
from tempfile import TemporaryDirectory
from copy import deepcopy
import sys
ROOT=Path(__file__).resolve().parents[1];folder=ROOT/'valuation'; checks=[]
def check(name,want,got):
    ok=math.isclose(want,got,rel_tol=1e-10,abs_tol=1e-10) if isinstance(want,float) else want==got
    checks.append(dict(test=name,expected=want,actual=got,passed=ok))
def fails(fn):
    try: fn();return False
    except ValueError:return True
def main():
    cut=date(2024,3,31);end=date(2023,12,31)
    with TemporaryDirectory() as name:
        work=Path(name);db=connect(work/'test.sqlite')
        for p in (folder/'filings').glob('*.csv'):ingest(db,p)
        check('Byte-identical repeat',0,ingest(db,folder/'filings/periods.csv'))
        accepted,states=snapshot(db,cut)
        check('Partition',61,len(states));check('Valid filings',58,len(accepted))
        check('Future information filtered',0,len(snapshot(db,date(2000,1,1))[0]))
        rows=[r for r in accepted if r['Issuer']=='ACME']
        res=ltm(rows,end);check('Actual quarters revenue',7000.,res['Revenue'])
        check('Annual + six month stub - prior comparable stub',5460.,ltm([r for r in accepted if r['Issuer']=='ARCO'],end)['Revenue'])
        ids=[r for r in rows if r['Kind']=='Q' and r['Start']>='2023']
        check('Missing quarter',None,ltm(ids[:-1],end))
        changed=deepcopy(ids);changed[0]['Currency']='GBP'
        check('Quarter currency change',None,ltm(changed,end))
        overlap=deepcopy(ids[0]);overlap.update(Start='2023-02-01',End='2023-04-30')
        check('Overlapping quarter',None,ltm(ids+[overlap],end))
        raw=csvrows(folder/'filings/periods.csv')[0]
        leap=dict(raw,Start='2023-03-01',End='2024-02-29',Published='2024-03-31',Kind='FY')
        check('Leap-day FY',leap['End'],normalize(leap)['End'])
        check('Fiscal-year transition requires stub',True,fails(lambda:normalize(dict(leap,Start='2023-01-01'))))
        check('Explicit transition stub','STUB',normalize(dict(leap,Start='2023-01-01',Kind='STUB'))['Kind'])
        for k,v in [('Revenue','TBD'),('Units','Unspecified'),('End','2023-02-30'),('Equity','0')]:
            check('Reject '+k,True,fails(lambda:normalize(dict(raw,**{k:v}))))
        check('Missing cash remains unavailable',None,normalize(dict(raw,Cash=''))['Cash'])
        check('Negative equity retained when balance balances',-1.,normalize(dict(raw,Assets='7999',Equity='-1'))['Equity'])
        check('Negative EBITDA multiple',None,divide(123.,-4.))
        check('Zero earnings multiple',None,divide(12.,0.))
        problems=[(k,v) for k,v in states.items() if v[0]=='Quarantined']
        rid=next(k for k,v in problems if 'Conflicting' in v[1]); invalid=next(k for k,v in problems if 'Conflicting' not in v[1])
        check('Approval requires review',True,fails(lambda:review(db,rid,'Approve','Ana','check')))
        review(db,rid,'Review','Ana','Compared original source documents');review(db,rid,'Approve','Ana','Selected source '+rid)
        check('Approved variant reprocessed','Accepted',snapshot(db,cut)[1][rid][0])
        review(db,invalid,'Review','Ana','Unusable revenue');check('Invalid approval blocked',True,fails(lambda:review(db,invalid,'Approve','Ana','invalid')))
        review(db,invalid,'Reject','Ana','Reported revenue is text');check('Rejected stays out','Rejected',snapshot(db,cut)[1][invalid][0])
        check('Append-only action count',4,db.execute('select count(*) from reviews').fetchone()[0])
        hashval=hashlib.sha256((folder/'filings/periods.csv').read_bytes()).hexdigest()
        check('Permanent record ID includes source digest',True,rows[0]['id'].startswith(hashval))
        future=dict(raw,Revenue='9000',Published='2025-04-01');output(work,'revision',META+FIELDS,[[future[k] for k in META+FIELDS]])
        ingest(db,work/'revision.csv');check('Revision hidden until published',7000.,ltm([r for r in snapshot(db,cut)[0] if r['Issuer']=='ACME'],end)['Revenue'])
        generate(db,folder,work/'out',cut,end,'EUR')
        target=next(r for r in csvrows(work/'out/Trading_Data.csv') if r['Issuer']=='ACME')
        check('EV with minority, preferred and other EV components',35185/1.08,float(target['EV']))
        check('Diluted equity value',32500/1.08,float(target['Equity_value']))
        check('No future quotes',0,generate(db,folder,work/'out',date(2024,3,1),end,'USD'))
        db.close()
    (ROOT/'docs/valuation_tests.json').write_text(json.dumps(checks,indent=2))
    print(f"{sum(c['passed'] for c in checks)}/{len(checks)} passed")
    for c in checks:
        if not c['passed']:print(c)
    return all(c['passed'] for c in checks)
if __name__=='__main__':sys.exit(not main())
