"""Isolated worker per size so peak working sets are comparable on Windows."""
from valuation_data import *
import ctypes,platform,subprocess,sys

def peak_mb():
    class Memory(ctypes.Structure):
        _fields_=[('cb',ctypes.c_ulong),('faults',ctypes.c_ulong)]+[(x,ctypes.c_size_t) for x in 'peak current paged pages nonpaged nonpages pagefile pagepeak'.split()]
    m=Memory();m.cb=ctypes.sizeof(m)
    if ctypes.windll.psapi.GetProcessMemoryInfo(ctypes.c_void_p(-1),ctypes.byref(m),m.cb):return round(m.peak/1048576,1)
def run(n):
    root=Path(__file__).resolve().parents[1]
    with tempfile.TemporaryDirectory() as tmp:
        folder=Path(tmp); path=folder/'stress.csv'; base=csvrows(root/'valuation/filings/periods.csv')[0]
        with path.open('w',newline='') as f:
            writer=csv.DictWriter(f,fieldnames=META+FIELDS);writer.writeheader()
            for i in range(n):
                record=dict(base,Issuer=f'S{i//4:06}',Kind='Q');start=date(2023,1+3*(i%4),1);end=months(start,3)-DAY
                record.update(Start=str(start),End=str(end),Published=str(end+timedelta(days=40)))
                if i%100==0:record['Revenue']='invalid'
                writer.writerow(record)
        db=connect(folder/'bench.sqlite'); result={'rows':n}; ingest(db,path,result)
        tick=perf_counter(); accepted,states=snapshot(db,date(2024,3,31));result['reconciliation_seconds']=perf_counter()-tick
        tick=perf_counter();by=defaultdict(list)
        for r in accepted:by[r['Issuer']].append(r)
        result['eligible_ltm']=sum(ltm(v,date(2023,12,31)) is not None for v in by.values())
        result['ltm_seconds']=perf_counter()-tick
        result.update(accepted=len(accepted),quarantined=len(states)-len(accepted),database_bytes=(folder/'bench.sqlite').stat().st_size,python_peak_MiB=peak_mb())
        db.close();print(json.dumps(result))
if __name__=='__main__':
    if len(sys.argv)>1:run(int(sys.argv[1]))
    else:
        rows=[json.loads(subprocess.check_output([sys.executable,__file__,str(n)],text=True)) for n in (1000,10000,50000)]
        report=dict(os=platform.platform(),processor=platform.processor(),logical_cpus=os.cpu_count(),python=sys.version,scope='SQLite import, normalization/validation, dispositions and LTM. Excludes Excel, PDF and fixture generation.',measurements=rows)
        path=Path(__file__).resolve().parents[1]/'docs/scale.json';path.write_text(json.dumps(report,indent=2));print(json.dumps(rows,indent=2))
