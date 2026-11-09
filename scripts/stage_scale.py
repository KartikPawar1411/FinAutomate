"""Create a disposable, exact-size portfolio including the demonstration target."""
from valuation_data import *
import shutil,sys
root=Path(__file__).resolve().parents[1];n=int(sys.argv[1]);folder=Path(sys.argv[2])
folder.mkdir(parents=True,exist_ok=True)
for file in ('market.csv','fx.csv'):shutil.copy(root/'valuation'/file,folder/file)
db=connect(folder/'bench.sqlite');timing={};totals=Counter();count=0
for file in sorted((root/'valuation/filings').glob('*.csv')):
    count+=ingest(db,file,timing);totals.update(timing)
base=csvrows(root/'valuation/filings/periods.csv')[0]
path=folder/'stress.csv'
with path.open('w',newline='') as f:
    writer=csv.DictWriter(f,fieldnames=META+FIELDS);writer.writeheader()
    for i in range(n-count):
        quarter_start=date(2023,1+(i%4)*3,1);end=months(quarter_start,3)-DAY
        writer.writerow(dict(base,Issuer=f'LOAD{i//4:06}',Start=str(quarter_start),End=str(end),Published=str(end+timedelta(days=40)),Revenue='bad' if i%100==0 else base['Revenue']))
ingest(db,path,timing);totals.update(timing)
t=perf_counter()
generate(db,folder,folder/'tables',date(2024,3,31),date(2023,12,31),'USD')
totals['reconcile_ltm_and_stage_seconds']=perf_counter()-t
counts=Counter(x['Status'] for x in csvrows(folder/'tables/Record_Lineage.csv'))
from benchmark_scale import peak_mb
print(json.dumps(dict(totals,rows=n,dispositions=dict(counts),python_peak_MiB=peak_mb())))
db.close()
