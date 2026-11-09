"""Deterministic synthetic valuation case. No observed security prices."""
from valuation_data import *
def build(folder):
    sources=folder/'filings';sources.mkdir(exist_ok=True,parents=True)
    rows=[];market=[]
    cases=[('ACME',1,1.0,'USD',12),('ARCO',.78,1.1,'USD',6),('BOLT',1.25,.9,'EUR',9),('CAST',.82,1.3,'USD',3)]
    for name,size,premium,ccy,fye in cases:
        def data(start,finish,kind,flows):
            balances=[4100,1500,30,45,10,500,5200,2900,15000,8000,7000,None,None]
            return [name,kind,str(start),str(finish),str(finish+timedelta(days=40)),ccy,'Millions','Synthetic valuation case v2']+[round(v*size,8) if v is not None else '' for v in flows+balances]
        series=[]
        for year in (2021,2022,2023):
            for q in range(4):
                start=date(year,1+3*q,1);end=months(start,3)-DAY
                rev={2021:5500,2022:6200,2023:7000}[year]*[.22,.24,.25,.29][q]
                margin_delta={'ACME':0,'ARCO':.025,'BOLT':-.035,'CAST':.06}[name]
                ebitda_margin=.20+.01*(year-2021)+margin_delta
                amortization=.02+.01*(year-2021)
                flows=[rev,rev*ebitda_margin,rev*(ebitda_margin-amortization),rev*(.13+.65*margin_delta),rev*amortization,rev*.065,rev*(.16+margin_delta)]
                series.append((start,end,flows))
                # One issuer demonstrates the annual plus comparable YTD bridge.
                if name!='ARCO' or end<=date(2022,12,31): rows.append(data(start,end,'Q',flows))
        for y in (2022,2023):
            finish=date(y,fye,monthrange(y,fye)[1]);start=months(finish+DAY,-12)
            subset=[f for a,b,f in series if start<=a and b<=finish]
            if len(subset)==4:rows.append(data(start,finish,'FY',list(map(sum,zip(*subset)))))
            if fye!=12:
                subset=[f for a,b,f in series if finish<a and b<=date(y,12,31)]
                rows.append(data(finish+DAY,date(y,12,31),'YTD',list(map(sum,zip(*subset)))))
        market.append([name,'2024-03-28',ccy,65*premium,'Industrial components','Europe' if ccy=='EUR' else 'North America','Component manufacturing','Synthetic market case v2'])
    output(sources,'periods',META+FIELDS,rows)
    base=list(rows[0]);base[0]='REVIEWCASE'
    revised=list(base);revised[META.index('Source')]='Synthetic conflicting extract';revised[len(META)]+=20
    invalid=list(base);invalid[0]='INVALIDCASE';invalid[len(META)]='TBD'
    output(sources,'review_cases',META+FIELDS,[base,revised,invalid])
    output(folder,'market','Issuer Date Currency Price Industry Geography BusinessModel Source'.split(),market)
    output(folder,'fx','From To Date Rate Source'.split(),[['EUR','USD','2024-03-28',1.08,'Synthetic FX assumption'],['USD','EUR','2024-03-28',1/1.08,'Synthetic FX assumption']])
if __name__=='__main__':build(Path(__file__).resolve().parents[1]/'valuation')
