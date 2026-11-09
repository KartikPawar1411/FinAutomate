"""Check Excel's saved results against the six source files. Python 3, standard library."""
from pathlib import Path
import csv
from decimal import Decimal as D, InvalidOperation
import math
import re
import sys
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NA, NM = "N/A", "N/M"

def require(condition, message):
    if not condition:
        raise AssertionError(message)

def as_num(v):
    return None if v == "" else D(v)

def quotient(a, b, nonnegative=False):
    if a is None or b is None:
        return NA
    if b <= 0 or (nonnegative and a < 0):
        return NM
    return a/b

def expected():
    filings = {}
    for path in sorted((ROOT/"sample_data").glob("company_*.csv")):
        with path.open(encoding="utf-8-sig", newline="") as stream:
            for row in csv.DictReader(stream):
                try:
                    vals = {k:as_num(row[k]) for k in list(row)[6:35]}
                except InvalidOperation:
                    require(row["Company_ID"]=="DBRA" and row["Reporting_Period"]=="FY2023","Unexpected malformed demo filing")
                    continue
                filings[row["Company_ID"],int(row["Reporting_Period"][2:])] = vals
    out = {}
    for (issuer,year), a in filings.items():
        prior = filings.get((issuer,year-1))
        rev,net,ebit,ebitda,eq = (a[k] for k in ("Revenue","Net_Income","EBIT","EBITDA","Total_Equity"))
        roe=roic=NA
        if prior:
            old = prior["Total_Equity"]
            roe = quotient(net,(old+eq)/2) if min(old,eq)>0 else NM
            capital = prior["Total_Debt"] + old - prior["Cash_Equivalents"]
            tax,ebt=a["Income_Tax"],a["EBT"]
            roic=quotient(ebit*(1-tax/ebt),capital) if ebt>0 and 0<=tax<=ebt else NM
        quick = None if a["Inventory"] is None else a["Current_Assets"]-a["Inventory"]
        growth=[]
        for k in ("Revenue","EBITDA","Net_Income"):
            val=quotient(a[k],prior[k]) if prior else NA
            growth.append(val-1 if isinstance(val,D) else val)
        out[issuer,"FY"+str(year)] = [quotient(x,rev) for x in (a["Gross_Profit"],ebitda,ebit,net)]+[
            roe,roic,quotient(a["Current_Assets"],a["Current_Liabilities"]),quotient(quick,a["Current_Liabilities"]),
            quotient(a["Total_Debt"],eq,True),quotient(a["Total_Debt"],ebitda,True),quotient(ebit,a["Interest_Expense"])]+growth+[
            quotient(a["Enterprise_Value"],rev,True),quotient(a["Enterprise_Value"],ebitda,True),quotient(a["Market_Cap"],net,True)]
    return out

def test(path):
    result=expected()
    with ZipFile(path) as z:
        require("xl/vbaProject.bin" in z.namelist(),"No embedded VBA")
        ss=ET.fromstring(z.read("xl/sharedStrings.xml"))
        strings=["".join(si.itertext()) for si in ss]
        wb=ET.fromstring(z.read("xl/workbook.xml"))
        sheets=[s.attrib["name"] for s in wb.findall("x:sheets/x:sheet",NS)]
        require(set("Home Dashboard Report Company_Comparison Financial_Ratios Financial_Statements Clean_Data Raw_Data Audit_Log Settings Instructions".split()) <= set(sheets),"Missing original sheet")
        ix=sheets.index("Financial_Ratios")+1
        root=ET.fromstring(z.read(f"xl/worksheets/sheet{ix}.xml"))
        def value(c):
            v=c.findtext("x:v","",NS)
            return strings[int(v)] if c.get("t")=="s" else v
        rows=root.findall("x:sheetData/x:row",NS)
        counted=0
        for row in rows[1:]:
            cells={re.sub(r"\d+$","",c.get("r")):c for c in row.findall("x:c",NS)}
            if "A" not in cells or not value(cells["A"]):
                continue
            key=value(cells["A"]),value(cells["B"])
            wanted=result.pop(key)
            for col,expect in zip("DEFGHIJKLMNOPQRST",wanted):
                c=cells[col]; actual=value(c)
                require(c.find("x:f",NS) is not None, f"Missing Excel formula {key} {col}")
                ok=math.isclose(float(expect),float(actual),rel_tol=1e-10,abs_tol=1e-11) if isinstance(expect,D) else expect==actual
                require(ok,f"{key} {col}: expected {expect}, Excel saved {actual}")
                counted+=1
        require(not result,"Missing filings: "+str(list(result)))
        for sheet in [n for n in z.namelist() if re.fullmatch(r"xl/worksheets/sheet\d+\.xml",n)]:
            require(not ET.fromstring(z.read(sheet)).findall('.//x:c[@t="e"]',NS),f"Excel error on {sheet}")
        print(f"PASS: {counted} Excel formula results independently reconciled; {counted//17} accepted filings, 17 metrics. No cached cell errors.")

if __name__=="__main__":
    test(Path(sys.argv[1]) if len(sys.argv)>1 else ROOT/"FinAutomate.xlsm")
