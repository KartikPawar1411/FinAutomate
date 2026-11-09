"""Render the measured native trial stages as a standalone vector chart."""
import json
from pathlib import Path
root=Path(__file__).resolve().parents[1]
rows=json.loads((root/'docs/scale_excel.json').read_text(encoding='utf-8-sig'))
series=[('import_io_sql_seconds','Import','#456780'),('normalize_and_validate_seconds','Validate','#79afaa'),('reconcile_ltm_and_stage_seconds','Reconcile and stage','#c8a459'),('refresh_workbook_seconds','Excel refresh','#16746e'),('pdf_seconds','PDF','#79868e')]
width=630
maximum=max(sum(r[k] for k,_,_ in series) for r in rows)
parts=['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 970 260" role="img" aria-label="Benchmark stage times in seconds"><rect width="970" height="260" fill="white"/><g font-family="Arial, sans-serif" fill="#203d50">','<text x="28" y="34" font-size="23">Processing time by filing volume</text>','<text x="28" y="58" font-size="14">One Windows Excel trial per size · four quoted issuers · seconds</text>']
for i,row in enumerate(rows):
    y=78+i*43;x=150;total=0
    parts.append(f'<text x="28" y="{y+18}">{row["rows"]:,} records</text>')
    for key,_,color in series:
        t=row[key];span=width*t/maximum;total+=t
        parts.append(f'<rect x="{x:.2f}" y="{y}" width="{span:.2f}" height="22" fill="{color}"/>')
        x+=span
    parts.append(f'<text x="{x+8:.1f}" y="{y+17}">{total:.2f} s</text>')
x=28
for _,label,color in series:
    parts.append(f'<rect x="{x}" y="221" width="14" height="14" fill="{color}"/><text x="{x+20}" y="233" font-size="14">{label}</text>');x+=len(label)*8+42
(root/'screenshots/Performance.svg').write_text(''.join(parts)+'</g></svg>',encoding='utf-8')
