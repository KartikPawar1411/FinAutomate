"""Build and verify the synthetic release. Run only against the shipped demo portfolio."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import subprocess,sys,tempfile
root=Path(__file__).resolve().parents[1]
out=root.parent/'FinAutomate-Platform.zip'
skip={'__pycache__','.git','node_modules'}
files=[p for p in sorted(root.rglob('*')) if p.is_file() and not skip.intersection(p.parts) and not p.name.startswith(('~$','build_','Before_reset_')) and p.relative_to(root).as_posix()!='valuation/records.sqlite']
with ZipFile(out,'w',ZIP_DEFLATED) as z:
    for p in files:z.write(p,root.name+'/'+p.relative_to(root).as_posix())
with ZipFile(out) as z,tempfile.TemporaryDirectory() as tmp:
    assert z.testzip() is None
    for p in files:assert z.read(root.name+'/'+p.relative_to(root).as_posix())==p.read_bytes()
    z.extractall(tmp)
    project=Path(tmp)/root.name
    for name in ('test_workbook.py','test_valuation.py'):
        subprocess.run([sys.executable,str(project/'scripts'/name)],cwd=project,check=True)
    subprocess.run([sys.executable,str(project/'scripts/valuation_data.py')],cwd=project,check=True)
    print('Extracted project refreshed and verified:',len(files),'files;',out.stat().st_size,'bytes')
