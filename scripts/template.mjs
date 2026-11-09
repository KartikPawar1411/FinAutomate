import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
const {Workbook,SpreadsheetFile}=await import(process.env.ARTIFACT_TOOL_URL || '@oai/artifact-tool');
process.chdir(fileURLToPath(new URL('..',import.meta.url)));
const wb=Workbook.create();
const columns=(await fs.readFile('sample_data/company_acme.csv','utf8')).split(/\r?\n/)[0].split(',');
const sheets=['Home','Dashboard','Report','Company_Comparison','Financial_Ratios','Financial_Statements','Clean_Data','Raw_Data','Audit_Log','Settings','Instructions'];
const notes=['FinAutomate','Company review','Financial report','Peer review','Financial ratios','Accepted statements','Cleaning ledger','Source ledger','Automation log','Run settings','Operator guide'];
sheets.forEach((name,i)=>{const s=wb.worksheets.add(name);const headers=['Raw_Data','Clean_Data','Financial_Statements'].includes(name)?columns:[notes[i]];
s.getRangeByIndexes(0,0,1,headers.length).values=[headers];const h=s.getRangeByIndexes(0,0,1,headers.length);h.format.fill='#17384B';h.format.font.color='#FFFFFF';h.format.font.bold=true;h.format.columnWidth=18;
if(name==='Home'||name==='Instructions'){s.getRange('A3').values=[['Open the assembled .xlsm to use the application.']];s.getRange('A4').values=[['Manual assembly: see docs/DEVELOPER_GUIDE.md.']];s.getRange('A1:A4').format.columnWidth=64;}
});
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'Home',range:'A1:A4',maxChars:650})).ndjson);
await (await SpreadsheetFile.exportXlsx(wb)).save('FinAutomate_Template.xlsx');
