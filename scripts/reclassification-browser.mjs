import {chromium} from 'playwright';
import {readFile,writeFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
const base=process.argv[2]||'http://localhost:5173',state=JSON.parse(await readFile('reports/verification-state.json','utf8'));
const browser=await chromium.launch({headless:true});
try {
const page=await browser.newPage({acceptDownloads:true});await page.goto(`${base}/#comparison`);
await page.evaluate(async state=>{await new Promise((resolve,reject)=>{const req=indexedDB.open('cbd-generator-local',1);req.onupgradeneeded=()=>req.result.createObjectStore('work');req.onsuccess=()=>{const db=req.result,tx=db.transaction('work','readwrite');tx.objectStore('work').put(state,'comparison-latest');tx.oncomplete=()=>{db.close();resolve()};tx.onerror=reject};req.onerror=reject})},state);
await page.reload();await page.getByRole('button',{name:'Download Matching Audit'}).waitFor({timeout:30000});
const row=page.locator('.material-table tbody tr').filter({has:page.locator('.reference-material',{hasText:'ES#N FD 160D'})});await row.getByLabel('Final Group').selectOption('OUTSHELL');await row.getByLabel('Final Group').selectOption('TRIMS');
await page.getByRole('button',{name:'Undo',exact:true}).click();if(await row.getByLabel('Final Group').inputValue()!=='OUTSHELL')throw Error('Undo failed');await page.getByRole('button',{name:'Redo',exact:true}).click();await page.waitForTimeout(600);await page.reload();await page.getByRole('button',{name:'Download Matching Audit'}).waitFor();
const results=[];
for(const [scope,name] of [['current','Export Current Style'],['all','Export All Styles']]){const event=page.waitForEvent('download');await page.getByRole('button',{name,exact:true}).click();const download=await event,path=`reports/${base.startsWith('https')?'live':'local'}-${scope}.xlsx`;await download.saveAs(path);const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(path);const sheet=wb.worksheets[1];let moved=false,unit=false;sheet.eachRow(r=>{if(String(r.getCell(14).value).includes('수동 그룹 변경 (OUTSHELL → TRIMS)'))moved=true;if(String(r.getCell(14).value).includes('Unit Changed')){unit=true;if(r.getCell(7).value!==null||r.getCell(10).value!==null)throw Error('Unit deltas not blank')}r.eachCell(c=>{if(/#VALUE!|#REF!|#DIV\/0!|#NAME\?/.test(String(c.value)))throw Error('Excel error')})});if(!moved||!unit)throw Error('Missing direction/unit change');if(Math.abs(sheet.getCell('B12').result-3.1731)>1e-8)throw Error('Total changed');results.push({scope,sheets:wb.worksheets.length,moved,unit,total:sheet.getCell('B12').result,path});}
await writeFile(`reports/${base.startsWith('https')?'live':'local'}-verification.json`,JSON.stringify({base,data:'Synthetic verification fixture; original CBD files unavailable',results},null,2));console.log(JSON.stringify(results));
} finally {await browser.close()}
