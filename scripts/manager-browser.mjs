import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import {chromium} from 'playwright';
import ExcelJS from 'exceljs';
const url=process.env.MANAGER_TEST_URL||'http://127.0.0.1:5173/#comparison';
const prefix=url.includes('pages.dev')?'manager-live':'manager-local';
const state=JSON.parse(await fs.readFile('reports/manager-state.json','utf8'));
const browser=await chromium.launch({headless:true});
const page=await browser.newPage({acceptDownloads:true,viewport:{width:1600,height:1000}}),errors=[];
page.on('pageerror',e=>errors.push(e.message));
await page.goto(url);await page.waitForLoadState('networkidle');
await page.evaluate(async state=>{
 const db=await new Promise((resolve,reject)=>{const r=indexedDB.open('cbd-generator-local',1);r.onupgradeneeded=()=>{if(!r.result.objectStoreNames.contains('work'))r.result.createObjectStore('work')};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)});
 await new Promise((resolve,reject)=>{const t=db.transaction('work','readwrite');t.objectStore('work').put(state,'comparison-latest');t.oncomplete=resolve;t.onerror=()=>reject(t.error)});db.close();
},state);
await page.reload();await page.getByRole('heading',{name:'Review Material Matches'}).waitFor();
const firstGroup=page.getByLabel('Final Group').first();await firstGroup.selectOption('TRIMS');
await page.getByRole('button',{name:'Undo',exact:true}).click();assert.equal(await page.getByLabel('Final Group').first().inputValue(),'OUTSHELL');
await page.getByRole('button',{name:'Redo',exact:true}).click();assert.equal(await page.getByLabel('Final Group').first().inputValue(),'TRIMS');
await page.waitForTimeout(800);await page.reload();await page.getByRole('heading',{name:'Review Material Matches'}).waitFor();assert.equal(await page.getByLabel('Final Group').first().inputValue(),'TRIMS');
await page.getByLabel('Material Filter').selectOption('group:OUTSHELL');
const downloads=[];
for(const [button,name,count] of [['Export Current Style – Manager Format','current',1],['Export All Styles – Manager Format','all',13],['Export Current Style','detail-current',2],['Export All Styles','detail-all',14]]){
 const pending=page.waitForEvent('download');await page.getByRole('button',{name:button,exact:true}).click();const download=await pending;
 const path=`reports/${prefix}-${name}.xlsx`;await download.saveAs(path);const book=new ExcelJS.Workbook();await book.xlsx.readFile(path);assert.equal(book.worksheets.length,count);
 if(!name.startsWith('detail')){
  for(const [i,sheet] of book.worksheets.entries()){assert.equal(sheet.getImages().length,2);assert.match(sheet.pageSetup.printArea,/^A1:Y\d+$/);const labor=sheet.getColumn(1).values.findIndex(v=>v==='Labor cost');for(const [id,col] of [[state.styleMatches[i].currentId,9],[state.styleMatches[i].referenceId,25]])assert.equal(sheet.getCell(labor,col).value,state.styles.find(s=>s.id===id)?.summary.remarks?.laborCost??null);}
  let found=false;book.worksheets[0].eachRow(row=>{if(String(row.getCell(25).value||'').includes('[GROUP MOVE: OUTSHELL → TRIMS]'))found=true});assert(found);
 }
 downloads.push({button,file:download.suggestedFilename(),path,sheets:count});
}
await page.screenshot({path:`reports/${prefix}-ui.png`,fullPage:true});
assert.deepEqual(errors,[]);await fs.writeFile(`reports/${prefix}-verification.json`,JSON.stringify({url,downloads,errors,groupMoveUndoRedoAndReload:true},null,2));
await browser.close();console.log(JSON.stringify({url,downloads,errors},null,2));
