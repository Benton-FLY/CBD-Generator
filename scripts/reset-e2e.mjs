import { chromium } from 'playwright';
import ExcelJS from 'exceljs';
import { strict as assert } from 'node:assert';

const baseUrl=process.env.E2E_URL||'http://127.0.0.1:4173';
const material=(id,group,cost,patch={})=>({id,item:id.toUpperCase(),width:'58',unit:'YD',group,included:true,baseCost:cost,adjustedCost:cost,baseUsage:1,adjustedUsage:1,baseLoss:.05,additionalLoss:.05,remark:'',split:false,sources:[],...patch});
const styles=[
  {id:'one',name:'F-16 GLOVE',sourceFile:'one.xlsx',sourceSheet:'BOM',laborRemark:'keep',finalFob:48.98,erpMaterialCost:6,materials:[material('shell','OUTSHELL',4,{adjustedCost:4.2,remark:'KEEP'}),material('trim','TRIMS',2,{adjustedUsage:1.5,included:false}),material('label','LABEL & PACKAGING',1,{additionalLoss:.12})]},
  {id:'two',name:'OTHER STYLE',sourceFile:'two.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('other','TRIMS',9,{adjustedCost:10})]},
];
const saved={version:1,styles,active:'one',settings:{exchangeRate:900,defaultLoss:.05},dict:{},selections:{},groupFilters:{},fobs:[],scope:'all',target:'cost',operation:'base-percent',bulkValue:5,savedAt:new Date().toISOString()};
const browser=await chromium.launch({headless:true});const page=await browser.newPage({acceptDownloads:true});
const row=name=>page.getByRole('row').filter({has:page.getByText(name,{exact:true})});
try{
  await page.goto(baseUrl,{waitUntil:'networkidle'});await page.evaluate(async value=>new Promise((resolve,reject)=>{const request=indexedDB.open('cbd-generator-local',1);request.onupgradeneeded=()=>request.result.createObjectStore('work');request.onerror=()=>reject(request.error);request.onsuccess=()=>{const db=request.result,tx=db.transaction('work','readwrite');tx.objectStore('work').put(value,'latest');tx.oncomplete=()=>{db.close();resolve()}}}),saved);await page.reload({waitUntil:'networkidle'});
  await page.getByRole('button',{name:'일괄조정 전체 초기화'}).click();assert.equal(await page.getByRole('radio',{name:'현재 스타일 전체 그룹'}).isChecked(),true);await page.getByRole('button',{name:'초기화 실행'}).click();await page.getByText(/3개 그룹, 3개 조정값/).waitFor();
  assert.equal(await row('SHELL').getByRole('spinbutton').first().inputValue(),'4.0000');assert.equal(await row('TRIM').getByRole('spinbutton').nth(1).inputValue(),'1.0000');assert.equal(await row('LABEL').getByRole('spinbutton').nth(2).inputValue(),'5.0000');assert.equal(await row('SHELL').getByRole('textbox').inputValue(),'KEEP');assert.equal(await page.getByLabel('Final FOB').inputValue(),'48.9800');
  await page.getByRole('button',{name:'Undo'}).click();await row('SHELL').getByRole('spinbutton').first().waitFor();assert.equal(await row('SHELL').getByRole('spinbutton').first().inputValue(),'4.2000');await page.getByRole('button',{name:'Redo'}).click();await page.waitForFunction(()=>[...document.querySelectorAll('tr')].find(row=>row.textContent?.includes('SHELL'))?.querySelector('input[type="number"]')?.value==='4.0000');
  const downloadEvent=page.waitForEvent('download');await page.getByRole('button',{name:/내부 현재/}).click();const download=await downloadEvent,book=new ExcelJS.Workbook();await book.xlsx.readFile(await download.path());const sheet=book.worksheets[0],shellRow=sheet.getRows(1,sheet.rowCount).find(r=>r.getCell(2).value==='SHELL');assert.equal(shellRow.getCell(5).value,4);assert.equal(shellRow.getCell(8).result,4.2);assert.match(shellRow.getCell(8).formula,/ROUND/);const totalRow=sheet.getRows(1,sheet.rowCount).find(r=>r.getCell(2).value==='Total material cost');assert.equal(totalRow.getCell(8).result,5.25);assert.match(totalRow.getCell(8).formula,/SUM/);
  await page.getByRole('button',{name:/OTHER STYLE/}).click();assert.equal(await row('OTHER').getByRole('spinbutton').first().inputValue(),'10.0000');await page.getByRole('button',{name:'일괄조정 전체 초기화'}).click();await page.getByRole('radio',{name:'모든 스타일 전체 그룹'}).check();await page.getByRole('button',{name:'초기화 실행'}).click();await page.getByText(/2개 스타일의 1개 그룹, 1개 조정값/).waitFor();assert.equal(await row('OTHER').getByRole('spinbutton').first().inputValue(),'9.0000');
  await page.waitForTimeout(650);await page.reload({waitUntil:'networkidle'});assert.equal(await row('OTHER').getByRole('spinbutton').first().inputValue(),'9.0000');console.log(`Reset E2E passed: ${baseUrl}`);
}finally{await browser.close()}
