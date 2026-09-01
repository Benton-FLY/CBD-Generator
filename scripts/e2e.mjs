import { chromium } from 'playwright';
import ExcelJS from 'exceljs';
import { strict as assert } from 'node:assert';

const baseUrl=process.env.E2E_URL||'http://127.0.0.1:4173';
const browser=await chromium.launch({headless:true});
const context=await browser.newContext({acceptDownloads:true});
const page=await context.newPage();

const material=(id,group,cost)=>({id,item:id.toUpperCase(),width:'58',unit:'YD',group,included:true,baseCost:cost,adjustedCost:cost,baseUsage:1,adjustedUsage:1,additionalLoss:.05,remark:'',split:false,sources:[]});
const styles=[
  {id:'style-one',name:'STYLE ONE 사전원가',sourceFile:'one.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('trim-1','TRIMS',1),material('trim-2','TRIMS',2),material('trim-3','TRIMS',3),material('shell-1','OUTSHELL',4),material('shell-2','OUTSHELL',5),material('shell-hidden','OUTSHELL',6)]},
  {id:'style-two',name:'STYLE TWO 사전원가',sourceFile:'two.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('other','TRIMS',9)]},
];
const saved={version:1,styles,active:'style-one',settings:{exchangeRate:900,defaultLoss:.05},dict:{},selections:{},groupFilters:{},fobs:[],scope:'all',target:'cost',operation:'base-percent',bulkValue:5,savedAt:new Date().toISOString()};

const itemRow=name=>page.getByRole('row').filter({has:page.getByText(name,{exact:true})});
const selectItem=async name=>itemRow(name).getByRole('checkbox').first().check();

try{
  await page.goto(baseUrl,{waitUntil:'networkidle'});
  await page.evaluate(async value=>{
    await new Promise((resolve,reject)=>{const request=indexedDB.open('cbd-generator-local',1);request.onupgradeneeded=()=>request.result.createObjectStore('work');request.onerror=()=>reject(request.error);request.onsuccess=()=>{const db=request.result,tx=db.transaction('work','readwrite');tx.objectStore('work').put(value,'latest');tx.oncomplete=()=>{db.close();resolve()};tx.onerror=()=>reject(tx.error)}});
  },saved);
  await page.reload({waitUntil:'networkidle'});

  const fob=page.getByLabel('Final FOB');
  assert.equal(await fob.getAttribute('type'),'text');
  assert.equal(await fob.getAttribute('inputmode'),'decimal');
  await fob.fill('48.98');await fob.press('Enter');
  assert.equal(await fob.inputValue(),'48.9800');
  await fob.focus();await page.mouse.wheel(0,100);assert.equal(await fob.inputValue(),'48.9800');

  const group=page.getByLabel('CBD Group 보기');
  await group.selectOption('TRIMS');for(const name of ['TRIM-1','TRIM-2','TRIM-3'])await selectItem(name);
  await group.selectOption('OUTSHELL');for(const name of ['SHELL-1','SHELL-2'])await selectItem(name);
  assert.equal(await page.getByRole('option',{name:'선택 항목 (5)'}).count(),1);
  await page.getByLabel('일괄 조정 대상').selectOption('selected');
  for(const name of ['TRIM-1','TRIM-2','TRIM-3','SHELL-1','SHELL-2'])assert.equal(await itemRow(name).count(),1);
  assert.equal(await itemRow('SHELL-HIDDEN').count(),0);
  assert.equal(await page.getByLabel('일괄 조정 값').inputValue(),'5.00');
  await page.getByLabel('일괄 조정 방식').selectOption('set');await page.getByRole('button',{name:'미리보기'}).click();
  assert.match(await page.locator('.toolbar').innerText(),/변경 전 \$22\.05 → 변경 후 \$32\.55/);
  await page.getByRole('button',{name:'적용'}).click();
  await page.waitForFunction(()=>document.querySelector('tr:nth-child(1) input[type="number"]')?.value==='5.0000');
  assert.equal(await itemRow('TRIM-1').getByRole('spinbutton').first().inputValue(),'5.0000');
  await itemRow('TRIM-1').getByRole('checkbox').first().click();assert.equal(await itemRow('TRIM-1').count(),0);

  await page.waitForTimeout(700);
  const downloadPromise=page.waitForEvent('download');await page.getByRole('button',{name:/Buyer 현재/}).click();const download=await downloadPromise;
  const path=await download.path();const workbook=new ExcelJS.Workbook();await workbook.xlsx.readFile(path);const sheet=workbook.worksheets[0];
  const fobRow=sheet.getRows(1,sheet.rowCount).find(row=>row.getCell(2).value==='FOB PRICE');
  assert.equal(fobRow.getCell(8).result,48.98);assert.equal(fobRow.getCell(8).numFmt,'$0.0000');
  assert.equal(workbook.getWorksheet('ERP MATERIAL CHECK'),undefined);assert.doesNotMatch(workbook.worksheets.flatMap(ws=>ws.getSheetValues()).flat(5).join(' '),/사전원가|INTERNAL USE ONLY/);
  const internalPromise=page.waitForEvent('download');await page.getByRole('button',{name:/내부 현재/}).click();const internalDownload=await internalPromise;
  const internalPath=await internalDownload.path(),internalBook=new ExcelJS.Workbook();await internalBook.xlsx.readFile(internalPath);assert.ok(internalBook.getWorksheet('ERP MATERIAL CHECK'));const internalSheet=internalBook.worksheets[0];const reviewLabels=internalSheet.getColumn(9).values.map(String);assert.ok(reviewLabels.includes('CBD 재료비 / FOB'));assert.ok(reviewLabels.includes('사전원가 재료비'));

  await page.reload({waitUntil:'networkidle'});assert.equal(await page.getByLabel('Final FOB').inputValue(),'48.9800');
  assert.equal(await page.getByText('STYLE TWO 사전원가',{exact:true}).count(),1);assert.equal(await page.getByRole('option',{name:'선택 항목 (4)'}).count(),1);
  assert.match(await page.locator('footer').innerText(),/자동 저장됨/);

  page.once('dialog',dialog=>dialog.dismiss());await page.getByRole('button',{name:/전체 초기화/}).click();assert.equal(await page.getByLabel('Final FOB').inputValue(),'48.9800');
  page.once('dialog',dialog=>dialog.accept());await page.getByRole('button',{name:/전체 초기화/}).click();await page.getByText('새 작업 시작').waitFor();
  const persisted=await page.evaluate(async()=>new Promise((resolve,reject)=>{const request=indexedDB.open('cbd-generator-local',1);request.onerror=()=>reject(request.error);request.onsuccess=()=>{const db=request.result,get=db.transaction('work').objectStore('work').get('latest');get.onsuccess=()=>{db.close();resolve(get.result??null)};get.onerror=()=>reject(get.error)}}));assert.equal(persisted,null);
  console.log(`E2E passed: ${baseUrl}`);
}finally{await browser.close()}
