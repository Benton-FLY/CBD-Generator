import { chromium, webkit } from 'playwright';
import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { strict as assert } from 'node:assert';

const baseUrl=process.env.E2E_URL||'http://127.0.0.1:4173';
const items=['PRINTED SOLID FLY RACING 25 JERSEY+JX-44 (POLY)','PRINTED FLY RACING 25 KINETIC 1 JERSEY+JX-300 (POLY) FR/SL','SOLID PRINTED FLY RACING JERSEY FABRIC','PRINTED SOLID JERSEY FABRIC','PRINTED JERSEY FABRIC'];
const rows=[['Structure','자재구분','Item','Width','Unit','소요량','환산단가'],...items.map((item,index)=>['BODY','자재',item,'58','YD',1,index+1]),['PACKING','자재','PRINTED LABEL','','PCS',1,.5],['FASTENER','자재','ZIPPER','','PCS',1,.25]];
const workbook=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet(rows);XLSX.utils.book_append_sheet(workbook,sheet,'BOM');const buffer=XLSX.write(workbook,{bookType:'xlsx',type:'buffer'});
const payload={name:'PRINTED TEST BOM.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer};
const red='rgb(220, 38, 38)',pale='rgb(254, 242, 242)';

async function run(browserType,name,checkExcel=false){
  const browser=await browserType.launch({headless:true}),context=await browser.newContext({acceptDownloads:true}),page=await context.newPage();
  try{
    await page.goto(baseUrl,{waitUntil:'networkidle'});await page.evaluate(async()=>new Promise((resolve,reject)=>{const request=indexedDB.deleteDatabase('cbd-generator-local');request.onsuccess=()=>resolve();request.onerror=()=>reject(request.error);request.onblocked=()=>resolve()}));await page.reload({waitUntil:'networkidle'});
    await page.getByLabel('BOM Excel 파일').setInputFiles(payload);await page.getByText('1개 스타일을 분석했습니다.').waitFor();
    for(const item of items)assert.equal(await page.getByLabel(`${item} CBD Group`).inputValue(),'OUTSHELL');assert.equal(await page.getByLabel('PRINTED LABEL CBD Group').inputValue(),'LABEL & PACKAGING');
    const filter=page.getByLabel('CBD Group 보기');await filter.selectOption('TRIMS');for(const item of items)assert.equal(await page.getByText(item,{exact:true}).count(),0);await filter.selectOption('OUTSHELL');for(const item of items)assert.equal(await page.getByText(item,{exact:true}).count(),1);
    if(checkExcel){const event=page.waitForEvent('download');await page.getByRole('button',{name:/Buyer 현재/}).click();const download=await event,book=new ExcelJS.Workbook();await book.xlsx.readFile(await download.path());const ws=book.worksheets[0],values=ws.getColumn(2).values.map(String);for(const item of items)assert.ok(values.includes(item));const subtotal=ws.getRows(1,ws.rowCount).find(row=>row.getCell(2).value==='OUTSHELL SUBTOTAL');assert.equal(subtotal.getCell(8).result,15.75);const total=ws.getRows(1,ws.rowCount).find(row=>row.getCell(2).value==='Total material cost');assert.equal(total.getCell(8).result,16.5375)}
    await filter.selectOption('LABEL & PACKAGING');const reviewSelect=page.getByLabel('PRINTED LABEL CBD Group');await reviewSelect.selectOption('NEEDS REVIEW');
    const sidebar=page.locator('.needs-review-status'),badge=page.locator('.review-count-badge');await sidebar.waitFor();assert.match(await sidebar.innerText(),/검토 필요 1건/);assert.match(await badge.innerText(),/검토 1/);
    for(const locator of [sidebar,reviewSelect,badge]){const style=await locator.evaluate(element=>({color:getComputedStyle(element).color,background:getComputedStyle(element).backgroundColor,fontWeight:getComputedStyle(element).fontWeight}));assert.equal(style.color,red);assert.equal(style.background,pale);assert.ok(Number(style.fontWeight)>=700)}
    let warning='';page.once('dialog',dialog=>{warning=dialog.message();dialog.dismiss()});await page.getByRole('button',{name:/Buyer 현재/}).click();await page.waitForTimeout(50);assert.equal(warning,'검토가 완료되지 않은 항목이 1건 있습니다. 그래도 Buyer CBD를 추출하시겠습니까?');
    await reviewSelect.selectOption('OUTSHELL');assert.equal(await page.locator('.needs-review-status').count(),0);assert.equal(await page.locator('.review-count-badge').count(),0);assert.equal(await reviewSelect.evaluate(element=>getComputedStyle(element).color)===red,false);
    console.log(`Classification/review E2E passed: ${name} ${baseUrl}`);
  }finally{await browser.close()}
}
await run(chromium,'Chromium',true);await run(webkit,'WebKit');
