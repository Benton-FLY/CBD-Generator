import { chromium, webkit } from 'playwright';
import XLSX from 'xlsx';
import { strict as assert } from 'node:assert';

const baseUrl=process.env.E2E_URL||'http://127.0.0.1:5173';
const workbookBuffer=bookType=>{const workbook=XLSX.utils.book_new(),sheet=XLSX.utils.aoa_to_sheet([['Structure','Item','Unit','소요량','환산단가'],['BODY','TEST FABRIC','YD',1,2.5]]);XLSX.utils.book_append_sheet(workbook,sheet,'BOM');return XLSX.write(workbook,{bookType,type:'buffer'})};
const xls=workbookBuffer('xls'),xlsx=workbookBuffer('xlsx');
const payload=(name,buffer,mimeType='')=>({name,mimeType,buffer});

async function run(browserType,name){
  const browser=await browserType.launch({headless:true}),context=await browser.newContext(),page=await context.newPage();const errors=[];
  page.on('pageerror',error=>errors.push(error.message));page.on('console',message=>{if(message.type()==='error'&&!message.text().includes('Failed to load resource'))errors.push(message.text())});
  try{
    await page.goto(baseUrl,{waitUntil:'networkidle'});await page.evaluate(async()=>new Promise((resolve,reject)=>{const request=indexedDB.deleteDatabase('cbd-generator-local');request.onsuccess=()=>resolve();request.onerror=()=>reject(request.error);request.onblocked=()=>resolve()}));await page.reload({waitUntil:'networkidle'});
    const box=page.getByTestId('start-bom-uploader'),input=page.getByLabel('BOM Excel 파일');
    const dom=await input.evaluate(element=>{const style=getComputedStyle(element);return{disabled:element.disabled,multiple:element.multiple,accept:element.accept,id:element.id,label:document.querySelector(`label[for="${element.id}"]`)!==null,pointerEvents:style.pointerEvents,zIndex:style.zIndex}});
    assert.equal(dom.disabled,false);assert.equal(dom.multiple,true);assert.equal(dom.label,true);assert.match(dom.accept,/application\/vnd\.ms-excel/);assert.notEqual(dom.pointerEvents,'none');

    let chooserPromise=page.waitForEvent('filechooser');await box.click({position:{x:5,y:5}});let chooser=await chooserPromise;await chooser.setFiles([]);
    chooserPromise=page.waitForEvent('filechooser');await box.getByText('BOM Excel 파일을 클릭하거나 이곳에 끌어다 놓으세요.').click();chooser=await chooserPromise;await chooser.setFiles([]);
    chooserPromise=page.waitForEvent('filechooser');await box.getByRole('button',{name:'파일 찾아보기'}).click();chooser=await chooserPromise;await chooser.setFiles([]);
    chooserPromise=page.waitForEvent('filechooser');await box.locator('label').press('Enter');chooser=await chooserPromise;await chooser.setFiles([]);
    chooserPromise=page.waitForEvent('filechooser');await box.locator('label').press('Space');chooser=await chooserPromise;await chooser.setFiles([]);

    await input.setInputFiles(payload('27 LITE PANT BOM.xls',xls,''));await page.getByText('1개 스타일을 분석했습니다.').waitFor();assert.equal(await page.getByText('LITE PANT',{exact:true}).count(),1);
    const workspace=page.getByTestId('workspace-bom-uploader'),workspaceInput=page.getByLabel('추가 BOM Excel 파일');
    chooserPromise=page.waitForEvent('filechooser');await workspace.getByText('BOM 추가',{exact:true}).click();chooser=await chooserPromise;await chooser.setFiles([]);
    chooserPromise=page.waitForEvent('filechooser');await workspace.getByRole('button',{name:'파일 선택'}).click();chooser=await chooserPromise;await chooser.setFiles([]);

    await workspaceInput.setInputFiles(payload('28 LITE PANT 사전원가.xlsx',xlsx,''));await page.getByText('1개 스타일을 분석했습니다.').waitFor();
    await workspaceInput.setInputFiles(payload('28 LITE PANT 사전원가.xlsx',xlsx,''));await page.getByText('1개 스타일을 분석했습니다.').waitFor();assert.equal(await page.locator('aside .style').count(),3);
    await workspaceInput.setInputFiles([payload('TEST BOM.XLSX',xlsx,''),payload('한글 공백 (BOM).xlsx',xlsx,'')]);await page.getByText('2개 스타일을 분석했습니다.').waitFor();assert.equal(await page.locator('aside .style').count(),5);

    const dropFile=payload('드롭 테스트 BOM.xls',xls,'');await workspace.evaluate((element,file)=>{const bytes=Uint8Array.from(atob(file.data),char=>char.charCodeAt(0)),transfer=new DataTransfer();transfer.items.add(new File([bytes],file.name,{type:''}));window.__bomDropTransfer=transfer;for(const type of ['dragenter','dragover'])element.dispatchEvent(new DragEvent(type,{bubbles:true,cancelable:true,dataTransfer:transfer}))},{name:dropFile.name,data:dropFile.buffer.toString('base64')});
    await workspace.evaluate(element=>new Promise((resolve,reject)=>requestAnimationFrame(()=>element.classList.contains('dragging')?resolve():reject(new Error('drop highlight missing')))));await workspace.evaluate(element=>element.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true,dataTransfer:window.__bomDropTransfer})));
    await page.getByText('1개 스타일을 분석했습니다.').waitFor();assert.equal(await workspace.evaluate(element=>element.classList.contains('dragging')),false);
    await workspace.evaluate(element=>{const transfer=new DataTransfer();transfer.items.add(new File(['pdf'],'견적 이미지.pdf',{type:'application/pdf'}));element.dispatchEvent(new DragEvent('dragenter',{bubbles:true,cancelable:true,dataTransfer:transfer}));element.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true,dataTransfer:transfer}))});await page.getByText(/견적 이미지\.pdf: Excel 파일/).waitFor();

    await page.waitForTimeout(700);const count=await page.locator('aside .style').count();await page.reload({waitUntil:'networkidle'});assert.equal(await page.locator('aside .style').count(),count);assert.match(await page.locator('footer').innerText(),/자동 저장됨/);
    assert.deepEqual(errors,[]);console.log(`Upload E2E passed: ${name} ${baseUrl}`);
  }finally{await browser.close()}
}

await run(chromium,'Chromium');
await run(webkit,'WebKit');
