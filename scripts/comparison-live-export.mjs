import {chromium} from 'playwright';
import ExcelJS from 'exceljs';

const browser=await chromium.launch({headless:true}),page=await browser.newPage({acceptDownloads:true});
await page.goto('https://cbd-generator.pages.dev/#comparison');
await page.locator('input[type=file]').nth(0).setInputFiles(['reference/FLY RACING 27 MX JERSEY CBD 2025.10.17.xlsx','reference/FLY RACING 27 BMX JERSEY CBD 2025.10.17.xlsx']);
await page.getByText('2 files · 16 styles',{exact:true}).waitFor({timeout:60000});
await page.locator('input[type=file]').nth(1).setInputFiles('reference/FLY RACING 28 MX BMX JERSEY CBD SHEET (내부) 2026.09.02.xlsx');
await page.getByRole('heading',{name:'Review Style Matches'}).waitFor({timeout:60000});
await page.getByRole('button',{name:'Confirm Style Matches'}).click();
await page.getByRole('heading',{name:'Review Material Matches'}).waitFor({timeout:60000});
await page.getByLabel('Material Filter').selectOption('group:OUTSHELL');

const downloadWorkbook=async(name)=>{const event=page.waitForEvent('download');await page.getByRole('button',{name,exact:true}).click();const download=await event,path=await download.path();if(!path)throw new Error('download path unavailable');const workbook=new ExcelJS.Workbook();await workbook.xlsx.readFile(path);return{workbook,fileName:download.suggestedFilename()}};
const validate=workbook=>{const formulas=[];for(const sheet of workbook.worksheets)sheet.eachRow(row=>row.eachCell(cell=>{if(cell.formula)formulas.push(cell.formula)}));if(!formulas.some(value=>/^SUM\(B\d+:B\d+\)$/.test(value)))throw new Error('formula-based TOTAL MATERIAL not found');if(formulas.some(value=>/ISNUMBER|SUMPRODUCT|OFFSET|INDIRECT|#VALUE!|#REF!|#DIV\/0!|#NAME\?/i.test(value)))throw new Error('complex or invalid formula found');return formulas.length};
const current=await downloadWorkbook('Export Current Style'),currentFormulas=validate(current.workbook);if(current.workbook.worksheets.length!==2)throw new Error(`current export contains ${current.workbook.worksheets.length} sheets`);const currentGroups=current.workbook.worksheets[1].getColumn(1).values.map(String);if(!currentGroups.includes('SEWING THREAD')||!currentGroups.some(value=>value.includes('SPECIAL PROCESS')))throw new Error('current export was incorrectly limited by OUTSHELL screen filter');
const all=await downloadWorkbook('Export All Styles'),allFormulas=validate(all.workbook);if(all.workbook.worksheets.length<=2)throw new Error('all-style export did not retain all styles');
console.log(`Cloudflare exports PASS: current=${current.workbook.worksheets.length} sheets/${currentFormulas} formulas (${current.fileName}), all=${all.workbook.worksheets.length} sheets/${allFormulas} formulas (${all.fileName})`);await browser.close();
