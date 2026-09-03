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
const event=page.waitForEvent('download');await page.getByRole('button',{name:'Export Comparison'}).click();
const download=await event,path=await download.path();if(!path)throw new Error('download path unavailable');
const workbook=new ExcelJS.Workbook();await workbook.xlsx.readFile(path);const formulas=[];
for(const sheet of workbook.worksheets)sheet.eachRow(row=>row.eachCell(cell=>{if(cell.formula)formulas.push(cell.formula)}));
if(workbook.worksheets.length<2||!formulas.some(value=>/^SUM\(B\d+:B\d+\)$/.test(value)))throw new Error('formula-based TOTAL MATERIAL not found');
if(formulas.some(value=>/ISNUMBER|SUMPRODUCT|OFFSET|INDIRECT|#VALUE!|#REF!|#DIV\/0!|#NAME\?/i.test(value)))throw new Error('complex or invalid formula found');
console.log(`Cloudflare export PASS: ${workbook.worksheets.length} sheets, ${formulas.length} readable formulas, ${download.suggestedFilename()}`);await browser.close();
