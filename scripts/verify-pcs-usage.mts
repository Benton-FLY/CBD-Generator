import ExcelJS from 'exceljs';
import { readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';
import { aggregate, detectHeader, rowsFromSheet } from '../src/parser';
import { createBuyerWorkbook } from '../src/exporter';
import { extendedCost, type StyleData } from '../src/types';

const sourcePath=process.argv[2];
if(!sourcePath)throw new Error('Usage: npm run verify:pcs -- <BOM.xls>');
const source=XLSX.read(await readFile(sourcePath),{type:'buffer'});
const rows=XLSX.utils.sheet_to_json<unknown[]>(source.Sheets[source.SheetNames[0]],{header:1,raw:true,defval:''});
const header=detectHeader(rows);
if(!header)throw new Error('BOM header not found');
const bomRows=rowsFromSheet(rows,header.mapping,header.row,{file:sourcePath,sheet:source.SheetNames[0]},{exchangeRate:900,defaultLoss:.05});
const materials=aggregate(bomRows,.05,{});
const expected=new Map([
  ['HEADER CARD',1],['SIZE STICKER',2],['J RING',1],['UPC BARCODE',3],['PP POLY BAG',1],
  ['E/STRING',2],['SIZE LABEL',2],['CARE LABEL',2],['S/T',1],
]);
const checked:{item:string;before:number;after:number;extended:number}[]=[];
for(const [label,usage] of expected){
  const material=materials.find(entry=>entry.item.toUpperCase().includes(label));
  if(!material)throw new Error(`${label} not found`);
  if(material.baseUsage!==usage||material.adjustedUsage!==usage)throw new Error(`${label}: expected ${usage}, got ${material.adjustedUsage}`);
  checked.push({item:material.item,before:material.sources.reduce((sum,row)=>sum+row.usage,0),after:material.adjustedUsage,extended:extendedCost(material)});
}
for(const bomRow of bomRows.filter(row=>['M','CONE','ROLL'].includes(row.unit.trim().toUpperCase()))){
  const material=materials.find(entry=>entry.sources.some(sourceRow=>sourceRow.id===bomRow.id));
  if(!material||material.baseUsage!==material.sources.reduce((sum,row)=>sum+row.usage,0))throw new Error(`${bomRow.item}: non-PCS usage changed`);
}
const style:StyleData={id:'verify',name:'28 KINETIC 1 GLOVE',sourceFile:sourcePath,sourceSheet:source.SheetNames[0],materials,laborRemark:''};
const workbook=await createBuyerWorkbook([style]);const buffer=await workbook.xlsx.writeBuffer();const reopened=new ExcelJS.Workbook();await reopened.xlsx.load(buffer);const sheet=reopened.worksheets[0];
for(const result of checked){let rowNumber=0;sheet.eachRow((row,index)=>{if(row.getCell(2).value===result.item)rowNumber=index});if(sheet.getCell(rowNumber,6).value!==result.after||(sheet.getCell(rowNumber,8).value as ExcelJS.CellFormulaValue).result!==result.extended)throw new Error(`${result.item}: Excel mismatch`)}
console.table(checked);
console.log(`Verified ${checked.length} PCS items; web data and Excel export match.`);
