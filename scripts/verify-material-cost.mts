import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { createBuyerWorkbook } from '../src/exporter';
import { parseBomFile } from '../src/parser';
import { extendedCost, groupSubtotal, totalMaterialCost, visibleGroups } from '../src/types';

const path=process.argv[2]||'reference/28 EVO GLOVE.xls';
const bytes=await readFile(path);
const file=new File([bytes],path.split('/').at(-1)!,{type:'application/vnd.ms-excel'});
const parsed=await parseBomFile(file,{exchangeRate:900,defaultLoss:.05},{});
const style=parsed.styles[0];
if(!style)throw new Error(`No style parsed from ${path}`);

const rows=style.materials.filter(m=>m.included&&visibleGroups.includes(m.group)).map(m=>({
  group:m.group,item:m.item,unitPrice:m.adjustedCost,consumption:m.adjustedUsage,loss:m.additionalLoss,
  rawExtendedCost:m.adjustedCost*m.adjustedUsage*(1+m.additionalLoss),roundedExtendedCost:extendedCost(m),
}));
const groups=visibleGroups.map(group=>({group,subtotal:groupSubtotal(style.materials.filter(m=>m.included&&m.group===group))}));

const generated=await createBuyerWorkbook([style],new Date('2026-09-01T00:00:00Z'));
const buffer=await generated.xlsx.writeBuffer();
const reopened=new ExcelJS.Workbook();
await reopened.xlsx.load(buffer);
const sheet=reopened.worksheets[0];
const formulas:{label:string;formula:string;result:unknown}[]=[];
sheet.eachRow(row=>{
  const value=row.getCell(8).value;
  if(value&&typeof value==='object'&&'formula' in value)formulas.push({label:String(row.getCell(2).value||''),formula:String(value.formula),result:value.result});
});

console.log(JSON.stringify({style:style.name,rows,groups,rawTotal:rows.reduce((sum,row)=>sum+row.rawExtendedCost,0),totalMaterialCost:totalMaterialCost(style.materials),exportFormulas:formulas},null,2));
