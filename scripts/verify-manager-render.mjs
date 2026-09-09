import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
const runs=JSON.parse(await fs.readFile('reports/manager-render/recalculation.json','utf8'));
let formulas=0;
for(const run of runs){
 const book=new ExcelJS.Workbook();await book.xlsx.readFile(run.file);
 for(let i=0;i<run.sheets.length;i++)for(const value of run.sheets[i].cells){
  const cell=book.worksheets[i].getCell(value.row,value.column);assert(cell.formula);assert.equal(value.formula.slice(1).replaceAll(';',','),cell.formula);formulas++;
  if(typeof cell.result==='number')assert(Math.abs(cell.result-value.value)<1e-9,`${cell.address}: ${cell.result} versus ${value.value}`);
  else assert.equal(value.text,cell.result);
 }
}
console.log(`PASS: ${formulas} cached formula results agree with independent LibreOffice recalculation.`);
