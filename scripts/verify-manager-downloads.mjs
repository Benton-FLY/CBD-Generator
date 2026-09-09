import ExcelJS from 'exceljs';
import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
const prefix=process.argv[2]||'manager-local',report=[];
for(const scope of ['current','all']){
 const manager=new ExcelJS.Workbook(),detail=new ExcelJS.Workbook();
 await manager.xlsx.readFile(`reports/${prefix}-${scope}.xlsx`);await detail.xlsx.readFile(`reports/${prefix}-detail-${scope}.xlsx`);
 for(let i=0;i<manager.worksheets.length;i++){
  const s=manager.worksheets[i],d=detail.worksheets[i+1];
  const find=label=>s.getColumn(1).values.findIndex(v=>v===label),total=find('Total material cost'),fob=find('FOB PRICE');
  for(const [mc,dc] of [[8,3],[24,2]]){
   assert(Math.abs(s.getCell(total,mc).result-d.getCell(12,dc).result)<1e-9);
   assert.equal(s.getCell(fob,mc).value,d.getCell(7,dc).result);
   for(const [j,group] of ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'].entries())assert(Math.abs(s.getCell(find(group+' SUBTOTAL'),mc).result-d.getCell(j+8,dc).result)<1e-9);
  }
  s.eachRow(row=>row.eachCell(cell=>{if(typeof cell.result==='number')assert(Number.isFinite(cell.result));if(cell.formula)assert(!/#VALUE!|#REF!|#DIV\/0!|#NAME\?|#N\/A|IFERROR/.test(cell.formula))}));
  report.push({scope,style:s.name,currentTotal:s.getCell(total,8).result,referenceTotal:s.getCell(total,24).result,currentFob:s.getCell(fob,8).value,referenceFob:s.getCell(fob,24).value,groupSubtotalsMatchDetailed:true});
 }
}
await fs.writeFile(`reports/${prefix}-financials.json`,JSON.stringify(report,null,2));
console.log(`PASS: ${prefix} Current/All downloaded Manager totals, every cost group, and FINAL FOB match the downloaded detailed Excel.`);
