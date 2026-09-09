import fs from 'node:fs/promises';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {managerFixture} from './manager-fixture';
import {buildManagerWorkbook} from '../src/comparison/managerExporter';
import {buildComparisonWorkbook} from '../src/comparison/exporter';
import {COST_GROUPS} from '../src/comparison/model';
const state=await managerFixture();
await fs.writeFile('reports/manager-state.json',JSON.stringify(state));
const current=buildManagerWorkbook(state,{stylePairIds:[state.styleMatches[0].id]}),all=buildManagerWorkbook(state),detail=buildComparisonWorkbook(state);
await current.xlsx.writeFile('reports/EVO_GLOVE_CBD_Comparison_Manager_Format.xlsx');await all.xlsx.writeFile('reports/All_Styles_CBD_Comparison_Manager_Format.xlsx');await detail.xlsx.writeFile('reports/manager-regression-detail.xlsx');
const reopened=new ExcelJS.Workbook();await reopened.xlsx.readFile('reports/All_Styles_CBD_Comparison_Manager_Format.xlsx');
const results=[];
for(const [i,sheet] of reopened.worksheets.entries()){
 assert.equal(sheet.getImages().length,2);assert.equal(sheet.pageSetup.printArea,`A1:Y${sheet.rowCount}`);
 const pair=state.styleMatches[i],ref=state.styles.find(s=>s.id===pair.referenceId)!,cur=state.styles.find(s=>s.id===pair.currentId)!;
 let totalRow=0,formulas=0;sheet.eachRow(row=>{if(row.getCell(1).value==='Total material cost')totalRow=row.number;row.eachCell(cell=>{if(cell.formula){formulas++;assert(!/#VALUE!|#REF!|#DIV\/0!|#NAME\?|#N\/A|IFERROR/.test(cell.formula));assert(!cell.formula.includes('!'));for(const match of cell.formula.matchAll(/\b([A-Z]+)(\d+)\b/g)){assert(Number(match[2])<=sheet.rowCount);assert.notEqual(match[0],cell.address)}}})});
 for(const [style,col,materialCol,detailCol] of [[cur,8,2,3],[ref,24,18,2]] as const){
  const rows=style.materials;
  const actual:string[]=[];sheet.eachRow(row=>{if(row.number>=6&&typeof row.getCell(materialCol).value==='string'&&!row.getCell(materialCol).isMerged)actual.push(String(row.getCell(materialCol).value))});
  assert.deepEqual(actual.sort(),rows.map(r=>r.material).sort());
  const expected=rows.filter(r=>COST_GROUPS.includes(r.group)).reduce((s,r)=>s+(r.extended??0),0);
  const total=sheet.getCell(totalRow,col).result;assert(Math.abs(Number(total)-expected)<1e-9);
  assert(Math.abs(Number(total)-Number(detail.worksheets[i+1].getCell(12,detailCol).result))<1e-9);
  assert.equal(sheet.getCell(totalRow+4,col).value,style.summary.finalFob??null);
 }
 results.push({style:sheet.name,rows:sheet.rowCount,formulas,referenceRows:ref.materials.length,currentRows:cur.materials.length,currentTotal:sheet.getCell(totalRow,8).result,referenceTotal:sheet.getCell(totalRow,24).result,sourceCurrentTotal:cur.summary.totalMaterialCost,sourceReferenceTotal:ref.summary.totalMaterialCost});
}
assert.equal(reopened.worksheets.length,13);assert.equal(current.worksheets.length,1);
await fs.writeFile('reports/manager-verification.json',JSON.stringify(results,null,2));console.log(JSON.stringify(results,null,2));
