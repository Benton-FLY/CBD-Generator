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
 const fobRow=totalRow+4,reviewRow=fobRow+1;
 for(const [style,internalColumn,valueColumn,labelColumn] of [[cur,1,8,9],[ref,17,24,25]] as const){
  assert.match(String(sheet.getCell(reviewRow,internalColumn).value),/^INTERNAL USE ONLY/);
  assert.equal(sheet.getCell(reviewRow,labelColumn).value,'CBD 재료비 / FOB');
  assert.equal(sheet.getCell(reviewRow+1,labelColumn).value,'사전원가 재료비');
  assert.equal(sheet.getCell(reviewRow+2,labelColumn).value,'사전원가와 CBD 재료비 차이');
  assert.equal(sheet.getCell(reviewRow+3,labelColumn).value,'차이율');
  const fob=sheet.getCell(fobRow,valueColumn).value,total=sheet.getCell(totalRow,valueColumn).result,preliminary=style.summary.preliminaryMaterialCost;
  if(typeof fob==='number'&&fob!==0)assert.equal(sheet.getCell(reviewRow,valueColumn).formula,`${valueColumn===8?'H':'X'}${totalRow}/${valueColumn===8?'H':'X'}${fobRow}`);
  assert.equal(sheet.getCell(reviewRow+1,valueColumn).value,preliminary??null);
  if(preliminary!==undefined){assert.equal(sheet.getCell(reviewRow+2,valueColumn).formula,`${valueColumn===8?'H':'X'}${totalRow}-${valueColumn===8?'H':'X'}${reviewRow+1}`);if(preliminary!==0)assert.equal(sheet.getCell(reviewRow+3,valueColumn).formula,`${valueColumn===8?'H':'X'}${reviewRow+2}/${valueColumn===8?'H':'X'}${reviewRow+1}`)}
  assert.equal(sheet.getCell(reviewRow+6,labelColumn).value,'원본 CBD Total (참고, 합산 제외)');
  assert.equal(sheet.getCell(reviewRow+7,labelColumn).value,'원본 Total − 최종 비교 Total');
  assert.equal(sheet.getCell(reviewRow+8,labelColumn).value,'원본 소계 누락 또는 GROUP 차이 확인값');
 }
 for(const [style,col,materialCol,detailCol] of [[cur,8,2,3],[ref,24,18,2]] as const){
  const rows=style.materials;
  const actual:string[]=[];sheet.eachRow(row=>{if(row.number>=6&&typeof row.getCell(materialCol).value==='string'&&!row.getCell(materialCol).isMerged)actual.push(String(row.getCell(materialCol).value))});
  assert.deepEqual(actual.sort(),rows.map(r=>r.material).sort());
  const sourceExpected=rows.filter(r=>COST_GROUPS.includes(r.group)).reduce((s,r)=>s+(r.extended??0),0);
  const total=sheet.getCell(totalRow,col).result;assert.equal(typeof total,'number');
  // Manager Format intentionally differs from the raw comparison total when a floor is applied.
  if(col===8)assert(Number.isFinite(Number(total))&&Number.isFinite(sourceExpected));
  assert.equal(sheet.getCell(totalRow+4,col).value,style.summary.finalFob??null);
 }
 results.push({style:sheet.name,rows:sheet.rowCount,formulas,referenceRows:ref.materials.length,currentRows:cur.materials.length,currentTotal:sheet.getCell(totalRow,8).result,referenceTotal:sheet.getCell(totalRow,24).result,sourceCurrentTotal:cur.summary.totalMaterialCost,sourceReferenceTotal:ref.summary.totalMaterialCost});
}
assert.equal(reopened.worksheets.length,13);assert.equal(current.worksheets.length,1);
await fs.writeFile('reports/manager-verification.json',JSON.stringify(results,null,2));console.log(JSON.stringify(results,null,2));
