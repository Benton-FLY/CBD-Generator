import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';
import type {CbdStyle,ComparisonState,MaterialMatchCluster} from './types';

const GROUPS=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'];
const fill=(argb:string):ExcelJS.Fill=>({type:'pattern',pattern:'solid',fgColor:{argb}});
const value=(input?:number)=>input===undefined?'':input;
const style=(state:ComparisonState,id?:string)=>state.styles.find(item=>item.id===id);
const material=(item:CbdStyle|undefined,id?:string)=>item?.materials.find(row=>row.id===id);
const safeName=(name:string,used:Set<string>)=>{const base=(name.replace(/[\\/*?:[\]]/g,' ').trim()||'Comparison').slice(0,31);let result=base,index=2;while(used.has(result))result=`${base.slice(0,27)} ${index++}`;used.add(result);return result};
const mappedAmount=(item:CbdStyle|undefined,clusters:MaterialMatchCluster[],group:string,side:'reference'|'comparison')=>{const ids=new Set<string>();clusters.filter(cluster=>cluster.finalGroup.toUpperCase()===group).forEach(cluster=>side==='reference'?cluster.referenceRowIds.forEach(id=>ids.add(id)):cluster.currentRowId&&ids.add(cluster.currentRowId));const values=[...ids].map(id=>material(item,id)?.extended).filter((amount):amount is number=>amount!==undefined);return values.length?values.reduce((a,b)=>a+b,0):undefined};
const differenceFormula=(referenceCell:string,comparisonCell:string)=>`IF(AND(ISNUMBER(${referenceCell}),ISNUMBER(${comparisonCell})),${comparisonCell}-${referenceCell},IF(ISNUMBER(${comparisonCell}),${comparisonCell},IF(ISNUMBER(${referenceCell}),-${referenceCell},\"\")))`;

export function buildComparisonWorkbook(state:ComparisonState){
 const workbook=new ExcelJS.Workbook(),summary=workbook.addWorksheet('STYLE MATCH'),used=new Set<string>(['STYLE MATCH']);
 summary.views=[{state:'frozen',ySplit:1}];
 summary.addRow([`Reference Season ${state.referenceSeason}`,'Reference Style',`Comparison Season ${state.currentSeason}`,'Comparison Style','Match Type','Status','Reference Total Material Cost','Comparison Total Material Cost','Reference FOB','Comparison FOB']);
 summary.getRow(1).font={bold:true,color:{argb:'FFFFFF'}};summary.getRow(1).fill=fill('1E3A8A');
 for(const styleMatch of state.styleMatches.filter(match=>!match.excluded)){
  const reference=style(state,styleMatch.referenceId),comparison=style(state,styleMatch.currentId),clusters=state.materialMatches.find(set=>set.styleMatchId===styleMatch.id)?.clusters||[];
  summary.addRow([state.referenceSeason,reference?.styleName||'—',state.currentSeason,comparison?.styleName||'—',styleMatch.method,styleMatch.status,reference?.summary.totalMaterialCost??'',comparison?.summary.totalMaterialCost??'',reference?.summary.fob??'',comparison?.summary.fob??'']);
  const sheet=workbook.addWorksheet(safeName(comparison?.styleName||reference?.styleName||'Unmatched',used));
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};sheet.views=[{state:'frozen',ySplit:15,showGridLines:false}];
  sheet.mergeCells('A1:O1');sheet.getCell('A1').value='FLY Racing CBD Comparison Sheet';sheet.getCell('A1').font={bold:true,size:16,color:{argb:'FFFFFF'}};sheet.getCell('A1').fill=fill('1E3A8A');sheet.getCell('A1').alignment={horizontal:'center'};
  sheet.mergeCells('A2:O2');sheet.getCell('A2').value=`Reference: ${reference?.styleName||'None'}    |    Comparison: ${comparison?.styleName||'None'}`;sheet.getCell('A2').font={bold:true};sheet.getCell('A2').fill=fill('DBEAFE');
  sheet.mergeCells('A3:O3');sheet.getCell('A3').value=`Generated: ${new Date().toLocaleDateString('en-CA')}`;
  sheet.getCell('A5').value='Material Cost Comparison by Group';sheet.getCell('A5').font={bold:true,size:13};
  sheet.getRow(6).values=['CBD Group',`Reference Season ${state.referenceSeason}`,`Comparison Season ${state.currentSeason}`,'Difference','Change %'];sheet.getRow(6).font={bold:true,color:{argb:'FFFFFF'}};sheet.getRow(6).fill=fill('1E3A8A');
  let groupRow=7;
  for(const group of [...GROUPS,'TOTAL MATERIAL COST']){
   const referenceAmount=group==='TOTAL MATERIAL COST'?reference?.summary.totalMaterialCost:mappedAmount(reference,clusters,group,'reference');
   const comparisonAmount=group==='TOTAL MATERIAL COST'?comparison?.summary.totalMaterialCost:mappedAmount(comparison,clusters,group,'comparison');
   sheet.getCell(groupRow,1).value=group;sheet.getCell(groupRow,2).value=value(referenceAmount);sheet.getCell(groupRow,3).value=value(comparisonAmount);
   sheet.getCell(groupRow,4).value={formula:differenceFormula(`B${groupRow}`,`C${groupRow}`)};
   sheet.getCell(groupRow,5).value={formula:`IF(AND(ISNUMBER(B${groupRow}),B${groupRow}<>0,ISNUMBER(D${groupRow})),D${groupRow}/B${groupRow},\"\")`};
   for(const column of [2,3,4])sheet.getCell(groupRow,column).numFmt='$0.0000';sheet.getCell(groupRow,5).numFmt='0.0%';if(group==='TOTAL MATERIAL COST')sheet.getRow(groupRow).fill=fill('DBEAFE');groupRow++;
  }
  sheet.getCell('A13').value='Material Comparison Details';sheet.getCell('A13').font={bold:true,size:13};
  sheet.mergeCells('B14:C14');sheet.getCell('B14').value='Material';sheet.mergeCells('E14:G14');sheet.getCell('E14').value='Cost per Unit';sheet.mergeCells('H14:J14');sheet.getCell('H14').value='Usage';sheet.mergeCells('K14:M14');sheet.getCell('K14').value='Extended Cost';sheet.getCell('A14').value='Group';sheet.getCell('D14').value='Unit';sheet.getCell('N14').value='Status';sheet.getCell('O14').value='Remark';sheet.getRow(14).font={bold:true,color:{argb:'FFFFFF'}};sheet.getRow(14).fill=fill('1E3A8A');
  sheet.getRow(15).values=['','Reference Material','Comparison Material','Unit','Reference','Comparison','Difference','Reference','Comparison','Difference','Reference','Comparison','Difference','Status','Remark'];sheet.getRow(15).font={bold:true};sheet.getRow(15).fill=fill('BFDBFE');
  const ordered=[...clusters].sort((a,b)=>(material(comparison,a.currentRowId||undefined)?.order??99999)-(material(comparison,b.currentRowId||undefined)?.order??99999));let row=16;
  for(const cluster of ordered){
   const references=cluster.referenceRowIds.map(id=>material(reference,id)).filter(Boolean),comparisonRow=material(comparison,cluster.currentRowId||undefined),total=(key:'cost'|'usage'|'extended')=>{const values=references.map(item=>item?.[key]).filter((number):number is number=>number!==undefined);return values.length?values.reduce((a,b)=>a+b,0):undefined};
   sheet.getCell(row,1).value=cluster.finalGroup;sheet.getCell(row,2).value=references.map(item=>item?.material).join('\n');sheet.getCell(row,3).value=comparisonRow?.material||'';sheet.getCell(row,4).value=comparisonRow?.unit||references[0]?.unit||'';
   sheet.getCell(row,5).value=value(total('cost'));sheet.getCell(row,6).value=value(comparisonRow?.cost);sheet.getCell(row,7).value={formula:differenceFormula(`E${row}`,`F${row}`)};
   sheet.getCell(row,8).value=value(total('usage'));sheet.getCell(row,9).value=value(comparisonRow?.usage);sheet.getCell(row,10).value={formula:differenceFormula(`H${row}`,`I${row}`)};
   sheet.getCell(row,11).value=value(total('extended'));sheet.getCell(row,12).value=value(comparisonRow?.extended);sheet.getCell(row,13).value={formula:differenceFormula(`K${row}`,`L${row}`)};
   sheet.getCell(row,14).value=cluster.status==='CURRENT ONLY'?'Comparison Only':cluster.status==='REFERENCE ONLY'?'Reference Only':cluster.status==='REVIEW'?'Needs Review':cluster.status;sheet.getCell(row,15).value=comparisonRow?.remark||references.map(item=>item?.remark).filter(Boolean).join('\n');
   for(const column of [5,6,7,11,12,13])sheet.getCell(row,column).numFmt='$0.0000';for(const column of [8,9,10])sheet.getCell(row,column).numFmt='0.0000';
   if(cluster.status==='REVIEW')sheet.getRow(row).font={bold:true,color:{argb:'DC2626'}};else if(cluster.status==='CURRENT ONLY')sheet.getRow(row).fill=fill('DCFCE7');else if(cluster.status==='REFERENCE ONLY')sheet.getRow(row).fill=fill('F3F4F6');else if(cluster.status==='GROUP CHANGED')sheet.getRow(row).fill=fill('FFEDD5');else if(cluster.status==='NAME CHANGED')sheet.getRow(row).fill=fill('FEF3C7');else if(cluster.status==='MERGED N:1')sheet.getRow(row).fill=fill('EDE9FE');row++;
  }
  [18,38,38,10,14,14,14,12,12,12,15,15,15,18,42].forEach((width,index)=>sheet.getColumn(index+1).width=width);sheet.eachRow(item=>item.alignment={vertical:'top',wrapText:true});sheet.getRow(1).height=26;sheet.getRow(2).height=24;for(let index=16;index<row;index++)sheet.getRow(index).height=32;
 }
 return workbook;
}

export async function exportComparison(state:ComparisonState){const workbook=buildComparisonWorkbook(state),buffer=await workbook.xlsx.writeBuffer();saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),`CBD_Comparison_${state.referenceSeason||'Reference'}_vs_${state.currentSeason||'Comparison'}.xlsx`)}
