import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';
import type {CbdMaterialRow,CbdStyle,ComparisonFinancials,ComparisonState,MaterialMatchCluster} from './types';

const DEFAULT_GROUPS=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'];
const fill=(argb:string):ExcelJS.Fill=>({type:'pattern',pattern:'solid',fgColor:{argb}});
const numeric=(input?:number)=>input===undefined?null:input;
const style=(state:ComparisonState,id?:string)=>state.styles.find(item=>item.id===id);
const material=(item:CbdStyle|undefined,id?:string)=>item?.materials.find(row=>row.id===id);
const safeName=(name:string,used:Set<string>)=>{const base=(name.replace(/[\\/*?:[\]]/g,' ').trim()||'Comparison').slice(0,31);let result=base,index=2;while(used.has(result))result=`${base.slice(0,27)} ${index++}`;used.add(result);return result};
const stripSeasonPrefix=(name:string|undefined,season:string)=>{const value=name||'없음',escaped=season.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');return value.replace(new RegExp(`^${escaped}(?=\\s+)\\s*`),'')};
const excelStatus=(status:string)=>{const key=status.toUpperCase();return key==='MATCH'||key==='EXACT'?'일치':key==='CODE MATCH'?'코드 일치':key==='NORMALIZED'?'정규화 일치':key==='TOKEN MATCH'?'핵심 토큰 일치':key==='SIMILAR'?'유사 일치':key==='ALIAS'?'별칭 일치':key==='NAME CHANGED'?'자재명 변경':key==='GROUP CHANGED'?'그룹 변경':key==='MANUAL'?'수동 연결':key==='REFERENCE ONLY'||key==='REFERENCE-ONLY'?'기준 시즌만 존재':key==='CURRENT ONLY'||key==='CURRENT-ONLY'?'비교 시즌만 존재':key==='REVIEW'?'검토 필요':key==='DUPLICATE'?'중복 · 검토 필요':key==='MERGED N:1'||key==='N:1'||key==='MANY TO ONE'?'다대일 연결':key==='UNMATCHED'?'미연결':status};
const sum=(values:(number|undefined)[])=>values.reduce<number>((total,item)=>total+(item??0),0);
const groupTotal=(item:CbdStyle|undefined,group:string)=>Object.entries(item?.groupTotals||{}).find(([key])=>key.trim().toUpperCase()===group.trim().toUpperCase())?.[1];
const formula=(text:string,result:number|string):ExcelJS.CellFormulaValue=>({formula:text,result});
const sheetReference=(name:string,cell:string)=>`'${name.replace(/'/g,"''")}'!${cell}`;
const applyHeaderStyle=(cell:ExcelJS.Cell)=>{cell.alignment={horizontal:'center',vertical:'middle',wrapText:false,shrinkToFit:false,indent:0}};
const applyNumericCellStyle=(cell:ExcelJS.Cell)=>{cell.alignment={horizontal:'right',vertical:'top',wrapText:false,shrinkToFit:false,indent:0}};
const applyTextCellStyle=(cell:ExcelJS.Cell,wrapText=false)=>{cell.alignment={horizontal:'left',vertical:'top',wrapText,shrinkToFit:false,indent:0}};
const applyCenteredCellStyle=(cell:ExcelJS.Cell,wrapText=false)=>{cell.alignment={horizontal:'center',vertical:'top',wrapText,shrinkToFit:false,indent:0}};
export const comparisonFinancials=(reference?:CbdStyle,comparison?:CbdStyle):ComparisonFinancials=>({referenceFob:reference?.summary.finalFob,comparisonFob:comparison?.summary.finalFob,referenceMaterialTotal:reference?.summary.totalMaterialCost,comparisonMaterialTotal:comparison?.summary.totalMaterialCost,referenceMaterialToFobRatio:reference?.summary.materialToFobRatio,comparisonMaterialToFobRatio:comparison?.summary.materialToFobRatio,referenceFobEvidence:reference?.summary.fobEvidence??'review-required',comparisonFobEvidence:comparison?.summary.fobEvidence??'review-required'});

type Subtotal={row:number;reference:number;comparison:number};
export type ComparisonExportOptions={scope:'current'|'all';stylePairIds?:string[]};

export function buildComparisonWorkbook(state:ComparisonState,options?:Pick<ComparisonExportOptions,'stylePairIds'>){
 const workbook=new ExcelJS.Workbook(),summary=workbook.addWorksheet('STYLE MATCH'),used=new Set<string>(['STYLE MATCH']);
 workbook.calcProperties.fullCalcOnLoad=true;
 summary.views=[{state:'frozen',ySplit:1}];
 summary.addRow([`기준 시즌 ${state.referenceSeason}`,'기준 STYLE',`비교 시즌 ${state.currentSeason}`,'비교 STYLE','매칭 방식','상태','기준 총 재료비','비교 총 재료비','기준 FOB','비교 FOB']);
 summary.getRow(1).font={bold:true,color:{argb:'FFFFFF'}};summary.getRow(1).fill=fill('1E3A8A');
 const selectedIds=options?.stylePairIds?.length?new Set(options.stylePairIds):null;
 for(const styleMatch of state.styleMatches.filter(match=>!match.excluded&&(!selectedIds||selectedIds.has(match.id)))){
  const reference=style(state,styleMatch.referenceId),comparison=style(state,styleMatch.currentId),financials=comparisonFinancials(reference,comparison),clusters=state.materialMatches.find(set=>set.styleMatchId===styleMatch.id)?.clusters||[],summaryRow=summary.addRow([state.referenceSeason,reference?.styleName||'—',state.currentSeason,comparison?.styleName||'—',styleMatch.method,excelStatus(styleMatch.status),null,null,financials.referenceFob??null,financials.comparisonFob??null]),sheet=workbook.addWorksheet(safeName(comparison?.styleName||reference?.styleName||'Unmatched',used));
  sheet.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};sheet.views=[{state:'frozen',ySplit:15,showGridLines:false}];
  sheet.mergeCells('A1:O1');sheet.getCell('A1').value='FLY Racing CBD 비교표';sheet.getCell('A1').font={bold:true,size:16,color:{argb:'FFFFFF'}};sheet.getCell('A1').fill=fill('1E3A8A');sheet.getCell('A1').alignment={horizontal:'center'};
  sheet.mergeCells('A2:O2');sheet.getCell('A2').value=`기준 시즌 ${state.referenceSeason} – ${stripSeasonPrefix(reference?.styleName,state.referenceSeason)}    |    비교 시즌 ${state.currentSeason} – ${stripSeasonPrefix(comparison?.styleName,state.currentSeason)}`;sheet.getCell('A2').font={bold:true};sheet.getCell('A2').fill=fill('DBEAFE');
  sheet.mergeCells('A3:O3');sheet.getCell('A3').value=`생성일: ${new Date().toLocaleDateString('ko-KR')}`;
  sheet.mergeCells('A5:E5');sheet.getCell('A5').value='FOB 및 그룹별 재료비 비교';sheet.getCell('A5').font={bold:true,size:13};sheet.getCell('A5').alignment={horizontal:'left'};
  sheet.mergeCells('A13:O13');sheet.getCell('A13').value='자재 비교 상세';sheet.getCell('A13').font={bold:true,size:13};
  for(const merged of [...sheet.model.merges]){const match=merged.match(/^([A-O])(1[45]):([A-O])(1[45])$/);if(match)sheet.unMergeCells(merged)}
  sheet.mergeCells('B14:C14');sheet.getCell('B14').value='자재';sheet.mergeCells('E14:G14');sheet.getCell('E14').value='단가';sheet.mergeCells('H14:J14');sheet.getCell('H14').value='소요량';sheet.mergeCells('K14:M14');sheet.getCell('K14').value='재료비';sheet.getCell('A14').value='그룹';sheet.getCell('D14').value='단위';sheet.getCell('N14').value='상태';sheet.getCell('O14').value='비고';sheet.getRow(14).font={bold:true,color:{argb:'FFFFFF'}};sheet.getRow(14).fill=fill('1E3A8A');
  sheet.getRow(15).values=['','기준 시즌 자재','비교 시즌 자재','단위','기준','비교','차액','기준','비교','차액','기준','비교','차액','상태','비고'];sheet.getRow(15).font={bold:true};sheet.getRow(15).fill=fill('BFDBFE');

  const present=new Set([...(reference?.groupOrder||[]),...(comparison?.groupOrder||[]),...Object.keys(reference?.groupTotals||{}),...Object.keys(comparison?.groupTotals||{}),...clusters.map(cluster=>cluster.finalGroup)].filter(Boolean)),groups=[...new Set([...DEFAULT_GROUPS,...present])],costGroups=DEFAULT_GROUPS;
  const subtotals=new Map<string,Subtotal>();let row=16;
  for(const group of groups){
   const grouped=clusters.filter(cluster=>cluster.finalGroup===group).sort((a,b)=>(material(comparison,a.currentRowId||undefined)?.order??99999)-(material(comparison,b.currentRowId||undefined)?.order??99999)),start=row;
   let referenceDetail=0,comparisonDetail=0;
   for(const cluster of grouped){
    const references=cluster.referenceRowIds.map(id=>material(reference,id)).filter((item):item is CbdMaterialRow=>!!item),comparisonRow=material(comparison,cluster.currentRowId||undefined),total=(key:'cost'|'usage'|'extended')=>references.length?sum(references.map(item=>item[key])):undefined;
    const referenceCost=total('cost'),comparisonCost=comparisonRow?.cost,referenceUsage=total('usage'),comparisonUsage=comparisonRow?.usage,referenceExtended=total('extended'),comparisonExtended=comparisonRow?.extended;
    referenceDetail+=referenceExtended??0;comparisonDetail+=comparisonExtended??0;
    sheet.getCell(row,1).value=group;sheet.getCell(row,2).value=references.map(item=>item.material).join('\n');sheet.getCell(row,3).value=comparisonRow?.material||null;sheet.getCell(row,4).value=comparisonRow?.unit||references[0]?.unit||null;
    sheet.getCell(row,5).value=numeric(referenceCost);sheet.getCell(row,6).value=numeric(comparisonCost);sheet.getCell(row,7).value=formula(`F${row}-E${row}`,(comparisonCost??0)-(referenceCost??0));
    sheet.getCell(row,8).value=numeric(referenceUsage);sheet.getCell(row,9).value=numeric(comparisonUsage);sheet.getCell(row,10).value=formula(`I${row}-H${row}`,(comparisonUsage??0)-(referenceUsage??0));
    sheet.getCell(row,11).value=numeric(referenceExtended);sheet.getCell(row,12).value=numeric(comparisonExtended);sheet.getCell(row,13).value=formula(`L${row}-K${row}`,(comparisonExtended??0)-(referenceExtended??0));
    sheet.getCell(row,14).value=excelStatus(cluster.status);sheet.getCell(row,15).value=comparisonRow?.remark||references.map(item=>item.remark).filter(Boolean).join('\n')||null;
    if(cluster.status==='REVIEW')sheet.getRow(row).font={bold:true,color:{argb:'DC2626'}};else if(cluster.status==='CURRENT ONLY')sheet.getRow(row).fill=fill('DCFCE7');else if(cluster.status==='REFERENCE ONLY')sheet.getRow(row).fill=fill('F3F4F6');else if(cluster.status==='GROUP CHANGED')sheet.getRow(row).fill=fill('FFEDD5');else if(cluster.status==='NAME CHANGED')sheet.getRow(row).fill=fill('FEF3C7');else if(cluster.status==='MERGED N:1')sheet.getRow(row).fill=fill('EDE9FE');row++;
   }
   const referenceOriginal=groupTotal(reference,group),comparisonOriginal=groupTotal(comparison,group),useReferenceOriginal=referenceOriginal!==undefined&&Math.abs(referenceOriginal-referenceDetail)>.00005,useComparisonOriginal=comparisonOriginal!==undefined&&Math.abs(comparisonOriginal-comparisonDetail)>.00005;
   let sourceRow:number|undefined;
   if(useReferenceOriginal||useComparisonOriginal){sourceRow=row++;sheet.getCell(sourceRow,1).value=`${group} 원본 CBD 소계`;sheet.getCell(sourceRow,11).value=useReferenceOriginal?referenceOriginal!:null;sheet.getCell(sourceRow,12).value=useComparisonOriginal?comparisonOriginal!:null;sheet.getCell(sourceRow,13).value=formula(`L${sourceRow}-K${sourceRow}`,(useComparisonOriginal?comparisonOriginal!:0)-(useReferenceOriginal?referenceOriginal!:0));sheet.getCell(sourceRow,15).value='원본 CBD의 그룹 소계 사용';sheet.getRow(sourceRow).fill=fill('FEF3C7');}
   const subtotalRow=row++,detailEnd=sourceRow?sourceRow-1:subtotalRow-1,referenceValue=useReferenceOriginal?referenceOriginal!:referenceDetail,comparisonValue=useComparisonOriginal?comparisonOriginal!:comparisonDetail;
   sheet.getCell(subtotalRow,1).value=`${group} 소계`;sheet.getCell(subtotalRow,11).value=formula(useReferenceOriginal?`K${sourceRow}`:`SUM(K${start}:K${detailEnd})`,referenceValue);sheet.getCell(subtotalRow,12).value=formula(useComparisonOriginal?`L${sourceRow}`:`SUM(L${start}:L${detailEnd})`,comparisonValue);sheet.getCell(subtotalRow,13).value=formula(`L${subtotalRow}-K${subtotalRow}`,comparisonValue-referenceValue);sheet.getRow(subtotalRow).font={bold:true};sheet.getRow(subtotalRow).fill=fill('DBEAFE');subtotals.set(group,{row:subtotalRow,reference:referenceValue,comparison:comparisonValue});
  }

  sheet.getRow(6).values=['CBD 그룹',`기준 시즌 ${state.referenceSeason}`,`비교 시즌 ${state.currentSeason}`,'차액','변동률'];sheet.getRow(6).font={bold:true,color:{argb:'FFFFFF'}};sheet.getRow(6).fill=fill('1E3A8A');
  const fobRow=7,fobReference=financials.referenceFob,fobComparison=financials.comparisonFob;sheet.getCell(fobRow,1).value=fobReference===undefined||fobComparison===undefined?'FOB (검토 필요)':'FOB';sheet.getCell(fobRow,2).value=formula(sheetReference('STYLE MATCH',`I${summaryRow.number}`),fobReference??'');sheet.getCell(fobRow,3).value=formula(sheetReference('STYLE MATCH',`J${summaryRow.number}`),fobComparison??'');sheet.getCell(fobRow,4).value=fobReference!==undefined&&fobComparison!==undefined?formula(`C${fobRow}-B${fobRow}`,fobComparison-fobReference):null;sheet.getCell(fobRow,5).value=fobReference!==undefined&&fobComparison!==undefined?formula(`IF(B${fobRow}=0,"",D${fobRow}/B${fobRow})`,fobReference===0?'':(fobComparison-fobReference)/fobReference):null;sheet.getRow(fobRow).font={bold:true};sheet.getRow(fobRow).fill=fill('DBEAFE');
  let groupRow=8;
  for(const group of costGroups){const subtotal=subtotals.get(group)!;sheet.getCell(groupRow,1).value=group;sheet.getCell(groupRow,2).value=formula(`K${subtotal.row}`,subtotal.reference);sheet.getCell(groupRow,3).value=formula(`L${subtotal.row}`,subtotal.comparison);sheet.getCell(groupRow,4).value=formula(`C${groupRow}-B${groupRow}`,subtotal.comparison-subtotal.reference);sheet.getCell(groupRow,5).value=formula(`IF(B${groupRow}=0,"",D${groupRow}/B${groupRow})`,subtotal.reference===0?'':(subtotal.comparison-subtotal.reference)/subtotal.reference);groupRow++}
  const totalRow=groupRow,firstGroupRow=8,lastGroupRow=groupRow-1,referenceTotal=sum(costGroups.map(group=>subtotals.get(group)!.reference)),comparisonTotal=sum(costGroups.map(group=>subtotals.get(group)!.comparison));
  sheet.getCell(totalRow,1).value='TOTAL MATERIAL COST';sheet.getCell(totalRow,2).value=formula(`SUM(B${firstGroupRow}:B${lastGroupRow})`,referenceTotal);sheet.getCell(totalRow,3).value=formula(`SUM(C${firstGroupRow}:C${lastGroupRow})`,comparisonTotal);sheet.getCell(totalRow,4).value=formula(`C${totalRow}-B${totalRow}`,comparisonTotal-referenceTotal);sheet.getCell(totalRow,5).value=formula(`IF(B${totalRow}=0,"",D${totalRow}/B${totalRow})`,referenceTotal===0?'':(comparisonTotal-referenceTotal)/referenceTotal);sheet.getRow(totalRow).fill=fill('DBEAFE');sheet.getRow(totalRow).font={bold:true};
  for(let index=7;index<=totalRow;index++){for(const column of [2,3,4])sheet.getCell(index,column).numFmt='$0.0000';sheet.getCell(index,5).numFmt='0.0%'}
  for(let index=16;index<row;index++){for(const column of [5,6,7,11,12,13])sheet.getCell(index,column).numFmt='$0.0000';for(const column of [8,9,10])sheet.getCell(index,column).numFmt='0.0000';sheet.getRow(index).height=32}
  [20,36,36,9,12,12,12,12,12,12,14,14,14,18,28].forEach((width,index)=>sheet.getColumn(index+1).width=width);
  applyTextCellStyle(sheet.getCell(6,1));sheet.getCell(6,1).alignment.vertical='middle';for(let column=2;column<=5;column++)applyHeaderStyle(sheet.getCell(6,column));
  for(let index=7;index<=totalRow;index++){applyTextCellStyle(sheet.getCell(index,1));sheet.getCell(index,1).alignment.vertical='middle';for(let column=2;column<=5;column++){applyNumericCellStyle(sheet.getCell(index,column));sheet.getCell(index,column).alignment.vertical='middle'}}
  for(let column=1;column<=15;column++){applyHeaderStyle(sheet.getCell(14,column));applyHeaderStyle(sheet.getCell(15,column))}sheet.getRow(14).height=22;sheet.getRow(15).height=22;
  for(let index=16;index<row;index++){applyTextCellStyle(sheet.getCell(index,1));for(const column of [2,3,15])applyTextCellStyle(sheet.getCell(index,column),true);applyCenteredCellStyle(sheet.getCell(index,4));for(let column=5;column<=13;column++)applyNumericCellStyle(sheet.getCell(index,column));applyCenteredCellStyle(sheet.getCell(index,14),true);const lines=Math.max(1,...[2,3,14,15].map(column=>String(sheet.getCell(index,column).value??'').split('\n').length));sheet.getRow(index).height=Math.min(60,Math.max(20,lines*15))}
  sheet.autoFilter=undefined;sheet.getRow(1).height=26;sheet.getRow(2).height=24;
  summaryRow.getCell(7).value=formula(sheetReference(sheet.name,`B${totalRow}`),referenceTotal);summaryRow.getCell(8).value=formula(sheetReference(sheet.name,`C${totalRow}`),comparisonTotal);
 }
 return workbook;
}

const filePart=(value:string)=>value.normalize('NFKC').replace(/[<>:"/\\|?*]+/g,' ').trim().replace(/[^\p{L}\p{N}._-]+/gu,'_').replace(/^_+|_+$/g,'')||'STYLE';
export async function exportComparison(state:ComparisonState,options:ComparisonExportOptions={scope:'all'}){const stylePairIds=options.scope==='current'?options.stylePairIds:undefined,workbook=buildComparisonWorkbook(state,{stylePairIds}),buffer=await workbook.xlsx.writeBuffer(),selected=options.scope==='current'?state.styleMatches.find(match=>stylePairIds?.includes(match.id)):undefined,currentStyle=style(state,selected?.currentId)?.styleName||style(state,selected?.referenceId)?.styleName,fileName=options.scope==='current'?`CBD_비교_${filePart(state.referenceSeason)}_vs_${filePart(state.currentSeason)}_${filePart(currentStyle||'STYLE')}.xlsx`:`CBD_Comparison_${state.referenceSeason||'Reference'}_vs_${state.currentSeason||'Comparison'}.xlsx`;saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),fileName)}
