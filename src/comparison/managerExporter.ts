import {comparisonRowIds} from './relations';
import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';
import {flyLogo} from './assets/fly-logo';
import {comparisonModel,COST_GROUPS} from './model';
import {unitChanged} from './groups';
import type {CbdMaterialRow,CbdStyle,ComparisonState,MaterialMatchCluster} from './types';
import type {ComparisonExportOptions} from './exporter';

export const MANAGER_WIDTHS=[15.25,56.5,9.125,6.125,7.875,6.25,6.25,13.25,23.375,1.25,9,6.625,9,6.625,9,1.25,15.25,56.5,9.125,6.125,7.875,6.25,6.25,13.25,23.375];
const COLORS={blue:'FF99CCFF',gray:'FFC0C0C0',yellow:'FFFFFF00',cyan:'FFCCFFFF',pink:'FFFFCCFF'};
const fill=(color:string):ExcelJS.Fill=>({type:'pattern',pattern:'solid',fgColor:{argb:color}});
const numberFormat='0.0000;[Red](0.0000);0.0000';
const percentFormat='0.0%;[Red](0.0%);0.0%';
const formula=(formula:string,result:number|string):ExcelJS.CellFormulaValue=>({formula,result});
const amount=(row?:CbdMaterialRow)=>row?.extended??0;
const safeSheetName=(name:string,used:Set<string>)=>{
 const base=(name.replace(/[\\/*?:[\]\x00-\x1f]/g,' ').replace(/^'+|'+$/g,'').trim()||'STYLE').slice(0,31).replace(/^'+|'+$/g,'')||'STYLE';
 let value=base,index=2;while(used.has(value.toLowerCase())){const suffix=` (${index++})`;value=base.slice(0,31-suffix.length)+suffix}used.add(value.toLowerCase());return value;
};
const filePart=(name:string)=>name.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').trim()||'STYLE';

/** Validates identity coverage before writing any monetary totals. No fuzzy matching here. */
export function managerRelations(reference:CbdStyle|undefined,comparison:CbdStyle|undefined,clusters:MaterialMatchCluster[]){
 const refs=new Map(reference?.materials.map(r=>[r.id,r])||[]),curs=new Map(comparison?.materials.map(r=>[r.id,r])||[]);
 const seenR=new Set<string>(),seenC=new Set<string>();
 const relations=clusters.map(cluster=>{
  const references=cluster.referenceRowIds.map(id=>{const row=refs.get(id);if(!row||seenR.has(id))throw new Error(`Invalid or repeated reference material: ${id}`);seenR.add(id);return row});
  if(cluster.currentRowIds&&cluster.currentRowId&&!cluster.currentRowIds.includes(cluster.currentRowId))throw new Error('Split relation must contain its anchor.');
  const currents=comparisonRowIds(cluster).map(id=>{const row=curs.get(id);if(!row||seenC.has(id))throw new Error(`Invalid or repeated comparison material: ${id}`);seenC.add(id);return row});
  return {cluster,references,current:currents[0] as CbdMaterialRow|undefined,currents};
 });
 // A wholly unmatched style still has a valid, empty opposite CBD.
 if(!reference||!comparison){
  for(const row of reference?.materials||[])if(!seenR.has(row.id))relations.push({cluster:{id:row.id,referenceRowIds:[row.id],currentRowId:null,relationType:'one-to-one',matchSource:'auto',finalGroup:row.group,status:'REFERENCE ONLY',confidence:0},references:[row],current:undefined,currents:[]});
  for(const row of comparison?.materials||[])if(!seenC.has(row.id))relations.push({cluster:{id:row.id,referenceRowIds:[],currentRowId:row.id,relationType:'one-to-one',matchSource:'auto',finalGroup:row.group,status:'CURRENT ONLY',confidence:0},references:[],current:row,currents:[row]});
 }else if(seenR.size!==refs.size||seenC.size!==curs.size)throw new Error('Material relationships are incomplete. Confirm material matches before export.');
 return relations;
}

export function buildManagerWorkbook(state:ComparisonState,options?:Pick<ComparisonExportOptions,'stylePairIds'>){
 const workbook=new ExcelJS.Workbook(),used=new Set<string>();
 workbook.calcProperties.fullCalcOnLoad=true;
 const logo=workbook.addImage({base64:flyLogo,extension:'png'});
 const selected=options?.stylePairIds?new Set(options.stylePairIds):undefined;
 for(const pair of state.styleMatches.filter(p=>!p.excluded&&(!selected||selected.has(p.id)))){
  const {reference,comparison,clusters}=comparisonModel(state,pair);
  const relations=managerRelations(reference,comparison,clusters);
  const sheet=workbook.addWorksheet(safeSheetName(comparison?.styleName||reference?.styleName||'STYLE',used));
  MANAGER_WIDTHS.forEach((width,i)=>sheet.getColumn(i+1).width=width);
  sheet.views=[{state:'normal',showGridLines:false,zoomScale:105}];
  sheet.pageSetup={paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,printTitlesRow:'1:5',margins:{left:.2,right:.2,top:.25,bottom:.25,header:0,footer:0}};
  for(const col of [0,16])sheet.addImage(logo,{tl:{col:col+.05,row:0},ext:{width:106,height:106*44/143},editAs:'oneCell'});
  sheet.getRow(1).height=18;sheet.getRow(2).height=22;
  sheet.mergeCells('C2:G2');sheet.mergeCells('S2:V2');
  for(const [item,season,label,title,date] of [[comparison,state.currentSeason,'B2','C2','I3'],[reference,state.referenceSeason,'R2','S2','Y3']] as const){
   sheet.getCell(label).value='Model Name:';
   const cell=sheet.getCell(title);cell.value=item?`${season} ${item.styleName.replace(new RegExp(`^${season.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}\\s+`),'')}`:null;
   cell.fill=fill(COLORS.gray);cell.font={name:'Arial',size:10,bold:true,color:{argb:title==='C2'?'FFFF0000':'FF000000'}};
   sheet.getCell(date).value=item?.sourceDate??null;
  }
  sheet.getCell('B4').value='REFERENCES';sheet.getCell('R4').value='REFERENCES';
  const headers=['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark'];
  for(const offset of [0,16])headers.forEach((text,i)=>sheet.getCell(5,offset+i+1).value=text);
  ['단가차이','인상%','소요량 차이','인상%','비용 차이'].forEach((text,i)=>sheet.getCell(5,i+11).value=text);
  sheet.getRow(5).height=30;
  const tint=(r:number,cols:number[],color:string)=>cols.forEach(c=>sheet.getCell(r,c).fill=fill(color));
  const difference=(r:number,out:string,left:string,right:string,a?:number,b?:number)=>{
   if(a===undefined&&b===undefined)return;
   sheet.getCell(`${out}${r}`).value=formula(`IF(${left}${r}="",IF(${right}${r}="","",-${right}${r}),IF(${right}${r}="",${left}${r},${left}${r}-${right}${r}))`,(a??0)-(b??0));
  };
  const rate=(r:number,out:string,delta:string,base:string,b?:number)=>{
   if(b===undefined||b===0)return;
   const d=sheet.getCell(`${delta}${r}`).result;if(typeof d!=='number')return;
   sheet.getCell(`${out}${r}`).value=formula(`IF(${base}${r}="","",IF(${base}${r}=0,"",${delta}${r}/${base}${r}))`,d/b);
  };
  const writeMaterial=(r:number,offset:number,item:CbdMaterialRow|undefined,cluster:MaterialMatchCluster)=>{
   if(!item)return;
   const moved=item.group.trim().toUpperCase()!==cluster.finalGroup;
   const remark=[item.remark,moved?`[GROUP MOVE: ${item.group} → ${cluster.finalGroup}]`:''].filter(Boolean).join('\n');
   const values=[displayGroup(cluster.finalGroup),item.material,item.size,item.unit,item.cost,item.usage,item.loss,item.extended,remark];
   // Preserve the common parser's verified Extended Cost, including PCS/fixed/special values.
   values.forEach((v,i)=>sheet.getCell(r,offset+i+1).value=v===undefined||v===''?null:v);
   if(moved)tint(r,[offset+9,15],COLORS.pink);
  };
  const groups=[...COST_GROUPS,...new Set(relations.map(r=>r.cluster.finalGroup).filter(g=>!COST_GROUPS.includes(g)&&!g.toUpperCase().startsWith('SPECIAL PROCESS'))),...new Set(['SPECIAL PROCESS (LIST ONLY)',...relations.map(r=>r.cluster.finalGroup).filter(g=>g.toUpperCase().startsWith('SPECIAL PROCESS'))])];
  // Normalize only the display block for multiline source spelling.
  const displayGroup=(g:string)=>g.toUpperCase().startsWith('SPECIAL PROCESS')?'SPECIAL PROCESS (LIST ONLY)':g;
  const subtotals:{group:string;row:number;current:number;reference:number}[]=[];
  let row=6;
  for(const group of [...new Set(groups.map(displayGroup))]){
   const listOnly=group.startsWith('SPECIAL PROCESS'),start=row;
   const position=(r:typeof relations[number])=>r.cluster.manualOrder??(r.references.length?Math.min(...r.references.map(m=>m.order)):Number.MAX_SAFE_INTEGER);
   const grouped=relations.filter(r=>displayGroup(r.cluster.finalGroup)===group).sort((a,b)=>position(a)-position(b));
   // Explicit manual placement takes priority, then reference order; new-only rows stay at the group end.
   let currentTotal=0,referenceTotal=0;
   for(const {cluster,references,current,currents} of grouped){
    const anchor=row,count=Math.max(1,references.length,currents.length),end=anchor+count-1;
    for(let i=0;i<count;i++,row++){
     const cur=currents[i],ref=references[i];
     writeMaterial(row,0,cur,cluster);writeMaterial(row,16,ref,cluster);
     currentTotal+=amount(cur);referenceTotal+=amount(ref);
     if(!listOnly&&count===1){
      if(!unitChanged(references,current)){
       difference(row,'K','E','U',cur?.cost,ref?.cost);rate(row,'L','K','U',ref?.cost);
       difference(row,'M','F','V',cur?.usage,ref?.usage);rate(row,'N','M','V',ref?.usage);
      }
      difference(row,'O','H','X',cur?.extended,ref?.extended);
     }
     if(!listOnly){
      if(!cur||!ref||cur.material!==ref.material)tint(row,[...(cur?[2]:[]),...(ref?[18]:[])],COLORS.pink);
      for(const [a,b,col] of [[cur?.cost,ref?.cost,11],[cur?.usage,ref?.usage,13],[cur?.extended,ref?.extended,15]] as const)
       if(a!==undefined&&b!==undefined&&Math.abs(a-b)>Math.max(.00005,Math.max(Math.abs(a),Math.abs(b))*.0001)&&sheet.getCell(row,col).value!==null)tint(row,[col],COLORS.pink);
      if(cluster.status==='REVIEW')tint(row,[...(cur?[9]:[]),...(ref?[25]:[])],COLORS.pink);
     }
    }
    if(!listOnly&&count>1)sheet.getCell(`O${anchor}`).value=formula(`SUM(H${anchor}:H${end})-SUM(X${anchor}:X${end})`,currents.reduce((s,r)=>s+amount(r),0)-references.reduce((s,r)=>s+amount(r),0));
   }
   // Keep a separator row on both sides; never fill an absent counterpart.
   row++;
   const subtotal=row++;
   for(const [item,label,col,total] of [[comparison,'A','H',currentTotal],[reference,'Q','X',referenceTotal]] as const){
    sheet.getCell(`${label}${subtotal}`).value=`${group} SUBTOTAL`;
    if(item&&!listOnly)sheet.getCell(`${col}${subtotal}`).value=formula(`SUM(${col}${start}:${col}${subtotal-1})`,total);
    const source=Object.entries(item?.groupTotals||{}).find(([key])=>displayGroup(key.trim().toUpperCase())===group)?.[1];
    if(!listOnly&&source!==undefined&&Math.abs(source-total)>.0001)sheet.getCell(`${col==='H'?'I':'Y'}${subtotal}`).value=`원본 CBD 소계 (참고): ${source.toFixed(4)}; 비교 합산 제외`;
   }
   tint(subtotal,[1,2,3,4,5,6,7,8,9,17,18,19,20,21,22,23,24,25],COLORS.yellow);
   if(!listOnly)difference(subtotal,'O','H','X',comparison?currentTotal:undefined,reference?referenceTotal:undefined);
   subtotals.push({group,row:subtotal,current:currentTotal,reference:referenceTotal});row++;
  }
  const totalRow=row,fobRow=row+4,costSubtotals=subtotals.filter(s=>COST_GROUPS.includes(s.group));
  const labels=['Total material cost','Labor cost','Overhead','Profit','FOB PRICE'];
  for(let i=0;i<labels.length;i++,row++){
   for(const [item,offset,col,side] of [[comparison,0,'H','current'],[reference,16,'X','reference']] as const){
    sheet.mergeCells(row,offset+1,row,offset+7);sheet.getCell(row,offset+1).value=labels[i];
    if(item){
     const key=['totalMaterialCost','laborCost','overhead','profit','finalFob'][i];sheet.getCell(row,offset+9).value=item.summary.remarks?.[key]??null;
     if(i===0)sheet.getCell(`${col}${row}`).value=formula(`SUM(${costSubtotals.map(s=>`${col}${s.row}`).join(',')})`,costSubtotals.reduce((v,s)=>v+s[side],0));
     else {const value=i===1?item.summary.laborCost:i===2?item.summary.overhead:i===3?item.summary.profit:item.summary.fobEvidence==='ratio-derived'?undefined:item.summary.finalFob;sheet.getCell(`${col}${row}`).value=value??null;}
    }
    tint(row,Array.from({length:9},(_,j)=>offset+j+1),COLORS.cyan);
   }
   const value=(col:string)=>{const cell=sheet.getCell(`${col}${row}`),v=cell.value;return typeof v==='number'?v:typeof cell.result==='number'?cell.result:undefined};
   difference(row,'O','H','X',value('H'),value('X'));
  }
  for(const [item,col,remark] of [[comparison,'H','I'],[reference,'X','Y']] as const){
   if(!item)continue;
   const fob=sheet.getCell(`${col}${fobRow}`).value;
   if(typeof fob==='number'&&fob!==0){
    const total=sheet.getCell(`${col}${totalRow}`).result as number;
    sheet.getCell(`${col}${row}`).value=formula(`IF(${col}${fobRow}=0,"",${col}${totalRow}/${col}${fobRow})`,total/fob);sheet.getCell(`${col}${row}`).numFmt=percentFormat;sheet.getCell(`${remark}${row}`).value='Material cost ratio';
   }
   const extras=[['사전원가 재료비',item.summary.erpMaterial],['사전원가와 CBD 재료비 차이',item.summary.difference],['차이율',item.summary.differenceRate]] as const;
   extras.forEach(([label,value],i)=>{if(value!==undefined){sheet.getCell(`${col}${row+i+1}`).value=value;sheet.getCell(`${remark}${row+i+1}`).value=label;if(i===2)sheet.getCell(`${col}${row+i+1}`).numFmt=percentFormat}});
   const sourceTotal=item.summary.totalMaterialCost,comparisonTotal=sheet.getCell(`${col}${totalRow}`).result as number;
   if(sourceTotal!==undefined&&Math.abs(sourceTotal-comparisonTotal)>.0001){
    sheet.getCell(`${col}${row+4}`).value=sourceTotal;sheet.getCell(`${remark}${row+4}`).value='원본 CBD Total (참고, 합산 제외)';
    sheet.getCell(`${col}${row+5}`).value=formula(`${col}${row+4}-${col}${totalRow}`,sourceTotal-comparisonTotal);sheet.getCell(`${remark}${row+5}`).value='원본 Total − 최종 비교 Total (원본 소계·반올림 차이)';
   }
  }
  const last=Math.max(fobRow,sheet.rowCount);
  for(let r=1;r<=last;r++)for(let c=1;c<=25;c++){
   if(c===10||c===16)continue;
   const cell=sheet.getCell(r,c);if(cell.isMerged&&cell.master.address!==cell.address)continue;const local=c>16?c-16:c;
   cell.font={name:'Arial',size:8,...cell.font};
   const text=local===2||local===9||r<5,center=local===1||local===3||local===4;
   cell.alignment={horizontal:r>=totalRow&&r<=fobRow&&cell.isMerged?'left':r===5?'center':text?'left':center?'center':'right',vertical:'middle',wrapText:r===5||text||center};
   if(r>=5){
    const black={argb:'FF000000'};
    cell.border={left:{style:c===1||c===17?'medium':'thin',color:black},right:{style:c===9||c===25?'medium':'thin',color:black},top:{style:r===5?'medium':'thin',color:black},bottom:{style:r===last?'medium':'thin',color:black}};
    if(!cell.numFmt)cell.numFmt=[7,12,14,23].includes(c)?percentFormat:numberFormat;
   }
   if(subtotals.some(s=>s.row===r)||(r>=totalRow&&r<=fobRow))cell.font={name:'Arial',size:8,bold:true};
   if(r===5){cell.fill=fill(COLORS.blue);cell.font={name:'Arial',size:10,bold:true};}
  }
  for(let r=6;r<=last;r++){
   let lines=1;
   for(const col of [1,2,3,4,9,17,18,19,20,25]){
    const cell=sheet.getCell(r,col);if(cell.isMerged)continue;
    const width=MANAGER_WIDTHS[col-1];
    const value=String(cell.value??'');const units=(s:string)=>[...s].reduce((n,c)=>n+(c.charCodeAt(0)>255?2:1),0);
    lines=Math.max(lines,value.split('\n').reduce((n,s)=>n+Math.max(1,Math.ceil(units(s)/Math.max(8,width*1.05))),0));
   }
   sheet.getRow(r).height=Math.min(409,Math.max(15,lines*12+4));
  }
  sheet.pageSetup.printArea=`A1:Y${last}`;
 }
 return workbook;
}

export async function exportManagerComparison(state:ComparisonState,options:ComparisonExportOptions={scope:'all'}){
 const ids=options.scope==='current'?options.stylePairIds:undefined;
 if(options.scope==='current'&&ids?.length!==1)throw new Error('Select one style to export.');
 const workbook=buildManagerWorkbook(state,{stylePairIds:ids});
 if(!workbook.worksheets.length)throw new Error('No styles are selected for comparison.');
 const pair=state.styleMatches.find(p=>ids?.includes(p.id));
 const name=state.styles.find(s=>s.id===pair?.currentId)?.styleName||state.styles.find(s=>s.id===pair?.referenceId)?.styleName||'STYLE';
 const buffer=await workbook.xlsx.writeBuffer();
 saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),options.scope==='current'?`${filePart(name)}_CBD_Comparison_Manager_Format.xlsx`:'All_Styles_CBD_Comparison_Manager_Format.xlsx');
}
