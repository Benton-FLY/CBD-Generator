import {readTabularWorkbook} from '../spreadsheet';
import {normalizeHeader,normalizeText} from './normalize';
import type {CbdMaterialRow,CbdStyle,ComparisonSide,CbdSummary,UploadedCbd} from './types';

const fields=[['group',['GROUPOF','GROUP']],['material',['MATERIAL','ITEM']],['size',['SIZE','WIDTH']],['unit',['UNIT']],['cost',['COSTPERUNIT','UNITCOST']],['usage',['USAGE']],['loss',['LOSS']],['extended',['EXTENDEDCOST','AMOUNT']],['remark',['REMARK','REMARKS']]] as const;
const num=(v:unknown)=>{if(typeof v==='number'&&Number.isFinite(v))return v;const raw=String(v??'');if(!raw.trim())return undefined;const n=Number(raw.replace(/[$,%\s,]/g,''));return Number.isFinite(n)?(/%/.test(raw)?n/100:n):undefined};
const findHeader=(rows:unknown[][])=>{for(let r=0;r<Math.min(rows.length,30);r++){const cells=rows[r].map(normalizeHeader),map:Record<string,number>={};for(const [key,aliases] of fields){const col=cells.findIndex(c=>aliases.some(a=>c===a||c.includes(a)));if(col>=0)map[key]=col}if(['group','material','unit','cost','usage','loss','extended','remark'].every(k=>map[k]!==undefined))return{row:r,map}}return null};
const styleName=(rows:unknown[][],sheet:string,file:string)=>{for(let r=0;r<Math.min(rows.length,30);r++)for(let c=0;c<(rows[r]?.length||0);c++){if(/MODEL\s*NAME|STYLE(?:\s*NAME)?/.test(normalizeText(rows[r][c]))){const right=String(rows[r][c+1]??'').trim();if(right)return right}}const repeated=new Map<string,number>();rows.slice(0,4).flatMap(r=>r||[]).map(v=>String(v??'').trim()).filter(v=>v.length>3).forEach(v=>repeated.set(v,(repeated.get(v)||0)+1));const title=[...repeated.entries()].filter(([v,n])=>n>=2&&!/FLY RACING CBD SHEET|DATE|GROUP|MATERIAL|COST|USAGE|REMARK/i.test(v)).sort((a,b)=>b[1]-a[1])[0]?.[0];if(title)return title;return rows.slice(0,8).flatMap(r=>r||[]).map(v=>String(v??'').trim()).filter(v=>v.length>4&&!/GROUP|MATERIAL|COST|USAGE|REMARK/i.test(v)).sort((a,b)=>b.length-a.length)[0]||sheet||file.replace(/\.xlsx?$/i,'')};
const label=(value:unknown)=>String(value??'').replace(/[\u00a0\r\n\t]+/g,' ').replace(/ +/g,' ').trim().toUpperCase();
const FOB_LABELS=new Set(['FINAL FOB','FOB PRICE','FINAL FOB PRICE','FOB']);
const EXCLUDED_FOB_LABELS=new Set(['CBD MATERIAL / FOB','CBD 재료비 / FOB','MATERIAL / FOB','FOB RATIO','DIFFERENCE','DIFFERENCE RATE','CHANGE %','MATERIAL RATE']);
type SummaryField='totalMaterialCost'|'laborCost'|'overhead'|'profit'|'finalFob'|'materialToFobRatio'|'erpMaterial'|'difference'|'differenceRate';
const summaryKey=(value:unknown):SummaryField|undefined=>{const text=label(value);if(FOB_LABELS.has(text))return'finalFob';if(text==='CBD MATERIAL / FOB'||text==='CBD 재료비 / FOB'||text==='MATERIAL / FOB'||text==='FOB RATIO'||text==='MATERIAL RATE')return'materialToFobRatio';if(text==='TOTAL MATERIAL COST')return'totalMaterialCost';if(text==='LABOR COST')return'laborCost';if(text==='OVERHEAD')return'overhead';if(text==='PROFIT')return'profit';if(/ERP/.test(text)&&/(PRE|사전)/.test(text))return'erpMaterial';if(text==='DIFFERENCE RATE'||text==='CHANGE %')return'differenceRate';if(text==='DIFFERENCE')return'difference';return undefined};
const summaryValue=(row:unknown[],preferredColumn:number)=>num(row[preferredColumn])??[...row].reverse().map(num).find((value):value is number=>value!==undefined);
const validateFob=(summary:CbdSummary)=>{summary.fobEvidence=summary.finalFob===undefined?'review-required':'explicit';if(summary.totalMaterialCost!==undefined&&summary.materialToFobRatio!==undefined&&summary.materialToFobRatio!==0){summary.calculatedFob=summary.totalMaterialCost/summary.materialToFobRatio;if(summary.finalFob!==undefined)summary.fobValidation=Math.abs(summary.finalFob-summary.calculatedFob)<=Math.max(.01,Math.abs(summary.finalFob)*.001)?'matched':'mismatch'}else summary.fobValidation='not-available'};

export async function parseCbdFiles(files:File[],side:ComparisonSide):Promise<{files:UploadedCbd[];styles:CbdStyle[];errors:string[]}>
{
 const parsedFiles:UploadedCbd[]=[],styles:CbdStyle[]=[],errors:string[]=[];
 for(const file of files){
  try{
   const fileId=`${side}-${file.name}-${file.size}-${file.lastModified}`,sheets=await readTabularWorkbook(file),styleIds:string[]=[];
   for(const sheet of sheets){
    const header=findHeader(sheet.rows);if(!header)continue;
    let group='',order=0;
    const materials:CbdMaterialRow[]=[],summary:CbdSummary={fobEvidence:'review-required'},groupOrder:string[]=[],groupTotals:Record<string,number>={},duplicates=new Map<string,number>();
    for(let r=header.row+1;r<sheet.rows.length;r++){
     const row=sheet.rows[r]||[],material=String(row[header.map.material]??'').trim(),groupCell=String(row[header.map.group]??'').trim(),extended=num(row[header.map.extended]);
     const labeled=row.map((value,column)=>({key:summaryKey(value),raw:label(value),column})).find(item=>item.key&&(!EXCLUDED_FOB_LABELS.has(item.raw)||item.key==='materialToFobRatio'||item.key==='difference'||item.key==='differenceRate'));
     if(labeled?.key){const value=summaryValue(row,header.map.extended);if(value!==undefined)summary[labeled.key]=value;continue}
     if(/SUBTOTAL/i.test(material)||/SUBTOTAL/i.test(groupCell)){const subtotal=extended??num(row.find(v=>num(v)!==undefined));if(subtotal!==undefined)groupTotals[(groupCell||group).replace(/SUBTOTAL/i,'').trim()]=subtotal;continue}
     // Legacy 27-season sheets use an unlabeled numeric row immediately after each group.
     if(!material&&!groupCell&&group&&extended!==undefined){groupTotals[group]=extended;continue}
     if(!material||/^REFERENCES?$|TOTAL MATERIAL|INTERNAL USE ONLY/i.test(material))continue;
     if(groupCell&&!/TOTAL/i.test(groupCell)){group=groupCell;if(!groupOrder.includes(group))groupOrder.push(group)}
     if(!group)group='NEEDS REVIEW';
     const duplicateKey=`${group}\u0000${material}`,sequence=(duplicates.get(duplicateKey)||0)+1;duplicates.set(duplicateKey,sequence);
     materials.push({id:`${side}::${fileId}::${sheet.name}::row-${r+1}::seq-${sequence}`,group,material,size:String(row[header.map.size]??''),unit:String(row[header.map.unit]??''),cost:num(row[header.map.cost]),usage:num(row[header.map.usage]),loss:num(row[header.map.loss]),extended,remark:String(row[header.map.remark]??''),width:String(row[header.map.size]??''),order:order++});
    }
    if(!materials.length)continue;
    validateFob(summary);const id=`${side}::${fileId}::${sheet.name}`;
    styles.push({id,side,fileName:file.name,sheetName:sheet.name,styleName:styleName(sheet.rows,sheet.name,file.name),materials,summary,groupOrder,groupTotals});styleIds.push(id);
   }
   parsedFiles.push({id:fileId,name:file.name,size:file.size,lastModified:file.lastModified,side,styleIds});
  }catch(error){errors.push(`${file.name}: ${error instanceof Error?error.message:'Parsing failed'}`)}
 }
 return{files:parsedFiles,styles,errors};
}
