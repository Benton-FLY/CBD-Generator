import { classify, normalizeText } from './classifier';
import { readTabularWorkbook } from './spreadsheet';
import { roundTo4, type AppSettings, type BomRow, type ClassificationDictionary, type ColumnMapping, type ErpMaterialAuditRow, type Material, type RowStatusDetail, type StyleData } from './types';

export const HEADER_ALIASES: Record<string,string[]> = {
  structure:['STRUCTURE'], materialType:['자재구분','MATERIAL TYPE'], sequence:['원순번','순번','SEQ','SEQUENCE'], serialNo:['일련번호','SERIAL NO','SERIAL NUMBER'], parentSerialNo:['부모일련번호','PARENT SERIAL NO','PARENT SERIAL NUMBER'], level:['레벨','LEVEL'], hasChildren:['하위YN','하위 YN','HAS CHILDREN'], itemNo:['ITEM#','ITEM NO','ITEM NO.'], item:['ITEM','자재명'], width:['WIDTH','SIZE'], color:['COLOR'], unit:['UNIT'], netUsage:['정소요량','NET USAGE'], bomLoss:['로스율','LOSS'], usage:['소요량','USAGE'], currency:['CURRENCY','통화'], rawPrice:['단가','PRICE'], convertedPrice:['환산단가','CONVERTED PRICE','USD PRICE'], ancillaryCost:['부대비용','ADDITIONAL COST','ANCILLARY COST'], materialCostAdjustment:['자재비용차액대체'], amount:['금액','AMOUNT'], specialFlag:['특수공정'], remark:['비고','REMARK']
};
const cleanHeader = (v: unknown) => normalizeText(v).replace(/[\s_.-]/g,'');
export function detectHeader(rows: unknown[][]): { row: number; mapping: ColumnMapping } | null {
  let bestRow = -1, bestScore = -1, bestMapping: ColumnMapping = {};
  rows.slice(0,30).forEach((cells,row) => {
    const mapping: ColumnMapping = {};
    cells.forEach((cell,col) => Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => {
      if (aliases.some(a => cleanHeader(a) === cleanHeader(cell))) mapping[field] = col;
    }));
    const score = ['item','unit','usage','netUsage','convertedPrice','rawPrice','structure'].filter(k => mapping[k] !== undefined).length;
    if (score > bestScore) { bestRow=row; bestMapping=mapping; bestScore=score; }
  });
  return bestScore >= 4 && bestMapping.item !== undefined ? {row:bestRow,mapping:bestMapping} : null;
}
const num = (v: unknown): number|undefined => { if (v === '' || v == null) return undefined; const n=Number(String(v).replace(/[$,%\s,]/g,'')); return Number.isFinite(n)?n:undefined; };
const text = (v: unknown) => String(v ?? '').trim();
export function rowsFromSheet(data: unknown[][], mapping: ColumnMapping, headerRow: number, meta: {file:string; sheet:string}, settings: AppSettings): BomRow[] {
  const get=(row:unknown[],field:string)=>mapping[field]===undefined?undefined:row[mapping[field]!];
  return data.slice(headerRow+1).map((row,i) => {
    const usageCol=num(get(row,'usage')), net=num(get(row,'netUsage')), loss=num(get(row,'bomLoss')) ?? 0;
    const usage=usageCol ?? (net === undefined ? 0 : net*(1+loss/100));
    const converted=num(get(row,'convertedPrice')), raw=num(get(row,'rawPrice')) ?? 0, currency=normalizeText(get(row,'currency'));
    const materialCostAdjustment=num(get(row,'materialCostAdjustment')) ?? 0,ancillaryCost=num(get(row,'ancillaryCost'))??0;
    const convertedPrice=(converted ?? (currency==='KRW' ? raw/settings.exchangeRate : raw))+materialCostAdjustment;
    return {id:`${meta.file}:${meta.sheet}:${headerRow+i+2}`,sourceFile:meta.file,sourceSheet:meta.sheet,sourceRow:headerRow+i+2,structure:text(get(row,'structure')),materialType:text(get(row,'materialType')),sequence:text(get(row,'sequence')),serialNo:text(get(row,'serialNo')),parentSerialNo:text(get(row,'parentSerialNo')),level:text(get(row,'level')),hasChildren:text(get(row,'hasChildren')),itemNo:text(get(row,'itemNo')),item:text(get(row,'item')),width:text(get(row,'width')),color:text(get(row,'color')),unit:text(get(row,'unit')),netUsage:net,bomLoss:loss,usage,currency,rawPrice:raw,convertedPrice,ancillaryCost,materialCostAdjustment,amount:num(get(row,'amount')),specialFlag:text(get(row,'specialFlag')),remark:text(get(row,'remark'))};
  }).filter(r => r.item && (r.usage !== 0 || r.convertedPrice !== 0 || r.materialType));
}
export function buildErpMaterialAudit(rows:BomRow[]):ErpMaterialAuditRow[]{
  const parentIds=new Set(rows.map(r=>normalizeText(r.parentSerialNo)).filter(Boolean));
  return rows.map(source=>{const isParent=normalizeText(source.hasChildren)==='Y'||Boolean(source.serialNo&&parentIds.has(normalizeText(source.serialNo)));const material=normalizeText(source.materialType)==='자재';const usedFallback=material&&!isParent&&source.amount===undefined;const included=material&&!isParent;const includedAmount=included?(source.amount??roundTo4(source.usage*(source.convertedPrice+(source.ancillaryCost??0)))):0;const reason=!material?`제외: 자재구분=${source.materialType||'공란'}`:isParent?'제외: 상위 복합자재':usedFallback?'검토: 금액 공란으로 fallback 계산':'포함: 말단 자재';return{source,isParent,included,includedAmount,usedFallback,reason};});
}
export function buyerRowsWithoutHierarchyDuplicates(rows:BomRow[],dict:ClassificationDictionary):BomRow[]{
  const childrenByParent=new Map<string,BomRow[]>();rows.forEach(r=>{const key=normalizeText(r.parentSerialNo);if(key)childrenByParent.set(key,[...(childrenByParent.get(key)||[]),r])});
  const descendantIds=new Set<string>();const visit=(serial:string)=>{for(const child of childrenByParent.get(normalizeText(serial))||[]){descendantIds.add(child.id);if(child.serialNo)visit(child.serialNo)}};rows.filter(r=>normalizeText(r.hasChildren)==='Y'||childrenByParent.has(normalizeText(r.serialNo))).forEach(r=>{if(r.serialNo)visit(r.serialNo)});
  const representatives=rows.filter(r=>!descendantIds.has(r.id));const listOnlyChildren=rows.filter(r=>descendantIds.has(r.id)&&classify(r,dict)==='SPECIAL PROCESS (LIST ONLY)');return [...representatives,...listOnlyChildren];
}
export const displayStyleName=(value:string)=>value.normalize('NFKC').replace(/\.(xlsx?|xls)$/i,'').replace(/\bBOM\b/ig,'').replace(/^\s*\d{2}\s+/,'').replace(/\s*사전원가\s*$/,'').replace(/_/g,' ').replace(/\s+/g,' ').trim();
const combined=(values:string[])=>[...new Set(values.map(v=>v.trim()).filter(Boolean))].join(' / ');
const keyOf=(r:BomRow)=>[normalizeText(r.item),normalizeText(r.width),normalizeText(r.unit)].join('|');
export function aggregate(rows: BomRow[], loss: number, dict: ClassificationDictionary): Material[] {
  const buckets=new Map<string,BomRow[]>();
  rows.forEach(r => {
    const special=classify(r,dict)==='SPECIAL PROCESS (LIST ONLY)';
    const key=special ? `${keyOf(r)}|${normalizeText(r.structure)}|${normalizeText(r.remark)}|${r.id}` : keyOf(r);
    buckets.set(key,[...(buckets.get(key)||[]),r]);
  });
  return [...buckets.entries()].map(([key,sources])=>{
    const usage=sources.reduce((s,r)=>s+r.usage,0);
    const weighted=usage ? sources.reduce((s,r)=>s+r.convertedPrice*r.usage,0)/usage : sources.reduce((s,r)=>s+r.convertedPrice,0)/sources.length;
    const first=sources[0], group=classify(first,dict);
    const originalRemark=combined(sources.map(r=>r.remark))||combined(sources.map(r=>r.structure));
    return {id:key,item:first.item,width:first.width,unit:first.unit,group,included:group!=='EXCLUDE',baseCost:weighted||0,adjustedCost:weighted||0,baseUsage:usage,adjustedUsage:usage,additionalLoss:loss,remark:originalRemark,originalRemark,remarkEdited:false,sources,split:false};
  });
}
export function statusDetails(materials:Material[]):RowStatusDetail[]{
  return materials.flatMap(m=>m.sources.flatMap(r=>{
    let disposition:RowStatusDetail['disposition']|undefined,reason='',result='';
    const type=normalizeText(r.materialType),item=normalizeText(r.item);
    if(type.includes('봉제공임')||/SEWING COST/.test(item)){disposition='separate';reason='Buyer CBD 상세목록 비공개 – Internal Sewing으로 반영';result='INTERNAL SEWING';}
    else if(type.includes('포장공임')||/PACKING COST/.test(item)){disposition='separate';reason='Buyer CBD 상세목록 비공개 – Internal Packing으로 반영';result='INTERNAL PACKING';}
    else if(!m.included){disposition='excluded';reason='사용자가 Include 해제';result='계산 및 목록에서 제외';}
    else if(m.group==='NEEDS REVIEW'){disposition='review';reason='분류 규칙을 찾지 못해 검토 필요';result='검토 대기';}
    else if(m.group==='SPECIAL PROCESS (LIST ONLY)'){disposition='separate';reason='SPECIAL PROCESS (LIST ONLY)에 별도 반영';result='INTERNAL SPECIAL PROCESS 및 Buyer 목록';}
    else if(m.group==='EXCLUDE'){disposition='excluded';reason='완전 제외';result='계산 및 목록에서 제외';}
    if(!disposition)return [];
    return [{id:r.id,disposition,sourceRow:r.sourceRow,itemNo:r.itemNo,item:r.item,structure:r.structure,materialType:r.materialType,unit:r.unit,convertedPrice:r.convertedPrice,materialCostAdjustment:r.materialCostAdjustment,remark:r.remark,result,reason}];
  }));
}
export async function parseBomFile(file: File, settings: AppSettings, dict: ClassificationDictionary): Promise<{styles:StyleData[]; unmapped?:{sheet:string; rows:unknown[][]}}> {
  const sheets=await readTabularWorkbook(file); const allRows:BomRow[]=[]; const sourceSheets:string[]=[];
  for(const sheet of sheets){
    const sheetName=sheet.name, rows=sheet.rows; const found=detectHeader(rows);
    if(!found) return {styles:[],unmapped:{sheet:sheetName,rows}};
    const bomRows=rowsFromSheet(rows,found.mapping,found.row,{file:file.name,sheet:sheetName},settings);
    if(!bomRows.length) continue;
    allRows.push(...bomRows); sourceSheets.push(sheetName);
  }
  if(!allRows.length)return {styles:[]};
  const name=displayStyleName(file.name)||displayStyleName(sourceSheets[0]),audit=buildErpMaterialAudit(allRows),erpMaterialCost=roundTo4(audit.reduce((sum,row)=>sum+row.includedAmount,0)); const materials=aggregate(buyerRowsWithoutHierarchyDuplicates(allRows,dict),settings.defaultLoss,dict);
  return {styles:[{id:`${file.name}:${Date.now()}`,name,sourceFile:file.name,sourceSheet:sourceSheets.join(', '),materials,statusDetails:statusDetails(materials),erpMaterialCost,erpAudit:audit,laborRemark:'It also includes the listed special process and packing costs in the factory.'}]};
}

export const splitMaterial = (m: Material): Material[] => m.sources.map((r,i)=>({ ...m,id:`${m.id}:split:${i}`,sources:[r],baseUsage:r.usage,adjustedUsage:r.usage,baseCost:r.convertedPrice,adjustedCost:r.convertedPrice,remark:r.remark||r.structure,originalRemark:r.remark||r.structure,remarkEdited:false,split:true }));
