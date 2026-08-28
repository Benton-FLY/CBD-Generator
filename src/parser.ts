import { classify, normalizeText } from './classifier';
import { readTabularWorkbook } from './spreadsheet';
import type { AppSettings, BomRow, ClassificationDictionary, ColumnMapping, Material, StyleData } from './types';

export const HEADER_ALIASES: Record<string,string[]> = {
  structure:['STRUCTURE'], materialType:['자재구분','MATERIAL TYPE'], sequence:['원순번','순번','SEQ','SEQUENCE'], itemNo:['ITEM#','ITEM NO','ITEM NO.'], item:['ITEM','자재명'], width:['WIDTH','SIZE'], color:['COLOR'], unit:['UNIT'], netUsage:['정소요량','NET USAGE'], bomLoss:['로스율','LOSS'], usage:['소요량','USAGE'], currency:['CURRENCY','통화'], rawPrice:['단가','PRICE'], convertedPrice:['환산단가','CONVERTED PRICE','USD PRICE'], materialCostAdjustment:['자재비용차액대체'], amount:['금액','AMOUNT'], specialFlag:['특수공정'], remark:['비고','REMARK']
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
    const materialCostAdjustment=num(get(row,'materialCostAdjustment')) ?? 0;
    const convertedPrice=(converted ?? (currency==='KRW' ? raw/settings.exchangeRate : raw))+materialCostAdjustment;
    return {id:`${meta.file}:${meta.sheet}:${headerRow+i+2}`,sourceFile:meta.file,sourceSheet:meta.sheet,sourceRow:headerRow+i+2,structure:text(get(row,'structure')),materialType:text(get(row,'materialType')),sequence:text(get(row,'sequence')),itemNo:text(get(row,'itemNo')),item:text(get(row,'item')),width:text(get(row,'width')),color:text(get(row,'color')),unit:text(get(row,'unit')),netUsage:net,bomLoss:loss,usage,currency,rawPrice:raw,convertedPrice,materialCostAdjustment,amount:num(get(row,'amount')),specialFlag:text(get(row,'specialFlag')),remark:text(get(row,'remark'))};
  }).filter(r => r.item && (r.usage !== 0 || r.convertedPrice !== 0 || r.materialType));
}
const slugStyle = (filename:string) => filename.replace(/\.(xlsx?|xls)$/i,'').replace(/\bBOM\b/ig,'').replace(/\b(?:V(?:ER)?\.?\s*)?\d+(?:\.\d+)*\b/ig,'').replace(/\b20\d{2}[._-]\d{1,2}[._-]\d{1,2}\b/g,'').replace(/[_-]+/g,' ').replace(/\s+/g,' ').trim();
const keyOf=(r:BomRow)=>[normalizeText(r.item),normalizeText(r.width),normalizeText(r.unit)].join('|');
export function aggregate(rows: BomRow[], loss: number, dict: ClassificationDictionary): Material[] {
  const buckets=new Map<string,BomRow[]>();
  rows.forEach(r => {
    const special=/특수공임|기타공임/.test(r.materialType);
    const key=special ? `${keyOf(r)}|${normalizeText(r.structure)}|${normalizeText(r.remark)}|${r.id}` : keyOf(r);
    buckets.set(key,[...(buckets.get(key)||[]),r]);
  });
  return [...buckets.entries()].map(([key,sources])=>{
    const usage=sources.reduce((s,r)=>s+r.usage,0);
    const weighted=usage ? sources.reduce((s,r)=>s+r.convertedPrice*r.usage,0)/usage : sources.reduce((s,r)=>s+r.convertedPrice,0)/sources.length;
    const first=sources[0], group=classify(first,dict);
    return {id:key,item:first.item,width:first.width,unit:first.unit,group,included:group!=='EXCLUDE',baseCost:weighted||0,adjustedCost:weighted||0,baseUsage:usage,adjustedUsage:usage,additionalLoss:loss,remark:first.remark,sources,split:false};
  });
}
export async function parseBomFile(file: File, settings: AppSettings, dict: ClassificationDictionary): Promise<{styles:StyleData[]; unmapped?:{sheet:string; rows:unknown[][]}}> {
  const sheets=await readTabularWorkbook(file); const styles:StyleData[]=[];
  for(const sheet of sheets){
    const sheetName=sheet.name, rows=sheet.rows; const found=detectHeader(rows);
    if(!found) return {styles,unmapped:{sheet:sheetName,rows}};
    const bomRows=rowsFromSheet(rows,found.mapping,found.row,{file:file.name,sheet:sheetName},settings);
    if(!bomRows.length) continue;
    const base=slugStyle(file.name)||sheetName; const name=sheets.length>1 ? `${base} - ${sheetName}` : base;
    styles.push({id:`${file.name}:${sheetName}:${Date.now()}`,name,sourceFile:file.name,sourceSheet:sheetName,materials:aggregate(bomRows,settings.defaultLoss,dict),laborRemark:'It also includes the listed special process and packing costs in the factory.'});
  }
  return {styles};
}

export const splitMaterial = (m: Material): Material[] => m.sources.map((r,i)=>({ ...m,id:`${m.id}:split:${i}`,sources:[r],baseUsage:r.usage,adjustedUsage:r.usage,baseCost:r.convertedPrice,adjustedCost:r.convertedPrice,split:true }));
