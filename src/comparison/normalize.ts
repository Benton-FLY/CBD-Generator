const dash=/[\u2010-\u2015\u2212_]+/g;
export const normalizeText=(value:unknown)=>String(value??'').normalize('NFKC').replace(/[\u00a0\r\n\t]+/g,' ').replace(dash,'-').replace(/\s+/g,' ').trim().toUpperCase();
export const normalizeHeader=(value:unknown)=>normalizeText(value).replace(/[&/.#\-\s]+/g,'');
export const normalizeStyle=(value:string)=>normalizeText(value)
  .replace(/\.(XLSX?|XLS)$/g,'').replace(/^\d{2,4}\s+(?=[A-Z])/,'')
  .replace(/\bCBD(?:\s+SHEET)?\b/g,'').replace(/[()（）]?\s*내부\s*[)）]?/g,'')
  .replace(/\s+/g,' ').trim();
export const styleKey=(value:string)=>normalizeStyle(value).replace(/\s*\(S\s*~\s*2XL\)\s*/g,' ').replace(/\s+/g,' ').trim();
const seasonPhrase=/\bFLY\s+RACING\s+(?:20)?\d{2}\s+/g;
export const normalizeMaterial=(value:string,removeSeason=false)=>{
  let result=normalizeText(value).replace(/\s*([+(),/])\s*/g,'$1').replace(/([()])\1+/g,'$1');
  if(removeSeason)result=result.replace(seasonPhrase,'FLY RACING ');
  return result;
};
export const itemCodes=(value:string):string[]=>Array.from(normalizeText(value).match(/(?:JX-\d+|#?[A-Z]+-[A-Z0-9]+|\b\d{4}\b|\b\d+S\/\d+\b|\b\d+D\/\d+\b)/g)||[]);
export const qualifiers=(value:string)=>(['POLYGIENE','PFAS FREE','WMN','YOUTH','BIG SIZE','FR','SL'] as string[]).filter(q=>normalizeText(value).includes(q));
