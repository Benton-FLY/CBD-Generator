const dash=/[\u2010-\u2015\u2212_]+/g;
const auxiliaryWords=new Set(['VT','NEW','LOW','PRICE','PFAS','FREE','SET']);
const productTypes=['HANG TAG','WOVEN LABEL','PRINT MATERIAL','WEBBING','BUCKLE','ZIPPER','ELASTIC','LABEL','STICKER','PATCH','FABRIC','THREAD','VELCRO','EYELET','SNAP','BUTTON','TAPE'];
const positions=['LEFT','RIGHT','FRONT','BACK','INNER','OUTER','A','B'];
const demographics=['WMN','WOMEN','YOUTH','PEE WEE'];

export const normalizeText=(value:unknown)=>String(value??'').normalize('NFKC').replace(/[\u00a0\r\n\t]+/g,' ').replace(dash,'-').replace(/\s+/g,' ').trim().toUpperCase();
export const normalizeHeader=(value:unknown)=>normalizeText(value).replace(/[&/.#\-\s]+/g,'');
export const normalizeStyle=(value:string)=>normalizeText(value).replace(/\.(XLSX?|XLS)$/g,'').replace(/^\d{2,4}\s+(?=[A-Z])/,'').replace(/\bCBD(?:\s+SHEET)?\b/g,'').replace(/[()（）]?\s*내부\s*[)）]?/g,'').replace(/\s+/g,' ').trim();
export const styleKey=(value:string)=>normalizeStyle(value).replace(/\s*\(S\s*~\s*2XL\)\s*/g,' ').replace(/\s+/g,' ').trim();

export interface MaterialAttributes {normalized:string;coreName:string;coreTokens:string[];auxiliaryTokens:string[];codes:string[];specs:string[];styles:string[];composition:string[];positions:string[];demographics:string[];productTypes:string[]}
const unique=(values:string[])=>[...new Set(values.filter(Boolean))].sort();
const canonical=(value:string)=>normalizeText(value).replace(/^VT[- ]+/,'VT ').replace(/\bHANG\s*TAG\b/g,'HANG TAG').replace(/\b(\d+(?:\.\d+)?)\s*M\s*\/\s*M\b/g,'$1MM').replace(/\b(\d+(?:\.\d+)?)\s*MM\b/g,'$1MM').replace(/\bAND\b|&/g,' AND ').replace(/\s*([(),/+])\s*/g,' $1 ').replace(/\s+-\s+/g,' ').replace(/\s+/g,' ').trim();

export function extractMaterialAttributes(value:string):MaterialAttributes {
  const normalized=canonical(value);
  const codes=unique([...(normalized.match(/#[A-Z0-9]+(?:\s+[A-Z]\d+)?/g)||[]),...(normalized.match(/\b(?:JX|VT|ES|PB|VSC)-[A-Z0-9-]+\b/g)||[])].map(code=>code.replace(/\s+/g,' ')));
  const specs=unique(normalized.match(/\b\d+(?:\.\d+)?MM\b/g)||[]),styles=unique(normalized.match(/\b[A-Z]+-\d+[A-Z0-9-]*\b/g)||[]);
  const composition=unique([...(normalized.match(/\b\d+\s+PAGES?\b/g)||[]),...(normalized.match(/\b\d+\s+EYELETS?\b/g)||[])]);
  const present=(terms:string[])=>terms.filter(term=>new RegExp(`(?:^|\\s)${term.replace(' ','\\s+')}(?:$|\\s)`).test(normalized));
  const foundProducts=present(productTypes).filter(type=>type!=='LABEL'||!normalized.includes('WOVEN LABEL')),foundPositions=present(positions),foundDemographics=present(demographics);
  const supplierTokens=unique(Array.from(normalized.matchAll(/\(([^()]*)\)/g)).flatMap(match=>{const content=match[1].trim();return /^(?:[A-Z]{1,12}|[A-Z]+\s+[A-Z]+)$/.test(content)&&!codes.some(code=>code.includes(content))&&!specs.some(spec=>content.includes(spec))?content.split(/\s+/):[]}));
  const seasonTokens=unique(Array.from(normalized.matchAll(/\bFLY RACING ((?:20)?\d{2})\b/g),match=>match[1]));
  const explicitAux=unique([...supplierTokens,...seasonTokens,...[...auxiliaryWords].filter(word=>new RegExp(`(?:^|\\s)${word}(?:$|\\s)`).test(normalized))]);
  const codeParts=new Set(codes.flatMap(code=>code.replace('#','').split(/\s+/)));
  const coreTokens=unique(normalized.replace(/[(),/+]/g,' ').split(/\s+/).filter(token=>token&&token!=='AND'&&!auxiliaryWords.has(token)&&!explicitAux.includes(token)&&!seasonTokens.includes(token)&&!codeParts.has(token)&&!/^\d{2}$/.test(token)));
  return {normalized,coreName:coreTokens.join(' '),coreTokens,auxiliaryTokens:explicitAux,codes,specs,styles,composition,positions:foundPositions,demographics:foundDemographics,productTypes:foundProducts};
}
export const normalizeMaterialName=(value:string)=>extractMaterialAttributes(value).normalized;
export const normalizeMaterial=(value:string,removeSeason=false)=>{const normalized=normalizeMaterialName(value);return removeSeason?normalized.replace(/\bFLY RACING (?:20)?\d{2}\b/g,'FLY RACING').replace(/\s+/g,' ').trim():normalized};
export const itemCodes=(value:string)=>extractMaterialAttributes(value).codes;
export const qualifiers=(value:string)=>{const a=extractMaterialAttributes(value);return a.demographics.concat(a.auxiliaryTokens)};
