import { readTabularWorkbook } from './spreadsheet';

const STYLE_HEADERS=['STYLE#','STYLE','STYLE NO','STYLE NO.','STYLE NUMBER','MODEL','MODEL NAME'];
const FOB_HEADERS=['FINAL FOB','FOB','FOB PRICE','FINAL PRICE','UNIT FOB'];
const headerKey=(v:unknown)=>String(v??'').normalize('NFKC').replace(/\u00a0/g,' ').toUpperCase().replace(/[\s#._-]+/g,'');
export const normalizeFobStyle=(value:unknown,removeSeason=false)=>{
  let out=String(value??'').normalize('NFKC').replace(/[–—−]/g,'-').replace(/\.(xlsx?|xls)$/i,'').replace(/\s*사전원가\s*$/,'').replace(/\bBOM\b/ig,'').replace(/\s+/g,' ').trim().toUpperCase();
  if(removeSeason)out=out.replace(/^\d{2}\s+/,'');
  return out;
};
const canonicalAlias=(value:string)=>normalizeFobStyle(value,true).replace(/\s*-\s*/g,' ').replace(/SONIC FLY LOGO|소닉/g,'SONIC').replace(/SUBLIMATED FLY LOGO|DIGITAL SUBLIMATION|디지전사/g,'SUBLIMATED').replace(/\s+/g,' ').trim();
export interface FobRecord { style:string; fob:number; sheet?:string; row?:number }
export interface FobMatch { record?:FobRecord; method?:'Exact'|'Normalized'|'Alias'; reason?:string; candidates:FobRecord[] }
export async function parseFobFile(file:File):Promise<FobRecord[]> { const sheets=await readTabularWorkbook(file); const out:FobRecord[]=[];
  for(const {name,rows} of sheets){let hi=-1,si=-1,fi=-1;
    for(let r=0;r<Math.min(30,rows.length);r++){const vals=rows[r].map(headerKey);const s=vals.findIndex(v=>STYLE_HEADERS.map(headerKey).includes(v));const f=vals.findIndex(v=>FOB_HEADERS.map(headerKey).includes(v));if(s>=0&&f>=0){hi=r;si=s;fi=f;break;}}
    if(hi>=0)rows.slice(hi+1).forEach((row,index)=>{const raw=String(row[si]??'').trim(),fob=Number(String(row[fi]??'').replace(/[$,\s]/g,''));if(raw&&Number.isFinite(fob))raw.split(/\r?\n/).map(v=>v.trim()).filter(Boolean).forEach(style=>out.push({style,fob,sheet:name,row:hi+index+2}));});
  }if(!out.length)throw new Error('상단 30행에서 STYLE과 FOB 헤더 및 가격 데이터를 찾지 못했습니다.');return out; }
export function matchFob(style:string,records:FobRecord[]):FobMatch{
  const exact=records.filter(r=>normalizeFobStyle(r.style)===normalizeFobStyle(style));if(exact.length===1)return{record:exact[0],method:'Exact',candidates:exact};
  const normalized=records.filter(r=>normalizeFobStyle(r.style,true)===normalizeFobStyle(style,true));if(normalized.length===1)return{record:normalized[0],method:'Normalized',candidates:normalized};
  const aliases=records.filter(r=>canonicalAlias(r.style)===canonicalAlias(style));if(aliases.length===1)return{record:aliases[0],method:'Alias',candidates:aliases};
  const query=new Set(canonicalAlias(style).split(' '));const scored=records.map(record=>{const candidate=new Set(canonicalAlias(record.style).split(' '));const common=[...query].filter(x=>candidate.has(x)).length;return{record,score:common/Math.max(query.size,candidate.size)}}).filter(x=>x.score>=.8).sort((a,b)=>b.score-a.score);
  if(scored.length===1)return{record:scored[0].record,method:'Normalized',candidates:scored.map(x=>x.record)};
  const candidates=scored.map(x=>x.record);return{candidates,reason:candidates.length>1?'유사 후보가 2개 이상이어서 수동 선택이 필요합니다.':'일치하는 STYLE alias가 없습니다.'};
}
export const exactFobMatch=(style:string,records:FobRecord[])=>matchFob(style,records).record;
export const fobCandidates=(style:string,records:FobRecord[])=>matchFob(style,records).candidates.slice(0,5);
