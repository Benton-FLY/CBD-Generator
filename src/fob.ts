import * as XLSX from 'xlsx';
import { normalizeText } from './classifier';

const STYLE_HEADERS=['STYLE','STYLE NAME','MODEL','MODEL NAME','스타일','품명'];
const FOB_HEADERS=['FOB','FOB PRICE','FINAL FOB','최종 FOB','PRICE'];
const norm=(s:unknown)=>normalizeText(s).replace(/\bBOM\b/g,'').replace(/[-_]+/g,' ').replace(/\s+/g,' ').trim();
export interface FobRecord { style:string; fob:number }
export async function parseFobFile(file:File):Promise<FobRecord[]> { const wb=XLSX.read(await file.arrayBuffer(),{type:'array'}); const out:FobRecord[]=[];
  for(const name of wb.SheetNames){const rows=XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name],{header:1,defval:''});let hi=-1,si=-1,fi=-1;
    for(let r=0;r<Math.min(30,rows.length);r++){const vals=rows[r].map(norm);const s=vals.findIndex(v=>STYLE_HEADERS.map(norm).includes(v));const f=vals.findIndex(v=>FOB_HEADERS.map(norm).includes(v));if(s>=0&&f>=0){hi=r;si=s;fi=f;break;}}
    if(hi>=0) rows.slice(hi+1).forEach(row=>{const style=String(row[si]??'').trim(),fob=Number(String(row[fi]??'').replace(/[$,\s]/g,''));if(style&&Number.isFinite(fob))out.push({style,fob});});
  } return out; }
export const exactFobMatch=(style:string,records:FobRecord[])=>records.find(r=>norm(r.style)===norm(style));
export const fobCandidates=(style:string,records:FobRecord[])=>records.filter(r=>{const a=norm(style),b=norm(r.style);return a.includes(b)||b.includes(a)||a.split(' ').filter(x=>b.includes(x)).length>=2;}).slice(0,5);
