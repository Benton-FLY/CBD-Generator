import fs from 'node:fs/promises';
import {File} from 'node:buffer';
import * as XLSX from 'xlsx';
import {parseCbdFiles} from '../src/comparison/parser';
import {resolveGroup} from '../src/comparison/groups';
import type {ComparisonState,MaterialMatchCluster} from '../src/comparison/types';

export async function managerFixture(){
 const files=await fs.readdir('reference'),path=`reference/${files.find(f=>f.endsWith('.xlsx'))}`;
 const source=XLSX.read(await fs.readFile(path),{type:'buffer'}),state:ComparisonState={version:2,referenceSeason:'27',currentSeason:'28',files:[],styles:[],styleMatches:[],materialMatches:[],step:3};
 for(const [side,offset] of [['current',0],['reference',16]] as const){
  const book=XLSX.utils.book_new();
  for(const name of source.SheetNames){
   const input=source.Sheets[name],output:XLSX.WorkSheet={};
   const range=XLSX.utils.decode_range(input['!ref']!);
   for(let r=0;r<=Math.min(range.e.r,100);r++)for(let c=0;c<9;c++){
    const original=input[XLSX.utils.encode_cell({r,c:c+offset})];if(original)output[XLSX.utils.encode_cell({r,c})]={t:original.t,v:typeof original.v==='string'?original.v.replace(/_x000d_/gi,'\r'):original.v};
   }
   output['!merges']=(input['!merges']||[]).filter(m=>m.s.c>=offset&&m.e.c<offset+9).map(m=>({s:{r:m.s.r,c:m.s.c-offset},e:{r:m.e.r,c:m.e.c-offset}}));
   output['!ref']='A1:I100';XLSX.utils.book_append_sheet(book,output,name);
  }
  const data=XLSX.write(book,{type:'buffer',bookType:'xlsx'});
  await fs.writeFile(`reports/manager-input-${side}.xlsx`,data);
  const parsed=await parseCbdFiles([new File([data],`${side}.xlsx`) as unknown as globalThis.File],side);
  if(parsed.errors.length)throw new Error(parsed.errors.join('\n'));
  state.styles.push(...parsed.styles);state.files.push(...parsed.files);
 }
 for(const name of source.SheetNames){
  const reference=state.styles.find(s=>s.side==='reference'&&s.sheetName===name)!,current=state.styles.find(s=>s.side==='current'&&s.sheetName===name)!;
  const pair={id:name,referenceId:reference.id,currentId:current.id,method:'Manual',confidence:1,status:'Manual' as const};state.styleMatches.push(pair);
  // The reference workbook itself supplies reviewer-confirmed row relationships.
  // This fixture records those explicit relationships; production never matches strings here.
  const rowNumber=(id:string)=>Number(id.match(/::row-(\d+)::/)![1]);
  const rows=[...new Set([...reference.materials,...current.materials].map(r=>rowNumber(r.id)))].sort((a,b)=>a-b);
  const clusters:MaterialMatchCluster[]=rows.map(r=>{
   const ref=reference.materials.find(m=>rowNumber(m.id)===r),cur=current.materials.find(m=>rowNumber(m.id)===r);
   return resolveGroup({id:`${name}:${r}`,referenceRowIds:ref?[ref.id]:[],currentRowId:cur?.id??null,relationType:'one-to-one',matchSource:'manual',manualLocked:true,finalGroup:cur?.group||ref!.group,status:!ref?'CURRENT ONLY':!cur?'REFERENCE ONLY':'MANUAL',confidence:1},reference.materials,current.materials);
  });
  state.materialMatches.push({styleMatchId:name,clusters,matcherVersion:4});
 }
 state.activeMatchId=state.styleMatches[0].id;
 return state;
}
