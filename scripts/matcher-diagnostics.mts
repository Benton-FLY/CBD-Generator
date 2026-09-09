import {writeFile} from 'node:fs/promises';
import {extractMaterialAttributes} from '../src/comparison/normalize';
import {scoreMaterialCandidate,assignMaterialMatches} from '../src/comparison/matcher';
const names=[['HANG TAG FLY RACING 26 RAYCE PANT','HANG TAG FLY RACING 25 RAYCE JERSEY','PCS','PCS'],['FLY RACING MAIN BUCKLE','FLY RACING MAIN BUCKLE SET','PCS','SET'],['NYLON WEBBING (30M/M)','NEW - NYLON WEBBING (30M/M) P','YD','M']];
const rows=names.map(([a,b,u,v],i)=>[a,b].map((material,j)=>({id:`${i}-${j}`,material,unit:j?v:u,group:'TRIMS',size:'',width:'',remark:'',order:i})));
const report=rows.map(([r,c])=>{const candidates=rows.map(x=>({name:x[1].material,...scoreMaterialCandidate(r,x[1])})).sort((a,b)=>b.score-a.score);return{reference:r,current:c,referenceAttributes:extractMaterialAttributes(r.material),comparisonAttributes:extractMaterialAttributes(c.material),candidates,margin:candidates[0].score-candidates[1].score}});
await writeFile(process.argv[2]||'reports/matcher-after.json',JSON.stringify({scope:'Provided representative names; original workbooks unavailable',report,clusters:assignMaterialMatches(rows.map(x=>x[0]),rows.map(x=>x[1]))},null,2));
