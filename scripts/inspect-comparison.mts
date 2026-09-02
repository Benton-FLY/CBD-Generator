import {readFile} from 'node:fs/promises';
import {parseCbdFiles} from '../src/comparison/parser';
import {matchStyles,matchMaterials} from '../src/comparison/matcher';
const paths=['reference/FLY RACING 27 MX JERSEY CBD 2025.10.17.xlsx','reference/FLY RACING 27 BMX JERSEY CBD 2025.10.17.xlsx','reference/FLY RACING 28 MX BMX JERSEY CBD SHEET (내부) 2026.09.02.xlsx'];
const files=await Promise.all(paths.map(async p=>new File([await readFile(p)],p.split('/').at(-1)!)));
const ref=await parseCbdFiles(files.slice(0,2),'reference'),cur=await parseCbdFiles([files[2]],'current');
console.log('reference',ref.styles.length,'files',ref.files.length,'errors',ref.errors);console.log('current',cur.styles.length,'files',cur.files.length,'errors',cur.errors);
const styles=[...ref.styles,...cur.styles],matches=matchStyles(styles);console.log('style matches',matches.length,'review',matches.filter(m=>m.status==='review'||m.status==='duplicate').length);for(const m of matches.filter(m=>styles.find(s=>s.id===m.referenceId)?.styleName?.includes('RAYCE')||styles.find(s=>s.id===m.currentId)?.styleName?.includes('RAYCE')))console.log(m,styles.find(s=>s.id===m.referenceId)?.styleName,styles.find(s=>s.id===m.currentId)?.styleName);
const rayceRef=ref.styles.find(s=>s.styleName==='27 RAYCE JERSEY')!,rayceCur=cur.styles.find(s=>s.styleName==='RAYCE JERSEY')!;
const base=rayceRef.materials.filter(m=>m.material.includes('JX-306 (POLY)')&&m.material.startsWith('PRINTED FLY')),merged=rayceCur.materials.find(m=>m.material.includes('PRINTED SOLID')&&m.material.includes('JX-306 (POLY) POLYGIENE'))!;
if(base.length!==2||!merged)throw new Error('RAYCE N:1 회귀 자재를 찾지 못했습니다.');
if(!base.every(m=>m.id.startsWith('reference::')&&m.id.includes('::row-')&&m.id.includes('::seq-'))||!merged.id.startsWith('current::'))throw new Error('stable row ID 형식 오류');
console.log('RAYCE N:1 fixture',base.map(m=>[m.cost,m.usage,m.extended]),[merged.cost,merged.usage,merged.extended]);
for(const sm of matches.filter(m=>m.referenceId&&m.currentId).slice(0,5)){const set=matchMaterials(sm,styles);console.log('materials',styles.find(s=>s.id===sm.referenceId)?.styleName,set.clusters.length,'review',set.clusters.filter(x=>x.status==='REVIEW').length)}
const printed=styles.flatMap(s=>s.materials.filter(m=>/^PRINTED/i.test(m.material))).slice(0,30);console.log('printed groups',printed.map(m=>[m.material,m.group]));
const excludedWords=/PRINTED LABEL|PRINTED STICKER|PRINTED CARTON|PRINT COST|SILICON PRINT COST|SUBLIMATION PRINT COST|HEAT TRANSFER|PRINT MATERIAL|PACKING|PATCH|BADGE/i;console.log('printed exclusions incorrectly OUTSHELL',styles.flatMap(s=>s.materials.filter(m=>excludedWords.test(m.material)&&m.group==='OUTSHELL')).length);
