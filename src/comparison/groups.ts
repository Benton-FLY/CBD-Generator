import type {CbdMaterialRow,MaterialMatchCluster} from './types';
export const normalizeGroup=(value:string)=>value.trim().toUpperCase().replace(/^TRIM$/,'TRIMS');
export function resolveGroup(cluster:MaterialMatchCluster,reference:CbdMaterialRow[],comparison:CbdMaterialRow[]):MaterialMatchCluster {
 const refs=reference.filter(r=>cluster.referenceRowIds.includes(r.id)),cur=comparison.find(r=>r.id===cluster.currentRowId);
 const referenceOriginalGroup=[...new Set(refs.map(r=>normalizeGroup(r.group)))].join(' + '),comparisonOriginalGroup=cur?normalizeGroup(cur.group):'';
 const manualGroupOverride=cluster.manualGroupOverride??((!!cluster.manualLocked||cluster.matchSource==='manual')&&normalizeGroup(cluster.finalGroup)!==(comparisonOriginalGroup||referenceOriginalGroup));
 const effectiveGroup=normalizeGroup(manualGroupOverride?cluster.finalGroup:comparisonOriginalGroup||referenceOriginalGroup||cluster.finalGroup);
 return {...cluster,referenceOriginalGroup,comparisonOriginalGroup,effectiveGroup,finalGroup:effectiveGroup,manualGroupOverride,groupAssignmentSource:manualGroupOverride?'manual':cur?'comparison':'referenceFallback'};
}
export const unitChanged=(refs:CbdMaterialRow[],cur?:CbdMaterialRow)=>!!cur&&refs.some(r=>r.unit.trim().toUpperCase()!==cur.unit.trim().toUpperCase());
export function relationPresentation(cluster:MaterialMatchCluster,refs:CbdMaterialRow[],cur?:CbdMaterialRow,ko=false){
 const c=resolveGroup(cluster,refs,cur?[cur]:[]),from=c.referenceOriginalGroup||c.comparisonOriginalGroup,to=c.effectiveGroup;
 const changed=from!==to||!!c.comparisonOriginalGroup&&c.comparisonOriginalGroup!==to;
 const status=[!refs.length?(ko?'비교 시즌만 존재':'Comparison Only'):!cur?(ko?'기준 시즌만 존재':'Reference Only'):c.status==='REVIEW'?'Needs Review':c.relationType==='many-to-one'?(ko?'다대일 연결':'Merged N:1'):c.matchSource==='manual'?(ko?'수동 연결':'Manual Match'):c.evidence?.method==='Context Match'?'Context Match':'Auto Matched'];
 if(unitChanged(refs,cur))status.push('Unit Changed');
 if(changed)status.push(`${ko?(c.manualGroupOverride?'수동 그룹 변경':from?.includes(' + ')?'그룹 통합':'그룹 변경'):(c.manualGroupOverride?'Manual Group Changed':'Group Changed')} (${from===to?c.comparisonOriginalGroup:from} → ${to})`);
 const notes=[ko?`기준 원본: ${c.referenceOriginalGroup||'—'} / 비교 원본: ${c.comparisonOriginalGroup||'—'} / 최종 비교 그룹: ${to} / ${c.manualGroupOverride?'수동 지정':'자동 지정'}`:`Reference original: ${c.referenceOriginalGroup||'—'} / Comparison original: ${c.comparisonOriginalGroup||'—'} / Effective group: ${to} / ${c.manualGroupOverride?'Manual':'Automatic'}`];
 if(unitChanged(refs,cur))notes.push(`Unit changed: ${[...new Set(refs.map(r=>r.unit))].join(' + ')} → ${cur!.unit}`);
 if(c.evidence?.reason)notes.push(c.evidence.reason);
 if(cur&&refs.some(r=>r.material!==cur.material))notes.push(ko?'자재명/부가 설명 변경':'Material name / auxiliary description changed');
 notes.push(...refs.map(r=>r.remark),cur?.remark||'');return {status:status.join(' · '),notes:notes.filter(Boolean).join('\n')};
}
export function groupAmounts(clusters:MaterialMatchCluster[],reference:CbdMaterialRow[],comparison:CbdMaterialRow[]){
 const totals:Record<string,{reference:number;comparison:number}>={};
 for(const raw of clusters){const c=resolveGroup(raw,reference,comparison),t=totals[c.effectiveGroup!]??={reference:0,comparison:0};for(const id of c.referenceRowIds)t.reference+=reference.find(r=>r.id===id)?.extended??0;t.comparison+=comparison.find(r=>r.id===c.currentRowId)?.extended??0;}return totals;
}
