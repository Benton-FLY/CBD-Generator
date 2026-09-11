import type {CbdMaterialRow,ManagerAppliedMaterial,MaterialMatchCluster} from './types';
import {comparisonRowIds} from './relations';

export const round4=(value:number|undefined)=>value===undefined?undefined:Math.round((value+Number.EPSILON)*10000)/10000;
const unit=(value:string)=>value.trim().toUpperCase().replace(/\s+/g,'');
export const managerUnitChanged=(oldRows:CbdMaterialRow[],newRow?:CbdMaterialRow)=>!!newRow&&oldRows.some(row=>unit(row.unit)!==unit(newRow.unit));

/** Manager-only decision. Source rows are never changed and comparison exports do not call this. */
export function applyManagerFloors(oldRows:CbdMaterialRow[],newRows:CbdMaterialRow[],cluster:MaterialMatchCluster):ManagerAppliedMaterial {
 const current=newRows.find(row=>comparisonRowIds(cluster).includes(row.id));
 const old=oldRows[0];
 if(!current)return {sourceOldUnitCost:old?.cost,sourceOldUsage:old?.usage,unitCostFloorApplied:false,usageFloorApplied:false,status:'reference-only',reason:'신규 대체 자재 또는 삭제 여부 확인 필요'};
 if(!old)return {sourceNewUnitCost:current.cost,sourceNewUsage:current.usage,appliedNewUnitCost:current.cost,appliedNewUsage:current.usage,unitCostFloorApplied:false,usageFloorApplied:false,status:cluster.status==='REVIEW'?'needs-review':'current-only',reason:cluster.status==='REVIEW'?'Needs Review':''};
 if(managerUnitChanged(oldRows,current))return {sourceOldUnitCost:old.cost,sourceOldUsage:old.usage,sourceNewUnitCost:current.cost,sourceNewUsage:current.usage,appliedNewUnitCost:current.cost,appliedNewUsage:current.usage,unitCostFloorApplied:false,usageFloorApplied:false,status:'needs-review',reason:`Needs Review – Unit Changed: ${oldRows.map(row=>row.unit).join(' + ')} → ${current.unit}`};
 // A split/merge has no safe per-row allocation. A manually locked relation is an explicit anchor;
 // automatic N:1/1:N relations remain review-only to prevent repeated floors and amounts.
 const isOneToOne=oldRows.length===1&&comparisonRowIds(cluster).length===1;
 if(!isOneToOne&&!cluster.manualLocked&&cluster.matchSource!=='manual')return {sourceOldUnitCost:old.cost,sourceOldUsage:old.usage,sourceNewUnitCost:current.cost,sourceNewUsage:current.usage,appliedNewUnitCost:current.cost,appliedNewUsage:current.usage,unitCostFloorApplied:false,usageFloorApplied:false,status:'needs-review',reason:'Needs Review – N:1/1:N 배분 근거 필요'};
 const oldCost=round4(old.cost),newCost=round4(current.cost),oldUsage=round4(old.usage),newUsage=round4(current.usage);
 const costHold=oldCost!==undefined&&newCost!==undefined&&newCost<oldCost;
 const usageHold=oldUsage!==undefined&&newUsage!==undefined&&newUsage<oldUsage;
 return {sourceOldUnitCost:old.cost,sourceOldUsage:old.usage,sourceNewUnitCost:current.cost,sourceNewUsage:current.usage,appliedNewUnitCost:costHold?old.cost:current.cost,appliedNewUsage:usageHold?old.usage:current.usage,unitCostFloorApplied:costHold,usageFloorApplied:usageHold,status:'matched'};
}
