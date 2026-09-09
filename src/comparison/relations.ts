import type {CbdMaterialRow,MaterialMatchCluster} from './types';
/** currentRowId is the anchor; currentRowIds, when present, is the complete split relation. */
export const comparisonRowIds=(cluster:MaterialMatchCluster)=>cluster.currentRowIds??(cluster.currentRowId?[cluster.currentRowId]:[]);
export const comparisonRows=(cluster:MaterialMatchCluster,rows:CbdMaterialRow[])=>comparisonRowIds(cluster).map(id=>rows.find(r=>r.id===id)).filter((r):r is CbdMaterialRow=>!!r);
