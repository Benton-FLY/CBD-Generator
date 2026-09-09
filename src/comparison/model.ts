import {resolveGroup} from './groups';
import type {ComparisonState,StyleMatch} from './types';

// Both Excel renderers consume these same persisted, confirmed relationships.
// This layer does not match, reprice, or mutate source materials.
export const COST_GROUPS=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'];
export function comparisonModel(state:ComparisonState,pair:StyleMatch){
 const reference=state.styles.find(s=>s.id===pair.referenceId);
 const comparison=state.styles.find(s=>s.id===pair.currentId);
 const clusters=(state.materialMatches.find(s=>s.styleMatchId===pair.id)?.clusters||[])
  .map(c=>resolveGroup(c,reference?.materials||[],comparison?.materials||[]));
 return {reference,comparison,clusters};
}
