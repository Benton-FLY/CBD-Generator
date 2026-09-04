export type ComparisonSide='reference'|'current';
export type MatchStatus='Exact'|'Normalized'|'Alias'|'Manual'|'reference-only'|'current-only'|'duplicate'|'review';
export interface CbdMaterialRow {id:string;group:string;material:string;size:string;unit:string;cost?:number;usage?:number;loss?:number;extended?:number;remark:string;width:string;order:number}
export type FobEvidence='explicit'|'ratio-derived'|'review-required';
export interface CbdSummary {totalMaterialCost?:number;laborCost?:number;overhead?:number;profit?:number;finalFob?:number;materialToFobRatio?:number;erpMaterial?:number;difference?:number;differenceRate?:number;fobEvidence?:FobEvidence;fobValidation?:'matched'|'mismatch'|'not-available';calculatedFob?:number}
export interface CbdStyle {id:string;side:ComparisonSide;fileName:string;sheetName:string;styleName:string;materials:CbdMaterialRow[];summary:CbdSummary;groupOrder:string[];groupTotals?:Record<string,number>}
export interface UploadedCbd {id:string;name:string;size:number;lastModified:number;side:ComparisonSide;styleIds:string[]}
export interface StyleMatch {id:string;referenceId?:string;currentId?:string;method:string;confidence:number;status:MatchStatus;excluded?:boolean}
export type MaterialStatus='MATCH'|'NAME CHANGED'|'GROUP CHANGED'|'CURRENT ONLY'|'REFERENCE ONLY'|'REVIEW'|'MANUAL'|'MERGED N:1';
export interface MaterialMatch {id:string;referenceId?:string;currentId?:string;finalGroup:string;referenceGroup?:string;currentGroup?:string;method:string;confidence:number;status:MaterialStatus}
export type MaterialMatchCluster={id:string;referenceRowIds:string[];currentRowId:string|null;relationType:'one-to-one'|'many-to-one';matchSource:'auto'|'manual';finalGroup:string;status:MaterialStatus;confidence:number};
export interface MaterialMatchSet {styleMatchId:string;clusters:MaterialMatchCluster[];matches?:MaterialMatch[]}
export interface ComparisonState {version:number;referenceSeason:string;currentSeason:string;files:UploadedCbd[];styles:CbdStyle[];styleMatches:StyleMatch[];materialMatches:MaterialMatchSet[];step:1|2|3|4|5;activeMatchId?:string;savedAt?:string;history?:{past:ComparisonState[];future:ComparisonState[]}}
export interface ComparisonFinancials {referenceFob?:number;comparisonFob?:number;referenceMaterialTotal?:number;comparisonMaterialTotal?:number;referenceMaterialToFobRatio?:number;comparisonMaterialToFobRatio?:number;referenceFobEvidence:FobEvidence;comparisonFobEvidence:FobEvidence}
