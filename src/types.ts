export const GROUPS = ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)','NEEDS REVIEW','EXCLUDE'] as const;
export type CbdGroup = typeof GROUPS[number];

export interface BomRow {
  id: string; sourceFile: string; sourceSheet: string; sourceRow: number;
  structure: string; materialType: string; sequence: string; serialNo?:string; parentSerialNo?:string; level?:string; hasChildren?:string; itemNo: string; item: string;
  width: string; color: string; unit: string; netUsage?: number; bomLoss?: number; usage: number;
  currency: string; rawPrice?: number; convertedPrice: number; ancillaryCost?:number; materialCostAdjustment: number; amount?: number; specialFlag: string; remark: string;
}
export interface ErpMaterialAuditRow { source:BomRow; isParent:boolean; included:boolean; includedAmount:number; usedFallback:boolean; reason:string }
export type RowDisposition = 'review'|'separate'|'excluded';
export interface RowStatusDetail {
  id:string; disposition:RowDisposition; sourceRow:number; itemNo:string; item:string; structure:string;
  materialType:string; unit:string; convertedPrice:number; materialCostAdjustment:number; remark:string;
  result:string; reason:string;
}
export interface Material {
  id: string; item: string; width: string; unit: string; group: CbdGroup; included: boolean;
  baseCost: number; adjustedCost: number; baseUsage: number; adjustedUsage: number; baseLoss?: number; additionalLoss: number;
  remark: string; originalRemark?:string; remarkEdited?:boolean; sources: BomRow[]; split: boolean;
}
export interface StyleData {
  id: string; name: string; sourceFile: string; sourceSheet: string; materials: Material[];
  finalFob?: number; fobMatch?: string; fobMatchMethod?:'Exact'|'Normalized'|'Alias'|'Manual'; fobUnmatchedReason?:string; threadSubtotal?: number;
  statusDetails?:RowStatusDetail[];
  erpMaterialCost?:number; erpAudit?:ErpMaterialAuditRow[];
  laborRemark: string;
}
export interface AppSettings { exchangeRate: number; defaultLoss: number; }
export type ClassificationDictionary = Record<string, CbdGroup>;
export interface ColumnMapping { [field: string]: number | undefined }

export const visibleGroups: CbdGroup[] = ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'];

/** Match Excel ROUND(value, 4) before any material subtotal is calculated. */
export const roundTo4 = (value: number) => {
  if (!Number.isFinite(value)) return 0;
  const absolute=Math.abs(value);
  return Math.sign(value)*Math.round((absolute+Number.EPSILON*Math.max(1,absolute))*10000)/10000;
};

export const extendedCost = (m: Material) => roundTo4(m.adjustedCost * m.adjustedUsage * (1 + m.additionalLoss));
export const groupSubtotal = (materials: Material[]) => roundTo4(materials.reduce((sum,m)=>sum+extendedCost(m),0));
export const totalMaterialCost = (materials: Material[]) => roundTo4(visibleGroups.reduce((total,group)=>
  total+groupSubtotal(materials.filter(m=>m.included&&m.group===group)),0));
