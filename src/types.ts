export const GROUPS = ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)','NEEDS REVIEW','EXCLUDE'] as const;
export type CbdGroup = typeof GROUPS[number];

export interface BomRow {
  id: string; sourceFile: string; sourceSheet: string; sourceRow: number;
  structure: string; materialType: string; sequence: string; itemNo: string; item: string;
  width: string; color: string; unit: string; netUsage?: number; bomLoss?: number; usage: number;
  currency: string; rawPrice?: number; convertedPrice: number; materialCostAdjustment: number; amount?: number; specialFlag: string; remark: string;
}
export interface Material {
  id: string; item: string; width: string; unit: string; group: CbdGroup; included: boolean;
  baseCost: number; adjustedCost: number; baseUsage: number; adjustedUsage: number; additionalLoss: number;
  remark: string; sources: BomRow[]; split: boolean;
}
export interface StyleData {
  id: string; name: string; sourceFile: string; sourceSheet: string; materials: Material[];
  finalFob?: number; fobMatch?: string; threadSubtotal?: number;
  laborRemark: string;
}
export interface AppSettings { exchangeRate: number; defaultLoss: number; }
export type ClassificationDictionary = Record<string, CbdGroup>;
export interface ColumnMapping { [field: string]: number | undefined }

export const extendedCost = (m: Material) => m.adjustedCost * m.adjustedUsage * (1 + m.additionalLoss);
export const visibleGroups: CbdGroup[] = ['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING'];
