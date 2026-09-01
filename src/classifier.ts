import type { BomRow, CbdGroup, ClassificationDictionary } from './types';

export const normalizeText = (v: unknown) => String(v ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
export const dictionaryKeys = (r: BomRow) => [r.itemNo && `NO:${normalizeText(r.itemNo)}`, `ITEM:${normalizeText(r.item)}`].filter(Boolean);

export function classify(r: BomRow, dictionary: ClassificationDictionary = {}): CbdGroup {
  const item = normalizeText(r.item), structure = normalizeText(r.structure), type = normalizeText(r.materialType);
  const unit = normalizeText(r.unit), itemNo=normalizeText(r.itemNo);
  for (const key of dictionaryKeys(r)) if (dictionary[key]) return dictionary[key];
  if (type.includes('금형및철형') || /철형|금형|C-COST/.test(structure) || itemNo==='00015627' || /CUTTING TOOL|CUTTING DIE|MOLD COST/.test(item)) return 'SPECIAL PROCESS (LIST ONLY)';
  // Specific base-material names take priority over broader PRINT/TRIMS rules.
  if (item.includes('CHIRON') || item.startsWith('DIGITAL')) return 'OUTSHELL';
  if (type.includes('봉제공임') || type.includes('포장공임') || /SEWING COST|PACKING COST/.test(item)) return 'EXCLUDE';
  if (type.includes('특수공임') || type.includes('기타공임') || /HEAT TRANSFER PRESS|SILICON PRINT COST|LASER CUTTING COST/.test(item)) return 'SPECIAL PROCESS (LIST ONLY)';
  if (structure.startsWith('포장-') || /LABEL|HANG\s*TAG|STICKER|BARCODE|POLY\s*BAG|DRY\s*SAC|SILICA\s*GEL/.test(item)) return 'LABEL & PACKAGING';
  if (structure.includes('재봉사') || structure.includes('THREAD') || /\bSPUN\b|NYLON\s+THREAD|KEVLAR\s+THREAD|\bSW\s*[*-]?\s*TH\b/.test(item) || (unit==='CONE' && /THREAD|SPUN|NYLON|KEVLAR/.test(item))) return 'SEWING THREAD';
  if (/FABRIC|LEATHER|KEVLAR|MESH|POLY \(|NYLON \(|OUTSHELL|SHELL/.test(item) && !/WEBBING|ZIPPER|TAPE|PATCH|PRINT MATERIAL/.test(item)) return 'OUTSHELL';
  if (item) return 'TRIMS';
  return 'NEEDS REVIEW';
}
