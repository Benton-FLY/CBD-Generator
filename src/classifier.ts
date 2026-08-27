import type { BomRow, CbdGroup, ClassificationDictionary } from './types';

export const normalizeText = (v: unknown) => String(v ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
export const dictionaryKeys = (r: BomRow) => [r.itemNo && `NO:${normalizeText(r.itemNo)}`, `ITEM:${normalizeText(r.item)}`].filter(Boolean);

export function classify(r: BomRow, dictionary: ClassificationDictionary = {}): CbdGroup {
  for (const key of dictionaryKeys(r)) if (dictionary[key]) return dictionary[key];
  const item = normalizeText(r.item), structure = normalizeText(r.structure), type = normalizeText(r.materialType);
  const seq = Number(String(r.sequence).replace(/[^0-9.]/g, ''));
  if (type.includes('봉제공임') || type.includes('포장공임') || /SEWING COST|PACKING COST/.test(item)) return 'EXCLUDE';
  if (structure.startsWith('포장-') || (seq >= 100 && seq <= 149) || /LABEL|HANG\s*TAG|STICKER|BARCODE|POLY\s*BAG|DRY\s*SAC|SILICA\s*GEL/.test(item)) return 'LABEL & PACKAGING';
  if (structure.startsWith('재봉사') || (seq >= 150 && seq <= 199) || /THREAD/.test(item)) return 'SEWING THREAD';
  if (type.includes('특수공임') || type.includes('기타공임') || /HEAT TRANSFER PRESS|SILICON PRINT COST|LASER CUTTING COST/.test(item)) return 'SPECIAL PROCESS (LIST ONLY)';
  if (/FABRIC|LEATHER|KEVLAR|MESH|POLY \(|NYLON \(|OUTSHELL|SHELL/.test(item) && !/WEBBING|ZIPPER|TAPE|PATCH|PRINT MATERIAL/.test(item)) return 'OUTSHELL';
  if (item) return 'TRIMS';
  return 'NEEDS REVIEW';
}
