import type { BomRow, CbdGroup, ClassificationDictionary, StyleData } from './types';

export const CLASSIFICATION_VERSION=2;
export const normalizeText = (v: unknown) => String(v ?? '').normalize('NFKC').replace(/[\u00a0\r\n]+/g,' ').replace(/[‐‑‒–—―−]/g,'-').trim().replace(/\s+/g, ' ').toUpperCase();
const matchText=(v:unknown)=>normalizeText(v).replace(/[-_/+.,:;()[\]{}]+/g,' ').replace(/\s+/g,' ').trim();
const PRINTED_OUTSHELL_PATTERNS=['PRINTED SOLID FLY RACING','SOLID PRINTED FLY RACING','PRINTED FLY RACING','PRINTED SOLID','PRINTED'];
const PRINTED_NON_FABRIC=/^(?:PRINTED\s+)?(?:LABEL|STICKER|CARTON|PACKING|PATCH|BADGE)\b|\b(?:PRINT(?:ED)?\s+MATERIAL|PRINT\s+COST|SILICON\s+PRINT\s+COST|SUBLIMATION\s+PRINT\s+COST|HEAT\s+TRANSFER|PACKING|PATCH|BADGE)\b/;
export const isPrintedOutshell=(r:Pick<BomRow,'item'|'structure'|'materialType'|'unit'|'width'>)=>{
  const item=matchText(r.item),structure=matchText(r.structure),type=matchText(r.materialType),unit=matchText(r.unit),width=matchText(r.width);
  if(!PRINTED_OUTSHELL_PATTERNS.some(pattern=>item.startsWith(pattern))||PRINTED_NON_FABRIC.test(item))return false;
  if(/공임|COST|PACK|LABEL|STICKER|CARTON/.test(type)||/포장|PACK|LABEL/.test(structure))return false;
  return /^(YD|M)$/.test(unit)||Boolean(width)||PRINTED_OUTSHELL_PATTERNS.some(pattern=>item.startsWith(pattern));
};
export const dictionaryKeys = (r: BomRow) => [r.itemNo && `NO:${normalizeText(r.itemNo)}`, `ITEM:${normalizeText(r.item)}`].filter(Boolean);

export function classify(r: BomRow, dictionary: ClassificationDictionary = {}): CbdGroup {
  const item = normalizeText(r.item), structure = normalizeText(r.structure), type = normalizeText(r.materialType);
  const unit = normalizeText(r.unit), itemNo=normalizeText(r.itemNo);
  for (const key of dictionaryKeys(r)) if (dictionary[key]) return dictionary[key];
  // Manual > special process > packaging > thread > printed fabric > other shell > trims > fallback.
  if (type.includes('금형및철형') || /철형|금형|C-COST/.test(structure) || itemNo==='00015627' || /CUTTING TOOL|CUTTING DIE|MOLD COST/.test(item)) return 'SPECIAL PROCESS (LIST ONLY)';
  if (type.includes('봉제공임') || type.includes('포장공임') || /SEWING COST|PACKING COST/.test(item)) return 'EXCLUDE';
  if (type.includes('특수공임') || type.includes('기타공임') || /HEAT TRANSFER PRESS|SILICON PRINT COST|LASER CUTTING COST/.test(item)) return 'SPECIAL PROCESS (LIST ONLY)';
  if (structure.startsWith('포장-') || /LABEL|HANG\s*TAG|STICKER|BARCODE|POLY\s*BAG|DRY\s*SAC|SILICA\s*GEL/.test(item)) return 'LABEL & PACKAGING';
  if (structure.includes('재봉사') || structure.includes('THREAD') || /\bSPUN\b|NYLON\s+THREAD|KEVLAR\s+THREAD|\bSW\s*[*-]?\s*TH\b/.test(item) || (unit==='CONE' && /THREAD|SPUN|NYLON|KEVLAR/.test(item))) return 'SEWING THREAD';
  if (isPrintedOutshell(r)) return 'OUTSHELL';
  if (item.includes('CHICRON') || item.includes('CHIRON') || item.startsWith('DIGITAL')) return 'OUTSHELL';
  if (/FABRIC|LEATHER|KEVLAR|MESH|POLY \(|NYLON \(|OUTSHELL|SHELL/.test(item) && !/WEBBING|ZIPPER|TAPE|PATCH|PRINT MATERIAL/.test(item)) return 'OUTSHELL';
  if (item) return 'TRIMS';
  return 'NEEDS REVIEW';
}

/**
 * Apply newly introduced priority rules to persisted auto-classification results.
 * A dictionary entry is created by the group editor and is therefore treated as
 * an explicit manual decision. Those materials must never be migrated.
 */
export function migrateStoredPriorityGroups(styles:StyleData[], dictionary:ClassificationDictionary):StyleData[] {
  return styles.map(style=>({...style,materials:style.materials.map(material=>{
    const priorityName=material.sources.some(source=>{const item=normalizeText(source.item);return item.includes('CHICRON')||item.includes('CHIRON')||isPrintedOutshell(source)});
    const manuallyClassified=material.sources.some(source=>dictionaryKeys(source).some(key=>dictionary[key]!==undefined));
    const manual=material.groupSource==='manual'||manuallyClassified;
    return priorityName&&!manual&&material.group==='TRIMS'
      ? {...material,group:'OUTSHELL',included:true,groupSource:'auto',classificationVersion:CLASSIFICATION_VERSION}
      : {...material,groupSource:manual?'manual':material.groupSource||'auto',classificationVersion:CLASSIFICATION_VERSION};
  }),classificationVersion:CLASSIFICATION_VERSION}));
}
