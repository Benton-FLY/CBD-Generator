import { describe, expect, it } from 'vitest';
import { classify, migrateStoredPriorityGroups, normalizeText } from './classifier';
import type { BomRow, Material, StyleData } from './types';

const row=(item:string):BomRow=>({id:item,sourceFile:'test.xlsx',sourceSheet:'BOM',sourceRow:1,structure:'',materialType:'자재',sequence:'',itemNo:'',item,width:'',color:'',unit:'YD',usage:1,currency:'USD',convertedPrice:1,materialCostAdjustment:0,specialFlag:'',remark:''});

describe('CBD group classification',()=>{
  it.each([
    'CHICRON TOUCH (E-SUEDE) 0.6MM (8G060-B000D)','CHICRON TOUCH','CHICRON 0.8MM','Chicron Touch',' chicron touch',
    'CHIRON 0.8MM (CN)','CHIRON 0.7MM','CHIRON BLACK','chiron 0.8mm',' Chiron 0.8MM ',
    'DIGITAL SOLID PRINTED','DIGITAL PRINTED','DIGITAL','DIGITAL PRINTED SOLID',
    'DIGITAL PRINTED MESH','DIGITAL 4WAY STRETCH',' digital   printed mesh ',
  ])('classifies %j as OUTSHELL',item=>expect(classify(row(item))).toBe('OUTSHELL'));

  it('normalizes case, surrounding whitespace, and repeated spaces',()=>{
    expect(normalizeText(' digital   solid printed ')).toBe('DIGITAL SOLID PRINTED');
    expect(normalizeText(' printed\u00a0solid\nfly racing ')).toBe('PRINTED SOLID FLY RACING');
  });

  it.each([
    'PRINTED SOLID FLY RACING 25 JERSEY+JX-44 (POLY)',
    'PRINTED FLY RACING 25 KINETIC 1 JERSEY+JX-300 (POLY) FR/SL',
    'SOLID PRINTED FLY RACING JERSEY FABRIC','PRINTED SOLID JERSEY FABRIC','PRINTED JERSEY FABRIC',
    'ＰＲＩＮＴＥＤ　ＦＬＹ－ＲＡＣＩＮＧ JERSEY',
  ])('classifies printed textile %j as OUTSHELL',item=>expect(classify(row(item))).toBe('OUTSHELL'));

  it.each(['PRINTED LABEL','PRINTED CARTON STICKER','PRINT MATERIAL','SILICON PRINT COST','SUBLIMATION PRINT COST','HEAT TRANSFER FLY LOGO'])
    ('does not classify printed non-fabric %j as OUTSHELL',item=>expect(classify(row(item))).not.toBe('OUTSHELL'));

  it('keeps a manual dictionary override ahead of the automatic rule',()=>{
    expect(classify(row('CHICRON TOUCH'),{'ITEM:CHICRON TOUCH':'TRIMS'})).toBe('TRIMS');
    expect(classify(row('CHIRON 0.8MM'),{'ITEM:CHIRON 0.8MM':'TRIMS'})).toBe('TRIMS');
    expect(classify(row('DIGITAL PRINTED MESH'),{'ITEM:DIGITAL PRINTED MESH':'TRIMS'})).toBe('TRIMS');
  });

  it('migrates the persisted F-16 auto result while protecting a saved manual group',()=>{
    const source=row('CHICRON TOUCH (E-SUEDE) 0.6MM (8G060-B000D)');
    const material:Material={id:'chicron',item:source.item,width:'',unit:'YD',group:'TRIMS',included:true,baseCost:1,adjustedCost:1,baseUsage:1,adjustedUsage:1,additionalLoss:.05,remark:'',sources:[source],split:false};
    const style:StyleData={id:'f16',name:'F-16 GLOVE',sourceFile:'28 F-16 GLOVE.xls',sourceSheet:'BOM',materials:[material],laborRemark:''};
    expect(migrateStoredPriorityGroups([style],{})[0].materials[0].group).toBe('OUTSHELL');
    expect(migrateStoredPriorityGroups([style],{'ITEM:CHICRON TOUCH (E-SUEDE) 0.6MM (8G060-B000D)':'TRIMS'})[0].materials[0].group).toBe('TRIMS');
  });

  it('migrates only auto-classified persisted PRINTED TRIMS rows',()=>{
    const source=row('PRINTED FLY RACING 25 KINETIC JERSEY');
    const base:Material={id:'printed',item:source.item,width:'58',unit:'YD',group:'TRIMS',included:true,baseCost:1,adjustedCost:1,baseUsage:1,adjustedUsage:1,additionalLoss:.05,remark:'',sources:[source],split:false};
    const style:StyleData={id:'s',name:'S',sourceFile:'s.xls',sourceSheet:'BOM',materials:[base],laborRemark:''};
    const migrated=migrateStoredPriorityGroups([style],{})[0];expect(migrated.materials[0].group).toBe('OUTSHELL');expect(migrated.materials[0].groupSource).toBe('auto');expect(migrated.classificationVersion).toBe(2);
    expect(migrateStoredPriorityGroups([{...style,materials:[{...base,groupSource:'manual'}]}],{})[0].materials[0].group).toBe('TRIMS');
  });
});
