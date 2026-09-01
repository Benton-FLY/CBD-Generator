import { describe, expect, it } from 'vitest';
import { classify, normalizeText } from './classifier';
import type { BomRow } from './types';

const row=(item:string):BomRow=>({id:item,sourceFile:'test.xlsx',sourceSheet:'BOM',sourceRow:1,structure:'',materialType:'자재',sequence:'',itemNo:'',item,width:'',color:'',unit:'YD',usage:1,currency:'USD',convertedPrice:1,materialCostAdjustment:0,specialFlag:'',remark:''});

describe('CBD group classification',()=>{
  it.each([
    'CHIRON 0.8MM (CN)','CHIRON 0.7MM','CHIRON BLACK','chiron 0.8mm',' Chiron 0.8MM ',
    'DIGITAL SOLID PRINTED','DIGITAL PRINTED','DIGITAL','DIGITAL PRINTED SOLID',
    'DIGITAL PRINTED MESH','DIGITAL 4WAY STRETCH',' digital   printed mesh ',
  ])('classifies %j as OUTSHELL',item=>expect(classify(row(item))).toBe('OUTSHELL'));

  it('normalizes case, surrounding whitespace, and repeated spaces',()=>{
    expect(normalizeText(' digital   solid printed ')).toBe('DIGITAL SOLID PRINTED');
  });

  it('keeps a manual dictionary override ahead of the automatic rule',()=>{
    expect(classify(row('CHIRON 0.8MM'),{'ITEM:CHIRON 0.8MM':'TRIMS'})).toBe('TRIMS');
    expect(classify(row('DIGITAL PRINTED MESH'),{'ITEM:DIGITAL PRINTED MESH':'TRIMS'})).toBe('TRIMS');
  });
});
