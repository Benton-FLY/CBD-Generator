import { describe, expect, it } from 'vitest';
import { createBuyerWorkbook, createInternalWorkbook } from './exporter';
import type { StyleData } from './types';

describe('buyer workbook',()=>{
  it('writes a directly entered FINAL FOB to FOB PRICE with four-decimal formatting',async()=>{
    const style:StyleData={id:'s',name:'DIRECT FOB',sourceFile:'bom.xlsx',sourceSheet:'BOM',laborRemark:'',finalFob:48.98,materials:[]};
    const workbook=await createBuyerWorkbook([style],new Date('2026-08-28T00:00:00Z'));
    const sheet=workbook.worksheets[0];
    const row=sheet.getRows(1,sheet.rowCount)!.find(r=>r.getCell(2).value==='FOB PRICE')!;
    expect(row.getCell(8).value).toEqual({formula:'48.98',result:48.98});
    expect(row.getCell(8).numFmt).toBe('$0.0000');
  });
});

describe('internal workbook ERP pre-cost',()=>{
  it('leaves ERP pre-cost, difference, and difference rate blank when no web value exists',async()=>{
    const style:StyleData={id:'missing',name:'MISSING ERP',sourceFile:'bom.xlsx',sourceSheet:'BOM',laborRemark:'',finalFob:10,materials:[]};
    const workbook=await createInternalWorkbook([style]);
    const sheet=workbook.worksheets[0];
    const rowFor=(label:string)=>sheet.getColumn(9).values.findIndex(value=>value===label);
    const erp=sheet.getCell(rowFor('사전원가 재료비'),8);
    const difference=sheet.getCell(rowFor('사전원가와 CBD 재료비 차이'),8);
    const rate=sheet.getCell(rowFor('차이율'),8);
    expect(erp.value).toBeNull();expect(erp.formula).toBeUndefined();
    expect(difference.result).toBe('');expect(rate.result).toBe('');
    expect(difference.formula).toContain(`${erp.address}=""`);
  });
});
