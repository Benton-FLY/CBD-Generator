import { describe, expect, it } from 'vitest';
import { createBuyerWorkbook } from './exporter';
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
