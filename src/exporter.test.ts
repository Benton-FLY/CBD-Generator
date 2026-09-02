import { describe, expect, it } from 'vitest';
import { createBuyerWorkbook, createInternalWorkbook } from './exporter';
import type { StyleData } from './types';

const material=(id:string,group:'OUTSHELL'|'TRIMS'|'NEEDS REVIEW',cost=1)=>({id,item:id,width:'58',unit:'YD',group,included:true,baseCost:cost,adjustedCost:cost,baseUsage:1,adjustedUsage:1,additionalLoss:0,remark:'',sources:[],split:false});

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
  it('keeps review rows out of Buyer output and highlights them only in internal output',async()=>{
    const style:StyleData={id:'review',name:'REVIEW',sourceFile:'bom.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('PRINTED FABRIC','OUTSHELL'),material('UNKNOWN','NEEDS REVIEW')]};
    const buyer=await createBuyerWorkbook([style]),internal=await createInternalWorkbook([style]);
    expect(buyer.worksheets[0].getColumn(2).values).not.toContain('UNKNOWN');
    const row=internal.worksheets[0].getRows(1,internal.worksheets[0].rowCount)!.find(r=>r.getCell(2).value==='UNKNOWN')!;
    expect(row.getCell(2).font.color?.argb).toBe('FFDC2626');expect(row.getCell(2).fill).toMatchObject({fgColor:{argb:'FFFEF2F2'}});
    const total=internal.worksheets[0].getRows(1,internal.worksheets[0].rowCount)!.find(r=>r.getCell(2).value==='Total material cost')!;expect(total.getCell(8).result).toBe(1);
  });
});
