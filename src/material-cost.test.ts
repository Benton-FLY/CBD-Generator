import { describe, expect, it } from 'vitest';
import { createBuyerWorkbook } from './exporter';
import { extendedCost, groupSubtotal, totalMaterialCost, type Material, type StyleData } from './types';

const material=(id:string,group:Material['group'],raw:number):Material=>({
  id,item:id,width:'',unit:'EA',group,included:true,baseCost:raw,adjustedCost:raw,
  baseUsage:1,adjustedUsage:1,additionalLoss:0,remark:'',sources:[],split:false,
});

describe('Excel-compatible material cost rounding',()=>{
  it('handles floating-point boundaries with Excel-style half-away-from-zero rounding',()=>{
    const positive=material('positive','TRIMS',1.23445),negative=material('negative','TRIMS',-1.23445);
    expect(extendedCost(positive)).toBe(1.2345);
    expect(extendedCost(negative)).toBe(-1.2345);
  });

  it('rounds every row before group subtotals and the material total',()=>{
    const materials=[
      material('shell-1','OUTSHELL',1.000049),
      material('shell-2','OUTSHELL',1.000049),
      material('trim-1','TRIMS',1.000049),
      material('trim-2','TRIMS',1.751549),
    ];
    expect(materials.reduce((sum,m)=>sum+m.adjustedCost*m.adjustedUsage,0)).toBeCloseTo(4.7517,4);
    expect(materials.map(extendedCost)).toEqual([1,1,1,1.7515]);
    expect(groupSubtotal(materials.filter(m=>m.group==='OUTSHELL'))).toBe(2);
    expect(groupSubtotal(materials.filter(m=>m.group==='TRIMS'))).toBe(2.7515);
    expect(totalMaterialCost(materials)).toBe(4.7515);
  });

  it('exports ROUND row formulas, SUM subtotals, and the same total result',async()=>{
    const materials=[material('shell','OUTSHELL',2.000049),material('trim','TRIMS',2.751549)];
    const style:StyleData={id:'rounding',name:'28 EVO GLOVE',sourceFile:'BOM.xls',sourceSheet:'BOM',materials,laborRemark:''};
    const workbook=await createBuyerWorkbook([style]);
    const sheet=workbook.worksheets[0];
    const formulas:{label:string;formula?:string;result?:unknown}[]=[];
    sheet.eachRow(row=>{const cell=row.getCell(8);formulas.push({label:String(row.getCell(2).value||''),formula:cell.formula,result:cell.result})});
    expect(formulas.filter(x=>['shell','trim'].includes(x.label)).map(x=>x.formula)).toEqual(['ROUND(E5*F5*(1+G5),4)','ROUND(E7*F7*(1+G7),4)']);
    expect(formulas.filter(x=>x.label.endsWith('SUBTOTAL')).map(x=>x.formula)).toEqual(['SUM(H5:H5)','SUM(H7:H7)']);
    expect(formulas.find(x=>x.label==='Total material cost')).toMatchObject({formula:'SUM(H6,H8)',result:4.7515});
  });
});
