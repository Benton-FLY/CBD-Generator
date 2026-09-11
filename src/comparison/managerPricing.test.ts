import {describe,expect,it} from 'vitest';
import {applyManagerFloors,round4} from './managerPricing';
import type {CbdMaterialRow,MaterialMatchCluster} from './types';

const row=(id:string,overrides:Partial<CbdMaterialRow>={}):CbdMaterialRow=>({id,group:'OUTSHELL',material:id,size:'40',unit:'YD',cost:1,usage:1,loss:0,width:'40',remark:'',order:0,...overrides});
const match=(extra:Partial<MaterialMatchCluster>={}):MaterialMatchCluster=>({id:'m',referenceRowIds:['old'],currentRowId:'new',relationType:'one-to-one',matchSource:'auto',finalGroup:'OUTSHELL',status:'EXACT',confidence:1,...extra});

describe('Manager Format prior-season floors',()=>{
 it('applies usage and cost independently and preserves source values',()=>{
  const result=applyManagerFloors([row('old',{cost:.1,usage:1})],[row('new',{cost:.12,usage:.8})],match());
  expect(result).toMatchObject({sourceOldUnitCost:.1,sourceOldUsage:1,sourceNewUnitCost:.12,sourceNewUsage:.8,appliedNewUnitCost:.12,appliedNewUsage:1,unitCostFloorApplied:false,usageFloorApplied:true});
 });
 it('does not hold values that are equal after four-place rounding',()=>{
  const result=applyManagerFloors([row('old',{cost:1.00004,usage:2.00004})],[row('new',{cost:1.00003,usage:2.00003})],match());
  expect(result.unitCostFloorApplied).toBe(false);expect(result.usageFloorApplied).toBe(false);expect(result.appliedNewUnitCost).toBe(1.00003);
 });
 it('rejects unit changes and automatic split/merge floors',()=>{
  expect(applyManagerFloors([row('old',{unit:'YD',cost:3})],[row('new',{unit:'M',cost:2})],match()).status).toBe('needs-review');
  expect(applyManagerFloors([row('old'),row('old2')],[row('new')],match({referenceRowIds:['old','old2'],relationType:'many-to-one'})).status).toBe('needs-review');
 });
 it('keeps a fully new material unchanged',()=>expect(applyManagerFloors([],[row('new',{cost:4,usage:.2})],match({referenceRowIds:[]}))).toMatchObject({appliedNewUnitCost:4,appliedNewUsage:.2,status:'current-only'}));
 it('round4 follows the workbook display precision',()=>expect(round4(1.23456)).toBe(1.2346));
});
