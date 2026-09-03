import ExcelJS from 'exceljs';
import {describe,expect,it} from 'vitest';
import {parseCbdFiles} from './parser';

describe('comparison CBD group subtotals',()=>{
 it('preserves labeled and legacy unlabeled source subtotal rows',async()=>{
  const workbook=new ExcelJS.Workbook(),sheet=workbook.addWorksheet('27 F-16 JERSEY');
  sheet.addRow(['Model Name','27 F-16 JERSEY']);
  sheet.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);
  sheet.addRow(['OUTSHELL','FABRIC','','YD',1,1,0,1,'']);
  sheet.addRow(['','','','','','','',1,'']);
  sheet.addRow(['SEWING THREAD','THREAD','','CONE','','','','']);
  sheet.addRow(['','','','','','','',.3,'']);
  sheet.addRow(['LABEL & PACKAGING SUBTOTAL','','','','','','',.4,'']);
  sheet.addRow(['','Total material cost','','','','','',1.7,'']);
  const data=await workbook.xlsx.writeBuffer(),file=new File([data as BlobPart],'27.xlsx'),result=await parseCbdFiles([file],'reference'),style=result.styles[0];
  expect(style.groupTotals).toMatchObject({OUTSHELL:1,'SEWING THREAD':.3,'LABEL & PACKAGING':.4});
  expect(style.summary.totalMaterialCost).toBe(1.7);
 });
});
