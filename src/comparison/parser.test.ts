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
 it('keeps an exact FINAL FOB separate from the material-to-FOB ratio and cross-validates it',async()=>{
  const workbook=new ExcelJS.Workbook(),sheet=workbook.addWorksheet('RAYCE PANT');
  sheet.addRow(['Model Name','RAYCE PANT']);sheet.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);sheet.addRow(['OUTSHELL','FABRIC','','YD',17.6925,1,0,17.6925,'']);
  sheet.addRow(['','Total material cost','','','','','',17.6925,'']);sheet.addRow(['','FINAL\nFOB','','','','','',36.63,'']);sheet.addRow(['INTERNAL USE ONLY','','','','','','',0.483005733,'CBD\u00a0Material /\nFOB']);
  const data=await workbook.xlsx.writeBuffer(),file=new File([data as BlobPart],'28.xlsx'),result=await parseCbdFiles([file],'current'),summary=result.styles[0].summary;
  expect(summary.finalFob).toBe(36.63);expect(summary.materialToFobRatio).toBe(0.483005733);expect(summary.fobEvidence).toBe('explicit');expect(summary.calculatedFob).toBeCloseTo(36.63,3);expect(summary.fobValidation).toBe('matched');
 });
 it('does not promote excluded FOB labels or a ratio to FINAL FOB',async()=>{const workbook=new ExcelJS.Workbook(),sheet=workbook.addWorksheet('MISSING FOB');sheet.addRow(['Model Name','MISSING FOB']);sheet.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);sheet.addRow(['OUTSHELL','FABRIC','','YD',1,1,0,17.6925,'']);sheet.addRow(['','Total material cost','','','','','',17.6925,'']);sheet.addRow(['','','','','','','',0.483005733,'FOB RATIO']);sheet.addRow(['','','','','','','',99,'NOT FINAL FOB']);const data=await workbook.xlsx.writeBuffer(),file=new File([data as BlobPart],'missing.xlsx'),summary=(await parseCbdFiles([file],'current')).styles[0].summary;expect(summary.finalFob).toBeUndefined();expect(summary.materialToFobRatio).toBe(0.483005733);expect(summary.fobEvidence).toBe('review-required')});
});
