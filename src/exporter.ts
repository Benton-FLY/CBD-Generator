import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { CbdGroup, StyleData } from './types';
import { extendedCost, groupSubtotal, roundTo4 } from './types';
import { displayStyleName } from './parser';

const ORDER: CbdGroup[]=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)'];
const BLUE='4472C4', YELLOW='FFF2CC', AQUA='DDEBF7', BORDER='FF000000', WHITE='FFFFFFFF';
const MONEY='$0.0000', DECIMAL='0.0000';
export const excelStyleName=(name:string)=>displayStyleName(name.replace(/\bMODEL\s+NAME\s*:\s*/ig,''))||'CBD';
const safeSheet=(name:string,used:Set<string>)=>{let base=excelStyleName(name).replace(/[\\/*?:[\]]/g,' ').trim().slice(0,31)||'CBD';let out=base,n=2;while(used.has(out)){const s=` (${n++})`;out=base.slice(0,31-s.length)+s;}used.add(out);return out;};
const borders={top:{style:'thin' as const,color:{argb:BORDER}},left:{style:'thin' as const,color:{argb:BORDER}},bottom:{style:'thin' as const,color:{argb:BORDER}},right:{style:'thin' as const,color:{argb:BORDER}}};

async function createWorkbook(styles:StyleData[],date=new Date(),internalReview=false):Promise<ExcelJS.Workbook>{
  const wb=new ExcelJS.Workbook();wb.creator='CBD Chart Generator';wb.created=date;
  wb.calcProperties.fullCalcOnLoad=true;
  const used=new Set<string>();
  for(const style of styles){
    const ws=wb.addWorksheet(safeSheet(style.name,used),{pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.5,bottom:.5,header:.2,footer:.2}}});
    ws.columns=[{width:21},{width:54},{width:10},{width:8},{width:13},{width:11},{width:9},{width:16},{width:23}];
    ws.mergeCells('A1:C3');const logo=ws.getCell('A1');logo.value='FLY Racing CBD Sheet';logo.font={name:'Arial',size:22,bold:true,color:{argb:BLUE}};logo.alignment={vertical:'middle',horizontal:'center'};
    ws.mergeCells('D1:I2');const model=ws.getCell('D1');model.value=excelStyleName(style.name);model.font={name:'Arial',size:20,bold:true,color:{argb:WHITE}};model.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};model.alignment={vertical:'middle',horizontal:'center',wrapText:true};model.border=borders;
    ws.mergeCells('G3:I3');const dc=ws.getCell('G3');dc.value=`Date : ${date.getFullYear()}.${String(date.getMonth()+1).padStart(2,'0')}.${String(date.getDate()).padStart(2,'0')}`;dc.font={name:'Arial',size:9};dc.alignment={horizontal:'right',vertical:'middle'};ws.getRow(1).height=28;ws.getRow(2).height=28;
    const header=ws.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);header.height=28;header.eachCell(c=>{c.font={name:'Arial',bold:true,color:{argb:WHITE}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.border=borders;});
    const subtotalCells:string[]=[];
    for(const group of ORDER){
      const items=style.materials.filter(m=>m.group===group&&m.included);if(!items.length)continue;
      const start=ws.rowCount+1;
      for(const m of items){
        const listOnly=group==='SPECIAL PROCESS (LIST ONLY)';
        const row=ws.addRow(['',m.item,listOnly?'':m.width,listOnly?'':m.unit,listOnly?null:m.adjustedCost,listOnly?null:m.adjustedUsage,listOnly?null:m.additionalLoss,null,m.remark]);row.height=Math.max(24,18*(String(m.remark||'').split(/\r?\n| \/ /).length));
        row.eachCell({includeEmpty:true},c=>{c.font={name:'Arial',size:10};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});
        if(!listOnly){const result=extendedCost(m);row.getCell(8).value={formula:`ROUND(E${row.number}*F${row.number}*(1+G${row.number}),4)`,result};row.getCell(5).numFmt=MONEY;row.getCell(6).numFmt=DECIMAL;row.getCell(7).numFmt='0%';row.getCell(8).numFmt=MONEY;}
      }
      const end=ws.rowCount;ws.mergeCells(start,1,end,1);const gc=ws.getCell(start,1);gc.value=group==='SPECIAL PROCESS (LIST ONLY)'?'SPECIAL PROCESS\n(LIST ONLY)':group;gc.alignment={textRotation:0,horizontal:'center',vertical:'middle',wrapText:true};gc.font={name:'Arial',bold:true};gc.border=borders;
      if(group!=='SPECIAL PROCESS (LIST ONLY)'){const result=groupSubtotal(items);const sub=ws.addRow(['',`${group} SUBTOTAL`,'','','','','',null,'']);sub.getCell(8).value={formula:`SUM(H${start}:H${end})`,result};sub.getCell(8).numFmt=MONEY;subtotalCells.push(`H${sub.number}`);sub.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:YELLOW}};c.font={name:'Arial',bold:true};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});}
    }
    const materialTotal=roundTo4(subtotalCells.reduce((sum,address)=>sum+(Number((ws.getCell(address).value as ExcelJS.CellFormulaValue).result)||0),0));
    let materialTotalRow=0,fobRow=0;
    const totals:[string,number|null,string,boolean][]=[['Total material cost',materialTotal,'',true],['Labor cost',null,style.laborRemark,false],['Overhead',null,'',false],['Profit',null,'',false],['FOB PRICE',style.finalFob??null,'',false]];
    totals.forEach(([label,value,remark,isFormula])=>{const row=ws.addRow(['',label,'','','','','',value,remark]);if(label==='Total material cost')materialTotalRow=row.number;if(label==='FOB PRICE')fobRow=row.number;if(isFormula)row.getCell(8).value={formula:`SUM(${subtotalCells.join(',')})`,result:materialTotal};else if(label==='FOB PRICE'&&value!==null)row.getCell(8).value={formula:String(value),result:value};row.getCell(8).numFmt=MONEY;row.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:AQUA}};c.font={name:'Arial',bold:true};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});});
    if(internalReview){
      const erp=typeof style.erpMaterialCost==='number'&&Number.isFinite(style.erpMaterialCost)?style.erpMaterialCost:undefined;
      const ratio=style.finalFob?materialTotal/style.finalFob:'',difference=erp===undefined?'':materialTotal-erp,differenceRate=erp?(materialTotal-erp)/erp:'';
      const ratioRow=ws.addRow(['INTERNAL REVIEW','','','','','','',{formula:`IFERROR(H${materialTotalRow}/H${fobRow},"")`,result:ratio},'CBD 재료비 / FOB']);
      const erpRow=ws.addRow(['','','','','','','',erp??null,'사전원가 재료비']);
      const differenceRow=ws.addRow(['','','','','','','',{formula:`IF(OR(H${materialTotalRow}="",H${erpRow.number}=""),"",H${materialTotalRow}-H${erpRow.number})`,result:difference},'사전원가와 CBD 재료비 차이']);
      const rateRow=ws.addRow(['','','','','','','',{formula:`IFERROR(H${differenceRow.number}/H${erpRow.number},"")`,result:differenceRate},'차이율']);
      [ratioRow,erpRow,differenceRow,rateRow].forEach((row,index)=>{row.getCell(8).numFmt=index===0||index===3?'0%':MONEY;row.getCell(8).font={name:'Arial',bold:true,color:{argb:BLUE}};row.getCell(8).alignment={horizontal:'right'};row.getCell(9).font={name:'Arial',bold:true,color:{argb:BLUE}};row.getCell(9).alignment={horizontal:'left'};row.eachCell({includeEmpty:true},cell=>cell.border=borders)});
      ratioRow.getCell(1).font={name:'Arial',bold:true,color:{argb:'FFC00000'}};ratioRow.getCell(1).value='INTERNAL USE ONLY – DO NOT SEND TO BUYER';
    }
    ws.autoFilter={from:{row:4,column:1},to:{row:ws.rowCount,column:9}};ws.views=[{state:'frozen',ySplit:4}];ws.pageSetup.printArea=`A1:I${ws.rowCount}`;
  }
  if(internalReview){
    const audit=wb.addWorksheet('ERP MATERIAL CHECK');audit.columns=[{width:28},{width:13},{width:14},{width:16},{width:9},{width:10},{width:16},{width:16},{width:48},{width:13},{width:18},{width:20},{width:14},{width:38},{width:16}];
    const header=audit.addRow(['STYLE KEY','원본 BOM 행 번호','일련번호','부모일련번호','레벨','하위YN','자재구분','Item#','Item','금액','상위 부모행 여부','사전원가 재료비 포함 여부','포함 금액','포함 또는 제외 사유','FALLBACK AMOUNT']);header.eachCell(cell=>{cell.font={bold:true,color:{argb:WHITE}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};cell.border=borders;cell.alignment={wrapText:true}});
    for(const style of styles)for(const entry of style.erpAudit||[]){const source=entry.source,row=audit.addRow([style.id,source.sourceRow,source.serialNo||'',source.parentSerialNo||'',source.level||'',source.hasChildren||'',source.materialType,source.itemNo,source.item,source.amount??null,entry.isParent,entry.included,null,entry.reason,entry.usedFallback?entry.includedAmount:null]);row.getCell(13).value={formula:`IF(AND(G${row.number}="자재",K${row.number}=FALSE),IF(ISNUMBER(J${row.number}),J${row.number},O${row.number}),0)`,result:entry.includedAmount};row.getCell(10).numFmt=DECIMAL;row.getCell(13).numFmt=DECIMAL;row.getCell(15).numFmt=DECIMAL;row.eachCell({includeEmpty:true},cell=>{cell.border=borders;cell.alignment={vertical:'middle',wrapText:true}})}
    audit.getColumn(15).hidden=true;audit.views=[{state:'frozen',ySplit:1}];audit.autoFilter={from:{row:1,column:1},to:{row:Math.max(1,audit.rowCount),column:14}};
  }
  return wb;
}
export const createBuyerWorkbook=(styles:StyleData[],date=new Date())=>createWorkbook(styles,date,false);
export const createInternalWorkbook=(styles:StyleData[],date=new Date())=>createWorkbook(styles,date,true);
export async function downloadBuyerExcel(styles:StyleData[],filename='Buyer_CBD.xlsx'){const wb=await createBuyerWorkbook(styles);const buffer=await wb.xlsx.writeBuffer();saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),filename);}
export async function downloadInternalExcel(styles:StyleData[],filename='Internal_CBD.xlsx'){const wb=await createInternalWorkbook(styles);const buffer=await wb.xlsx.writeBuffer();saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),filename);}
