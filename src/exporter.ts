import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { CbdGroup, Material, StyleData } from './types';
import { extendedCost } from './types';

const ORDER: CbdGroup[]=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)'];
const BLUE='4472C4', YELLOW='FFF2CC', AQUA='DDEBF7', BORDER='FF000000';
const safeSheet=(name:string,used:Set<string>)=>{ let base=name.replace(/[\\/*?:[\]]/g,' ').trim().slice(0,31)||'CBD'; let out=base,n=2; while(used.has(out)){const s=` (${n++})`;out=base.slice(0,31-s.length)+s;} used.add(out); return out; };
const borders={top:{style:'thin' as const,color:{argb:BORDER}},left:{style:'thin' as const,color:{argb:BORDER}},bottom:{style:'thin' as const,color:{argb:BORDER}},right:{style:'thin' as const,color:{argb:BORDER}}};
const groupTotal=(items:Material[],group:CbdGroup,style:StyleData)=> group==='SEWING THREAD' && style.threadSubtotal!==undefined ? style.threadSubtotal : items.filter(m=>m.group===group&&m.included).reduce((s,m)=>s+extendedCost(m),0);

export async function createBuyerWorkbook(styles: StyleData[], date=new Date()): Promise<ExcelJS.Workbook> {
  const wb=new ExcelJS.Workbook(); wb.creator='CBD Chart Generator'; wb.created=date; const used=new Set<string>();
  for(const style of styles){
    const ws=wb.addWorksheet(safeSheet(style.name,used),{pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.5,bottom:.5,header:.2,footer:.2}}});
    ws.columns=[{width:15.2},{width:59.9},{width:9.1},{width:6.1},{width:8.2},{width:6.2},{width:6.2},{width:13.2},{width:23.4}];
    ws.mergeCells('A1:C3'); ws.getCell('A1').value='FLY Racing'; ws.getCell('A1').font={name:'Arial',size:22,bold:true,color:{argb:BLUE}}; ws.getCell('A1').alignment={vertical:'middle',horizontal:'center'};
    ws.mergeCells('D1:I1'); ws.getCell('D1').value=`Model Name : ${style.name}`; ws.mergeCells('D2:F2'); ws.getCell('D2').value=`Date : ${date.getFullYear()}.${String(date.getMonth()+1).padStart(2,'0')}.${String(date.getDate()).padStart(2,'0')}`; ws.mergeCells('G2:I2'); ws.getCell('G2').value='REFERENCES';
    const header=ws.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);
    header.height=28; header.eachCell(c=>{c.font={name:'Arial',bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.border=borders;});
    for(const group of ORDER){
      const items=style.materials.filter(m=>m.group===group&&m.included); if(!items.length) continue; const start=ws.rowCount+1;
      for(const m of items){
        const listOnly=group==='SPECIAL PROCESS (LIST ONLY)', thread=group==='SEWING THREAD';
        const row=ws.addRow(['',m.item,m.width,m.unit,thread||listOnly?null:m.adjustedCost,thread||listOnly?null:m.adjustedUsage,thread||listOnly?null:m.additionalLoss,thread||listOnly?null:extendedCost(m),m.remark]);
        row.height=24; row.eachCell({includeEmpty:true},c=>{c.font={name:'Arial',size:10};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});
        if(!thread&&!listOnly){row.getCell(5).numFmt='$0.0000';row.getCell(6).numFmt='0.000';row.getCell(7).numFmt='0%';row.getCell(8).numFmt='$0.0000';}
      }
      const end=ws.rowCount; ws.mergeCells(start,1,end,1); const gc=ws.getCell(start,1);gc.value=group;gc.alignment={textRotation:90,horizontal:'center',vertical:'middle',wrapText:true};gc.font={name:'Arial',bold:true};gc.border=borders;
      if(group!=='SPECIAL PROCESS (LIST ONLY)'){
        const sub=ws.addRow(['',`${group} SUBTOTAL`,'','','','','',groupTotal(items,group,style),'']); sub.getCell(8).numFmt='$0.0000';sub.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:YELLOW}};c.font={name:'Arial',bold:true};c.border=borders;});
      }
    }
    const materialTotal=ORDER.slice(0,4).reduce((s,g)=>s+groupTotal(style.materials,g,style),0);
    [['Total material cost',materialTotal,''],['Labor cost',null,style.laborRemark],['Overhead',null,''],['Profit',null,''],['FOB PRICE',style.finalFob??null,'']].forEach(([label,value,remark])=>{
      const row=ws.addRow(['',label,'','','','','',value,remark]); row.getCell(8).numFmt='$0.00'; row.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:AQUA}};c.font={name:'Arial',bold:true};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});
    });
    ws.autoFilter={from:{row:4,column:1},to:{row:ws.rowCount,column:9}}; ws.views=[{state:'frozen',ySplit:4}]; ws.pageSetup.printArea=`A1:I${ws.rowCount}`;
  }
  return wb;
}
export async function downloadBuyerExcel(styles:StyleData[],filename='Buyer_CBD.xlsx'){const wb=await createBuyerWorkbook(styles);const buffer=await wb.xlsx.writeBuffer();saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),filename);}
