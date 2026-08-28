import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import type { CbdGroup, StyleData } from './types';
import { extendedCost } from './types';

const ORDER: CbdGroup[]=['OUTSHELL','TRIMS','SEWING THREAD','LABEL & PACKAGING','SPECIAL PROCESS (LIST ONLY)'];
const BLUE='4472C4', YELLOW='FFF2CC', AQUA='DDEBF7', BORDER='FF000000', WHITE='FFFFFFFF';
const MONEY='$#,##0.0000', DECIMAL='0.0000';
const safeSheet=(name:string,used:Set<string>)=>{let base=name.replace(/[\\/*?:[\]]/g,' ').trim().slice(0,31)||'CBD';let out=base,n=2;while(used.has(out)){const s=` (${n++})`;out=base.slice(0,31-s.length)+s;}used.add(out);return out;};
const borders={top:{style:'thin' as const,color:{argb:BORDER}},left:{style:'thin' as const,color:{argb:BORDER}},bottom:{style:'thin' as const,color:{argb:BORDER}},right:{style:'thin' as const,color:{argb:BORDER}}};
const rounded=(n:number)=>Math.round((n+Number.EPSILON)*10000)/10000;

export async function createBuyerWorkbook(styles:StyleData[],date=new Date()):Promise<ExcelJS.Workbook>{
  const wb=new ExcelJS.Workbook();wb.creator='CBD Chart Generator';wb.created=date;
  wb.calcProperties.fullCalcOnLoad=true;
  const used=new Set<string>();
  for(const style of styles){
    const ws=wb.addWorksheet(safeSheet(style.name,used),{pageSetup:{paperSize:9,orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.25,right:.25,top:.5,bottom:.5,header:.2,footer:.2}}});
    ws.columns=[{width:21},{width:54},{width:10},{width:8},{width:13},{width:11},{width:9},{width:16},{width:23}];
    ws.mergeCells('A1:C3');const logo=ws.getCell('A1');logo.value='FLY Racing';logo.font={name:'Arial',size:22,bold:true,color:{argb:BLUE}};logo.alignment={vertical:'middle',horizontal:'center'};
    ws.mergeCells('D1:I2');const model=ws.getCell('D1');model.value=`MODEL NAME : ${style.name}`;model.font={name:'Arial',size:20,bold:true,color:{argb:WHITE}};model.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};model.alignment={vertical:'middle',horizontal:'center',wrapText:true};model.border=borders;
    ws.mergeCells('G3:I3');const dc=ws.getCell('G3');dc.value=`Date : ${date.getFullYear()}.${String(date.getMonth()+1).padStart(2,'0')}.${String(date.getDate()).padStart(2,'0')}`;dc.font={name:'Arial',size:9};dc.alignment={horizontal:'right',vertical:'middle'};ws.getRow(1).height=28;ws.getRow(2).height=28;
    const header=ws.addRow(['Group of','Material','Size','Unit','Cost per Unit','Usage','Loss','Extended Cost','Remark']);header.height=28;header.eachCell(c=>{c.font={name:'Arial',bold:true,color:{argb:WHITE}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:BLUE}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true};c.border=borders;});
    const subtotalCells:string[]=[];
    for(const group of ORDER){
      const items=style.materials.filter(m=>m.group===group&&m.included);if(!items.length)continue;
      const start=ws.rowCount+1,detailCosts:number[]=[];
      for(const m of items){
        const listOnly=group==='SPECIAL PROCESS (LIST ONLY)';
        const row=ws.addRow(['',m.item,listOnly?'':m.width,listOnly?'':m.unit,listOnly?null:m.adjustedCost,listOnly?null:m.adjustedUsage,listOnly?null:m.additionalLoss,null,m.remark]);row.height=24;
        row.eachCell({includeEmpty:true},c=>{c.font={name:'Arial',size:10};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});
        if(!listOnly){const result=rounded(extendedCost(m));row.getCell(8).value={formula:`ROUND(E${row.number}*F${row.number}*(1+G${row.number}),4)`,result};row.getCell(5).numFmt=MONEY;row.getCell(6).numFmt=DECIMAL;row.getCell(7).numFmt='0%';row.getCell(8).numFmt=MONEY;detailCosts.push(result);}
      }
      const end=ws.rowCount;ws.mergeCells(start,1,end,1);const gc=ws.getCell(start,1);gc.value=group==='SPECIAL PROCESS (LIST ONLY)'?'SPECIAL PROCESS\n(LIST ONLY)':group;gc.alignment={textRotation:0,horizontal:'center',vertical:'middle',wrapText:true};gc.font={name:'Arial',bold:true};gc.border=borders;
      if(group!=='SPECIAL PROCESS (LIST ONLY)'){const result=rounded(detailCosts.reduce((s,n)=>s+n,0));const sub=ws.addRow(['',`${group} SUBTOTAL`,'','','','','',null,'']);sub.getCell(8).value={formula:`SUM(H${start}:H${end})`,result};sub.getCell(8).numFmt=MONEY;subtotalCells.push(`H${sub.number}`);sub.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:YELLOW}};c.font={name:'Arial',bold:true};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});}
    }
    const materialTotal=rounded(subtotalCells.reduce((sum,address)=>sum+(Number((ws.getCell(address).value as ExcelJS.CellFormulaValue).result)||0),0));
    const totals:[string,number|null,string,boolean][]=[['Total material cost',materialTotal,'',true],['Labor cost',null,style.laborRemark,false],['Overhead',null,'',false],['Profit',null,'',false],['FOB PRICE',style.finalFob??null,'',false]];
    totals.forEach(([label,value,remark,isFormula])=>{const row=ws.addRow(['',label,'','','','','',value,remark]);if(isFormula)row.getCell(8).value={formula:`SUM(${subtotalCells.join(',')})`,result:materialTotal};row.getCell(8).numFmt=MONEY;row.eachCell({includeEmpty:true},c=>{c.fill={type:'pattern',pattern:'solid',fgColor:{argb:AQUA}};c.font={name:'Arial',bold:true};c.border=borders;c.alignment={vertical:'middle',wrapText:true};});});
    ws.autoFilter={from:{row:4,column:1},to:{row:ws.rowCount,column:9}};ws.views=[{state:'frozen',ySplit:4}];ws.pageSetup.printArea=`A1:I${ws.rowCount}`;
  }
  return wb;
}
export async function downloadBuyerExcel(styles:StyleData[],filename='Buyer_CBD.xlsx'){const wb=await createBuyerWorkbook(styles);const buffer=await wb.xlsx.writeBuffer();saveAs(new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),filename);}
