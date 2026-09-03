import {readFile,writeFile} from 'node:fs/promises';
import ExcelJS from 'exceljs';
import {parseCbdFiles} from '../src/comparison/parser';
import {matchMaterials,matchStyles} from '../src/comparison/matcher';
import {buildComparisonWorkbook} from '../src/comparison/exporter';
import type {ComparisonState} from '../src/comparison/types';

const paths=['reference/FLY RACING 27 MX JERSEY CBD 2025.10.17.xlsx','reference/FLY RACING 28 MX BMX JERSEY CBD SHEET (내부) 2026.09.02.xlsx'];
const files=await Promise.all(paths.map(async path=>new File([await readFile(path)],path.split('/').at(-1)!)));
const reference=await parseCbdFiles([files[0]],'reference'),comparison=await parseCbdFiles([files[1]],'current'),styles=[...reference.styles,...comparison.styles],styleMatches=matchStyles(styles),materialMatches=styleMatches.filter(match=>match.referenceId&&match.currentId).map(match=>matchMaterials(match,styles));
const target=styleMatches.find(match=>styles.find(item=>item.id===match.referenceId)?.styleName==='27 F-16 JERSEY (S ~ 2XL)'&&styles.find(item=>item.id===match.currentId)?.styleName==='F-16 JERSEY');
if(!target)throw new Error('F-16 JERSEY style match not found');
console.log('parsed group totals',styles.find(item=>item.id===target.referenceId)?.groupTotals,styles.find(item=>item.id===target.currentId)?.groupTotals);
const state:ComparisonState={version:2,referenceSeason:'27',currentSeason:'28',files:[],styles,styleMatches:[target],materialMatches:materialMatches.filter(set=>set.styleMatchId===target.id),step:5};
const workbook=buildComparisonWorkbook(state),buffer=await workbook.xlsx.writeBuffer(),path='/tmp/CBD_Comparison_27_vs_28_verified.xlsx';await writeFile(path,Buffer.from(buffer));
const reopened=new ExcelJS.Workbook();await reopened.xlsx.readFile(path);const sheet=reopened.worksheets[1],rowFor=(label:string,max=sheet.rowCount)=>{let found=0;sheet.getColumn(1).eachCell((cell,row)=>{if(!found&&row<=max&&cell.value===label)found=row});if(!found)throw new Error(`${label} row not found`);return found},result=(cell:string)=>Number((sheet.getCell(cell).value as ExcelJS.CellFormulaValue).result);
const expected={OUTSHELL:[3.5863002,5.2047],TRIMS:[.0399,.0399],'SEWING THREAD':[.3,.3],'LABEL & PACKAGING':[.642868338,.3998]} as const;
for(const [group,values] of Object.entries(expected)){const row=rowFor(group,12),subtotal=rowFor(`${group} 소계`);if(sheet.getCell(`B${row}`).formula!==`K${subtotal}`||sheet.getCell(`C${row}`).formula!==`L${subtotal}`)throw new Error(`${group} top summary is not linked to its subtotal`);if(Math.abs(result(`B${row}`)-values[0])>.000001||Math.abs(result(`C${row}`)-values[1])>.000001)throw new Error(`${group} amount mismatch`)}
const sewingSource=rowFor('SEWING THREAD 원본 CBD 소계');if(sheet.getCell(`O${sewingSource}`).value!=='원본 CBD의 그룹 소계 사용')throw new Error('source subtotal note missing');
const total=rowFor('TOTAL MATERIAL COST',12),referenceTotal=result(`B${total}`),comparisonTotal=result(`C${total}`);if(Math.abs(referenceTotal-4.569068538)>.000001||Math.abs(comparisonTotal-5.9444)>.000001)throw new Error('TOTAL MATERIAL mismatch');
const formulas:string[]=[];sheet.eachRow(row=>row.eachCell(cell=>{if(cell.formula)formulas.push(cell.formula)}));if(formulas.some(item=>/ISNUMBER|SUMPRODUCT|OFFSET|INDIRECT|#VALUE!|#REF!|#DIV\/0!|#NAME\?/i.test(item)))throw new Error('complex or invalid formula found');
console.log(JSON.stringify({path,referenceGroupTotal:referenceTotal,comparisonGroupTotal:comparisonTotal,sewingThread:[result(`B${rowFor('SEWING THREAD',12)}`),result(`C${rowFor('SEWING THREAD',12)}`)],totalFormulas:[sheet.getCell(`B${total}`).formula,sheet.getCell(`C${total}`).formula],detailFormula:sheet.getCell('M16').formula,fullCalcOnLoad:workbook.calcProperties.fullCalcOnLoad},null,2));
