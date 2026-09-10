import {describe,it,expect} from 'vitest';
import ExcelJS from 'exceljs';
import {buildManagerWorkbook,MANAGER_WIDTHS} from './managerExporter';
import {buildComparisonWorkbook} from './exporter';
import {resolveGroup,groupAmounts} from './groups';
import {matchMaterials} from './matcher';
import {upgradeComparison} from './persistence';
import type {CbdMaterialRow,ComparisonState,MaterialMatchCluster} from './types';
const material=(id:string,extended:number,group='OUTSHELL'):CbdMaterialRow=>({id,material:id,extended,group,cost:2,usage:extended/2,loss:.05,unit:'YD',size:'40',width:'40',remark:'Original remark',order:0});
const fixture=()=>{
 const references=[material('ref-a',2),material('ref-b',3),material('ref-only',1)],currents=[material('new-a',4,'TRIMS'),material('new-only',2)];
 const clusters:MaterialMatchCluster[]=[{id:'merged',referenceRowIds:['ref-a','ref-b'],currentRowId:'new-a',finalGroup:'TRIMS',manualLocked:true,manualGroupOverride:true,matchSource:'manual',relationType:'many-to-one',status:'MERGED N:1',confidence:1},{id:'ref',referenceRowIds:['ref-only'],currentRowId:null,finalGroup:'OUTSHELL',matchSource:'auto',relationType:'one-to-one',status:'REFERENCE ONLY',confidence:0},{id:'cur',referenceRowIds:[],currentRowId:'new-only',finalGroup:'OUTSHELL',matchSource:'auto',relationType:'one-to-one',status:'CURRENT ONLY',confidence:0}];
 const state:ComparisonState={version:2,referenceSeason:'27',currentSeason:'28',files:[],styles:[{id:'r',side:'reference',fileName:'r',sheetName:'TEST',styleName:'27 TEST',materials:references,summary:{finalFob:10,fobEvidence:'explicit'},groupOrder:['OUTSHELL']},{id:'c',side:'current',fileName:'c',sheetName:'TEST',styleName:'28 TEST',materials:currents,summary:{finalFob:12,fobEvidence:'explicit'},groupOrder:['OUTSHELL','TRIMS']}],styleMatches:[{id:'pair',referenceId:'r',currentId:'c',method:'Manual',status:'Manual',confidence:1}],materialMatches:[{styleMatchId:'pair',clusters}],step:5};return state;
};
const find=(s:ExcelJS.Worksheet,col:number,value:string)=>s.getColumn(col).values.findIndex(v=>v===value);
const result=(s:ExcelJS.Worksheet,row:number,col:number)=>s.getCell(row,col).result;
describe('Manager Format renderer',()=>{
 it('preserves source identity, blank counterparts, N:1 anchor sums and canonical financials after reopening',async()=>{
  const state=fixture(),before=structuredClone(state),book=buildManagerWorkbook(state),reopened=new ExcelJS.Workbook();await reopened.xlsx.load(await book.xlsx.writeBuffer());const sheet=reopened.worksheets[0];
  expect(state).toEqual(before);expect(reopened.worksheets).toHaveLength(1);expect(sheet.getCell('C2').value).toBe('28 TEST');expect(sheet.getCell('S2').value).toBe('27 TEST');
  const anchor=find(sheet,2,'new-a');expect(sheet.getCell(anchor,18).value).toBe('ref-a');expect(sheet.getCell(anchor+1,18).value).toBe('ref-b');expect(sheet.getCell(anchor,15).formula).toBe(`SUM(H${anchor}:H${anchor+1})-SUM(X${anchor}:X${anchor+1})`);expect(result(sheet,anchor,15)).toBe(-1);
  for(let c=1;c<=9;c++)expect(sheet.getCell(anchor+1,c).value).toBeNull();for(let c=11;c<=14;c++)expect(sheet.getCell(anchor,c).value).toBeNull();expect(sheet.getCell(anchor+1,15).value).toBeNull();
  const removed=find(sheet,18,'ref-only'),added=find(sheet,2,'new-only');for(let c=1;c<=9;c++)expect(sheet.getCell(removed,c).value).toBeNull();for(let c=17;c<=25;c++)expect(sheet.getCell(added,c).value).toBeNull();expect(result(sheet,removed,15)).toBe(-1);expect(result(sheet,removed,12)).toBe(-1);expect(result(sheet,added,15)).toBe(2);expect(sheet.getCell(added,12).value).toBeNull();
  expect(sheet.getCell(anchor,25).value).toBe('Original remark\n[GROUP MOVE: OUTSHELL → TRIMS]');expect(sheet.getCell(anchor+1,25).value).toContain('OUTSHELL → TRIMS');
  const trims=find(sheet,1,'TRIMS SUBTOTAL'),outshell=find(sheet,1,'OUTSHELL SUBTOTAL'),total=find(sheet,1,'Total material cost');expect(result(sheet,trims,24)).toBe(5);expect(result(sheet,trims,8)).toBe(4);expect(result(sheet,outshell,24)).toBe(1);expect(result(sheet,total,24)).toBe(6);expect(result(sheet,total,8)).toBe(6);
  const detailed=buildComparisonWorkbook(state).worksheets[1];expect(result(sheet,total,24)).toBe(result(detailed,12,2));expect(result(sheet,total,8)).toBe(result(detailed,12,3));expect(sheet.getCell(total+4,8).value).toBe(12);expect(sheet.getCell(total+4,24).value).toBe(10);
 });
 it('expands a confirmed 1:N split without repeating the reference amount',()=>{
  const state=fixture();state.materialMatches[0].clusters=[{...state.materialMatches[0].clusters[0],referenceRowIds:['ref-a'],currentRowId:'new-a',currentRowIds:['new-a','new-only'],relationType:'one-to-many'},...state.styles[0].materials.slice(1).map(r=>({...state.materialMatches[0].clusters[1],id:r.id,referenceRowIds:[r.id]}))];
  const sheet=buildManagerWorkbook(state).worksheets[0],anchor=find(sheet,2,'new-a');expect(sheet.getCell(anchor+1,2).value).toBe('new-only');expect(sheet.getCell(anchor+1,24).value).toBeNull();expect(result(sheet,anchor,15)).toBe(4);for(let c=11;c<=14;c++)expect(sheet.getCell(anchor,c).value).toBeNull();
  expect(groupAmounts(state.materialMatches[0].clusters,state.styles[0].materials,state.styles[1].materials).TRIMS.comparison).toBe(6);const rematched=matchMaterials(state.styleMatches[0],state.styles,state.materialMatches[0]);expect(rematched.clusters.find(c=>c.id==='merged')?.currentRowIds).toEqual(['new-a','new-only']);
 });
 it('caches unchanged numeric rates as finite zero, including all-zero totals',async()=>{
  const state=fixture();state.materialMatches[0].clusters[0].referenceRowIds=['ref-a'];state.materialMatches[0].clusters[0].relationType='one-to-one';state.styles[0].materials.splice(1,1);state.styles[1].materials[0].usage=state.styles[0].materials[0].usage;const book=buildManagerWorkbook(state),sheet=book.worksheets[0],r=find(sheet,2,'new-a');expect(sheet.getCell(r,12).result).toBe(0);expect(sheet.getCell(r,14).result).toBe(0);sheet.eachRow(row=>row.eachCell(cell=>{if(typeof cell.result==='number')expect(Number.isFinite(cell.result)).toBe(true)}));
  for(const style of state.styles)for(const row of style.materials)row.extended=0;const zero=buildManagerWorkbook(state).worksheets[0],total=find(zero,1,'Total material cost');expect(zero.getCell(total,15).result).toBe(0);expect(zero.getCell(total+5,8).result).toBe(0);
 });
 it('preserves source summary remarks without inventing missing financial values',()=>{const state=fixture();state.styles[1].summary.remarks={laborCost:'Includes heat transfer, laser and packing cost'};const sheet=buildManagerWorkbook(state).worksheets[0],r=find(sheet,1,'Labor cost');expect(sheet.getCell(r,9).value).toBe(state.styles[1].summary.remarks.laborCost);expect(sheet.getCell(r,8).value).toBeNull();});
 it('keeps explicit move provenance across group resolution, persistence upgrade and auto rematching',()=>{
  const state=fixture(),c=resolveGroup(state.materialMatches[0].clusters[0],state.styles[0].materials,state.styles[1].materials);expect(c).toMatchObject({referenceOriginalGroup:'OUTSHELL',comparisonOriginalGroup:'TRIMS',effectiveGroup:'TRIMS',groupChangedFrom:'OUTSHELL',groupChangedTo:'TRIMS',groupChangeSource:'manual'});
  state.materialMatches[0].clusters[0]=c;const restored=upgradeComparison(JSON.parse(JSON.stringify(state))),rerun=matchMaterials(state.styleMatches[0],state.styles,restored.materialMatches[0]);expect(rerun.clusters.find(x=>x.id===c.id)).toMatchObject(c);expect(groupAmounts(rerun.clusters,state.styles[0].materials,state.styles[1].materials).TRIMS).toEqual({reference:5,comparison:4});
 });
 it('leaves unit differences, zero denominators, empty measures and list-only comparisons blank',()=>{
  const state=fixture();state.materialMatches[0].clusters[0].referenceRowIds=['ref-a'];state.styles[0].materials.splice(1,1);state.styles[0].materials[0].unit='PCS';state.materialMatches[0].clusters[0].relationType='one-to-one';state.styles[0].materials[1].cost=0;state.styles[0].materials[1].usage=undefined;
  const sheet=buildManagerWorkbook(state).worksheets[0],anchor=find(sheet,2,'new-a'),removed=find(sheet,18,'ref-only');for(let c=11;c<=14;c++)expect(sheet.getCell(anchor,c).value).toBeNull();expect(result(sheet,anchor,15)).toBe(2);expect(sheet.getCell(removed,12).value).toBeNull();expect(sheet.getCell(removed,13).value).toBeNull();
  state.styles[1].materials[0].group='SPECIAL PROCESS\n(LIST ONLY)';state.materialMatches[0].clusters[0].finalGroup='SPECIAL PROCESS\n(LIST ONLY)';const list=buildManagerWorkbook(state).worksheets[0],r=find(list,2,'new-a');for(let c=11;c<=15;c++)expect(list.getCell(r,c).value).toBeNull();
 });
 it('uses template styles, two real logos, A:Y print area and safe unique sheet names',()=>{
  const state=fixture();state.styles[1].styleName='a/'.repeat(25);state.styleMatches.push({...state.styleMatches[0],id:'other'});state.materialMatches.push({...state.materialMatches[0],styleMatchId:'other'});
  const sheets=buildManagerWorkbook(state).worksheets;expect(new Set(sheets.map(s=>s.name)).size).toBe(2);
  for(const sheet of sheets){expect(sheet.name.length).toBeLessThanOrEqual(31);expect(sheet.name).not.toMatch(/[\\/*?:[\]]/);expect(sheet.getImages()).toHaveLength(2);expect(sheet.columns.map(c=>c.width)).toEqual(MANAGER_WIDTHS);expect(sheet.views[0]).toMatchObject({state:'frozen',ySplit:5});expect(sheet.pageSetup).toMatchObject({orientation:'landscape',paperSize:9,fitToWidth:1,fitToHeight:0,printArea:`A1:Y${sheet.rowCount}`});expect(sheet.getCell('A5').fill).toMatchObject({fgColor:{argb:'FF99CCFF'}});expect(sheet.getCell('C2').fill).toMatchObject({fgColor:{argb:'FFC0C0C0'}});expect(sheet.getCell('E6').numFmt).toContain('[Red](');expect(sheet.getCell('B6').alignment.wrapText).toBe(true);expect(sheet.getCell('E6').alignment.wrapText).toBe(false);}
  expect(buildManagerWorkbook(state,{stylePairIds:['other']}).worksheets).toHaveLength(1);
 });
 it('exports one-sided styles, rejects repeated and missing source rows, and never derives FOB from ratios',()=>{
  const state=fixture();state.styleMatches[0].referenceId=undefined;state.materialMatches=[];const sheet=buildManagerWorkbook(state).worksheets[0];expect(sheet.getCell('S2').value).toBeNull();const total=find(sheet,1,'Total material cost');expect(sheet.getCell(total,24).value).toBeNull();
  const broken=fixture();broken.materialMatches[0].clusters[0].referenceRowIds.push('ref-a');expect(()=>buildManagerWorkbook(broken)).toThrow(/repeated/);broken.materialMatches[0].clusters=[];expect(()=>buildManagerWorkbook(broken)).toThrow(/incomplete/);
  state.styles[1].summary={materialToFobRatio:.5,finalFob:12,fobEvidence:'ratio-derived'};const ratio=buildManagerWorkbook(state).worksheets[0];expect(ratio.getCell(find(ratio,1,'FOB PRICE'),8).value).toBeNull();
 });
 it('uses the displayed final material total for PEE WEE and RAYCE internal-review ratios',()=>{
  const state=fixture(),reference=state.styles[0],current=state.styles[1];
  reference.materials=[material('ref',0)];current.materials=[material('pee',5.2438)];state.materialMatches[0].clusters=[{id:'pee',referenceRowIds:['ref'],currentRowId:'pee',relationType:'one-to-one',matchSource:'manual',finalGroup:'OUTSHELL',status:'MANUAL',confidence:1}];
  current.styleName='PEE WEE PANT';current.summary={finalFob:16.98,preliminaryMaterialCost:4.8948,fobEvidence:'explicit'};const pee=buildManagerWorkbook(state).worksheets[0],total=find(pee,1,'Total material cost'),fob=find(pee,1,'FOB PRICE'),review=fob+1;
  expect(pee.getCell(review,8).formula).toBe(`H${total}/H${fob}`);expect(pee.getCell(review,8).result).toBeCloseTo(5.2438/16.98,8);expect(pee.getCell(review+2,8).result).toBeCloseTo(.349,8);expect(pee.getCell(review+3,8).result).toBeCloseTo(.349/4.8948,8);
  current.styleName='RAYCE PANT';current.materials=[material('rayce',17.053)];state.materialMatches[0].clusters[0]={...state.materialMatches[0].clusters[0],currentRowId:'rayce'};current.summary={finalFob:36.63,preliminaryMaterialCost:undefined,fobEvidence:'explicit'};const rayce=buildManagerWorkbook(state).worksheets[0],rayceFob=find(rayce,1,'FOB PRICE'),rayceReview=rayceFob+1;
  expect(rayce.getCell(rayceReview,8).formula).toBe(`H${find(rayce,1,'Total material cost')}/H${rayceFob}`);expect(rayce.getCell(rayceReview,8).result).toBeCloseTo(17.053/36.63,8);expect(rayce.getCell(rayceReview+1,8).value).toBeNull();
 });
});
