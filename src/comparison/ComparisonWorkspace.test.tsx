// @vitest-environment jsdom
import 'fake-indexeddb/auto';
import {afterEach,beforeEach,describe,expect,it} from 'vitest';
import {cleanup,render,screen,waitFor} from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import ComparisonWorkspace from './ComparisonWorkspace';
import {clearComparison,loadComparison,saveComparison} from './persistence';
import type {CbdStyle,ComparisonState} from './types';

const style=(id:string,side:'reference'|'current',name:string):CbdStyle=>({id,side,fileName:`${id}.xlsx`,sheetName:name,styleName:name,materials:[{id:`${side}::file::${name}::row-6::seq-1`,group:'OUTSHELL',material:`${name} MATERIAL`,size:'',unit:'YD',cost:1,usage:1,extended:1,remark:'',width:'',order:0}],summary:{totalMaterialCost:1},groupOrder:['OUTSHELL']});
const reference=style('r','reference','27 RAYCE JERSEY'),comparison=style('c','current','RAYCE JERSEY');
const state=(step:2|3):ComparisonState=>({version:2,referenceSeason:'27',currentSeason:'28',files:[],styles:[reference,comparison],styleMatches:[{id:'sm',referenceId:'r',currentId:'c',method:'Normalized',confidence:1,status:'Normalized'}],materialMatches:step===3?[{styleMatchId:'sm',clusters:[{id:'cluster',referenceRowIds:[reference.materials[0].id],currentRowId:comparison.materials[0].id,relationType:'one-to-one',matchSource:'auto',finalGroup:'OUTSHELL',status:'MATCH',confidence:1}]}]:[],step,activeMatchId:'sm'});

beforeEach(clearComparison);afterEach(async()=>{cleanup();await clearComparison()});
describe('CBD Comparison navigation and controls',()=>{
 it('keeps export with Undo and Redo in the sticky material toolbar and opens preview steps',async()=>{await saveComparison(state(3));const user=userEvent.setup();render(<ComparisonWorkspace onBack={()=>{}}/>);await screen.findByRole('heading',{name:'CBD Comparison'});const exportButton=screen.getByRole('button',{name:'Export Comparison'});expect(exportButton.closest('.matching-toolbar')).toBeTruthy();expect((exportButton as HTMLButtonElement).type).toBe('button');await user.click(screen.getByRole('button',{name:/4\. Comparison Preview/}));expect(screen.getByRole('heading',{name:'Comparison Preview'})).toBeTruthy();await user.click(screen.getByRole('button',{name:/5\. Export Comparison/}));expect(screen.getByText('Review the selected comparison before generating the Excel workbook.')).toBeTruthy()});
 it('records Exclude All as one atomic history action',async()=>{await saveComparison(state(2));const user=userEvent.setup();render(<ComparisonWorkspace onBack={()=>{}}/>);const master=await screen.findByRole('checkbox',{name:/Exclude All|Clear All Exclusions/}) as HTMLInputElement,initial=master.checked;await user.click(master);expect(master.checked).toBe(!initial);await waitFor(async()=>{const saved=await loadComparison();expect(saved?.history?.past).toHaveLength(1);expect(!!saved?.history?.past[0].styleMatches[0].excluded).toBe(initial)},{timeout:2000})});
});
