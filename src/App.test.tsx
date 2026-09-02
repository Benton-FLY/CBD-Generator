// @vitest-environment jsdom
import 'fake-indexeddb/auto';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { cleanup, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import App from './App';
import { clearSavedWork, loadSavedWork, saveWork } from './persistence';
import type { CbdGroup, Material, StyleData } from './types';

const material=(id:string,group:CbdGroup,cost:number):Material=>({id,item:id.toUpperCase(),width:'58',unit:'YD',group,included:true,baseCost:cost,adjustedCost:cost,baseUsage:1,adjustedUsage:1,additionalLoss:.05,remark:'',split:false,sources:[]});
const styles:StyleData[]=[
  {id:'style-one',name:'STYLE ONE 사전원가',sourceFile:'one.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('trim-1','TRIMS',1),material('trim-2','TRIMS',2),material('trim-3','TRIMS',3),material('shell-1','OUTSHELL',4),material('shell-2','OUTSHELL',5),material('shell-hidden','OUTSHELL',6)]},
  {id:'style-two',name:'STYLE TWO 사전원가',sourceFile:'two.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[material('other','TRIMS',9)]},
];
const seed=()=>saveWork({styles,active:'style-one',settings:{exchangeRate:900,defaultLoss:.05},dict:{},selections:{},groupFilters:{},fobs:[],scope:'all',target:'cost',operation:'base-percent',bulkValue:5});
const row=(name:string)=>screen.getByText(name).closest('tr') as HTMLTableRowElement;
const selectionBox=(name:string)=>within(row(name)).getAllByRole('checkbox')[0];
const adjustedCost=(name:string)=>within(row(name)).getAllByRole('spinbutton')[0] as HTMLInputElement;

beforeEach(async()=>{await clearSavedWork();localStorage.clear();vi.restoreAllMocks()});
afterEach(()=>cleanup());

describe('CBD workspace UI',()=>{
  it('restores synchronized bulk and material group state per style across repeated switches',async()=>{
    await seed();const user=userEvent.setup();render(<App/>);await screen.findByLabelText('Final FOB');
    const bulk=screen.getByLabelText('일괄 조정 대상') as HTMLSelectElement;
    const filter=screen.getByLabelText('CBD Group 보기') as HTMLSelectElement;
    const field=screen.getByLabelText('일괄 조정 필드') as HTMLSelectElement;
    const operation=screen.getByLabelText('일괄 조정 방식') as HTMLSelectElement;
    await user.selectOptions(field,'usage');await user.selectOptions(operation,'set');
    await user.selectOptions(bulk,'TRIMS');
    expect(filter.value).toBe('TRIMS');expect(screen.queryByText('SHELL-1')).toBeNull();expect(screen.getByText('TRIM-1')).toBeTruthy();
    await user.click(screen.getByRole('button',{name:/STYLE TWO/}));
    expect(bulk.value).toBe('all');expect(filter.value).toBe('all');expect(field.value).toBe('cost');expect(operation.value).toBe('base-percent');expect(screen.getByText('OTHER')).toBeTruthy();
    await user.selectOptions(bulk,'TRIMS');
    await user.click(screen.getByRole('button',{name:/STYLE ONE/}));
    expect(bulk.value).toBe('TRIMS');expect(filter.value).toBe('TRIMS');expect(field.value).toBe('usage');expect(operation.value).toBe('set');expect(screen.queryByText('SHELL-1')).toBeNull();
    await user.click(screen.getByRole('button',{name:/STYLE TWO/}));
    expect(bulk.value).toBe('TRIMS');expect(filter.value).toBe('TRIMS');expect(screen.getByText('OTHER')).toBeTruthy();
    await user.click(screen.getByRole('button',{name:/STYLE ONE/}));
    expect(bulk.value).toBe('TRIMS');expect(filter.value).toBe('TRIMS');
  });

  it('accepts a decimal FINAL FOB, formats it to four decimals, and supports undo/redo',async()=>{
    await seed();const user=userEvent.setup();render(<App/>);
    const fob=await screen.findByLabelText('Final FOB') as HTMLInputElement;
    expect(fob.type).toBe('text');expect(fob.inputMode).toBe('decimal');
    await user.click(fob);await user.paste('48.98');await user.tab();
    expect(fob.value).toBe('48.9800');expect(screen.getByText(/차액 \$26\.9300/)).toBeTruthy();
    await user.click(screen.getByRole('button',{name:/Undo/}));expect(fob.value).toBe('');
    await user.click(screen.getByRole('button',{name:/Redo/}));expect(fob.value).toBe('48.9800');
  });

  it('keeps cross-group selections and applies a two-decimal bulk value to only those rows',async()=>{
    await seed();const user=userEvent.setup();render(<App/>);await screen.findByLabelText('Final FOB');
    const group=screen.getByLabelText('CBD Group 보기');
    await user.selectOptions(group,'TRIMS');for(const name of ['TRIM-1','TRIM-2','TRIM-3'])await user.click(selectionBox(name));
    await user.selectOptions(group,'OUTSHELL');for(const name of ['SHELL-1','SHELL-2'])await user.click(selectionBox(name));
    expect(screen.getByRole('option',{name:'선택 항목 (5)'})).toBeTruthy();
    await user.selectOptions(screen.getByLabelText('일괄 조정 대상'),'selected');
    for(const name of ['TRIM-1','TRIM-2','TRIM-3','SHELL-1','SHELL-2'])expect(row(name)).toBeTruthy();
    expect(screen.queryByText('SHELL-HIDDEN')).toBeNull();expect((screen.getByLabelText('일괄 조정 값') as HTMLInputElement).value).toBe('5.00');
    await user.selectOptions(screen.getByLabelText('일괄 조정 방식'),'set');await user.click(screen.getByRole('button',{name:'적용'}));
    expect(adjustedCost('TRIM-1').value).toBe('5.0000');expect(adjustedCost('SHELL-2').value).toBe('5.0000');
    await user.click(selectionBox('TRIM-1'));expect(screen.queryByText('TRIM-1')).toBeNull();
    await user.selectOptions(screen.getByLabelText('일괄 조정 대상'),'all');expect(adjustedCost('SHELL-HIDDEN').value).toBe('6.0000');
  });

  it('auto-saves two styles and restores edits and selections after remount',async()=>{
    await seed();const user=userEvent.setup();const first=render(<App/>);const fob=await screen.findByLabelText('Final FOB') as HTMLInputElement;
    await user.type(fob,'48.98');await user.keyboard('{Enter}');await user.selectOptions(screen.getByLabelText('CBD Group 보기'),'TRIMS');await user.click(selectionBox('TRIM-1'));
    await waitFor(async()=>expect((await loadSavedWork())?.styles[0].finalFob).toBe(48.98),{timeout:2500});
    first.unmount();render(<App/>);expect((await screen.findByLabelText('Final FOB') as HTMLInputElement).value).toBe('48.9800');
    expect(screen.getByText('STYLE TWO 사전원가')).toBeTruthy();expect(screen.getByRole('option',{name:'선택 항목 (1)'})).toBeTruthy();expect(await screen.findByText(/자동 저장됨/)).toBeTruthy();
  });

  it('updates the persisted F-16 CHICRON auto result on restore and protects a manual TRIMS override',async()=>{
    const item='CHICRON TOUCH (E-SUEDE) 0.6MM (8G060-B000D)',source={id:'f16-row',sourceFile:'28 F-16 GLOVE.xls',sourceSheet:'BOM',sourceRow:12,structure:'PALM',materialType:'자재',sequence:'',itemNo:'8G060-B000D',item,width:'',color:'BLACK',unit:'YD',usage:1,currency:'USD',convertedPrice:1,materialCostAdjustment:0,specialFlag:'',remark:''};
    const chicron:Material={...material('f16-chicron','TRIMS',1),item,sources:[source]};
    const f16:StyleData={id:'f16',name:'F-16 GLOVE',sourceFile:'28 F-16 GLOVE.xls',sourceSheet:'BOM',laborRemark:'',materials:[chicron]};
    const saveF16=(dict:Record<string,CbdGroup>)=>saveWork({styles:[f16],active:'f16',settings:{exchangeRate:900,defaultLoss:.05},dict,selections:{},groupFilters:{},fobs:[],scope:'all',target:'cost',operation:'base-percent',bulkValue:5});
    await saveF16({});const first=render(<App/>);const group=within((await screen.findByText(item)).closest('tr')!).getByRole('combobox') as HTMLSelectElement;
    expect(group.value).toBe('OUTSHELL');first.unmount();await clearSavedWork();
    await saveF16({'ITEM:CHICRON TOUCH (E-SUEDE) 0.6MM (8G060-B000D)':'TRIMS'});render(<App/>);const manual=within((await screen.findByText(item)).closest('tr')!).getByRole('combobox') as HTMLSelectElement;
    expect(manual.value).toBe('TRIMS');
  });

  it('resets current or all styles atomically while preserving non-adjustment edits and auto-saving',async()=>{
    const adjusted=(id:string,group:CbdGroup,cost:number,patch:Partial<Material>):Material=>({...material(id,group,cost),baseLoss:.05,...patch});
    const resetStyles:StyleData[]=[
      {id:'style-one',name:'F-16 GLOVE',sourceFile:'one.xlsx',sourceSheet:'BOM',laborRemark:'keep labor',finalFob:48.98,materials:[
        adjusted('shell-item','OUTSHELL',4,{adjustedCost:4.2,remark:'keep remark'}),
        adjusted('trim-item','TRIMS',2,{adjustedUsage:1.5,included:false}),
        adjusted('label-item','LABEL & PACKAGING',1,{additionalLoss:.12,group:'LABEL & PACKAGING'}),
      ]},
      {id:'style-two',name:'OTHER STYLE',sourceFile:'two.xlsx',sourceSheet:'BOM',laborRemark:'',materials:[adjusted('other','TRIMS',9,{adjustedCost:10})]},
    ];
    await saveWork({styles:resetStyles,active:'style-one',settings:{exchangeRate:900,defaultLoss:.05},dict:{},selections:{},groupFilters:{},fobs:[],scope:'all',target:'cost',operation:'base-percent',bulkValue:5});
    const user=userEvent.setup();const view=render(<App/>);await screen.findByLabelText('Final FOB');
    await user.click(screen.getByRole('button',{name:'일괄조정 전체 초기화'}));
    expect((screen.getByRole('radio',{name:'현재 스타일 전체 그룹'}) as HTMLInputElement).checked).toBe(true);
    expect(screen.getByText(/F-16 GLOVE의 모든 CBD Group/)).toBeTruthy();
    await user.click(screen.getByRole('button',{name:'초기화 실행'}));
    expect(await screen.findByText(/F-16 GLOVE의 3개 그룹, 3개 조정값/)).toBeTruthy();
    expect(adjustedCost('SHELL-ITEM').value).toBe('4.0000');
    expect((within(row('TRIM-ITEM')).getAllByRole('spinbutton')[1] as HTMLInputElement).value).toBe('1.0000');
    expect((within(row('LABEL-ITEM')).getAllByRole('spinbutton')[2] as HTMLInputElement).value).toBe('5.0000');
    expect((within(row('SHELL-ITEM')).getByRole('textbox') as HTMLTextAreaElement).value).toBe('keep remark');
    expect((within(row('TRIM-ITEM')).getAllByRole('checkbox')[1] as HTMLInputElement).checked).toBe(false);
    expect((screen.getByLabelText('Final FOB') as HTMLInputElement).value).toBe('48.9800');
    await user.click(screen.getByRole('button',{name:/Undo/}));expect(adjustedCost('SHELL-ITEM').value).toBe('4.2000');
    await user.click(screen.getByRole('button',{name:/Redo/}));expect(adjustedCost('SHELL-ITEM').value).toBe('4.0000');
    await user.click(screen.getByRole('button',{name:/OTHER STYLE/}));expect(adjustedCost('OTHER').value).toBe('10.0000');
    await user.click(screen.getByRole('button',{name:'일괄조정 전체 초기화'}));await user.click(screen.getByRole('radio',{name:'모든 스타일 전체 그룹'}));
    expect(screen.getByText(/현재 등록된 2개 스타일/)).toBeTruthy();await user.click(screen.getByRole('button',{name:'초기화 실행'}));
    expect(await screen.findByText(/2개 스타일의 1개 그룹, 1개 조정값/)).toBeTruthy();expect(adjustedCost('OTHER').value).toBe('9.0000');
    await user.click(screen.getByRole('button',{name:/Undo/}));expect(adjustedCost('OTHER').value).toBe('10.0000');
    await user.click(screen.getByRole('button',{name:/Redo/}));expect(adjustedCost('OTHER').value).toBe('9.0000');
    await waitFor(async()=>expect((await loadSavedWork())?.styles[1].materials[0].adjustedCost).toBe(9),{timeout:2500});
    view.unmount();render(<App/>);await screen.findByLabelText('Final FOB');await user.click(screen.getByRole('button',{name:/OTHER STYLE/}));expect(adjustedCost('OTHER').value).toBe('9.0000');
    await user.click(screen.getByRole('button',{name:'일괄조정 전체 초기화'}));await user.click(screen.getByRole('button',{name:'초기화 실행'}));expect(await screen.findByText('초기화할 조정값이 없습니다.')).toBeTruthy();
  });

  it('cancels reset without data loss and clears UI plus IndexedDB after confirmation',async()=>{
    await seed();const user=userEvent.setup();render(<App/>);await screen.findByLabelText('Final FOB');
    vi.spyOn(window,'confirm').mockReturnValueOnce(false);await user.click(screen.getByRole('button',{name:'전체 초기화'}));
    expect(screen.getByText('STYLE ONE 사전원가')).toBeTruthy();expect(await loadSavedWork()).not.toBeNull();
    vi.spyOn(window,'confirm').mockReturnValueOnce(true);await user.click(screen.getByRole('button',{name:'전체 초기화'}));
    expect(await screen.findByText('새 작업 시작')).toBeTruthy();expect(await loadSavedWork()).toBeNull();
  });
});
