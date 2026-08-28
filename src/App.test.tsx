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

  it('cancels reset without data loss and clears UI plus IndexedDB after confirmation',async()=>{
    await seed();const user=userEvent.setup();render(<App/>);await screen.findByLabelText('Final FOB');
    vi.spyOn(window,'confirm').mockReturnValueOnce(false);await user.click(screen.getByRole('button',{name:/전체 초기화/}));
    expect(screen.getByText('STYLE ONE 사전원가')).toBeTruthy();expect(await loadSavedWork()).not.toBeNull();
    vi.spyOn(window,'confirm').mockReturnValueOnce(true);await user.click(screen.getByRole('button',{name:/전체 초기화/}));
    expect(await screen.findByText('새 작업 시작')).toBeTruthy();expect(await loadSavedWork()).toBeNull();
  });
});
