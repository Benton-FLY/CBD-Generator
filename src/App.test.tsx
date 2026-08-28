import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import App from './App';
import type { StyleData } from './types';

const style: StyleData = {
  id: 'lite',
  name: 'LITE PANT 사전원가',
  sourceFile: 'LITE PANT BOM.xlsx',
  sourceSheet: 'BOM',
  finalFob: 12.5,
  laborRemark: '',
  materials: [
    {id:'fabric',item:'FABRIC',width:'58',unit:'YD',group:'OUTSHELL',included:true,baseCost:48.98,adjustedCost:48.98,baseUsage:1.01,adjustedUsage:1.01,additionalLoss:.05,remark:'',split:false,sources:[]},
    {id:'zipper',item:'ZIPPER',width:'',unit:'EA',group:'TRIMS',included:true,baseCost:5.6,adjustedCost:5.6,baseUsage:1,adjustedUsage:1,additionalLoss:.05,remark:'',split:false,sources:[]},
  ],
};

afterEach(()=>vi.unstubAllGlobals());

describe('CBD workspace UI',()=>{
  it('renders multi-file BOM add, dynamic groups, redo, readonly FOB, and four-decimal values',()=>{
    vi.stubGlobal('localStorage',{getItem:(key:string)=>key==='cbd-generator-work-v1'?JSON.stringify({styles:[style],active:style.id,settings:{exchangeRate:900,defaultLoss:.05}}):null,setItem:()=>{}});
    const html=renderToStaticMarkup(<App/>);
    expect(html).toContain('BOM 추가');
    expect(html).toMatch(/type="file"[^>]*multiple=""[^>]*accept="\.xls,\.xlsx"/);
    expect(html).toContain('Redo');
    expect(html).toContain('CBD Group 보기');
    expect(html).toContain('OUTSHELL');
    expect(html).toContain('TRIMS');
    expect(html).toContain('$12.5000');
    expect(html).not.toContain('placeholder="직접 입력"');
    expect(html).not.toContain('Buyer-visible Target');
    expect(html).not.toContain('선택 자재 비례 배분');
    expect(html).toContain('value="48.9800"');
    expect(html).toContain('value="1.0100"');
    expect(html).toContain('$51.9433');
  });
});
