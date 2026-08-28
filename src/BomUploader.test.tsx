// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { cleanup, fireEvent, render, screen } from '@testing-library/react';
import BomUploader, { BOM_ACCEPT, isExcelFile } from './BomUploader';

afterEach(cleanup);

describe('BOM uploader',()=>{
  it('uses an explicitly associated, enabled standard file input',()=>{
    render(<BomUploader onFiles={()=>{}}/>);const input=screen.getByLabelText('BOM Excel 파일') as HTMLInputElement;
    expect(input.disabled).toBe(false);expect(input.multiple).toBe(true);expect(input.accept).toBe(BOM_ACCEPT);
    expect(document.querySelector(`label[for="${input.id}"]`)).not.toBeNull();expect(getComputedStyle(input).pointerEvents).not.toBe('none');
  });
  it('accepts Excel names case-insensitively even without a MIME type',()=>{
    for(const name of ['27 LITE PANT BOM.xls','28 LITE PANT 사전원가.xlsx','TEST BOM.XLSX'])expect(isExcelFile(new File([],name,{type:''}))).toBe(true);
    expect(isExcelFile(new File([],'not-a-bom.pdf',{type:'application/pdf'}))).toBe(false);
  });
  it('handles drag events, item fallback, and resets the drag state',()=>{
    const received:File[][]=[];render(<BomUploader onFiles={files=>received.push(files)}/>);const box=screen.getByTestId('start-bom-uploader');const file=new File([],'한글 BOM.XLSX',{type:''});
    const transfer={files:[],items:[{kind:'file',getAsFile:()=>file}],dropEffect:'none'} as unknown as DataTransfer;
    fireEvent.dragEnter(box,{dataTransfer:transfer});expect(box.classList.contains('dragging')).toBe(true);expect(transfer.dropEffect).toBe('copy');
    fireEvent.dragOver(box,{dataTransfer:transfer});fireEvent.drop(box,{dataTransfer:transfer});expect(received[0]).toEqual([file]);expect(box.classList.contains('dragging')).toBe(false);
  });
});
