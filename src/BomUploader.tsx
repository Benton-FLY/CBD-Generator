import { useId, useRef, useState, type DragEvent, type KeyboardEvent, type MouseEvent } from 'react';
import { Upload } from 'lucide-react';

export const BOM_ACCEPT='.xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
export const isExcelFile=(file:Pick<File,'name'>)=>/\.(xls|xlsx)$/i.test(file.name);

type Props={onFiles:(files:File[])=>void;compact?:boolean};

export default function BomUploader({onFiles,compact=false}:Props){
  const generatedId=useId(),inputId=`bom-files-${generatedId.replace(/:/g,'')}`;
  const inputRef=useRef<HTMLInputElement>(null),dragDepth=useRef(0);
  const [dragging,setDragging]=useState(false);
  const openPicker=()=>inputRef.current?.click();
  const keyboardOpen=(event:KeyboardEvent<HTMLElement>)=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();openPicker()}};
  const dragStart=(event:DragEvent<HTMLDivElement>)=>{event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect='copy';dragDepth.current+=1;setDragging(true)};
  const dragOver=(event:DragEvent<HTMLDivElement>)=>{event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect='copy'};
  const dragLeave=(event:DragEvent<HTMLDivElement>)=>{event.preventDefault();event.stopPropagation();dragDepth.current=Math.max(0,dragDepth.current-1);if(!dragDepth.current)setDragging(false)};
  const drop=(event:DragEvent<HTMLDivElement>)=>{event.preventDefault();event.stopPropagation();event.dataTransfer.dropEffect='copy';dragDepth.current=0;setDragging(false);let files=Array.from(event.dataTransfer.files||[]);if(!files.length)files=Array.from(event.dataTransfer.items||[]).map(item=>item.kind==='file'?item.getAsFile():null).filter((file):file is File=>Boolean(file));onFiles(files)};
  const containerClick=(event:MouseEvent<HTMLDivElement>)=>{if(event.target===event.currentTarget)openPicker()};
  return <div className={`bom-uploader ${compact?'compact':''} ${dragging?'dragging':''}`} data-testid={compact?'workspace-bom-uploader':'start-bom-uploader'} onClick={containerClick} onKeyDown={keyboardOpen} onDragEnter={dragStart} onDragOver={dragOver} onDragLeave={dragLeave} onDrop={drop}>
    <label htmlFor={inputId} tabIndex={0} onKeyDown={keyboardOpen}>
      <Upload size={compact?16:24}/><b>{compact?'BOM 추가':'BOM Excel 파일을 클릭하거나 이곳에 끌어다 놓으세요.'}</b>{!compact&&<span>.xls / .xlsx · 여러 파일 동시 선택 가능</span>}
    </label>
    <button type="button" onClick={event=>{event.stopPropagation();openPicker()}}>{compact?'파일 선택':'파일 찾아보기'}</button>
    <input ref={inputRef} id={inputId} className="accessible-file-input" aria-label={compact?'추가 BOM Excel 파일':'BOM Excel 파일'} type="file" accept={BOM_ACCEPT} multiple onChange={event=>{const files=Array.from(event.currentTarget.files||[]);event.currentTarget.value='';onFiles(files)}}/>
  </div>;
}
