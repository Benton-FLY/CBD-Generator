import { useId, useRef, useState, type DragEvent } from 'react';
import { FileSpreadsheet } from 'lucide-react';
import { BOM_ACCEPT, isExcelFile } from './BomUploader';

export default function FobUploader({onFile,fileName,summary}:{onFile:(file?:File,error?:string)=>void;fileName?:string;summary?:string}){
  const id=`fob-${useId().replace(/:/g,'')}`,input=useRef<HTMLInputElement>(null),[dragging,setDragging]=useState(false);
  const accept=(files:File[])=>{const file=files[0];if(!file)return;if(!isExcelFile(file))return onFile(undefined,'Excel 파일(.xls 또는 .xlsx)만 업로드할 수 있습니다.');onFile(file)};
  const drop=(event:DragEvent<HTMLDivElement>)=>{event.preventDefault();setDragging(false);accept(Array.from(event.dataTransfer.files||[]))};
  return <div className={`fob-uploader ${dragging?'dragging':''}`} onDragOver={e=>{e.preventDefault();e.dataTransfer.dropEffect='copy'}} onDragEnter={e=>{e.preventDefault();setDragging(true)}} onDragLeave={e=>{e.preventDefault();setDragging(false)}} onDrop={drop}>
    <label htmlFor={id}><FileSpreadsheet size={16}/><b>FOB 파일</b></label><button type="button" onClick={()=>input.current?.click()}>파일 선택</button>
    <input ref={input} id={id} className="accessible-file-input" aria-label="FOB Excel 파일" type="file" accept={BOM_ACCEPT} onChange={e=>{accept(Array.from(e.currentTarget.files||[]));e.currentTarget.value=''}}/>
    {fileName&&<small title={fileName}>{fileName}</small>}{summary&&<em>{summary}</em>}
  </div>;
}
