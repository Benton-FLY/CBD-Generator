import type { AppSettings, ClassificationDictionary, StyleData } from './types';
import type { FobRecord } from './fob';

const DB_NAME='cbd-generator-local';
const STORE_NAME='work';
const WORK_KEY='latest';
export const WORK_VERSION=1;

export interface SavedWork {
  version:number;
  styles:StyleData[];
  active:string;
  settings:AppSettings;
  dict:ClassificationDictionary;
  selections:Record<string,string[]>;
  groupFilters:Record<string,string>;
  /** Style-specific bulk-adjustment target. Optional for existing v1 saves. */
  scopes?:Record<string,string>;
  /** Style-specific bulk field, operation, and value. Optional for existing v1 saves. */
  bulkStates?:Record<string,{target:string;operation:string;value:number}>;
  fobs:FobRecord[];
  fobFile?:{name:string;size:number;lastModified:number;parsedCount:number};
  sidebarScrollTop?:number;
  scope:string;
  target:string;
  operation:string;
  bulkValue:number;
  savedAt:string;
}

const openDb=()=>new Promise<IDBDatabase>((resolve,reject)=>{
  const request=indexedDB.open(DB_NAME,1);
  request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(STORE_NAME))request.result.createObjectStore(STORE_NAME)};
  request.onsuccess=()=>resolve(request.result);
  request.onerror=()=>reject(request.error||new Error('IndexedDB를 열지 못했습니다.'));
});

const requestResult=<T,>(request:IDBRequest<T>)=>new Promise<T>((resolve,reject)=>{request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error||new Error('브라우저 저장소 작업에 실패했습니다.'))});

export async function loadSavedWork():Promise<SavedWork|null>{
  const db=await openDb();
  try{
    const value=await requestResult(db.transaction(STORE_NAME,'readonly').objectStore(STORE_NAME).get(WORK_KEY)) as SavedWork|undefined;
    return value?.version===WORK_VERSION?value:null;
  }finally{db.close()}
}

export async function saveWork(value:Omit<SavedWork,'version'|'savedAt'>):Promise<string>{
  const db=await openDb();
  const savedAt=new Date().toISOString();
  try{await requestResult(db.transaction(STORE_NAME,'readwrite').objectStore(STORE_NAME).put({...value,version:WORK_VERSION,savedAt},WORK_KEY));return savedAt}finally{db.close()}
}

export async function clearSavedWork():Promise<void>{
  const db=await openDb();
  try{await requestResult(db.transaction(STORE_NAME,'readwrite').objectStore(STORE_NAME).delete(WORK_KEY))}finally{db.close()}
}
