import type {ComparisonState} from './types';
const DB='cbd-generator-local',STORE='work',KEY='comparison-latest';
const open=()=>new Promise<IDBDatabase>((resolve,reject)=>{const request=indexedDB.open(DB,1);request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(STORE))request.result.createObjectStore(STORE)};request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error)});
const done=<T,>(request:IDBRequest<T>)=>new Promise<T>((resolve,reject)=>{request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error)});
export async function loadComparison(){const db=await open();try{return await done(db.transaction(STORE).objectStore(STORE).get(KEY)) as ComparisonState|undefined}finally{db.close()}}
export async function saveComparison(state:ComparisonState){const db=await open();try{await done(db.transaction(STORE,'readwrite').objectStore(STORE).put({...state,version:1,savedAt:new Date().toISOString()},KEY))}finally{db.close()}}
export async function clearComparison(){const db=await open();try{await done(db.transaction(STORE,'readwrite').objectStore(STORE).delete(KEY))}finally{db.close()}}
