import test from 'node:test';
import assert from 'node:assert/strict';
const api=await import('../src/manager-store.js').catch(error=>{if(error.code==='ERR_MODULE_NOT_FOUND')return{};throw error;});
const createStore=options=>{assert.equal(typeof api.createManagerStore,'function','manager store is implemented');return api.createManagerStore(options);};
const record=(name='经理.xlsx')=>({schemaVersion:1,name,sha256:'a'.repeat(64),uploadedAt:'2026-10-29T01:02:03.000Z',bytes:new Uint8Array([80,75,3,4]),profile:{schemaVersion:1,name,sha256:'a'.repeat(64),records:[{id:'manager-1-row-2',name:'甲',biography:'简介',workStart:'',investmentStart:'',sheet:'经理',row:2}],warnings:[]}});
const later=fn=>setTimeout(fn,0);
// IndexedDB is absent in Node: model its async structured-clone boundary and
// transaction commit separately, so successful put alone cannot satisfy save.
function fakeIndexedDB(){
 let exists=false;const stores=new Map(),control={failCommit:false};
 const db={objectStoreNames:{contains:name=>stores.has(name)},close(){},createObjectStore(name){stores.set(name,new Map());},transaction(name,mode){
  const staged=structuredClone(stores.get(name));const tx={error:null,aborted:false,abort(){this.aborted=true;later(()=>this.onabort?.());}};
  const request=action=>{const req={};later(()=>{if(tx.aborted)return;req.result=action();req.onsuccess?.();later(()=>{if(tx.aborted)return;if(control.failCommit&&mode==='readwrite'){control.failCommit=false;tx.error=new DOMException('Full','QuotaExceededError');tx.abort();}else{if(mode==='readwrite')stores.set(name,staged);tx.oncomplete?.();}});});return req;};
  tx.objectStore=()=>({keyPath:null,autoIncrement:false,get:key=>request(()=>structuredClone(staged.get(key))),put:(value,key)=>{const copy=structuredClone(value);return request(()=>{staged.set(key,copy);return key;});},delete:key=>request(()=>staged.delete(key))});return tx;
 }};
 return{control,corrupt(fn){for(const store of stores.values())for(const[key,value]of store)store.set(key,fn(value));},open(){const req={result:db,transaction:{abort(){}}};later(()=>{if(!exists){exists=true;req.onupgradeneeded?.({oldVersion:0});}req.onsuccess?.();});return req;}};
}
test('saved manager bytes and parsed records survive a fresh store instance and clear',async()=>{
 const indexedDB=fakeIndexedDB(),store=createStore({indexedDB});assert.equal(await store.load(),null);
 const original=record(),pending=store.save(original);original.bytes[0]=0;original.profile.records[0].biography='changed';await pending;
 assert.deepEqual(await createStore({indexedDB}).get(),record());
 const restored=await store.load();restored.bytes[0]=1;restored.profile.records[0].biography='changed';assert.deepEqual(await store.load(),record());
 await store.put(record('新经理.xlsx'));assert.equal((await store.get()).name,'新经理.xlsx');await store.clear();assert.equal(await createStore({indexedDB}).load(),null);
});
test('a failed replacement or clear preserves the previous committed workbook',async()=>{
 const indexedDB=fakeIndexedDB(),store=createStore({indexedDB});await store.save(record());
 indexedDB.control.failCommit=true;await assert.rejects(store.save(record('新经理.xlsx')),/存储|保存|空间/);assert.deepEqual(await store.load(),record());
 indexedDB.control.failCommit=true;await assert.rejects(store.clear(),/存储|清除|空间/);assert.deepEqual(await store.load(),record());
});
test('invalid and corrupt caches fail explicitly without replacing a usable workbook',async()=>{
 const indexedDB=fakeIndexedDB(),store=createStore({indexedDB});await store.save(record());
 for(const item of [{...record(),name:'经理.docx'},{...record(),bytes:new Uint8Array()},{...record(),profile:null},{...record(),profile:{...record().profile,records:[]}},{...record(),sha256:'bad'}])await assert.rejects(store.save(item),/基金经理|格式|文件/);
 assert.deepEqual(await store.load(),record());indexedDB.corrupt(r=>({...r,bytes:null}));await assert.rejects(store.load(),/缓存/);await store.clear();assert.equal(await store.load(),null);
});
test('unavailable local storage reports a readable failure',async()=>{
 for(const indexedDB of [null,{}, {open(){throw new DOMException('Denied','SecurityError');}}])await assert.rejects(createStore({indexedDB}).load(),/浏览器|存储|缓存/);
});
