import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
const source=await readFile(new URL('../coi-serviceworker.js',import.meta.url),'utf8').catch(error=>{if(error.code==='ENOENT')return'';throw error;});
function worker(fetchImpl=async()=>new Response('original body',{headers:{'Content-Type':'text/plain','X-Original':'keep'}})){
  const handlers=new Map();let skips=0,claims=0;
  const scope={location:{origin:'https://example.github.io'},
    addEventListener:(type,handler)=>handlers.set(type,handler),
    skipWaiting:async()=>{skips++;},clients:{claim:async()=>{claims++;}},
  };
  vm.runInNewContext(source,{self:scope,fetch:fetchImpl,URL,Headers,Response});
  const dispatch=(type,event)=>{assert.equal(typeof handlers.get(type),'function',`worker handles ${type}`);handlers.get(type)(event);};
  return{dispatch,counts:()=>({skips,claims})};
}

test('install and activate immediately adopt the worker without reading or writing caches',async()=>{
  const f=worker();const pending=[];
  f.dispatch('install',{waitUntil:promise=>pending.push(promise)});
  f.dispatch('activate',{waitUntil:promise=>pending.push(promise)});
  await Promise.all(pending);assert.deepEqual(f.counts(),{skips:1,claims:1});
});

test('same-origin content gains isolation headers while body, response status and original headers are retained',async()=>{
  let captured;
  const request=new Request('https://example.github.io/product-material-web/index.html');
  const f=worker(async value=>{assert.equal(value,request);return new Response('document body',{status:203,statusText:'Non-Authoritative',headers:{'Content-Type':'text/html','X-Original':'retained'}});});
  f.dispatch('fetch',{request,respondWith:promise=>{captured=promise;}});
  const response=await captured;
  assert.equal(response.headers.get('Cross-Origin-Opener-Policy'),'same-origin');
  assert.equal(response.headers.get('Cross-Origin-Embedder-Policy'),'require-corp');
  assert.equal(response.headers.get('X-Original'),'retained');
  assert.equal(response.headers.get('Content-Type'),'text/html');
  assert.equal(response.status,203);assert.equal(response.statusText,'Non-Authoritative');assert.equal(await response.text(),'document body');
});

test('cross-origin CDN requests, non-GET requests and unsupported cache modes are left to the browser',()=>{
  const f=worker(()=>{throw new Error('should not fetch intercepted request');});
  for(const request of [new Request('https://cdn.example/libreoffice.wasm'),new Request('https://example.github.io/upload',{method:'POST',body:'private'}),{url:'https://example.github.io/file',method:'GET',cache:'only-if-cached',mode:'no-cors'}]){
    let handled=false;f.dispatch('fetch',{request,respondWith(){handled=true;}});assert.equal(handled,false);
  }
});

test('opaque and cross-origin redirected responses pass through unchanged',async()=>{
  for(const original of [{status:0,type:'opaque'},{status:200,url:'https://cdn.example/redirected',type:'cors'}]){
    const f=worker(async()=>original);let pending;
    f.dispatch('fetch',{request:new Request('https://example.github.io/redirect'),respondWith:promise=>{pending=promise;}});
    assert.equal(await pending,original);
  }
});

test('network failures propagate and empty responses stay valid',async()=>{
  const failure=new Error('offline');const f=worker(async()=>{throw failure;});let pending;
  f.dispatch('fetch',{request:new Request('https://example.github.io/index.html'),respondWith:promise=>{pending=promise;}});
  await assert.rejects(pending,error=>error===failure);
  const empty=worker(async()=>new Response(null,{status:204}));
  empty.dispatch('fetch',{request:new Request('https://example.github.io/empty'),respondWith:promise=>{pending=promise;}});
  const response=await pending;assert.equal(response.status,204);assert.equal(response.body,null);
});
