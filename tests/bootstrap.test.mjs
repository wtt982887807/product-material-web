import test from 'node:test';
import assert from 'node:assert/strict';
const module=await import('../src/bootstrap.js').catch(error=>{if(error.code==='ERR_MODULE_NOT_FOUND')return{};throw error;});
const start=options=>{
  assert.equal(typeof module.startApplication,'function','browser bootstrap entry point is implemented');
  return module.startApplication(options);
};
const moduleUrl='https://example.github.io/product-material-web/src/bootstrap.js';
function pageFixture({isolated=false,storage=new Map()}={}){
  const listeners=new Map(),documentListeners=new Map();
  const inputs=[{disabled:false,files:[]},{disabled:true,files:[]}];
  const hint={textContent:'original hint'};
  const worker={controller:null,
    addEventListener(type,listener){listeners.set(type,listener);},
    removeEventListener(type,listener){if(listeners.get(type)===listener)listeners.delete(type);},
    async register(url,options){this.registration={url,options};this.claim(url);return{};},
    claim(url='https://example.github.io/product-material-web/coi-serviceworker.js'){
      this.controller={scriptURL:url};listeners.get('controllerchange')?.();
    },
    get ready(){throw new Error('ready must not be awaited without a bound');},
  };
  const page={crossOriginIsolated:isolated,isSecureContext:true,navigator:{serviceWorker:worker},
    location:{origin:'https://example.github.io',reload(){page.reloads++;}},reloads:0,
    sessionStorage:{getItem:key=>storage.get(key)||null,setItem:(key,value)=>storage.set(key,value),removeItem:key=>storage.delete(key)},
    document:{querySelectorAll:()=>inputs,getElementById:()=>hint,
      addEventListener:(type,listener)=>documentListeners.set(type,listener),
      removeEventListener:(type,listener)=>{if(documentListeners.get(type)===listener)documentListeners.delete(type);},
    },
  };
  return{page,worker,inputs,hint,storage,listeners,documentListeners};
}

test('first bootstrap disables file selection, registers at project root and reloads only after its worker controls the page',async()=>{
  const f=pageFixture();let appLoads=0,resolveRegistration;
  f.worker.register=(url,options)=>{
    f.worker.registration={url,options};
    assert.ok(f.inputs.every(input=>input.disabled));
    return new Promise(resolve=>{resolveRegistration=resolve;});
  };
  const pending=start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:100});
  await Promise.resolve();
  assert.equal(appLoads,0);assert.equal(f.page.reloads,0);
  assert.equal(f.worker.registration.url,'https://example.github.io/product-material-web/coi-serviceworker.js');
  assert.deepEqual(f.worker.registration.options,{scope:'https://example.github.io/product-material-web/',updateViaCache:'none'});
  let prevented=false;
  f.documentListeners.get('drop')({preventDefault(){prevented=true;},stopImmediatePropagation(){}});
  assert.ok(prevented);
  resolveRegistration({});f.worker.claim();await pending;
  assert.equal(f.page.reloads,1);assert.equal(appLoads,0);
  assert.ok(f.inputs.every(input=>input.disabled));
});

test('after one failed isolation reload the app still starts with original disabled states and no second reload',async()=>{
  const storage=new Map();
  const first=pageFixture({storage});
  await start({window:first.page,moduleUrl,loadApp:async()=>{},timeoutMs:20});
  const next=pageFixture({storage});let appLoads=0;
  await start({window:next.page,moduleUrl,loadApp:async()=>{appLoads++;assert.deepEqual(next.inputs.map(input=>input.disabled),[false,true]);},timeoutMs:20});
  assert.equal(next.page.reloads,0);assert.equal(appLoads,1);
  assert.equal(next.hint.textContent,'original hint');assert.equal(next.documentListeners.size,0);
});

test('an already isolated page starts without a reload and still checks its worker without the HTTP cache',async()=>{
  const f=pageFixture({isolated:true});let appLoads=0;
  await start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:20});
  assert.equal(f.page.reloads,0);assert.equal(appLoads,1);
  assert.equal(f.worker.registration.options.updateViaCache,'none');
  assert.deepEqual(f.inputs.map(input=>input.disabled),[false,true]);
});

test('unsupported or insecure browsers continue using the ordinary app',async()=>{
  for(const alter of [f=>{delete f.page.navigator.serviceWorker;},f=>{f.page.isSecureContext=false;},f=>{delete f.page.crossOriginIsolated;}]){
    const f=pageFixture();alter(f);let appLoads=0;
    await start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:20});
    assert.equal(appLoads,1);assert.equal(f.page.reloads,0);assert.equal(f.worker.registration,undefined);
  }
});

test('registration rejection, hung registration and absent controller all have a bounded fallback',async()=>{
  for(const register of [async()=>{throw new Error('denied');},()=>new Promise(()=>{}),async()=>({})]){
    const f=pageFixture();f.worker.register=register;let appLoads=0;
    await start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:10});
    assert.equal(appLoads,1);assert.equal(f.page.reloads,0);
    assert.deepEqual(f.inputs.map(input=>input.disabled),[false,true]);assert.equal(f.listeners.size,0);
    f.worker.claim();assert.equal(f.page.reloads,0,'a late controller cannot discard selected files');
  }
});

test('an unrelated controller cannot trigger a reload, and localhost uses its root scope',async()=>{
  const f=pageFixture();f.worker.register=async()=>{f.worker.claim('https://example.github.io/other.js');return{};};let appLoads=0;
  await start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:10});
  assert.equal(appLoads,1);assert.equal(f.page.reloads,0);
  const local=pageFixture();local.page.location.origin='http://localhost:8766';
  await start({window:local.page,moduleUrl:'http://localhost:8766/src/bootstrap.js',loadApp:async()=>{},timeoutMs:10});
  assert.equal(local.worker.registration.url,'http://localhost:8766/coi-serviceworker.js');
  assert.equal(local.worker.registration.options.scope,'http://localhost:8766/');
});

test('blocked session storage or already selected files skip auto reload to protect current work',async()=>{
  for(const alter of [f=>{f.page.sessionStorage.setItem=()=>{throw new Error('blocked');};},f=>{f.inputs[0].files=[{name:'selected.docx'}];}]){
    const f=pageFixture();alter(f);let appLoads=0;
    await start({window:f.page,moduleUrl,loadApp:async()=>{appLoads++;},timeoutMs:20});
    assert.equal(appLoads,1);assert.equal(f.page.reloads,0);
  }
});
