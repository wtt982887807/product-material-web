async function prepareIsolation(page,moduleUrl,timeoutMs,inputs){
  const serviceWorker=page.navigator?.serviceWorker;
  if(!page.isSecureContext||!serviceWorker||typeof page.crossOriginIsolated!=='boolean')return false;
  const workerUrl=new URL('../coi-serviceworker.js',moduleUrl).href;
  const scope=new URL('../',moduleUrl).href;
  if(new URL(workerUrl).origin!==page.location.origin)return false;
  const register=()=>serviceWorker.register(workerUrl,{scope,updateViaCache:'none'});
  // Check for worker updates without delaying an already isolated application.
  if(page.crossOriginIsolated){try{Promise.resolve(register()).catch(()=>{});}catch{}return false;}
  let timer,checkController;
  try{
    const controlled=new Promise(resolve=>{
      checkController=()=>{if(serviceWorker.controller?.scriptURL===workerUrl)resolve();};
      serviceWorker.addEventListener('controllerchange',checkController);
      checkController();
    });
    const ready=await Promise.race([
      Promise.all([register(),controlled]).then(()=>true),
      new Promise(resolve=>{timer=setTimeout(()=>resolve(false),timeoutMs);}),
    ]);
    if(!ready||inputs.some(input=>input.files?.length))return false;
    // The marker contains no file data. If it cannot be stored, do not risk a
    // refresh loop; ordinary DOCX work remains available in this browser.
    const key='product-material-coi-reloaded:'+scope;
    if(page.sessionStorage.getItem(key))return false;
    page.sessionStorage.setItem(key,'1');
    if(page.sessionStorage.getItem(key)!=='1')return false;
    page.location.reload();
    return true;
  }catch{return false;}
  finally{
    clearTimeout(timer);
    if(checkController)serviceWorker.removeEventListener('controllerchange',checkController);
  }
}

/** Prepare browser isolation before the application accepts any local files. */
export async function startApplication({window:page=globalThis.window,moduleUrl=import.meta.url,loadApp=()=>import('./app.js'),timeoutMs=5000}={}){
  const document=page.document;
  const inputs=[...document.querySelectorAll('input[type="file"]')];
  const disabled=inputs.map(input=>input.disabled);
  const hint=document.getElementById('selection-hint'),previousHint=hint?.textContent;
  const blockDrop=event=>{event.preventDefault();event.stopImmediatePropagation?.();};
  inputs.forEach(input=>{input.disabled=true;});
  if(hint)hint.textContent='正在准备本地转换环境，请稍候…';
  document.addEventListener('dragover',blockDrop,true);
  document.addEventListener('drop',blockDrop,true);
  let reloading=false;
  try{reloading=await prepareIsolation(page,moduleUrl,timeoutMs,inputs);}
  finally{
    // Keep inputs blocked until navigation when a reload has been scheduled.
    if(!reloading){
      inputs.forEach((input,index)=>{input.disabled=disabled[index];});
      if(hint)hint.textContent=previousHint;
      document.removeEventListener('dragover',blockDrop,true);
      document.removeEventListener('drop',blockDrop,true);
    }
  }
  if(!reloading)await loadApp();
}

if(typeof window!=='undefined'&&typeof document!=='undefined'){
  startApplication().catch(error=>{
    const message=document.getElementById('error-message'),panel=document.getElementById('error-panel');
    if(message)message.textContent='页面未能完成加载，请刷新后重试。';
    if(panel)panel.hidden=false;
    console.error(error);
  });
}
