import {currentMonth,deliveryName,materialDeliveryName,uniqueDeliveryNames} from './naming.js?v=20261029-ppt-profiles';
import {currentProductionDate,resolveProductionDate,productionMonth} from './production-date.js?v=20261029-manager';

const $=id=>document.getElementById(id);
const supplementKinds=['date','benchmark','code'];
const performanceKeys=['threshold1','ratio1','threshold2','ratio2'];
const state={contract:null,templates:[],task:null,filter:'focus',busy:false,urls:[],month:currentMonth(),productionDate:currentProductionDate(),manualDate:undefined,monthValid:true,started:new Date(),run:0,company:null,companyVariant:'general',companySessions:{},companyBytes:null,companyTemplate:null,companyDownloadUrl:null,companyIsDefault:false,companyReady:false,companyPersisted:false,companyStored:false,companyWarning:'',store:null,manager:null,managerBytes:null,managerTemplate:null,managerDownloadUrl:null,managerReady:false,managerIsDefault:true,managerPersisted:false,managerStored:false,managerStore:null,managerNotice:'',applyManagerProfile:null,reviewApi:null,calendar:null,calendarMatches:[],supplementPreview:null};
const el=(tag,className,text)=>{const n=document.createElement(tag);if(className)n.className=className;if(text!==undefined)n.textContent=text;return n;};
const readableSize=bytes=>bytes>=1048576?(bytes/1048576).toFixed(1)+' MB':Math.max(1,Math.round(bytes/1024))+' KB';
const nextPaint=()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
state.introductionsOnly=false;state.applyPptManagerProfile=null;state.finishProfiles=null;
function showError(error){$('error-message').textContent=error?.message||String(error);$('error-panel').hidden=false;}
function clearDownloads(){for(const url of state.urls)URL.revokeObjectURL(url);state.urls=[];$('download-list').replaceChildren();$('generation-status').hidden=true;$('generation-status').textContent='';}
function changeStep(name){for(const key of ['upload','review','download']){$('step-'+key).classList.toggle('active',key===name);$('step-'+key).classList.toggle('done',name==='download'&&key!=='download'||name==='review'&&key==='upload');}}
function updateMonth(reset=false){state.started=new Date();if(reset){state.manualDate=undefined;state.productionDate=currentProductionDate(state.started);}state.productionDate=resolveProductionDate({manualValue:state.manualDate??state.productionDate,now:state.started});state.month=productionMonth(state.productionDate);state.monthValid=true;$('month-value').value=state.productionDate;const [year,month,day]=state.productionDate.split('-');$('production-date-display').textContent=`${year}年${Number(month)}月${Number(day)}日`;}
async function changeMonth(reset=false){
 if(state.busy)return;
 clearDownloads();
 try{if(!reset&&!$('month-value').value)throw new Error('请选择完整的制作日期。');state.manualDate=reset?undefined:$('month-value').value;updateMonth(reset);$('error-panel').hidden=true;if(state.task){state.task.month=state.month;state.task.productionDate=state.productionDate;if(state.manager&&state.applyManagerProfile)state.applyManagerProfile(state.task,state.manager,{productionDate:state.productionDate});if(state.manager&&state.applyPptManagerProfile)state.applyPptManagerProfile(state.task,state.manager,{productionDate:state.productionDate});state.finishProfiles?.(state.task);clearDownloads();changeStep('review');renderReview();if(state.task.updateMode==='introductions'&&state.task.materials.length)await generateFiles({autoDownload:false});}}
 catch(error){state.monthValid=false;showError(error);}refreshSelection();
}
$('month-value').addEventListener('change',()=>changeMonth());
$('month-reset').onclick=()=>changeMonth(true);
function refreshSelection(){
 for(const [key,files] of [['contract-files',state.contract?[state.contract]:[]],['template-files',state.templates]]){
  const box=$(key);box.replaceChildren();
  for(const file of files){const row=el('div','file-entry');row.append(el('span','',file.name),el('small','',readableSize(file.size)));const remove=el('button','', '×');remove.type='button';remove.setAttribute('aria-label','移除 '+file.name);remove.disabled=state.busy;remove.onclick=()=>{if(key==='contract-files'){state.contract=null;$('contract-input').value='';}else state.templates=state.templates.filter(f=>f!==file);invalidate();refreshSelection();};row.append(remove);box.append(row);}
 }
 $('analyze-button').disabled=state.busy||!state.companyReady||!state.managerReady||!state.monthValid||(!state.introductionsOnly&&!state.contract)||state.templates.length===0;
 $('generate-button').disabled=state.busy||!state.monthValid||!state.task||!state.task.materials.length;
 $('introductions-only').disabled=state.busy;$('contract-drop').hidden=state.introductionsOnly;
 $('upload-description').textContent=state.introductionsOnly?'上传待更新的一页通或 PPT，使用当前公司资料与经理 Excel，无需产品合同。':'一份目标产品合同，以及至少一份模板。';
 $('scope-description').textContent=state.introductionsOnly?'仅更新匹配到的公司和经理介绍，保留各模板产品名称及产品要素。':'产品要素＋一页通和 PPT 中匹配到的公司介绍、经理介绍。海报仍仅更新产品要素；历史业绩保持原样。';
 $('company-input').disabled=state.busy||!state.companyReady;$('company-clear').disabled=state.busy||!state.companyReady;
 $('company-variant').disabled=state.busy||!state.companyReady;$('company-download').setAttribute('aria-disabled',String(state.busy||!state.companyReady));
 $('manager-input').disabled=state.busy||!state.managerReady;$('manager-clear').disabled=state.busy||!state.managerReady;$('manager-download').setAttribute('aria-disabled',String(state.busy||!state.managerReady));
 $('selection-hint').textContent=state.introductionsOnly?`仅更新公司和经理介绍 · 已选择 ${state.templates.length} 份模板，无需合同。`:state.contract&&state.templates.length?`已选择合同和 ${state.templates.length} 份物料模板。`:'选好文件后，开始核对。';
}
$('introductions-only').checked=false;
$('introductions-only').addEventListener('change',()=>{if(state.busy)return;state.introductionsOnly=Boolean($('introductions-only').checked);state.filter=state.introductionsOnly?'all':'focus';invalidate();refreshSelection();$('upload-section').hidden=false;});
function invalidate(){discardSupplementPreview();state.task=null;state.calendar=null;state.calendarMatches=[];$('calendar-input').value='';$('calendar-row').replaceChildren();$('calendar-row').hidden=true;$('calendar-row-label').hidden=true;$('calendar-row-detail').replaceChildren();$('calendar-status').textContent='Excel 在本机读取。根据合同产品匹配，保留来源行供核对。';for(const kind of supplementKinds)$('supplement-'+kind).value='';for(const key of performanceKeys)$('performance-'+key).value='';$('performance-mode').value='single';syncPerformanceMode();$('supplement-preview').hidden=true;$('supplement-notice').hidden=true;$('review-section').hidden=true;clearDownloads();changeStep('upload');$('error-panel').hidden=true;}
const companyVersionLabel=()=>state.companyVariant==='ccb'?'建行版':'通用版';
const companySessionKeys=['company','companyBytes','companyTemplate','companyIsDefault','companyPersisted','companyStored','companyWarning'];
function rememberCompany(){state.companySessions[state.companyVariant]=Object.fromEntries(companySessionKeys.map(k=>[k,state[k]]));}
function showCompany(){
 const prefix=state.company?(state.companyIsDefault?`网页内置：${state.company.name}。未上传新版时自动使用。`:`当前版本：${state.company.name}。${state.companyPersisted?'已保存在此浏览器，下次自动沿用。':'仅本次使用，尚未保存。'}`):'公司介绍暂未加载，本轮只更新产品要素。';
 $('company-status').textContent=companyVersionLabel()+' · '+prefix+(state.companyWarning?' '+state.companyWarning:'');$('company-clear').hidden=(!state.company||state.companyIsDefault)&&!state.companyStored;
 $('company-variant').value=state.companyVariant;
 if(state.companyDownloadUrl)URL.revokeObjectURL(state.companyDownloadUrl);state.companyDownloadUrl=null;
 const link=$('company-download');link.hidden=!state.company;link.textContent='下载当前模板';link.setAttribute('aria-label','下载当前公司介绍模板（'+companyVersionLabel()+'）');
 if(state.companyBytes){state.companyDownloadUrl=URL.createObjectURL(new Blob([state.companyBytes],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}));link.href=state.companyDownloadUrl;link.download=state.company.name;}
 else{link.href=state.companyTemplate?.url||'';link.download=state.companyTemplate?.name||'';}
 if(state.companyReady)rememberCompany();
}
async function useDefaultCompany(){const {loadDefaultCompanyProfile,getDefaultCompanyTemplate}=await import('./default-company-profile.js?v=20261029-company-versions');state.company=loadDefaultCompanyProfile(state.companyVariant);state.companyTemplate=getDefaultCompanyTemplate(state.companyVariant);state.companyBytes=null;state.companyIsDefault=true;state.companyPersisted=false;}
async function restoreCompany(){
 try{
  const {createProfileStore}=await import('./profile-store.js?v=20261029-company-versions');state.store=createProfileStore({variant:state.companyVariant});
  if(state.companySessions[state.companyVariant]){Object.assign(state,state.companySessions[state.companyVariant]);return;}
  state.company=null;state.companyBytes=null;state.companyTemplate=null;state.companyStored=false;state.companyWarning='';await useDefaultCompany();const record=await state.store.load();
  if(record){
   state.companyStored=true;const {parseCompanyProfile}=await import('./company-profile.js?v=20261029-company-versions');const profile=await parseCompanyProfile({name:record.name,bytes:record.bytes},{variant:state.companyVariant});
   if(profile.sha256!==record.sha256)throw new Error('已保存的公司介绍校验不一致，请重新上传。');
   const {isSupersededCompanyProfile}=await import('./default-company-profile.js?v=20261029-company-versions');
   if(isSupersededCompanyProfile(profile)){
    state.companyWarning='已自动采用规模日期更正后的公司介绍。';
    try{state.companyStored=!await state.store.clearIfSha256(profile.sha256);}catch{state.companyWarning+='旧版本缓存暂未清除，但不会覆盖更正后的内置版本。';}
   }else{
    const {prepareCompanyTemplate}=await import('./company-template.js?v=20261029-company-versions');const prepared=await prepareCompanyTemplate({name:record.name,bytes:record.bytes},{variant:state.companyVariant});
    const unchanged=prepared.name===record.name&&prepared.bytes.length===record.bytes.length&&prepared.bytes.every((b,i)=>b===record.bytes[i]);
    const current=unchanged?profile:await parseCompanyProfile(prepared,{variant:state.companyVariant});state.company=current;state.companyBytes=prepared.bytes;state.companyIsDefault=false;state.companyPersisted=true;
   }
  }
 }
 catch(error){if(error.code==='PROFILE_CACHE_INVALID')state.companyStored=true;state.companyWarning='未能恢复已保存版本：'+(error.message||String(error));}
 finally{state.companyReady=true;showCompany();refreshSelection();}
}
async function selectCompany(files){
 if(state.busy)return;if(!state.companyReady){showError(new Error('正在读取上次保存的公司介绍，请稍后再上传。'));return;}const list=[...files];if(!list.length)return;
 if(list.length!==1||! /\.docx$/i.test(list[0].name)){showError(new Error('公司介绍请选择一份 DOCX 文件。'));return;}
 const file=list[0];if(file.size>60*1024*1024){showError(new Error('公司介绍文件请控制在60 MB以内。'));return;}
 setBusy(true);progress('正在读取公司介绍','识别'+companyVersionLabel()+'的两种版式，成功后仅替换此版本。');await nextPaint();
 try{const {parseCompanyProfile}=await import('./company-profile.js?v=20261029-company-versions');const {prepareCompanyTemplate}=await import('./company-template.js?v=20261029-company-versions');const prepared=await prepareCompanyTemplate({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())},{variant:state.companyVariant}),bytes=prepared.bytes;const profile=await parseCompanyProfile(prepared,{variant:state.companyVariant});let persisted=false;state.companyWarning='';
  try{if(!state.store){const {createProfileStore}=await import('./profile-store.js?v=20261029-company-versions');state.store=createProfileStore({variant:state.companyVariant});}await state.store.save({schemaVersion:1,name:prepared.name,sha256:profile.sha256,uploadedAt:new Date().toISOString(),bytes,profile});persisted=true;state.companyStored=true;}
  catch(error){state.companyWarning='此版本本次可用，但未能保存；下次请重新上传。'+(error.message||'');}
  state.company=profile;state.companyBytes=bytes;state.companyIsDefault=false;state.companyPersisted=persisted;invalidate();showCompany();
 }catch(error){showError(new Error('公司介绍未更新，仍保留此前版本。'+(error.message||String(error))));}
 finally{$('company-input').value='';setBusy(false);$('progress-panel').hidden=true;}
}
$('company-input').addEventListener('change',e=>selectCompany(e.target.files));
$('company-variant').addEventListener('change',async()=>{
 const variant=$('company-variant').value;if(state.busy||!state.companyReady||!['general','ccb'].includes(variant)){$('company-variant').value=state.companyVariant;return;}if(variant===state.companyVariant)return;
 rememberCompany();state.companyVariant=variant;state.companyReady=false;setBusy(true);invalidate();try{await restoreCompany();}finally{setBusy(false);}
});
$('company-download').onclick=e=>{if(state.busy||!state.companyReady||!state.company)e.preventDefault();};
$('company-clear').onclick=async()=>{if(state.busy)return;setBusy(true);try{if(state.companyStored){if(!state.store)throw new Error('尚未连接本机保存空间，请刷新后再清除。');await state.store.clear();}state.companyStored=false;await useDefaultCompany();state.companyWarning='';invalidate();showCompany();}catch(error){showError(error);}finally{setBusy(false);}};
function showManager(){
 $('manager-status').textContent=(state.manager?(state.managerIsDefault?'网页内置：':'当前版本：')+state.manager.name+(state.managerPersisted?' · 已保存在此浏览器，下次自动沿用。':'。'):'本次未加载经理资料。')+(state.managerNotice?' '+state.managerNotice:'');
 $('manager-clear').hidden=state.managerIsDefault&&!state.managerStored;
 if(state.managerDownloadUrl)URL.revokeObjectURL(state.managerDownloadUrl);state.managerDownloadUrl=null;
 const link=$('manager-download');link.hidden=!state.manager;link.textContent='下载当前模板';
 if(state.managerBytes){state.managerDownloadUrl=URL.createObjectURL(new Blob([state.managerBytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));link.href=state.managerDownloadUrl;link.download=state.manager.name;}
 else{link.href=state.managerTemplate?.url||'';link.download=state.managerTemplate?.name||'';}
}
async function useDefaultManager(){const {DEFAULT_MANAGER_PROFILE,DEFAULT_MANAGER_TEMPLATE}=await import('./default-manager-profile.js?v=20261029-manager');state.manager=DEFAULT_MANAGER_PROFILE;state.managerTemplate=DEFAULT_MANAGER_TEMPLATE;state.managerBytes=null;state.managerIsDefault=true;state.managerPersisted=false;}
async function restoreManager(){
 try{
  await useDefaultManager();const {applyManagerProfile}=await import('./manager-profile.js?v=20261029-manager');state.applyManagerProfile=applyManagerProfile;
  try{const [{createManagerStore},{parseManagerProfile}]=await Promise.all([import('./manager-store.js?v=20261029-manager'),import('./manager-data.js?v=20261029-manager')]);state.managerStore=createManagerStore();const record=await state.managerStore.load();
   if(record){state.managerStored=true;const profile=await parseManagerProfile({name:record.name,bytes:record.bytes});if(profile.sha256!==record.sha256)throw new Error('cache');state.manager=profile;state.managerBytes=record.bytes;state.managerIsDefault=false;state.managerPersisted=true;}
  }catch(error){if(error.code==='MANAGER_CACHE_INVALID')state.managerStored=true;state.managerNotice='本次沿用网页内置资料。';}
 }catch{state.managerNotice='经理资料尚未准备完成，开始核对时会自动重试。';}
 finally{state.managerReady=true;showManager();refreshSelection();}
}
async function selectManager(files){
 if(state.busy||!state.managerReady)return;const list=Array.from(files||[]);if(!list.length)return;
 if(list.length!==1||!/\.xlsx$/i.test(list[0].name)||list[0].size>30*1024*1024){state.managerNotice='请选择一份30 MB以内的 XLSX 文件，当前模板保留。';showManager();return;}
 setBusy(true);progress('正在读取经理介绍','按姓名识别完整简介，保留当前模板的版式。');await nextPaint();
 try{
  const file=list[0],bytes=new Uint8Array(await file.arrayBuffer());const {parseManagerProfile}=await import('./manager-data.js?v=20261029-manager');const profile=await parseManagerProfile({name:file.name,bytes});let persisted=false;state.managerNotice='';
  try{if(!state.managerStore){const {createManagerStore}=await import('./manager-store.js?v=20261029-manager');state.managerStore=createManagerStore();}await state.managerStore.save({schemaVersion:1,name:file.name,sha256:profile.sha256,uploadedAt:new Date().toISOString(),bytes,profile});persisted=true;}catch{state.managerNotice='已用于本次更新，下次使用时请重新上传。';}
  state.manager=profile;state.managerBytes=bytes;state.managerIsDefault=false;state.managerPersisted=persisted;state.managerStored=persisted||state.managerStored;invalidate();showManager();
 }catch{state.managerNotice='未读取到有效的姓名与简介，继续使用当前模板。';showManager();}
 finally{$('manager-input').value='';setBusy(false);$('progress-panel').hidden=true;}
}
$('manager-input').addEventListener('change',e=>selectManager(e.target.files));
$('manager-download').onclick=e=>{if(state.busy||!state.managerReady||!state.manager)e.preventDefault();};
$('manager-clear').onclick=async()=>{if(state.busy||!state.managerReady)return;setBusy(true);try{await state.managerStore?.clear();await useDefaultManager();state.managerStored=false;state.managerNotice='';invalidate();showManager();}catch{state.managerNotice='当前模板仍保留，可继续使用。';showManager();}finally{setBusy(false);}};
function selectFiles(kind,files){
 if(state.busy)return;
 const list=[...files];if(!list.length)return;
 const allowed=kind==='contract'?/\.(docx|doc)$/i:/\.(docx|doc|rtf|pptx)$/i;
 if(list.some(f=>!allowed.test(f.name))){showError(new Error(kind==='contract'?'请选择 DOCX 或 DOC 合同。':'模板支持 DOCX、DOC、RTF 和 PPTX 文件。'));return;}
 if(list.some(f=>f.size>60*1024*1024)){showError(new Error('单个文件请控制在60 MB以内。'));return;}
 if(kind==='contract'){if(list.length!==1){showError(new Error('每次请选择一份主合同。'));return;}state.contract=list[0];}
 else{for(const file of list)if(!state.templates.some(f=>f.name===file.name&&f.size===file.size&&f.lastModified===file.lastModified))state.templates.push(file);}
 invalidate();refreshSelection();
}
$('contract-input').addEventListener('change',e=>selectFiles('contract',e.target.files));
$('templates-input').addEventListener('change',e=>selectFiles('templates',e.target.files));
for(const [id,kind]of [['contract-drop','contract'],['templates-drop','templates'],['company-drop','company'],['manager-drop','manager']]){
 const box=$(id);box.addEventListener('dragover',e=>{e.preventDefault();box.classList.add('dragging');});box.addEventListener('dragleave',()=>box.classList.remove('dragging'));box.addEventListener('drop',e=>{e.preventDefault();box.classList.remove('dragging');if(kind==='company')return selectCompany(e.dataTransfer.files);else if(kind==='manager')return selectManager(e.dataTransfer.files);else selectFiles(kind,e.dataTransfer.files);});
}
function progress(title,detail){$('progress-panel').hidden=false;$('progress-title').textContent=title;$('progress-detail').textContent=detail;}
function setBusy(busy){state.busy=busy;$('calendar-drop').setAttribute('aria-busy',String(busy));if(busy)$('calendar-drop').classList.remove('dragging');for(const id of ['contract-input','templates-input','company-input','company-clear','manager-input','manager-clear','month-value','month-reset','reset-button','calendar-input','calendar-row',...supplementKinds.flatMap(k=>['supplement-'+k,'supplement-parse-'+k]),...performanceKeys.map(k=>'performance-'+k),'performance-mode','supplement-parse-performance','supplement-text-tab','supplement-excel-tab'])$(id).disabled=busy;document.querySelectorAll('#facts-list input,#facts-list textarea,#facts-list button,#supplement-preview button').forEach(n=>n.disabled=busy);refreshSelection();}
$('dismiss-error').onclick=()=>$('error-panel').hidden=true;
$('reset-button').onclick=()=>{state.run++;state.contract=null;state.templates=[];$('contract-input').value='';$('templates-input').value='';invalidate();updateMonth(true);refreshSelection();$('upload-section').hidden=false;window.scrollTo({top:0,behavior:'smooth'});};
$('analyze-button').onclick=async()=>{
 if(state.busy||!state.companyReady||!state.managerReady||(!state.introductionsOnly&&!state.contract)||!state.templates.length)return;try{updateMonth();}catch(error){showError(error);return;}const run=++state.run;invalidate();setBusy(true);progress(state.introductionsOnly?'正在准备介绍更新':'正在读取合同',state.introductionsOnly?'使用当前保存的公司资料和经理 Excel。':'建立本次合同索引，保留文字出处。');await nextPaint();
 try{
  if(!state.manager||!state.applyManagerProfile){await restoreManager();if(!state.manager||!state.applyManagerProfile)throw new Error('经理介绍功能暂未加载完成，请刷新网页后重新核对。');}
  const [{readWordContract,readWordTemplate},reviewApi,profileApi]=await Promise.all([import('./word-input.js?v=20261027-custom-investment'),state.introductionsOnly?Promise.resolve({}):import('./review-engine.js?v=20261029-risk-custom'),import('./introduction-task.js?v=20261029-ppt-profiles')]);state.reviewApi=reviewApi;state.finishProfiles=profileApi.finalizeIntroductionTask;
  const contract=state.introductionsOnly?null:await readWordContract({name:state.contract.name,bytes:new Uint8Array(await state.contract.arrayBuffer())},{onProgress:progress});
  const templates=[],inputIssues=[],seen=new Set();
  for(const [i,file]of state.templates.entries()){
   progress(`正在准备模板 ${i+1} / ${state.templates.length}`,`接受旧修订，按需转换 DOCX：${file.name}`);await nextPaint();
   try{
    const template=await readWordTemplate({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())},{onProgress:progress});
    const hash=template.sourceSha256||template.sha256;
    if(contract&&hash===contract.sha256){inputIssues.push({sourceName:file.name,reason:'contract',message:'与合同是同一份文件，已从模板中跳过。'});continue;}
    if(seen.has(hash)){inputIssues.push({sourceName:file.name,reason:'duplicate',message:'与已选模板内容相同，已合并重复项。'});continue;}
    seen.add(hash);templates.push(template);
   }catch(error){inputIssues.push({sourceName:file.name,reason:'failed',message:error.message||String(error)});}
  }
  if(!templates.length)throw new Error('没有可更新的模板。'+inputIssues.map(i=>`${i.sourceName}：${i.message}`).join('\n'));
  progress(state.introductionsOnly?'正在定位介绍内容':'正在统一核对产品要素',state.introductionsOnly?'分别匹配公司主体及每位经理，保留模板产品名称。':'合并共同需求，查找合同依据，再定位各物料的改动。');await nextPaint();
  const task=state.introductionsOnly?profileApi.createIntroductionTask(templates,{now:state.started,month:state.month,productionDate:state.productionDate}):await reviewApi.analyzeMaterials(contract,templates,{now:state.started,month:state.month});task.now=state.started;task.productionDate=state.productionDate;task.inputIssues=[...(task.inputIssues||[]),...inputIssues];
  if(state.company){progress('正在核对公司介绍',companyVersionLabel()+' · '+state.company.name);await nextPaint();const [{applyCompanyProfile},{applyPptCompanyProfile}]=await Promise.all([import('./company-profile.js?v=20261029-company-versions'),import('./ppt-company-profile.js?v=20261029-ppt-profiles')]);await applyCompanyProfile(task,state.company);applyPptCompanyProfile(task,state.company);}
  if(state.manager&&state.applyManagerProfile){progress('正在更新经理介绍',state.manager.name);await nextPaint();const {applyPptManagerProfile}=await import('./ppt-manager-profile.js?v=20261029-ppt-profiles');state.applyPptManagerProfile=applyPptManagerProfile;state.applyManagerProfile(task,state.manager,{productionDate:state.productionDate});applyPptManagerProfile(task,state.manager,{productionDate:state.productionDate});}
  profileApi.finalizeIntroductionTask(task);
  if(run!==state.run)return;
  state.task=task;const feeGroups=groups().filter(g=>g.kind==='performance');if(feeGroups.length===1&&groupFacts(feeGroups[0]).some(f=>f.key==='performance_R2'))$('performance-mode').value='tiered';syncPerformanceMode();renderReview();$('upload-section').hidden=true;$('review-section').hidden=false;changeStep('review');
 }catch(error){showError(error);}finally{if(run===state.run){setBusy(false);$('progress-panel').hidden=true;}}
 if(run===state.run&&state.task?.updateMode==='introductions'){if(state.task.materials.length)await generateFiles({autoDownload:false});else generationStatus('未找到与当前资料匹配的介绍，相关模板已跳过。','warning');}
};
function factChanged(fact){return state.task.materials.some(m=>m.operations.some(o=>o.kind==='replace'&&o.factIds?.includes(fact.id)));}
function factStatus(fact){return fact.status==='verified'?(factChanged(fact)?'changed':'verified'):(fact.status==='conflict'?'conflict':'pending');}
function materialName(m){try{return materialDeliveryName(m,{targetName:state.task.target.name,month:state.task.month});}catch(error){return m.template.name+'（命名待检查）';}}
function sourceTitle(evidence){if(String(evidence.blockId).startsWith('manager-'))return '经理介绍 Excel 依据';if(String(evidence.blockId).startsWith('user-custom'))return '用户自定义内容';if(String(evidence.blockId).startsWith('company-'))return '最新公司介绍依据';if(String(evidence.blockId).startsWith('user-calendar')){const s=evidence.source||{};return '产品日历补充'+(s.file?'：'+s.file+' · '+(s.sheet||'')+' · 第'+s.row+'行':'');}if(String(evidence.blockId).startsWith('user-confirm'))return '人工确认';if(evidence.blockId==='user-supplement')return '手动补充';if(String(evidence.blockId).startsWith('template-'))return '模板中的说明';return '合同依据';}
function renderFact(fact){
 const row=el('article','fact-row');row.dataset.factId=fact.id;const status=factStatus(fact);const head=el('div','fact-heading');head.append(el('strong','',fact.label||fact.key));head.append(el('span','badge '+status,({changed:'已修改',verified:'原样保留',pending:'保留待核',conflict:'需要复核'})[status]));const count=new Set((fact.occurrences||[]).map(o=>o.materialId+'|'+o.paragraphId)).size;head.append(el('small','',`关联 ${count} 处`));row.append(head);
 const value=fact.value===null||fact.value===undefined?'保留各模板原值':String(fact.value);row.append(el('p','fact-value',value.length>220?value.slice(0,220)+'…':value));if(fact.status!=='verified')row.append(el('p','fact-reason',fact.reasonType==='contract_missing'?'合同未给出具体值，保留原值。':'需要人工核对，当前未自动确认。'));row.append(el('p','fact-reason',fact.reason||''));
 const details=el('details');details.append(el('summary','', '查看依据与物料位置'));
 for(const evidence of (fact.evidence||[]).slice(0,8)){details.append(el('p','',sourceTitle(evidence)));details.append(el('blockquote','',evidence.quote||''));}
 if(!fact.evidence?.length)details.append(el('p','','尚无可唯一确认的合同依据，当前保留模板值。'));
 for(const m of state.task.materials){const fields=m.fields.filter(f=>f.factIds?.includes(fact.id));for(const field of fields){const occurrence=el('div','occurrence');occurrence.append(el('strong','',m.template.name+' · '+field.label));const before=field.before||'';const after=field.after??before;if(before!==after){const b=el('p');b.append(el('del','',before));const a=el('p');a.append(el('ins','',after));occurrence.append(b,a);}else occurrence.append(el('p','',before));details.append(occurrence);}}
 row.append(details);
 if(fact.status!=='verified'){
  const form=el('form','supplement');const input=el('input');input.placeholder='有明确补充资料时，可在这里填写';input.setAttribute('aria-label','补充'+(fact.label||fact.key));const button=el('button','','确认补充');button.type='submit';form.append(input,button);form.onsubmit=async e=>{e.preventDefault();if(!input.value.trim())return;try{const {setSupplement}=await import('./review-engine.js?v=20261029-risk-custom');state.task=await setSupplement(state.task,fact.id,input.value);clearDownloads();$('error-panel').hidden=true;changeStep('review');renderReview();}catch(error){showError(error);}};row.append(form);
 }
 return row;
}
function renderManagerFact(facts){
 const fact=facts[0],ids=new Set(facts.map(f=>f.id));
 const row=el('article','fact-row manager-review');const heading=el('div','fact-heading');heading.append(el('h4','',fact.label),el('span','badge verified',facts.some(factChanged)?'已更新':'已核对'));row.append(heading);
 row.append(el('p','supplement-note','按经理 Excel 完整更新；Word 审阅仅标出差异。制作日期：'+state.productionDate));
 row.append(el('p','manager-biography',fact.value||''));
 if(fact.reason)row.append(el('p','supplement-note',fact.reason));
 for(const calculation of fact.calculations||[])row.append(el('p','supplement-note',`${calculation.label}：${calculation.startDate} 至 ${calculation.productionDate}，已满 ${calculation.years} 年。`));
 const details=el('details');details.append(el('summary','','查看 Excel 依据与物料位置'));
 for(const evidence of fact.evidence||[]){const source=evidence.source||{};details.append(el('p','',source.file?[source.file,source.sheet,source.row?'第 '+source.row+' 行':''].filter(Boolean).join(' · '):'经理介绍 Excel 依据'),el('blockquote','manager-biography',evidence.quote||''));}
 for(const material of state.task.materials)for(const field of material.fields.filter(f=>f.factIds?.some(id=>ids.has(id)))){const place=el('div','occurrence');place.append(el('strong','',material.template.name+' · '+field.label));place.append(el('p','supplement-note','原简介'),el('p','manager-biography',field.before||''));if(field.before!==field.after)place.append(el('p','supplement-note','更新后完整简介'),el('p','manager-biography',field.after??fact.value??''));details.append(place);}
 row.append(details);return row;
}
function isConfirmed(f){return f.userConfirmed===true||f.source==='user_confirmation';}
function canSupplement(f){return f.status!=='verified'||f.source==='user'||!!f.redemptionTier;}
function discardSupplementPreview(){state.supplementPreview=null;if(state.supplementFeedback){state.supplementFeedback.hidden=true;state.supplementFeedback.replaceChildren();state.supplementFeedback=null;}$('supplement-preview').hidden=true;}
function openSupplementFeedback(box){discardSupplementPreview();state.supplementFeedback=box;box.replaceChildren();box.hidden=false;$('error-panel').hidden=true;$('supplement-notice').hidden=true;}
function showSupplementIssue(error,box){openSupplementFeedback(box);box.append(el('p','supplement-warning',error?.message||String(error)));box.scrollIntoView({block:'nearest'});}
function reviewEdited(message){clearDownloads();discardSupplementPreview();$('error-panel').hidden=true;changeStep('review');renderReview();$('supplement-notice').textContent=message;$('supplement-notice').hidden=false;}
function groups(){
 const base=state.reviewApi?.getSupplementGroups?state.reviewApi.getSupplementGroups(state.task):state.task.facts.map(f=>({id:f.id,label:f.label,kind:'text',factIds:[f.id],status:f.status}));
 const cards=new Map(),normalize=s=>String(s??'').normalize('NFKC').replace(/[\s·•,，。;；:：、【】\[\]"“”'‘’]/g,'');
 for(const group of base){
  if(group.kind!=='text'){cards.set(group.id,group);continue;}
  const fs=groupFacts(group),scope={...group.scope};
  const managerGroup=fs.length&&fs.every(f=>f.source==='manager');
  if(managerGroup){delete scope.page;delete scope.pages;}
  // Only a condition that literally represents the original wording is ignored
  // for display. Shares, channels and semantic conditions stay independent.
  if(scope.condition&&fs.every(f=>(state.task._records||[]).some(r=>r.f===f&&normalize(r.atom.before)===scope.condition)))delete scope.condition;
  const identity=JSON.stringify([group.label,[...new Set(fs.map(f=>f.key))].sort(),Object.keys(scope).sort().map(k=>[k,scope[k]]),...(managerGroup?[fs.map(f=>[f.value,(f.evidence||[]).map(e=>e.blockId)])]:[])]);
  const existing=cards.get(identity);
  if(existing){existing.members.push(group);existing.factIds.push(...group.factIds);}
  else cards.set(identity,{...group,id:'review-'+group.id,scope,factIds:[...group.factIds],members:[group]});
 }
 return [...cards.values()].map(g=>g.members?.length===1?g.members[0]:g);
}
function groupFacts(group){return group.factIds.map(id=>state.task.facts.find(f=>f.id===id)).filter(Boolean);}
function sectionReview(group){const review=state.reviewApi?.getSectionReview?.(state.task,group.factIds);return review?{...review,factIds:group.factIds}:null;}
function sectionSignature(review){
 const ids=new Set(review.factIds||[]),wording=state.task.materials.flatMap(m=>m.fields.filter(f=>f.factIds?.some(id=>ids.has(id))).map(f=>[m.materialId,f.id,f.before,f.after]));
 return JSON.stringify([review.evidence,review.materials,wording]);
}
function sectionConfirmed(review){return !!review&&state.task.sectionAcknowledgements?.[review.id]===sectionSignature(review);}
function sectionPending(review){return !!review?.needsReview&&!sectionConfirmed(review)&&!review.factIds?.some(id=>state.task.facts.find(f=>f.id===id)?.customSectionOverrideId);}
function sectionCanAdopt(review){return ['investment_ratios','investment_strategy','opening_fixed_monthdays'].includes(review?.key)&&!!(review.replacement?.lines?.length||review.materials?.some(m=>m.replacement?.lines?.length));}
function sectionAction(review){return review.key==='investment_strategy'?(review.materials?.some(m=>m.replacement?.removedTitles?.length)?'采用合同策略标题':'补齐策略标题'):review.key==='opening_fixed_monthdays'?'采用合同开放期':'采用合同投资比例';}
function sectionAdoptedText(review){return review?.key==='investment_strategy'?'已采用合同策略标题':review?.key==='opening_fixed_monthdays'?'已采用合同开放期规则':'已采用完整合同投资比例';}
const sectionStatus={same:'已对应',local:'数值差异',structure:'口径或条件不同',missing:'未在此物料体现',uncertain:'需核对',summary:'摘要表达，需核对'};
function renderSectionReview(row,review){
 const confirmed=sectionConfirmed(review),adopted=state.task.sectionOverrides?.[review.id],box=el('div','section-review'+(sectionPending(review)?' needs-review':''));
 box.append(el('p','section-review-summary',adopted?sectionAdoptedText(review)+'，'+review.materials.map(m=>`已对应 ${m.matched}/${m.total} 项`).filter((s,i,a)=>a.indexOf(s)===i).join('；')+'。':review.summary));
 box.append(el('p','fact-reason',adopted?'已更新至以下物料；其他要素保持当前结果。':confirmed?'已核对差异，保留当前表述':review.preferCompleteEvidence?'展开下方完整合同依据，按整节核对；不会把合同全文重复填入各个旧段落。':review.needsReview?(review.key==='investment_strategy'?'可预览补齐缺失标题，确认后写入物料；只使用策略小标题。':review.key==='opening_fixed_monthdays'?'当前是月内固定多日期规则。请核对日期、顺延口径和附加条件；可预览采用合同完整规则。':'请逐项核对未体现或口径不同的内容；模板可能采用摘要表述，未自动整段替换。'):'依据已完整列出；继续沿用现有的局部修改方式。'));
 const compare=el('details');compare.append(el('summary','','逐项查看合同与当前物料'));
 for(const material of review.materials||[]){
  compare.append(el('p','section-material',material.name+` · 已对应 ${material.matched}/${material.total} 项`));
  for(const item of [...(material.items||[]),...(material.extra||[])]){
   const entry=el('div','section-comparison');entry.append(el('strong','',(sectionStatus[item.status]||'需核对')+' · '+(item.title||'补充表述')));
   if(item.contractText)entry.append(el('p','','合同：'+item.contractText));
   entry.append(el('p','','当前物料：'+(item.templateText||'尚未找到明确对应表述')));
   if(item.reason)entry.append(el('p','fact-reason',item.reason));compare.append(entry);
  }
 }
 if(!review.preferCompleteEvidence)box.append(compare);
 for(const note of review.notes||[]) {const detail=el('details');detail.append(el('summary','','条件性说明 · '+note.title),el('blockquote','section-evidence',note.contractText));box.append(detail);}
 if(adopted){
  const undo=el('button','text-button','撤销采用合同内容');undo.type='button';undo.onclick=()=>{if(state.busy)return;try{state.task=state.reviewApi.undoContractSection(state.task,review.id);reviewEdited('已撤销整段采用，恢复原有局部修改；其他补充保持不变。');}catch(error){showError(error);}};box.append(el('p','fact-reason',sectionAdoptedText(review)+'，Word 保留本轮审阅修订。'),undo);
 }else if(sectionCanAdopt(review)&&review.needsReview){
  const action=sectionAction(review),previewButton=el('button','','预览'+action),previewBox=el('div','section-adoption-preview');previewButton.type='button';previewBox.hidden=true;
  previewButton.onclick=()=>{if(state.busy)return;try{
   const preview=state.reviewApi.previewContractSection(state.task,review.factIds);previewBox.replaceChildren();previewBox.hidden=false;
   previewBox.append(el('p','supplement-note','确认后仅更新以下物料的'+review.label+'。其他要素保持当前结果；Word 使用审阅修订。'));
   for(const entry of preview.entries){previewBox.append(el('strong','',entry.name),el('p','','当前：'),el('blockquote','section-evidence',entry.before),el('p','','拟采用合同内容：'),el('blockquote','section-evidence',entry.after));for(const warning of entry.warnings||[])previewBox.append(el('p','supplement-warning',warning));}
   const apply=el('button','primary-button','确认'+action),cancel=el('button','text-button','取消');apply.type=cancel.type='button';
   apply.onclick=()=>{if(state.busy)return;try{state.task=state.reviewApi.applyContractSection(state.task,preview);reviewEdited(sectionAdoptedText(review)+'，并同步相关物料。可在本卡片撤销。');}catch(error){showError(error);}};cancel.onclick=()=>{previewBox.hidden=true;previewBox.replaceChildren();};previewBox.append(apply,cancel);
  }catch(error){showError(error);}};box.append(previewButton,previewBox);
 }
 if(review.needsReview){
  const actions=el('div','fact-actions'),button=el('button','text-button',confirmed?'撤销差异确认':'已核对差异，保留当前表述');button.type='button';button.setAttribute('aria-label',button.textContent+'：'+review.label);
  button.onclick=()=>{if(state.busy)return;const current=sectionReview({factIds:review.factIds});if(!current)return;state.task.sectionAcknowledgements??={};if(sectionConfirmed(current))delete state.task.sectionAcknowledgements[current.id];else state.task.sectionAcknowledgements[current.id]=sectionSignature(current);reviewEdited('已记录章节差异核对结果。物料当前文字与本轮修订保持不变。');};
  actions.append(button);box.append(actions);
 }
 row.append(box);
}
function scopedGroupLabel(group){const s=group.scope||{},parts=[s.share?s.share+'类份额':'',s.channel,s.issuer==='distributor'?'代销口径':''].filter(Boolean);return group.label+(parts.length?' · '+parts.join(' · '):'');}
function readableFact(f){const labels={opening_cycle:['开放周期','个月'],opening_days:['开放天数','个交易日'],opening_monthday:['每月开放日','日'],opening_weekday:['开放星期',''],opening_monthly_weekday:['月内开放安排','']};const mapped=labels[f.key];if(mapped)return mapped[0]+'：'+f.value+mapped[1];if(f.key==='opening_announced_date'){const parts=(f.scope?.condition||'').split('|'),label=({participation_exit:'参与和退出',participation_only:'仅参与'})[parts[3]]||({opening:'本次开放日',sale:'本次开放日',reservation:'本次预约日'})[parts[0]]||'本次开放日';return label+'：'+f.value;}return f.label+'：'+f.value;}
const examples={date:'2026年11月23日至2026年11月27日',performance:'2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%',benchmark:'2.25%',code:'000123',opening:'每满18个月，开放5个交易日',text:'填写本次确认使用的内容'};
function supplementForm(group,label){
 const keys=groupFacts(group).map(f=>f.key),fee=keys.some(k=>k.endsWith('_fee'));
 const hint=fee?'可填数字修改对应费率，或粘贴整段文字，预览后完整替换本项。':keys.includes('risk_rating')?'例如 R2 或中低风险。':keys.includes('lockup_days')?'例如 180天。':keys.includes('duration_years')?'例如 10年。':keys.some(k=>/^(minimum|additional)_amount$/.test(k))?'请填写金额和单位，例如 100万元。':'';
 const multiline=group.kind==='performance'||fee||keys.some(k=>/^(investment_|opening$|opening_fixed_monthdays|lockup$)/.test(k));
 const form=el('form','supplement card-supplement'),input=el(multiline?'textarea':'input');input.placeholder=hint||examples[group.kind]||examples.text;input.setAttribute('aria-label','补充'+label);input.addEventListener('input',discardSupplementPreview);if(multiline)input.rows=2;
 const button=el('button','','预览补充'),feedback=el('div','supplement-preview card-supplement-feedback');button.type='submit';feedback.hidden=true;feedback.setAttribute('role','status');feedback.setAttribute('aria-live','polite');form.append(input,button,feedback);
 form.onsubmit=e=>{e.preventDefault();if(state.busy)return;discardSupplementPreview();if(!input.value.trim()){showSupplementIssue('请填写本次确认使用的内容，再预览补充；当前物料保持不变。',feedback);return;}try{showSupplementPreview((fee&&state.reviewApi.previewCardSupplement?state.reviewApi.previewCardSupplement:state.reviewApi.previewSupplement)(state.task,group.id,input.value),{label,box:feedback});}catch(error){showSupplementIssue(error,feedback);}};
 return form;
}
function customSectionFor(group){const id=groupFacts(group).find(f=>f.customSectionOverrideId)?.customSectionOverrideId;return state.task.customSectionOverrides?.[id];}
function canCustomizeSection(group){const keys=[...new Set(groupFacts(group).map(f=>f.key))];return keys.length===1&&['investment_scope','investment_ratios','investment_strategy','investment_restrictions'].includes(keys[0]);}
function renderInvestmentEvidence(details,section){
 const key=s=>String(s).normalize('NFKC').replace(/\s+/g,''),versions=[];
 for(const e of section.evidence||[]){const existing=versions.find(v=>key(v.quote)===key(e.quote));if(existing)existing.locations+=(e.duplicateCount||0)+1;else versions.push({...e,locations:(e.duplicateCount||0)+1});}
 if(!versions.length)return;
 details.append(el('h4','','合同依据'),el('p','supplement-note','来源：'+state.task.contract.name+(versions.length>1?'。不同位置原文分别保留，以下顺序不代表适用优先级。':'')));
 const lines=e=>String(e.quote).split(/\n/).map(s=>s.trim()).filter(s=>s&&!/^(?:[（(]?[一二三四五六七八九十\d]+[）)、.．]?\s*)?投资(?:范围|比例)[:：]?\s*$/.test(s));
 const first=versions[0],firstLines=lines(first),firstKeys=new Set(firstLines.map(key));
 for(const [i,e]of versions.entries()){
  const count='完整合同依据 · 共 '+e.itemCount+' 项'+(e.noteCount?'＋'+e.noteCount+' 段补充说明':'')+(e.locations>1?' · '+e.locations+' 处原文一致，合并展示':'');
  if(!i){details.append(el('p','',count),el('blockquote','section-evidence',e.quote));continue;}
  const current=lines(e),currentKeys=new Set(current.map(key)),same=current.filter(s=>firstKeys.has(key(s))).length,missing=firstLines.filter(s=>!currentKeys.has(key(s))),added=current.filter(s=>!firstKeys.has(key(s)));
  const other=el('details','alternate-section-evidence');other.append(el('summary','','其他位置的合同依据 · '+(same?'与上方'+same+'段相同':'内容不同，展开核对')));
  if(missing.length)other.append(el('p','supplement-note','与上方相比，本处未含：'+missing.join('\n')));
  if(added.length)other.append(el('p','supplement-note','本处另有：'+added.join('\n')));
  other.append(el('p','',count),el('blockquote','section-evidence',e.quote));details.append(other);
 }
}
function customSectionControl(group,label){
 const override=customSectionFor(group),details=el('details','custom-section-control');details.append(el('summary','','自定义替换'));
 details.append(el('p','supplement-note','粘贴本项需要保留的完整内容。确认后整项替换，未包含的旧条目会删除；相关物料一起更新，Word 保留审阅修订。'));
 const form=el('div','supplement card-supplement'),input=el('textarea'),button=el('button','','预览整项替换'),feedback=el('div','supplement-preview card-supplement-feedback');
 input.rows=5;input.placeholder='粘贴完整的'+label+'，可包含多行或多个条目';input.setAttribute('aria-label','自定义替换'+label);input.value=override?.text||'';input.addEventListener('input',discardSupplementPreview);
 button.type='button';feedback.hidden=true;feedback.setAttribute('role','status');feedback.setAttribute('aria-live','polite');form.append(input,button,feedback);details.append(form);
 button.onclick=()=>{if(state.busy)return;discardSupplementPreview();try{
  const preview=state.reviewApi.previewCustomSection(state.task,group.factIds,input.value);openSupplementFeedback(feedback);state.supplementPreview=preview;
  feedback.append(el('h4','','整项替换预览 · '+label),el('p','supplement-note','以下各处将全部采用你填写的内容。旧条目不再保留；这属于用户自定义内容，不是合同自动提取结论。'));
  for(const entry of preview.entries)feedback.append(el('strong','',entry.name),el('p','','当前内容：'),el('blockquote','section-evidence',entry.before),el('p','','替换后完整内容：'),el('blockquote','section-evidence',entry.after));
  const apply=el('button','primary-button','确认整项替换'),cancel=el('button','text-button','取消');apply.type=cancel.type='button';
  apply.onclick=()=>{if(state.busy||state.supplementPreview!==preview)return;try{state.task=state.reviewApi.applyCustomSection(state.task,preview);reviewEdited('已整项替换'+label+'并同步相关物料。Word 保留审阅修订，可撤销。');}catch(error){showSupplementIssue(error,feedback);}};
  cancel.onclick=discardSupplementPreview;feedback.append(apply,cancel);feedback.scrollIntoView({block:'nearest'});
 }catch(error){showSupplementIssue(error,feedback);}};
 if(override){const undo=el('button','text-button','撤销自定义替换');undo.type='button';undo.onclick=()=>{if(state.busy)return;try{state.task=state.reviewApi.undoCustomSection(state.task,override.id);reviewEdited('已撤销'+label+'自定义替换，恢复此前核对结果。');}catch(error){showError(error);}};details.append(undo);}
 return details;
}
function renderGroup(group,section=sectionReview(group)){
 const fs=groupFacts(group),custom=customSectionFor(group),customizable=canCustomizeSection(group),pending=fs.some(f=>f.status!=='verified'),confirmed=fs.some(isConfirmed),changed=fs.some(factChanged),adopted=fs.some(f=>f.sectionOverrideId),noFee=fs.some(f=>f.key==='performance_policy'&&f.policyMode==='none'&&f.status==='verified'),status=custom?'changed':pending||sectionPending(section)?'pending':confirmed||sectionConfirmed(section)?'confirmed':changed?'changed':'verified';
 const label=scopedGroupLabel(group),row=el('article','fact-row');row.dataset.groupId=group.id;const head=el('div','fact-heading');head.append(el('strong','',label),el('span','badge '+status,({pending:'保留待核',confirmed:'已确认无误',changed:'已修改',verified:'原样保留'})[status]));const count=new Set(fs.flatMap(f=>(f.occurrences||[]).map(o=>o.materialId+'|'+o.paragraphId))).size;head.append(el('small','',`关联 ${count} 处`));row.append(head);
 const values=[...new Set(fs.filter(f=>f.value!==null&&f.value!==undefined).map(f=>f.feeTextOverrideId?String(f.value):fs.length>1||group.kind==='opening'?readableFact(f):String(f.value)))];row.append(el('p','fact-value',custom?custom.text:adopted?sectionAdoptedText(section):confirmed?(values.length?values.join('\n')+'\n其余原值已确认无误':'已确认保留各模板原值'):section?.preferCompleteEvidence?'建议按完整合同条款核对':values.length?values.join('\n'):'保留各模板原值'));
 const reasons=[...new Set(fs.map(f=>f.reason).filter(Boolean))];if(!section?.preferCompleteEvidence)for(const reason of reasons)row.append(el('p','fact-reason',reason));
 if(custom)row.append(el('p','fact-reason','已自定义替换'));else if(section)renderSectionReview(row,section);
 const details=el('details');details.append(el('summary','',section?'查看完整合同依据与物料位置':'查看依据与物料位置'));
 if(customizable&&['investment_scope','investment_ratios'].includes(section?.key)&&section?.evidence?.length)renderInvestmentEvidence(details,section);
 else for(const evidence of section?.evidence||[]){const label=`完整合同依据 · 共 ${evidence.itemCount} 项`+(evidence.noteCount?`＋${evidence.noteCount} 段补充说明`:'')+(evidence.duplicateCount>0?` · 合同中 ${evidence.duplicateCount+1} 处内容一致，已合并展示`:'');if(evidence.variantRole==='other_version'){const alternate=el('details');alternate.append(el('summary','','其他版本依据 · 共 '+evidence.itemCount+' 项'),el('p','',label),el('blockquote','section-evidence',evidence.quote));details.append(alternate);}else details.append(el('p','',label),el('blockquote','section-evidence',evidence.quote));}
 const seenEvidence=new Set();for(const f of fs)for(const e of(f.evidence||[])){if(section?.evidence?.length&&sourceTitle(e)==='合同依据')continue;const k=e.blockId+'|'+e.quote;if(seenEvidence.has(k))continue;seenEvidence.add(k);details.append(el('p','',sourceTitle(e)),el('blockquote','',e.quote||''));}
 if(customizable)details.append(el('h4','','物料原文与当前结果'));
 const factIds=new Set(group.factIds);for(const m of state.task.materials)for(const field of m.fields.filter(f=>f.factIds?.some(id=>factIds.has(id)))){const occurrence=el('div','occurrence');occurrence.append(el('strong','',m.template.name+' · '+field.label));if(noFee){occurrence.append(el('p','',field.after||'已移除旧计提说明；采用合同不提取业绩报酬的结论。'));}else if(field.before!==field.after){const b=el('p'),a=el('p');b.append(el('del','',field.before));a.append(el('ins','',field.after??field.before));occurrence.append(b,a);}else occurrence.append(el('p','',field.before));details.append(occurrence);}row.append(details);
 if(customizable||!adopted&&!(sectionCanAdopt(section)&&section?.needsReview)&&!section?.preferCompleteEvidence&&group.members?.length>1&&fs.some(canSupplement)){
  const edits=el('details');edits.append(el('summary','','分别补充原文'),el('p','supplement-note','各处原话术分别保留。补充只更新所选原文，不覆盖其他段落。'));
  if(custom||adopted)edits.append(el('p','supplement-note','当前已整项替换。如需分别补充原文，请先撤销整项替换。'));
  for(const [i,member]of (custom||adopted?[]:group.members||[group]).entries()){
   if(!customizable&&!groupFacts(member).some(canSupplement))continue;
   const records=(state.task._records||[]).filter(r=>member.factIds.includes(r.f.id));
   const originals=records.length?records.map(r=>({name:r.m.template.name,text:r.atom.before})):state.task.materials.flatMap(m=>m.fields.filter(f=>f.factIds?.some(id=>member.factIds.includes(id))).map(f=>({name:m.template.name,text:f.before})));
   edits.append(el('p','fact-reason','原文 '+(i+1)+' · '+[...new Set(originals.map(o=>o.name))].join('、')));
   for(const text of new Set(originals.map(o=>o.text)))edits.append(el('p','',text));
   edits.append(supplementForm(member,label+' · 原文 '+(i+1)));
  }
  row.append(edits);
 }else if(!adopted&&!(sectionCanAdopt(section)&&section?.needsReview)&&!section?.preferCompleteEvidence&&!group.members&&(fs.some(canSupplement)||['date','performance','benchmark','code','opening'].includes(group.kind)))row.append(supplementForm(group,label));
 if(customizable)row.append(customSectionControl(group,label));
 const feeTextId=fs.find(f=>f.feeTextOverrideId)?.feeTextOverrideId;if(feeTextId){const undo=el('button','text-button','撤销整段修改');undo.type='button';undo.onclick=()=>{if(state.busy)return;try{state.task=state.reviewApi.undoFeeTextSupplement(state.task,feeTextId);reviewEdited('已撤销整段修改，恢复此前核对结果。');}catch(error){showError(error);}};row.append(undo);}
 if(pending||confirmed){const actions=el('div','fact-actions'),confirm=el('button','text-button',confirmed&&!pending?'撤销确认':'确认无误');confirm.type='button';confirm.setAttribute('aria-label',(confirmed&&!pending?'撤销确认':'确认无误')+'：'+label);confirm.onclick=()=>{if(state.busy)return;try{const undo=confirmed&&!pending;state.task=state.reviewApi.confirmFacts(state.task,group.factIds,!undo);reviewEdited(undo?'已撤销确认，对应原值恢复待核标记。':'已确认原值无误。PPT 和海报不再添加对应待核标黄或批注，Word 已有审阅修订继续保留。');}catch(error){showError(error);}};actions.append(confirm,el('span','','确认原值；保留 Word 本轮修订'));row.append(actions);}
 return row;
}
function renderFacts(){discardSupplementPreview();const box=$('facts-list');box.replaceChildren();let count=0;for(const group of groups()){const fs=groupFacts(group),section=sectionReview(group),pending=fs.some(f=>f.status!=='verified')||sectionPending(section);if(!(state.filter==='all'||state.filter==='focus'&&(pending||sectionConfirmed(section)||fs.some(f=>factChanged(f)||isConfirmed(f)||f.source==='user'||f.source==='user_custom_section'))||state.filter==='pending'&&pending))continue;box.append(fs[0]?.source==='manager'?renderManagerFact(fs):state.reviewApi?.getSupplementGroups?renderGroup(group,section):renderFact(fs[0]));count++;}if(!count)box.append(el('div','empty-state','当前筛选下没有要素。'));document.querySelectorAll('[data-filter]').forEach(b=>b.classList.toggle('selected',b.dataset.filter===state.filter));}
function showSupplementPreview(preview,origin={}){
 const box=origin.box||$('supplement-preview');openSupplementFeedback(box);state.supplementPreview=preview;box.append(el('h4','','识别结果'+(origin.label?' · '+origin.label:'')));
 if(origin.note)box.append(el('p','supplement-note',origin.note));const unique=new Set();for(const update of preview.updates||[]){const fact=state.task.facts.find(f=>f.id===update.factId),text=readableFact({...fact,...update});if(!unique.has(text)){box.append(el('p','parsed-value',text));unique.add(text);}}
 for(const warning of preview.warnings||[])box.append(el('p','supplement-warning',typeof warning==='string'?warning:warning.message||String(warning)));
 for(const note of preview.notes||[])box.append(el('p','supplement-note',typeof note==='string'?note:note.message||String(note)));
 for(const change of preview.changes||[]){if(change.materialName)box.append(el('strong','',change.materialName));box.append(el('p','supplement-note','当前：'+change.before),el('p','parsed-value','拟修改为：'+change.after));}
 if(!preview.updates?.length){box.append(el('p','supplement-note','没有可直接应用的字段。可调整文字，或在对应要素下补充；当前核对结果保持原样。'));box.scrollIntoView({block:'nearest'});return;}
 const actions=el('div','preview-actions'),apply=el('button','primary-button','确认并应用补充'),cancel=el('button','text-button','取消');apply.type=cancel.type='button';apply.onclick=()=>{if(state.busy||state.supplementPreview!==preview)return;try{state.task=state.reviewApi.applySupplement(state.task,preview);reviewEdited('已应用补充资料并同步相关物料。Word 中的修改继续保留为审阅修订。');}catch(error){showSupplementIssue(error,box);}};cancel.onclick=discardSupplementPreview;actions.append(apply,cancel);box.append(actions);box.scrollIntoView({block:'nearest'});
}
function switchSupplement(mode){if(state.busy)return;for(const kind of ['text','excel']){$('supplement-'+kind+'-panel').hidden=kind!==mode;$('supplement-'+kind+'-tab').classList.toggle('selected',kind===mode);$('supplement-'+kind+'-tab').setAttribute('aria-pressed',String(kind===mode));}discardSupplementPreview();}
$('supplement-text-tab').onclick=()=>switchSupplement('text');$('supplement-excel-tab').onclick=()=>switchSupplement('excel');
function syncPerformanceMode(){
 const tiered=$('performance-mode').value==='tiered';
 $('performance-threshold2-wrap').hidden=$('performance-ratio2-wrap').hidden=!tiered;
 $('performance-threshold1-label').textContent=tiered?'第一个基准（%）':'计提基准（%）';$('performance-ratio1-label').textContent=tiered?'第一档提取比例（%）':'提取比例（%）';
 $('performance-pattern').textContent=tiered?'超过第一个基准、不超过第二个基准的部分按第一档比例计提；超过第二个基准的部分按第二档比例计提。':'超过计提基准的收益部分，按提取比例计提。仅替换模板对应数字。';
}
$('performance-mode').value='single';syncPerformanceMode();
$('performance-mode').addEventListener('change',()=>{discardSupplementPreview();syncPerformanceMode();});
for(const key of performanceKeys)$('performance-'+key).addEventListener('input',discardSupplementPreview);
$('supplement-parse-performance').onclick=()=>{
 if(state.busy||!state.task)return;discardSupplementPreview();$('supplement-notice').hidden=true;
 const mode=$('performance-mode').value,keys=mode==='tiered'?performanceKeys:['threshold1','ratio1'],values=Object.fromEntries(keys.map(k=>[k,$('performance-'+k).value]));
 if(keys.every(k=>!values[k].trim())){$('error-panel').hidden=true;$('supplement-notice').textContent='未填写业绩报酬数字，核对结果保持原样。';$('supplement-notice').hidden=false;return;}
 try{showSupplementPreview(state.reviewApi.previewPerformanceNumbers(state.task,{mode,...values}),{label:'业绩报酬',note:'仅替换模板对应的计提基准与提取比例，保留原话术；其他输入行保持不动。'});}catch(error){showError(error);}
};
for(const kind of supplementKinds){
 const input=$('supplement-'+kind);input.addEventListener('input',discardSupplementPreview);
 $('supplement-parse-'+kind).onclick=()=>{
  if(state.busy||!state.task)return;discardSupplementPreview();$('supplement-notice').hidden=true;
  if(!input.value.trim()){$('error-panel').hidden=true;$('supplement-notice').textContent='本行未填写内容，核对结果保持原样。';$('supplement-notice').hidden=false;return;}
  try{showSupplementPreview(state.reviewApi.previewSupplementField(state.task,kind,input.value),{label:({date:'本次开放日',performance:'业绩报酬',benchmark:'业绩比较基准',code:'产品代码',opening:'开放期'})[kind],note:'仅应用本行识别到的内容，其他输入行保持不动。'});}catch(error){showError(error);}
 };
}
function previewCalendarRow(){
 discardSupplementPreview();const row=state.calendar?.rows.find(r=>r.id===$('calendar-row').value),box=$('calendar-row-detail');box.replaceChildren();if(!row)return;
 const source=row.source||{};box.append(el('strong','',row.productName),el('p','supplement-note',`${source.file||state.calendar.name||'产品日历'} · ${source.sheet||''} · 第 ${source.row} 行`));
 const matched=state.calendarMatches.includes(row.id);if(!matched)box.append(el('p','supplement-warning','所选行与合同产品不同。请核对本次要采用的数值后再应用，目标产品仍以合同为准。'));
 const range=state.calendarOpeningRange(row);
 const pairs=[['产品代码',row.productCode],['本次开放日',range.value],['业绩报酬',row.performanceText],['业绩比较基准',row.benchmark]];for(const [label,value]of pairs)if(value)box.append(el('p','',label+'：'+value));
 if(range.warning)box.append(el('p','supplement-warning',range.warning));
 else if(range.end&&!range.start)box.append(el('p','supplement-note','开放截止日：'+range.end+'；按已明确的开放天数与股票交易日反推区间，请核对下方预览。'));
 if(row.openingDateNote)box.append(el('p','supplement-note',row.openingDateNote));for(const warning of row.warnings||[])box.append(el('p','supplement-warning',typeof warning==='string'?warning:warning.message||String(warning)));
 const text=state.calendarSupplementText(row);
 try{const preview=state.reviewApi.previewSupplementText(state.task,text);preview.source={...source,productName:row.productName};showSupplementPreview(preview,{label:'Excel 产品行',note:matched?'请核对下列数值，确认后同步到相关物料。':'这些数值来自另一产品行，仅在确认适用时应用。'});}catch(error){box.append(el('p','supplement-warning',error.message||String(error)));}
}
$('calendar-row').addEventListener('change',previewCalendarRow);
function clearCalendarImport(){state.calendar=null;state.calendarMatches=[];discardSupplementPreview();$('calendar-row').replaceChildren();$('calendar-row').value='';$('calendar-row').hidden=true;$('calendar-row-label').hidden=true;$('calendar-row-detail').replaceChildren();$('calendar-input').value='';}
async function selectCalendar(files){
 if(state.busy||!state.task)return;const list=Array.from(files||[]);if(!list.length)return;
 clearCalendarImport();$('supplement-notice').hidden=true;$('error-panel').hidden=true;$('calendar-status').textContent='此文件未导入，已有补充与核对结果保留。';
 if(list.length!==1){showError(new Error('每次请拖入或选择一份 XLSX 产品日历。'));return;}
 const file=list[0];if(!/\.xlsx$/i.test(file.name)||file.size>30*1024*1024){showError(new Error('请选择30 MB以内的 XLSX 产品日历。'));return;}
 setBusy(true);$('calendar-status').textContent='正在本机读取产品日历…';await nextPaint();try{const {readProductCalendar,matchCalendarRows,calendarSupplementText,calendarOpeningRange}=await import('./calendar-import.js?v=20261027-custom-investment');state.calendarSupplementText=calendarSupplementText;state.calendarOpeningRange=calendarOpeningRange;const calendar=await readProductCalendar({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())});state.calendar=calendar;const matches=matchCalendarRows(calendar,state.task.target.name);state.calendarMatches=matches.map(r=>r.id);const select=$('calendar-row');select.replaceChildren();const blank=el('option','','请选择产品行');blank.value='';select.append(blank);for(const row of [...matches,...calendar.rows.filter(r=>!state.calendarMatches.includes(r.id))]){const option=el('option','',(state.calendarMatches.includes(row.id)?'匹配产品 · ':'')+row.productName+(row.productCode?' · '+row.productCode:'')+' · 第'+row.source.row+'行');option.value=row.id;select.append(option);}select.hidden=false;$('calendar-row-label').hidden=false;$('calendar-row-detail').replaceChildren();$('calendar-status').textContent=`已读取 ${calendar.rows.length} 条产品记录。`+(matches.length===1?'已定位到合同产品，请核对解析结果。':matches.length>1?'同一产品有多条记录，请选择本次所用行。':'未找到合同产品的对应行，可手动选择并核对，或使用文字补充。')+(calendar.warnings?.length?' '+calendar.warnings.join('；'):'');select.value=matches.length===1?matches[0].id:'';if(select.value)previewCalendarRow();}
 catch(error){$('calendar-status').textContent='此文件未导入，已有补充与核对结果保留。';showError(error);}finally{$('calendar-input').value='';setBusy(false);}
}
$('calendar-input').addEventListener('change',e=>selectCalendar(e.target.files));
const calendarDrop=$('calendar-drop');
const isFileDrag=e=>Array.from(e.dataTransfer?.types||[]).includes('Files')||e.dataTransfer?.files?.length;
calendarDrop.addEventListener('dragover',e=>{if(!isFileDrag(e))return;e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect=state.busy?'none':'copy';if(!state.busy&&state.task)calendarDrop.classList.add('dragging');});
calendarDrop.addEventListener('dragleave',e=>{if(!calendarDrop.contains(e.relatedTarget))calendarDrop.classList.remove('dragging');});
calendarDrop.addEventListener('drop',e=>{e.preventDefault();calendarDrop.classList.remove('dragging');return selectCalendar(e.dataTransfer?.files);});
function renderReview(){
 const task=state.task;$('target-name').textContent=task.target.name;$('review-subtitle').textContent=`${task.templates.length} 份模板 · 同口径要素共享核对 · 制作日期 ${task.productionDate||state.productionDate} · 文件名月份 ${task.month}`+(state.company?' · 公司介绍（'+companyVersionLabel()+'）：'+state.company.name:'');
 const only=task.updateMode==='introductions';$('supplement-panel').hidden=only;$('review-eyebrow').textContent=only?'本次更新范围':'本次目标产品';$('facts-heading').textContent=only?'介绍核对':'要素核对';$('delivery-explanation').textContent=only?'仅更新公司和经理介绍；Word 保留本轮审阅修订，PPT 保留对应介绍页。':'黄色表示原值仍需确认，不影响生成审阅稿。下载后可在 Word 中查看修订与批注。';
 if(only)$('review-subtitle').textContent=`${task.templates.length} 份模板 · 保留各模板原产品名称 · 制作日期 ${task.productionDate||state.productionDate} · 文件名月份 ${task.month}`;
 $('count-facts').textContent=groups().length;$('count-updates').textContent=task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0);$('count-pending').textContent=groups().filter(g=>groupFacts(g).some(f=>f.status!=='verified')||sectionPending(sectionReview(g))).length;$('count-materials').textContent=task.materials.length;
 const warnings=[...(task.warnings||[]),...(task.inputIssues||[]).map(i=>`${i.sourceName}：${i.reason==='failed'?'未处理，':''}${i.message}`),...task.materials.flatMap(m=>(m.warnings||[]).map(w=>m.template.name+'：'+(typeof w==='string'?w:w.message||String(w))))];$('warnings').replaceChildren();for(const warning of [...new Set(warnings)])$('warnings').append(el('p','',typeof warning==='string'?warning:warning.message||String(warning)));
 const deliveryNames=uniqueDeliveryNames(task.materials.map(materialName));
 $('delivery-list').replaceChildren();for(const [index,m] of task.materials.entries()){const item=el('div','delivery-item');const hasProfiles=['company','manager'].some(key=>m.profilePages?.[key]?.length),kindLabel=m.kind==='ppt'?(only?'PPT · 公司及经理介绍':hasProfiles?'PPT · 产品要素及介绍':'PPT · 产品要素'):({word:'WORD · 一页通',poster:'PPT · 海报'})[m.kind]||'物料';item.append(el('span','kind',kindLabel),el('strong','',deliveryNames[index]));const page=m.kind==='ppt'?`仅输出源第 ${(m.selectedPages||[]).join('、')} 页`:(m.kind==='poster'?'保留海报全部页面':'接受旧修订，保留本轮新修订');item.append(el('p','',page));$('delivery-list').append(item);}renderFacts();
}
document.querySelectorAll('[data-filter]').forEach(b=>b.onclick=()=>{state.filter=b.dataset.filter;renderFacts();});
function downloadLink(bytes,mime,name,parent){const url=URL.createObjectURL(new Blob([bytes],{type:mime}));state.urls.push(url);const link=el('a','',name);link.href=url;link.download=name;parent.append(link);return link;}
function generationStatus(message,kind='progress'){const box=$('generation-status');box.hidden=false;box.dataset.kind=kind;box.textContent=message;}
function deliveryIssues(task,result){return [...(task.inputIssues||[]).filter(i=>i.reason==='failed'),...(result.failures||[])];}
function deliveryNotes(task,result){return [...(task.warnings||[]).map(w=>typeof w==='string'?w:w.message||String(w)),...(task.inputIssues||[]).map(i=>`${i.sourceName}：${i.reason==='failed'?'未生成，':''}${i.message}`),...task.materials.flatMap(m=>(m.warnings||[]).map(w=>m.template.name+'：'+(typeof w==='string'?w:w.message||String(w)))),...result.files.flatMap(f=>(f.report?.repairs||[]).map(r=>f.name+'：'+(typeof r==='string'?r:r.message||r.kind))),...(result.failures||[]).map(f=>`未生成 ${f.sourceName}：${f.message}`)];}
function reviewText(task,result){
 const lines=['物料审阅清单',`目标产品：${task.target.name}`,`制作日期：${task.productionDate||state.productionDate}`,`文件名月份：${task.month}`,`经理介绍：${state.manager?.name||'未使用'}`,`公司介绍：${state.company?companyVersionLabel()+' · '+state.company.name:'未使用'}`,'',task.updateMode==='introductions'?'本次仅更新公司和经理介绍；保留各模板产品名称和产品要素。':'本次更新产品要素及匹配到的公司介绍、经理介绍；历史业绩保持原样。','黄色表示原值仍需确认，不代表合同已核实。','','输出文件：',...result.files.map(f=>f.name),'','核对事项：'];
 const sections=groups().map(group=>({group,section:sectionReview(group)})).filter(x=>x.section),coveredIds=new Set(sections.flatMap(x=>x.group.factIds));
 for(const f of task.facts){lines.push(`${f.label}｜${f.status==='verified'?(isConfirmed(f)?'用户确认原值':f.source==='user_calendar'?'产品日历补充':f.source==='user_custom_section'?'用户自定义整项替换':f.source==='user'?'用户补充':f.source==='company'?'公司介绍依据':f.source==='manager'?'经理介绍 Excel 依据':'合同核对'):'保留待核'}｜${f.value??'保留各模板原值'}`,f.reason||'');for(const e of(f.evidence||[]).filter(e=>!coveredIds.has(f.id)||sourceTitle(e)!=='合同依据').slice(0,2))lines.push(sourceTitle(e)+'：'+e.quote);lines.push('');}
 for(const {group,section}of sections){
  const custom=customSectionFor(group);
  if(custom){lines.push('',section.label+'｜用户自定义整项替换','以下内容由用户确认，不是合同自动提取结论：',custom.text,...custom.entries.map(e=>'应用位置：'+e.name));for(const evidence of section.evidence||[])lines.push('合同原文供核对：',evidence.quote);continue;}
  lines.push('',section.label+'｜'+(sectionPending(section)?'章节覆盖待核':sectionConfirmed(section)?'已核对差异，保留当前表述':'完整章节核对'),section.summary);
  for(const material of section.materials||[]){lines.push(material.name+`：已对应 ${material.matched}/${material.total} 项`);for(const item of [...(material.items||[]),...(material.extra||[])])lines.push((sectionStatus[item.status]||'需核对')+'｜'+(item.title||''),'合同：'+(item.contractText||'未找到明确对应条款'),'当前物料：'+(item.templateText||'未找到明确对应表述'),item.reason||'');}
  for(const evidence of section.evidence||[])lines.push('完整合同依据：',evidence.quote);
 }
 const notes=deliveryNotes(task,result);if(notes.length)lines.push('','文件处理提示：',...notes);
 return lines.join('\r\n');
}
async function generateFiles({autoDownload=true}={}){
 if(!state.task||state.busy||!state.monthValid)return;setBusy(true);$('error-panel').hidden=true;clearDownloads();generationStatus('正在生成物料，请稍候…');$('generate-button').textContent='正在生成…';await nextPaint();
 try{
  const {writeTask}=await import('./office-writer.js?v=20261029-ppt-profiles');const result=await writeTask(state.task);generationStatus(`已完成 ${result.files.length} 份物料，正在准备下载包…`);const zip=new globalThis.JSZip();for(const file of result.files)zip.file(file.name,file.bytes);zip.file('审阅清单.txt','\uFEFF'+reviewText(state.task,result));const bundle=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  const issues=deliveryIssues(state.task,result),box=$('download-list');box.append(el('p','download-success',`已生成 ${result.files.length} 份物料，可下载审阅。`));const bundleLink=downloadLink(bundle,'application/zip',`${state.task.target.shortName||'产品物料'}_${state.task.month}.zip`,box);for(const file of result.files)downloadLink(file.bytes,file.mime,file.name,box);changeStep('download');
  for(const note of deliveryNotes(state.task,result))box.append(el('p','delivery-note',note));
  generationStatus((issues.length?`已生成 ${result.files.length} 份，${issues.length} 份未生成，原因见下方。`:'已生成。')+(autoDownload?'若未自动开始下载，请点击下方下载包或单个文件。':'请直接点击下方下载包或单个文件。'),issues.length?'warning':'success');box.scrollIntoView({block:'nearest'});
  // A browser may decline an automatic download after asynchronous work. The
  // visible direct links must remain usable regardless of that browser policy.
  if(autoDownload)try{bundleLink.click();}catch{}
 }catch(error){generationStatus('生成失败：'+(error.message||String(error)),'error');showError(error);}finally{setBusy(false);$('generate-button').textContent='生成并下载物料';$('progress-panel').hidden=true;}
}
$('generate-button').onclick=()=>generateFiles();
updateMonth(true);refreshSelection();const appReady=Promise.all([restoreCompany(),restoreManager()]);
