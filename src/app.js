import {currentMonth,deliveryName,resolveProductionMonth,monthInputValue} from './naming.js';

const $=id=>document.getElementById(id);
const state={contract:null,templates:[],task:null,filter:'focus',busy:false,urls:[],month:currentMonth(),manualMonth:undefined,monthValid:true,started:new Date(),run:0,company:null,companyReady:false,companyPersisted:false,companyStored:false,companyWarning:'',store:null};
const el=(tag,className,text)=>{const n=document.createElement(tag);if(className)n.className=className;if(text!==undefined)n.textContent=text;return n;};
const readableSize=bytes=>bytes>=1048576?(bytes/1048576).toFixed(1)+' MB':Math.max(1,Math.round(bytes/1024))+' KB';
const nextPaint=()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
function showError(error){$('error-message').textContent=error?.message||String(error);$('error-panel').hidden=false;}
function clearDownloads(){for(const url of state.urls)URL.revokeObjectURL(url);state.urls=[];$('download-list').replaceChildren();}
function changeStep(name){for(const key of ['upload','review','download']){$('step-'+key).classList.toggle('active',key===name);$('step-'+key).classList.toggle('done',name==='download'&&key!=='download'||name==='review'&&key==='upload');}}
function updateMonth(reset=false){state.started=new Date();if(reset)state.manualMonth=undefined;state.month=resolveProductionMonth({manualValue:state.manualMonth,now:state.started});state.monthValid=true;$('month-value').value=monthInputValue(state.month);}
function changeMonth(reset=false){
 if(state.busy)return;
 clearDownloads();
 try{state.manualMonth=reset?undefined:$('month-value').value;updateMonth(reset);$('error-panel').hidden=true;if(state.task){state.task.month=state.month;clearDownloads();changeStep('review');renderReview();}}
 catch(error){state.monthValid=false;showError(error);}refreshSelection();
}
$('month-value').addEventListener('change',()=>changeMonth());
$('month-reset').onclick=()=>changeMonth(true);
function refreshSelection(){
 for(const [key,files] of [['contract-files',state.contract?[state.contract]:[]],['template-files',state.templates]]){
  const box=$(key);box.replaceChildren();
  for(const file of files){const row=el('div','file-entry');row.append(el('span','',file.name),el('small','',readableSize(file.size)));const remove=el('button','', '×');remove.type='button';remove.setAttribute('aria-label','移除 '+file.name);remove.disabled=state.busy;remove.onclick=()=>{if(key==='contract-files'){state.contract=null;$('contract-input').value='';}else state.templates=state.templates.filter(f=>f!==file);invalidate();refreshSelection();};row.append(remove);box.append(row);}
 }
 $('analyze-button').disabled=state.busy||!state.companyReady||!state.monthValid||!state.contract||state.templates.length===0;
 $('generate-button').disabled=state.busy||!state.monthValid||!state.task;
 $('company-input').disabled=state.busy||!state.companyReady;$('company-clear').disabled=state.busy||!state.companyReady;
 $('selection-hint').textContent=state.contract&&state.templates.length?`已选择合同和 ${state.templates.length} 份物料模板。`:'选好文件后，开始核对。';
}
function invalidate(){state.task=null;$('review-section').hidden=true;clearDownloads();changeStep('upload');$('error-panel').hidden=true;}
function showCompany(){
 const prefix=state.company?`当前版本：${state.company.name}。${state.companyPersisted?'已保存在此浏览器，下次自动沿用。':'仅本次使用，尚未保存。'}`:'尚未上传公司介绍，本轮只更新产品要素。';
 $('company-status').textContent=prefix+(state.companyWarning?' '+state.companyWarning:'');$('company-clear').hidden=!state.company&&!state.companyStored;
}
async function restoreCompany(){
 try{const {createProfileStore}=await import('./profile-store.js');state.store=createProfileStore();const record=await state.store.load();if(record){state.companyStored=true;const {parseCompanyProfile}=await import('./company-profile.js');const profile=await parseCompanyProfile({name:record.name,bytes:record.bytes});if(profile.sha256!==record.sha256)throw new Error('已保存的公司介绍校验不一致，请重新上传。');state.company=profile;state.companyPersisted=true;}}
 catch(error){if(error.code==='PROFILE_CACHE_INVALID')state.companyStored=true;state.companyWarning='未能恢复已保存版本：'+(error.message||String(error));}
 finally{state.companyReady=true;showCompany();refreshSelection();}
}
async function selectCompany(files){
 if(state.busy)return;if(!state.companyReady){showError(new Error('正在读取上次保存的公司介绍，请稍后再上传。'));return;}const list=[...files];if(!list.length)return;
 if(list.length!==1||! /\.docx$/i.test(list[0].name)){showError(new Error('公司介绍请选择一份 DOCX 文件。'));return;}
 const file=list[0];if(file.size>60*1024*1024){showError(new Error('公司介绍文件请控制在60 MB以内。'));return;}
 setBusy(true);progress('正在读取公司介绍','识别两种版式，成功后替换此浏览器保存的版本。');await nextPaint();
 try{const {parseCompanyProfile}=await import('./company-profile.js');const bytes=new Uint8Array(await file.arrayBuffer());const profile=await parseCompanyProfile({name:file.name,bytes});let persisted=false;state.companyWarning='';
  try{if(!state.store){const {createProfileStore}=await import('./profile-store.js');state.store=createProfileStore();}await state.store.save({schemaVersion:1,name:file.name,sha256:profile.sha256,uploadedAt:new Date().toISOString(),bytes,profile});persisted=true;state.companyStored=true;}
  catch(error){state.companyWarning='此版本本次可用，但未能保存；下次请重新上传。'+(error.message||'');}
  state.company=profile;state.companyPersisted=persisted;invalidate();showCompany();
 }catch(error){showError(new Error('公司介绍未更新，仍保留此前版本。'+(error.message||String(error))));}
 finally{$('company-input').value='';setBusy(false);$('progress-panel').hidden=true;}
}
$('company-input').addEventListener('change',e=>selectCompany(e.target.files));
$('company-clear').onclick=async()=>{if(state.busy)return;setBusy(true);try{if(state.companyStored){if(!state.store)throw new Error('尚未连接本机保存空间，请刷新后再清除。');await state.store.clear();}state.company=null;state.companyStored=false;state.companyPersisted=false;state.companyWarning='';invalidate();showCompany();}catch(error){showError(error);}finally{setBusy(false);}};
function selectFiles(kind,files){
 if(state.busy)return;
 const list=[...files];if(!list.length)return;
 const allowed=kind==='contract'?/\.(docx|doc)$/i:/\.(docx|pptx)$/i;
 if(list.some(f=>!allowed.test(f.name))){showError(new Error(kind==='contract'?'请选择 DOCX 或 DOC 合同。':'模板支持 DOCX 和 PPTX 文件。'));return;}
 if(list.some(f=>f.size>60*1024*1024)){showError(new Error('单个文件请控制在60 MB以内。'));return;}
 if(kind==='contract'){if(list.length!==1){showError(new Error('每次请选择一份主合同。'));return;}state.contract=list[0];}
 else{for(const file of list)if(!state.templates.some(f=>f.name===file.name&&f.size===file.size&&f.lastModified===file.lastModified))state.templates.push(file);}
 invalidate();refreshSelection();
}
$('contract-input').addEventListener('change',e=>selectFiles('contract',e.target.files));
$('templates-input').addEventListener('change',e=>selectFiles('templates',e.target.files));
for(const [id,kind]of [['contract-drop','contract'],['templates-drop','templates'],['company-drop','company']]){
 const box=$(id);box.addEventListener('dragover',e=>{e.preventDefault();box.classList.add('dragging');});box.addEventListener('dragleave',()=>box.classList.remove('dragging'));box.addEventListener('drop',e=>{e.preventDefault();box.classList.remove('dragging');if(kind==='company')selectCompany(e.dataTransfer.files);else selectFiles(kind,e.dataTransfer.files);});
}
function progress(title,detail){$('progress-panel').hidden=false;$('progress-title').textContent=title;$('progress-detail').textContent=detail;}
function setBusy(busy){state.busy=busy;for(const id of ['contract-input','templates-input','company-input','company-clear','month-value','month-reset','reset-button'])$(id).disabled=busy;refreshSelection();}
$('dismiss-error').onclick=()=>$('error-panel').hidden=true;
$('reset-button').onclick=()=>{state.run++;state.contract=null;state.templates=[];$('contract-input').value='';$('templates-input').value='';invalidate();updateMonth(true);refreshSelection();$('upload-section').hidden=false;window.scrollTo({top:0,behavior:'smooth'});};
$('analyze-button').onclick=async()=>{
 if(state.busy||!state.companyReady)return;try{updateMonth();}catch(error){showError(error);return;}const run=++state.run;invalidate();setBusy(true);progress('正在读取合同','建立本次合同索引，保留文字出处。');await nextPaint();
 try{
  const [{readContract,readTemplate},{analyzeMaterials}]=await Promise.all([import('./office-reader.js'),import('./review-engine.js')]);
  const contract=await readContract({name:state.contract.name,bytes:new Uint8Array(await state.contract.arrayBuffer())});
  const templates=[];
  for(const [i,file]of state.templates.entries()){progress(`正在读取模板 ${i+1} / ${state.templates.length}`,file.name);await nextPaint();templates.push(await readTemplate({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())}));}
  if(new Set(templates.map(t=>t.sha256)).size!==templates.length)throw new Error('选择了内容相同的重复模板，请移除重复文件。');
  if(templates.some(t=>t.sha256===contract.sha256))throw new Error('合同不能同时作为待更新模板，请检查文件选择。');
  progress('正在统一核对产品要素','合并共同需求，查找合同依据，再定位各物料的改动。');await nextPaint();
  const task=await analyzeMaterials(contract,templates,{now:state.started,month:state.month});task.now=state.started;
  if(state.company){progress('正在核对一页通公司介绍',state.company.name);await nextPaint();const {applyCompanyProfile}=await import('./company-profile.js');await applyCompanyProfile(task,state.company);}
  if(run!==state.run)return;
  state.task=task;renderReview();$('upload-section').hidden=true;$('review-section').hidden=false;changeStep('review');
 }catch(error){showError(error);}finally{if(run===state.run){setBusy(false);$('progress-panel').hidden=true;}}
};
function factChanged(fact){return state.task.materials.some(m=>m.operations.some(o=>o.kind==='replace'&&o.factIds?.includes(fact.id)));}
function factStatus(fact){return fact.status==='verified'?(factChanged(fact)?'changed':'verified'):(fact.status==='conflict'?'conflict':'pending');}
function materialName(m){try{return deliveryName(m.template.name,{targetName:state.task.target.name,templateProductName:m.templateProductName,kind:m.kind,month:state.task.month});}catch(error){return m.template.name+'（命名待检查）';}}
function sourceTitle(evidence){if(String(evidence.blockId).startsWith('company-'))return '最新公司介绍依据';if(evidence.blockId==='user-supplement')return '手动补充';if(String(evidence.blockId).startsWith('template-'))return '模板中的说明';return '合同依据';}
function renderFact(fact){
 const row=el('article','fact-row');row.dataset.factId=fact.id;const status=factStatus(fact);const head=el('div','fact-heading');head.append(el('strong','',fact.label||fact.key));head.append(el('span','badge '+status,({changed:'已修改',verified:'原样保留',pending:'保留待核',conflict:'需要复核'})[status]));const count=new Set((fact.occurrences||[]).map(o=>o.materialId+'|'+o.paragraphId)).size;head.append(el('small','',`关联 ${count} 处`));row.append(head);
 const value=fact.value===null||fact.value===undefined?'保留各模板原值':String(fact.value);row.append(el('p','fact-value',value.length>220?value.slice(0,220)+'…':value));if(fact.status!=='verified')row.append(el('p','fact-reason',fact.reasonType==='contract_missing'?'合同未给出具体值，保留原值。':'需要人工核对，当前未自动确认。'));row.append(el('p','fact-reason',fact.reason||''));
 const details=el('details');details.append(el('summary','', '查看依据与物料位置'));
 for(const evidence of (fact.evidence||[]).slice(0,8)){details.append(el('p','',sourceTitle(evidence)));details.append(el('blockquote','',evidence.quote||''));}
 if(!fact.evidence?.length)details.append(el('p','','尚无可唯一确认的合同依据，当前保留模板值。'));
 for(const m of state.task.materials){const fields=m.fields.filter(f=>f.factIds?.includes(fact.id));for(const field of fields){const occurrence=el('div','occurrence');occurrence.append(el('strong','',m.template.name+' · '+field.label));const before=field.before||'';const after=field.after??before;if(before!==after){const b=el('p');b.append(el('del','',before));const a=el('p');a.append(el('ins','',after));occurrence.append(b,a);}else occurrence.append(el('p','',before));details.append(occurrence);}}
 row.append(details);
 if(fact.status!=='verified'){
  const form=el('form','supplement');const input=el('input');input.placeholder='有明确补充资料时，可在这里填写';input.setAttribute('aria-label','补充'+(fact.label||fact.key));const button=el('button','','确认补充');button.type='submit';form.append(input,button);form.onsubmit=async e=>{e.preventDefault();if(!input.value.trim())return;try{const {setSupplement}=await import('./review-engine.js');state.task=await setSupplement(state.task,fact.id,input.value);clearDownloads();$('error-panel').hidden=true;changeStep('review');renderReview();}catch(error){showError(error);}};row.append(form);
 }
 return row;
}
function renderFacts(){const box=$('facts-list');box.replaceChildren();const facts=state.task.facts.filter(f=>state.filter==='all'||state.filter==='focus'&&(f.status!=='verified'||factChanged(f))||state.filter==='pending'&&f.status!=='verified');for(const fact of facts)box.append(renderFact(fact));if(!facts.length)box.append(el('div','empty-state','当前筛选下没有要素。'));document.querySelectorAll('[data-filter]').forEach(b=>b.classList.toggle('selected',b.dataset.filter===state.filter));}
function renderReview(){
 const task=state.task;$('target-name').textContent=task.target.name;$('review-subtitle').textContent=`${task.templates.length} 份模板 · 同口径要素共享核对 · 文件名月份 ${task.month}`+(state.company?' · 公司介绍：'+state.company.name:'');
 $('count-facts').textContent=task.facts.length;$('count-updates').textContent=task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0);$('count-pending').textContent=task.facts.filter(f=>f.status!=='verified').length;$('count-materials').textContent=task.materials.length;
 const warnings=[...(task.warnings||[]),...task.materials.flatMap(m=>(m.warnings||[]).map(w=>m.template.name+'：'+(typeof w==='string'?w:w.message||String(w))))];$('warnings').replaceChildren();for(const warning of [...new Set(warnings)])$('warnings').append(el('p','',typeof warning==='string'?warning:warning.message||String(warning)));
 $('delivery-list').replaceChildren();for(const m of task.materials){const item=el('div','delivery-item');item.append(el('span','kind',({word:'WORD · 一页通',ppt:'PPT · 产品要素',poster:'PPT · 海报'})[m.kind]||'物料'),el('strong','',materialName(m)));const page=m.kind==='ppt'?`仅输出源第 ${(m.selectedPages||[]).join('、')} 页`:(m.kind==='poster'?'保留海报全部页面':'接受旧修订，保留本轮新修订');item.append(el('p','',page));$('delivery-list').append(item);}renderFacts();
}
document.querySelectorAll('[data-filter]').forEach(b=>b.onclick=()=>{state.filter=b.dataset.filter;renderFacts();});
function downloadLink(bytes,mime,name,parent,auto=false){const url=URL.createObjectURL(new Blob([bytes],{type:mime}));state.urls.push(url);const link=el('a','',name);link.href=url;link.download=name;parent.append(link);if(auto)link.click();return link;}
function reviewText(task,result){
 const lines=['物料审阅清单',`目标产品：${task.target.name}`,`制作月份：${task.month}`,`公司介绍：${state.company?.name||'未使用'}`,'','本次更新产品要素及匹配到的一页通公司介绍；历史业绩及人员介绍保持原样。','黄色表示原值仍需确认，不代表合同已核实。','','输出文件：',...result.files.map(f=>f.name),'','核对事项：'];
 for(const f of task.facts){lines.push(`${f.label}｜${f.status==='verified'?(f.source==='user'?'用户补充':f.source==='company'?'公司介绍依据':'合同核对'):'保留待核'}｜${f.value??'保留各模板原值'}`,f.reason||'');for(const e of(f.evidence||[]).slice(0,2))lines.push(sourceTitle(e)+'：'+e.quote);lines.push('');}
 return lines.join('\r\n');
}
$('generate-button').onclick=async()=>{
 if(!state.task||state.busy||!state.monthValid)return;setBusy(true);$('error-panel').hidden=true;clearDownloads();progress('正在生成物料','按已核对的差异回写，保留模板格式。');await nextPaint();
 try{
  const {writeTask}=await import('./office-writer.js');const result=await writeTask(state.task);const zip=new globalThis.JSZip();for(const file of result.files)zip.file(file.name,file.bytes);zip.file('审阅清单.txt','\uFEFF'+reviewText(state.task,result));const bundle=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  const box=$('download-list');box.append(el('p','download-success',`已生成 ${result.files.length} 份物料，可下载审阅。`));downloadLink(bundle,'application/zip',`${state.task.target.shortName||'产品物料'}_${state.task.month}.zip`,box,true);for(const file of result.files)downloadLink(file.bytes,file.mime,file.name,box);changeStep('download');
 }catch(error){showError(error);}finally{setBusy(false);$('progress-panel').hidden=true;}
};
updateMonth(true);refreshSelection();restoreCompany();
