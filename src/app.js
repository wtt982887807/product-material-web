import {currentMonth,deliveryName,resolveProductionMonth,monthInputValue,uniqueDeliveryNames} from './naming.js?v=20261026-drop';

const $=id=>document.getElementById(id);
const supplementKinds=['date','benchmark','code','opening'];
const performanceKeys=['threshold1','ratio1','threshold2','ratio2'];
const state={contract:null,templates:[],task:null,filter:'focus',busy:false,urls:[],month:currentMonth(),manualMonth:undefined,monthValid:true,started:new Date(),run:0,company:null,companyIsDefault:false,companyReady:false,companyPersisted:false,companyStored:false,companyWarning:'',store:null,reviewApi:null,calendar:null,calendarMatches:[],supplementPreview:null};
const el=(tag,className,text)=>{const n=document.createElement(tag);if(className)n.className=className;if(text!==undefined)n.textContent=text;return n;};
const readableSize=bytes=>bytes>=1048576?(bytes/1048576).toFixed(1)+' MB':Math.max(1,Math.round(bytes/1024))+' KB';
const nextPaint=()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
function showError(error){$('error-message').textContent=error?.message||String(error);$('error-panel').hidden=false;}
function clearDownloads(){for(const url of state.urls)URL.revokeObjectURL(url);state.urls=[];$('download-list').replaceChildren();$('generation-status').hidden=true;$('generation-status').textContent='';}
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
function invalidate(){state.task=null;state.calendar=null;state.calendarMatches=[];state.supplementPreview=null;$('calendar-input').value='';$('calendar-row').replaceChildren();$('calendar-row').hidden=true;$('calendar-row-label').hidden=true;$('calendar-row-detail').replaceChildren();$('calendar-status').textContent='Excel 在本机读取。根据合同产品匹配，保留来源行供核对。';for(const kind of supplementKinds)$('supplement-'+kind).value='';for(const key of performanceKeys)$('performance-'+key).value='';$('performance-mode').value='single';syncPerformanceMode();$('supplement-preview').hidden=true;$('supplement-notice').hidden=true;$('review-section').hidden=true;clearDownloads();changeStep('upload');$('error-panel').hidden=true;}
function showCompany(){
 const prefix=state.company?(state.companyIsDefault?`网页内置：${state.company.name}。未上传新版时自动使用。`:`当前版本：${state.company.name}。${state.companyPersisted?'已保存在此浏览器，下次自动沿用。':'仅本次使用，尚未保存。'}`):'公司介绍暂未加载，本轮只更新产品要素。';
 $('company-status').textContent=prefix+(state.companyWarning?' '+state.companyWarning:'');$('company-clear').hidden=(!state.company||state.companyIsDefault)&&!state.companyStored;
}
async function useDefaultCompany(){const {loadDefaultCompanyProfile}=await import('./default-company-profile.js');state.company=loadDefaultCompanyProfile();state.companyIsDefault=true;state.companyPersisted=false;}
async function restoreCompany(){
 try{
  await useDefaultCompany();const {createProfileStore}=await import('./profile-store.js');state.store=createProfileStore();const record=await state.store.load();
  if(record){
   state.companyStored=true;const {parseCompanyProfile}=await import('./company-profile.js');const profile=await parseCompanyProfile({name:record.name,bytes:record.bytes});
   if(profile.sha256!==record.sha256)throw new Error('已保存的公司介绍校验不一致，请重新上传。');
   const {isSupersededCompanyProfile}=await import('./default-company-profile.js');
   if(isSupersededCompanyProfile(profile)){
    state.companyWarning='已自动采用规模日期更正后的公司介绍。';
    try{state.companyStored=!await state.store.clearIfSha256(profile.sha256);}catch{state.companyWarning+='旧版本缓存暂未清除，但不会覆盖更正后的内置版本。';}
   }else{state.company=profile;state.companyIsDefault=false;state.companyPersisted=true;}
  }
 }
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
  state.company=profile;state.companyIsDefault=false;state.companyPersisted=persisted;invalidate();showCompany();
 }catch(error){showError(new Error('公司介绍未更新，仍保留此前版本。'+(error.message||String(error))));}
 finally{$('company-input').value='';setBusy(false);$('progress-panel').hidden=true;}
}
$('company-input').addEventListener('change',e=>selectCompany(e.target.files));
$('company-clear').onclick=async()=>{if(state.busy)return;setBusy(true);try{if(state.companyStored){if(!state.store)throw new Error('尚未连接本机保存空间，请刷新后再清除。');await state.store.clear();}state.companyStored=false;await useDefaultCompany();state.companyWarning='';invalidate();showCompany();}catch(error){showError(error);}finally{setBusy(false);}};
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
for(const [id,kind]of [['contract-drop','contract'],['templates-drop','templates'],['company-drop','company']]){
 const box=$(id);box.addEventListener('dragover',e=>{e.preventDefault();box.classList.add('dragging');});box.addEventListener('dragleave',()=>box.classList.remove('dragging'));box.addEventListener('drop',e=>{e.preventDefault();box.classList.remove('dragging');if(kind==='company')selectCompany(e.dataTransfer.files);else selectFiles(kind,e.dataTransfer.files);});
}
function progress(title,detail){$('progress-panel').hidden=false;$('progress-title').textContent=title;$('progress-detail').textContent=detail;}
function setBusy(busy){state.busy=busy;$('calendar-drop').setAttribute('aria-busy',String(busy));if(busy)$('calendar-drop').classList.remove('dragging');for(const id of ['contract-input','templates-input','company-input','company-clear','month-value','month-reset','reset-button','calendar-input','calendar-row',...supplementKinds.flatMap(k=>['supplement-'+k,'supplement-parse-'+k]),...performanceKeys.map(k=>'performance-'+k),'performance-mode','supplement-parse-performance','supplement-text-tab','supplement-excel-tab'])$(id).disabled=busy;document.querySelectorAll('#facts-list input,#facts-list textarea,#facts-list button,#supplement-preview button').forEach(n=>n.disabled=busy);refreshSelection();}
$('dismiss-error').onclick=()=>$('error-panel').hidden=true;
$('reset-button').onclick=()=>{state.run++;state.contract=null;state.templates=[];$('contract-input').value='';$('templates-input').value='';invalidate();updateMonth(true);refreshSelection();$('upload-section').hidden=false;window.scrollTo({top:0,behavior:'smooth'});};
$('analyze-button').onclick=async()=>{
 if(state.busy||!state.companyReady)return;try{updateMonth();}catch(error){showError(error);return;}const run=++state.run;invalidate();setBusy(true);progress('正在读取合同','建立本次合同索引，保留文字出处。');await nextPaint();
 try{
  const [{readWordContract,readWordTemplate},reviewApi]=await Promise.all([import('./word-input.js'),import('./review-engine.js?v=20261026-numbers')]);state.reviewApi=reviewApi;const {analyzeMaterials}=reviewApi;
  const contract=await readWordContract({name:state.contract.name,bytes:new Uint8Array(await state.contract.arrayBuffer())},{onProgress:progress});
  const templates=[],inputIssues=[],seen=new Set();
  for(const [i,file]of state.templates.entries()){
   progress(`正在准备模板 ${i+1} / ${state.templates.length}`,`接受旧修订，按需转换 DOCX：${file.name}`);await nextPaint();
   try{
    const template=await readWordTemplate({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())},{onProgress:progress});
    const hash=template.sourceSha256||template.sha256;
    if(hash===contract.sha256){inputIssues.push({sourceName:file.name,reason:'contract',message:'与合同是同一份文件，已从模板中跳过。'});continue;}
    if(seen.has(hash)){inputIssues.push({sourceName:file.name,reason:'duplicate',message:'与已选模板内容相同，已合并重复项。'});continue;}
    seen.add(hash);templates.push(template);
   }catch(error){inputIssues.push({sourceName:file.name,reason:'failed',message:error.message||String(error)});}
  }
  if(!templates.length)throw new Error('没有可更新的模板。'+inputIssues.map(i=>`${i.sourceName}：${i.message}`).join('\n'));
  progress('正在统一核对产品要素','合并共同需求，查找合同依据，再定位各物料的改动。');await nextPaint();
  const task=await analyzeMaterials(contract,templates,{now:state.started,month:state.month});task.now=state.started;task.inputIssues=inputIssues;
  if(state.company){progress('正在核对一页通公司介绍',state.company.name);await nextPaint();const {applyCompanyProfile}=await import('./company-profile.js');await applyCompanyProfile(task,state.company);}
  if(run!==state.run)return;
  state.task=task;const feeGroups=groups().filter(g=>g.kind==='performance');if(feeGroups.length===1&&groupFacts(feeGroups[0]).some(f=>f.key==='performance_R2'))$('performance-mode').value='tiered';syncPerformanceMode();renderReview();$('upload-section').hidden=true;$('review-section').hidden=false;changeStep('review');
 }catch(error){showError(error);}finally{if(run===state.run){setBusy(false);$('progress-panel').hidden=true;}}
};
function factChanged(fact){return state.task.materials.some(m=>m.operations.some(o=>o.kind==='replace'&&o.factIds?.includes(fact.id)));}
function factStatus(fact){return fact.status==='verified'?(factChanged(fact)?'changed':'verified'):(fact.status==='conflict'?'conflict':'pending');}
function materialName(m){try{return deliveryName(m.template.name,{targetName:state.task.target.name,templateProductName:m.templateProductName,kind:m.kind,month:state.task.month});}catch(error){return m.template.name+'（命名待检查）';}}
function sourceTitle(evidence){if(String(evidence.blockId).startsWith('company-'))return '最新公司介绍依据';if(String(evidence.blockId).startsWith('user-calendar')){const s=evidence.source||{};return '产品日历补充'+(s.file?'：'+s.file+' · '+(s.sheet||'')+' · 第'+s.row+'行':'');}if(String(evidence.blockId).startsWith('user-confirm'))return '人工确认';if(evidence.blockId==='user-supplement')return '手动补充';if(String(evidence.blockId).startsWith('template-'))return '模板中的说明';return '合同依据';}
function renderFact(fact){
 const row=el('article','fact-row');row.dataset.factId=fact.id;const status=factStatus(fact);const head=el('div','fact-heading');head.append(el('strong','',fact.label||fact.key));head.append(el('span','badge '+status,({changed:'已修改',verified:'原样保留',pending:'保留待核',conflict:'需要复核'})[status]));const count=new Set((fact.occurrences||[]).map(o=>o.materialId+'|'+o.paragraphId)).size;head.append(el('small','',`关联 ${count} 处`));row.append(head);
 const value=fact.value===null||fact.value===undefined?'保留各模板原值':String(fact.value);row.append(el('p','fact-value',value.length>220?value.slice(0,220)+'…':value));if(fact.status!=='verified')row.append(el('p','fact-reason',fact.reasonType==='contract_missing'?'合同未给出具体值，保留原值。':'需要人工核对，当前未自动确认。'));row.append(el('p','fact-reason',fact.reason||''));
 const details=el('details');details.append(el('summary','', '查看依据与物料位置'));
 for(const evidence of (fact.evidence||[]).slice(0,8)){details.append(el('p','',sourceTitle(evidence)));details.append(el('blockquote','',evidence.quote||''));}
 if(!fact.evidence?.length)details.append(el('p','','尚无可唯一确认的合同依据，当前保留模板值。'));
 for(const m of state.task.materials){const fields=m.fields.filter(f=>f.factIds?.includes(fact.id));for(const field of fields){const occurrence=el('div','occurrence');occurrence.append(el('strong','',m.template.name+' · '+field.label));const before=field.before||'';const after=field.after??before;if(before!==after){const b=el('p');b.append(el('del','',before));const a=el('p');a.append(el('ins','',after));occurrence.append(b,a);}else occurrence.append(el('p','',before));details.append(occurrence);}}
 row.append(details);
 if(fact.status!=='verified'){
  const form=el('form','supplement');const input=el('input');input.placeholder='有明确补充资料时，可在这里填写';input.setAttribute('aria-label','补充'+(fact.label||fact.key));const button=el('button','','确认补充');button.type='submit';form.append(input,button);form.onsubmit=async e=>{e.preventDefault();if(!input.value.trim())return;try{const {setSupplement}=await import('./review-engine.js?v=20261026-numbers');state.task=await setSupplement(state.task,fact.id,input.value);clearDownloads();$('error-panel').hidden=true;changeStep('review');renderReview();}catch(error){showError(error);}};row.append(form);
 }
 return row;
}
function isConfirmed(f){return f.userConfirmed===true||f.source==='user_confirmation';}
function discardSupplementPreview(){state.supplementPreview=null;$('supplement-preview').hidden=true;}
function reviewEdited(message){clearDownloads();discardSupplementPreview();$('error-panel').hidden=true;changeStep('review');renderReview();$('supplement-notice').textContent=message;$('supplement-notice').hidden=false;}
function groups(){return state.reviewApi?.getSupplementGroups?state.reviewApi.getSupplementGroups(state.task):state.task.facts.map(f=>({id:f.id,label:f.label,kind:'text',factIds:[f.id],status:f.status}));}
function groupFacts(group){return group.factIds.map(id=>state.task.facts.find(f=>f.id===id)).filter(Boolean);}
function scopedGroupLabel(group){const s=group.scope||{},parts=[s.share?s.share+'类份额':'',s.channel,s.issuer==='distributor'?'代销口径':''].filter(Boolean);return group.label+(parts.length?' · '+parts.join(' · '):'');}
function readableFact(f){const labels={opening_cycle:['开放周期','个月'],opening_days:['开放天数','个交易日'],opening_monthday:['每月开放日','日'],opening_weekday:['开放星期',''],opening_monthly_weekday:['月内开放安排','']};const mapped=labels[f.key];if(mapped)return mapped[0]+'：'+f.value+mapped[1];if(f.key==='opening_announced_date'){const parts=(f.scope?.condition||'').split('|'),label=({participation_exit:'参与和退出',participation_only:'仅参与'})[parts[3]]||({opening:'本次开放日',sale:'本次开放日',reservation:'本次预约日'})[parts[0]]||'本次开放日';return label+'：'+f.value;}return f.label+'：'+f.value;}
const examples={date:'2026年11月23日至2026年11月27日',performance:'2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%',benchmark:'2.25%',code:'000123',opening:'每满18个月，开放5个交易日',text:'填写本次确认使用的内容'};
function renderGroup(group){
 const fs=groupFacts(group),pending=fs.some(f=>f.status!=='verified'),confirmed=fs.some(isConfirmed),changed=fs.some(factChanged),status=pending?'pending':confirmed?'confirmed':changed?'changed':'verified';
 const label=scopedGroupLabel(group),row=el('article','fact-row');row.dataset.groupId=group.id;const head=el('div','fact-heading');head.append(el('strong','',label),el('span','badge '+status,({pending:'保留待核',confirmed:'已确认无误',changed:'已修改',verified:'原样保留'})[status]));const count=new Set(fs.flatMap(f=>(f.occurrences||[]).map(o=>o.materialId+'|'+o.paragraphId))).size;head.append(el('small','',`关联 ${count} 处`));row.append(head);
 const values=[...new Set(fs.filter(f=>f.value!==null&&f.value!==undefined).map(f=>fs.length>1||group.kind==='opening'?readableFact(f):String(f.value)))];row.append(el('p','fact-value',confirmed?(values.length?values.join('\n')+'\n其余原值已确认无误':'已确认保留各模板原值'):values.length?values.join('\n'):'保留各模板原值'));
 const reasons=[...new Set(fs.map(f=>f.reason).filter(Boolean))];for(const reason of reasons)row.append(el('p','fact-reason',reason));
 const details=el('details');details.append(el('summary','','查看依据与物料位置'));const seenEvidence=new Set();for(const f of fs)for(const e of(f.evidence||[])){const k=e.blockId+'|'+e.quote;if(seenEvidence.has(k))continue;seenEvidence.add(k);details.append(el('p','',sourceTitle(e)),el('blockquote','',e.quote||''));}
 const factIds=new Set(group.factIds);for(const m of state.task.materials)for(const field of m.fields.filter(f=>f.factIds?.some(id=>factIds.has(id)))){const occurrence=el('div','occurrence');occurrence.append(el('strong','',m.template.name+' · '+field.label));if(field.before!==field.after){const b=el('p'),a=el('p');b.append(el('del','',field.before));a.append(el('ins','',field.after??field.before));occurrence.append(b,a);}else occurrence.append(el('p','',field.before));details.append(occurrence);}row.append(details);
 if(pending||['date','performance','benchmark','code','opening'].includes(group.kind)){
  const form=el('form','supplement'),input=el(group.kind==='performance'?'textarea':'input');input.placeholder=examples[group.kind]||examples.text;input.setAttribute('aria-label','补充'+label);input.addEventListener('input',discardSupplementPreview);if(input.tagName==='TEXTAREA'||group.kind==='performance')input.rows=2;const button=el('button','','预览补充');button.type='submit';form.append(input,button);form.onsubmit=e=>{e.preventDefault();if(state.busy||!input.value.trim())return;try{showSupplementPreview(state.reviewApi.previewSupplement(state.task,group.id,input.value),{label});}catch(error){showError(error);}};row.append(form);
 }
 if(pending||confirmed){const actions=el('div','fact-actions'),confirm=el('button','text-button',confirmed&&!pending?'撤销确认':'确认无误');confirm.type='button';confirm.setAttribute('aria-label',(confirmed&&!pending?'撤销确认':'确认无误')+'：'+label);confirm.onclick=()=>{if(state.busy)return;try{const undo=confirmed&&!pending;state.task=state.reviewApi.confirmFacts(state.task,group.factIds,!undo);reviewEdited(undo?'已撤销确认，对应原值恢复待核标记。':'已确认原值无误。PPT 和海报不再添加对应待核标黄或批注，Word 已有审阅修订继续保留。');}catch(error){showError(error);}};actions.append(confirm,el('span','','确认原值；保留 Word 本轮修订'));row.append(actions);}
 return row;
}
function renderFacts(){const box=$('facts-list');box.replaceChildren();let count=0;for(const group of groups()){const fs=groupFacts(group);if(!(state.filter==='all'||state.filter==='focus'&&fs.some(f=>f.status!=='verified'||factChanged(f)||isConfirmed(f))||state.filter==='pending'&&fs.some(f=>f.status!=='verified')))continue;box.append(state.reviewApi?.getSupplementGroups?renderGroup(group):renderFact(fs[0]));count++;}if(!count)box.append(el('div','empty-state','当前筛选下没有要素。'));document.querySelectorAll('[data-filter]').forEach(b=>b.classList.toggle('selected',b.dataset.filter===state.filter));}
function showSupplementPreview(preview,origin={}){
 state.supplementPreview=preview;const box=$('supplement-preview');box.replaceChildren();box.hidden=false;$('error-panel').hidden=true;$('supplement-notice').hidden=true;box.append(el('h4','','识别结果'+(origin.label?' · '+origin.label:'')));
 if(origin.note)box.append(el('p','supplement-note',origin.note));const unique=new Set();for(const update of preview.updates||[]){const fact=state.task.facts.find(f=>f.id===update.factId),text=readableFact({...fact,...update});if(!unique.has(text)){box.append(el('p','parsed-value',text));unique.add(text);}}
 for(const warning of preview.warnings||[])box.append(el('p','supplement-warning',typeof warning==='string'?warning:warning.message||String(warning)));
 for(const note of preview.notes||[])box.append(el('p','supplement-note',typeof note==='string'?note:note.message||String(note)));
 if(!preview.updates?.length){box.append(el('p','supplement-note','没有可直接应用的字段。可调整文字，或在对应要素下补充；当前核对结果保持原样。'));return;}
 const actions=el('div','preview-actions'),apply=el('button','primary-button','确认并应用补充'),cancel=el('button','text-button','取消');apply.type=cancel.type='button';apply.onclick=()=>{if(state.busy||state.supplementPreview!==preview)return;try{state.task=state.reviewApi.applySupplement(state.task,preview);reviewEdited('已应用补充资料并同步相关物料。Word 中的修改继续保留为审阅修订。');}catch(error){showError(error);}};cancel.onclick=()=>{state.supplementPreview=null;box.hidden=true;};actions.append(apply,cancel);box.append(actions);box.scrollIntoView({block:'nearest'});
}
function switchSupplement(mode){if(state.busy)return;for(const kind of ['text','excel']){$('supplement-'+kind+'-panel').hidden=kind!==mode;$('supplement-'+kind+'-tab').classList.toggle('selected',kind===mode);$('supplement-'+kind+'-tab').setAttribute('aria-pressed',String(kind===mode));}state.supplementPreview=null;$('supplement-preview').hidden=true;}
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
 state.supplementPreview=null;$('supplement-preview').hidden=true;const row=state.calendar?.rows.find(r=>r.id===$('calendar-row').value),box=$('calendar-row-detail');box.replaceChildren();if(!row)return;
 const source=row.source||{};box.append(el('strong','',row.productName),el('p','supplement-note',`${source.file||state.calendar.name||'产品日历'} · ${source.sheet||''} · 第 ${source.row} 行`));
 const matched=state.calendarMatches.includes(row.id);if(!matched)box.append(el('p','supplement-warning','所选行与合同产品不同。请核对本次要采用的数值后再应用，目标产品仍以合同为准。'));
 const range=state.calendarOpeningRange(row);
 const pairs=[['产品代码',row.productCode],['本次开放日',range.value],['开放期',row.openingPeriod],['业绩报酬',row.performanceText],['业绩比较基准',row.benchmark]];for(const [label,value]of pairs)if(value)box.append(el('p','',label+'：'+value));
 if(range.warning)box.append(el('p','supplement-warning',range.warning));
 else if(row.openingEnd&&!range.start)box.append(el('p','supplement-note','仅有档期结束日期：'+row.openingEnd+'；尚无起始日期，不自动套用为开放区间。'));
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
 setBusy(true);$('calendar-status').textContent='正在本机读取产品日历…';await nextPaint();try{const {readProductCalendar,matchCalendarRows,calendarSupplementText,calendarOpeningRange}=await import('./calendar-import.js?v=20261026-openings');state.calendarSupplementText=calendarSupplementText;state.calendarOpeningRange=calendarOpeningRange;const calendar=await readProductCalendar({name:file.name,bytes:new Uint8Array(await file.arrayBuffer())});state.calendar=calendar;const matches=matchCalendarRows(calendar,state.task.target.name);state.calendarMatches=matches.map(r=>r.id);const select=$('calendar-row');select.replaceChildren();const blank=el('option','','请选择产品行');blank.value='';select.append(blank);for(const row of [...matches,...calendar.rows.filter(r=>!state.calendarMatches.includes(r.id))]){const option=el('option','',(state.calendarMatches.includes(row.id)?'匹配产品 · ':'')+row.productName+(row.productCode?' · '+row.productCode:'')+' · 第'+row.source.row+'行');option.value=row.id;select.append(option);}select.hidden=false;$('calendar-row-label').hidden=false;$('calendar-row-detail').replaceChildren();$('calendar-status').textContent=`已读取 ${calendar.rows.length} 条产品记录。`+(matches.length===1?'已定位到合同产品，请核对解析结果。':matches.length>1?'同一产品有多条记录，请选择本次所用行。':'未找到合同产品的对应行，可手动选择并核对，或使用文字补充。')+(calendar.warnings?.length?' '+calendar.warnings.join('；'):'');select.value=matches.length===1?matches[0].id:'';if(select.value)previewCalendarRow();}
 catch(error){$('calendar-status').textContent='此文件未导入，已有补充与核对结果保留。';showError(error);}finally{$('calendar-input').value='';setBusy(false);}
}
$('calendar-input').addEventListener('change',e=>selectCalendar(e.target.files));
const calendarDrop=$('calendar-drop');
const isFileDrag=e=>Array.from(e.dataTransfer?.types||[]).includes('Files')||e.dataTransfer?.files?.length;
calendarDrop.addEventListener('dragover',e=>{if(!isFileDrag(e))return;e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect=state.busy?'none':'copy';if(!state.busy&&state.task)calendarDrop.classList.add('dragging');});
calendarDrop.addEventListener('dragleave',e=>{if(!calendarDrop.contains(e.relatedTarget))calendarDrop.classList.remove('dragging');});
calendarDrop.addEventListener('drop',e=>{e.preventDefault();calendarDrop.classList.remove('dragging');return selectCalendar(e.dataTransfer?.files);});
function renderReview(){
 const task=state.task;$('target-name').textContent=task.target.name;$('review-subtitle').textContent=`${task.templates.length} 份模板 · 同口径要素共享核对 · 文件名月份 ${task.month}`+(state.company?' · 公司介绍：'+state.company.name:'');
 $('count-facts').textContent=groups().length;$('count-updates').textContent=task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0);$('count-pending').textContent=groups().filter(g=>groupFacts(g).some(f=>f.status!=='verified')).length;$('count-materials').textContent=task.materials.length;
 const warnings=[...(task.warnings||[]),...(task.inputIssues||[]).map(i=>`${i.sourceName}：${i.reason==='failed'?'未处理，':''}${i.message}`),...task.materials.flatMap(m=>(m.warnings||[]).map(w=>m.template.name+'：'+(typeof w==='string'?w:w.message||String(w))))];$('warnings').replaceChildren();for(const warning of [...new Set(warnings)])$('warnings').append(el('p','',typeof warning==='string'?warning:warning.message||String(warning)));
 const deliveryNames=uniqueDeliveryNames(task.materials.map(materialName));
 $('delivery-list').replaceChildren();for(const [index,m] of task.materials.entries()){const item=el('div','delivery-item');item.append(el('span','kind',({word:'WORD · 一页通',ppt:'PPT · 产品要素',poster:'PPT · 海报'})[m.kind]||'物料'),el('strong','',deliveryNames[index]));const page=m.kind==='ppt'?`仅输出源第 ${(m.selectedPages||[]).join('、')} 页`:(m.kind==='poster'?'保留海报全部页面':'接受旧修订，保留本轮新修订');item.append(el('p','',page));$('delivery-list').append(item);}renderFacts();
}
document.querySelectorAll('[data-filter]').forEach(b=>b.onclick=()=>{state.filter=b.dataset.filter;renderFacts();});
function downloadLink(bytes,mime,name,parent){const url=URL.createObjectURL(new Blob([bytes],{type:mime}));state.urls.push(url);const link=el('a','',name);link.href=url;link.download=name;parent.append(link);return link;}
function generationStatus(message,kind='progress'){const box=$('generation-status');box.hidden=false;box.dataset.kind=kind;box.textContent=message;}
function deliveryIssues(task,result){return [...(task.inputIssues||[]).filter(i=>i.reason==='failed'),...(result.failures||[])];}
function deliveryNotes(task,result){return [...(task.warnings||[]).map(w=>typeof w==='string'?w:w.message||String(w)),...(task.inputIssues||[]).map(i=>`${i.sourceName}：${i.reason==='failed'?'未生成，':''}${i.message}`),...task.materials.flatMap(m=>(m.warnings||[]).map(w=>m.template.name+'：'+(typeof w==='string'?w:w.message||String(w)))),...result.files.flatMap(f=>(f.report?.repairs||[]).map(r=>f.name+'：'+(typeof r==='string'?r:r.message||r.kind))),...(result.failures||[]).map(f=>`未生成 ${f.sourceName}：${f.message}`)];}
function reviewText(task,result){
 const lines=['物料审阅清单',`目标产品：${task.target.name}`,`制作月份：${task.month}`,`公司介绍：${state.company?.name||'未使用'}`,'','本次更新产品要素及匹配到的一页通公司介绍；历史业绩及人员介绍保持原样。','黄色表示原值仍需确认，不代表合同已核实。','','输出文件：',...result.files.map(f=>f.name),'','核对事项：'];
 for(const f of task.facts){lines.push(`${f.label}｜${f.status==='verified'?(isConfirmed(f)?'用户确认原值':f.source==='user_calendar'?'产品日历补充':f.source==='user'?'用户补充':f.source==='company'?'公司介绍依据':'合同核对'):'保留待核'}｜${f.value??'保留各模板原值'}`,f.reason||'');for(const e of(f.evidence||[]).slice(0,2))lines.push(sourceTitle(e)+'：'+e.quote);lines.push('');}
 const notes=deliveryNotes(task,result);if(notes.length)lines.push('','文件处理提示：',...notes);
 return lines.join('\r\n');
}
$('generate-button').onclick=async()=>{
 if(!state.task||state.busy||!state.monthValid)return;setBusy(true);$('error-panel').hidden=true;clearDownloads();generationStatus('正在生成物料，请稍候…');$('generate-button').textContent='正在生成…';await nextPaint();
 try{
  const {writeTask}=await import('./office-writer.js?v=20261026-drop');const result=await writeTask(state.task);generationStatus(`已完成 ${result.files.length} 份物料，正在准备下载包…`);const zip=new globalThis.JSZip();for(const file of result.files)zip.file(file.name,file.bytes);zip.file('审阅清单.txt','\uFEFF'+reviewText(state.task,result));const bundle=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  const issues=deliveryIssues(state.task,result),box=$('download-list');box.append(el('p','download-success',`已生成 ${result.files.length} 份物料，可下载审阅。`));const bundleLink=downloadLink(bundle,'application/zip',`${state.task.target.shortName||'产品物料'}_${state.task.month}.zip`,box);for(const file of result.files)downloadLink(file.bytes,file.mime,file.name,box);changeStep('download');
  for(const note of deliveryNotes(state.task,result))box.append(el('p','delivery-note',note));
  generationStatus((issues.length?`已生成 ${result.files.length} 份，${issues.length} 份未生成，原因见下方。`:'已生成。')+'若未自动开始下载，请点击下方下载包或单个文件。',issues.length?'warning':'success');box.scrollIntoView({block:'nearest'});
  // A browser may decline an automatic download after asynchronous work. The
  // visible direct links must remain usable regardless of that browser policy.
  try{bundleLink.click();}catch{}
 }catch(error){generationStatus('生成失败：'+(error.message||String(error)),'error');showError(error);}finally{setBusy(false);$('generate-button').textContent='生成并下载物料';$('progress-panel').hidden=true;}
};
updateMonth(true);refreshSelection();restoreCompany();
