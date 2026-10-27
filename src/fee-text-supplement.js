// Explicit user wording is a separate operation from automatic numeric matching.
const unique=a=>[...new Set(a)];
const clean=s=>String(s??'').replace(/[\r\n\t]+/g,' ').replace(/ {2,}/g,' ').trim();
const savedKeys=['value','status','source','reason','reasonType','evidence','userConfirmed','_confirmationBefore'];
const snapshot=f=>Object.fromEntries(savedKeys.map(k=>[k,f[k]]));
export function assertFeeActionScope(text,scope={}){
 if(!scope.feeActions)return;
 const labels=[...String(text).normalize('NFKC').matchAll(/(?:认[、/]?申购|认购|申购|参与|赎回|退出)(?:(?:费(?:用|率)?)?\s*[/、及和与]\s*(?:认购|申购|参与|赎回|退出))*费(?:用|率)?/g)].map(m=>m[0]);
 if(!labels.length)return;
 const actions=unique(labels.flatMap(s=>[...s.replace(/认[、/]?申购/g,'认购/申购').matchAll(/认购|申购|参与|赎回|退出/g)].map(m=>m[0]==='申购'?'参与':m[0]==='赎回'?'退出':m[0])));
 const canonical=['认购','参与','退出'].filter(a=>actions.includes(a)).join('|');
 if(canonical!==scope.feeActions)throw new Error('填写的认购、参与或退出费率口径与本卡片不一致，请保留本项完整费率名称，或只填写数字。');
}
function labelKeys(text,labels,trust){
 const hits=labels.flatMap(([key,re])=>[...text.matchAll(new RegExp(re.source,'g'))].map(m=>({key,start:m.index,end:m.index+m[0].replace(/率$/,'').length})));
 return unique(hits.filter(h=>!hits.some(o=>o!==h&&o.start<=h.start&&o.end>=h.end&&(o.start<h.start||o.end>h.end))).map(h=>h.key==='trust_total_fee'&&!trust?'management_fee':h.key==='management_fee'&&trust?'trust_manager_fee':h.key));
}
function checkExplicitScope(text,scope){
 assertFeeActionScope(text,scope);
 const raw=text.normalize('NFKC'),shares=unique([...raw.matchAll(/([A-Z])类(?:份额)?/g)].map(m=>m[1]));
 const actors=unique([...raw.matchAll(/(代销机构|销售机构|销售银行|管理人|受托人)(?:收取的?|的)?[^，,。；;]{0,10}(?:费率|管理费|退出费|赎回费|报酬)/g)].map(m=>/代销|销售/.test(m[1])?'distributor':'manager'));
 const channels=unique([...raw.matchAll(/渠道[:：]\s*([^，,。；;:\n]{1,24})/g)].map(m=>m[1].trim()));
 if(shares.some(s=>s!==scope.share)||actors.some(s=>s!==scope.issuer)||channels.some(s=>s!==scope.channel))throw new Error('填写文字中的份额、收费主体或渠道与本卡片口径不同，请在对应卡片补充。');
}

export function previewFeeText(task,group,text,labels){
 const raw=clean(text),facts=group.factIds.map(id=>task.facts.find(f=>f.id===id)),keys=unique(facts.map(f=>f?.key));
 if(keys.length!==1||!keys[0]?.endsWith('_fee')||!raw||!/[\u4e00-\u9fff]/.test(raw))throw new Error('请填写本项完整费率文字。');
 const key=keys[0],trust=/信托/.test(task.target.name),foreign=labelKeys(raw,labels,trust).filter(k=>k!==key);checkExplicitScope(raw,group.scope||facts[0].scope||{});
 if(foreign.length)throw new Error('本次只修改'+group.label+'，整段文字中含其他费用，请仅填写本项内容。');
 if(/[-−－]\s*\d+(?:\.\d+)?\s*[%％]/.test(raw)||[...raw.matchAll(/\d+(?:\.\d+)?\s*[%％]/g)].some(m=>Number(m[0].replace(/[%％\s]/g,''))>100))throw new Error('费率应在0至100%之间，请核对完整文字。');
 const ids=new Set(group.factIds),entries=[];
 for(const m of task.materials)for(const field of m.fields.filter(f=>f.factIds.some(id=>ids.has(id)))){
  const ps=field.paragraphIds.map(id=>m.template.paragraphs.find(p=>p.id===id)).filter(p=>p?.text?.trim()),pids=new Set(ps.map(p=>p.id));
  if(!ps.length)continue;
  if(field.factIds.some(id=>!ids.has(id))||m.fields.some(f=>f!==field&&f.paragraphIds.some(id=>pids.has(id)))||labelKeys(field.before,labels,trust).some(k=>k!==key))throw new Error('该处与其他要素或分档口径共用内容，不能整段覆盖。请在对应的独立要素位置补充。');
  if(field.sectionOverrideId)throw new Error('此项已采用合同内容，请先撤销再补充。');
  entries.push({materialId:m.materialId,materialName:m.template.name,fieldId:field.id,factIds:field.factIds,before:field.after,after:raw,changes:ps.map((p,i)=>({paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after:i===0?raw:''}))});
 }
 if(!entries.length)throw new Error('未找到本项可以更新的原文。');
 const data={overrideId:"fee-text-"+facts[0].id,groupId:group.id,text:raw,feeText:true,key,entries,factState:facts.map(f=>({id:f.id,...snapshot(f),feeTextOverrideId:f.feeTextOverrideId})),updates:facts.map(f=>({factId:f.id,key,label:group.label,value:raw})),changes:entries.map(e=>({materialId:e.materialId,materialName:e.materialName,label:group.label,before:e.before,after:e.after})),warnings:[],notes:['本次按完整文字替换上述要素内容，不要求与模板旧分档一致。请核对预览后确认；Word 保留审阅修订。']};
 return{...data,signature:JSON.stringify(data)};
}

export function applyFeeText(task,preview,current){
 if(preview.signature!==current.signature||JSON.stringify({...preview,signature:undefined})!==JSON.stringify({...current,signature:undefined}))throw new Error('本项原文或核对结果已变化，请重新预览。');
 task.feeTextOverrides??={};const old=task.feeTextOverrides[current.overrideId],facts=current.updates.map(u=>task.facts.find(f=>f.id===u.factId));
 const override={...current,factSnapshots:old?.factSnapshots||facts.map(f=>({id:f.id,values:snapshot(f)}))};task.feeTextOverrides[current.overrideId]=override;
 for(const f of facts){f.feeTextOverrideId=current.overrideId;f.value=current.text;f.status='verified';f.source='user';f.reason='由用户确认整段文字；不是合同自动提取的依据。';f.reasonType='user_confirmed';f.evidence=[{blockId:'user-supplement',quote:current.text}];delete f.userConfirmed;delete f._confirmationBefore;}
 for(const e of current.entries)task.materials.find(m=>m.materialId===e.materialId).fields.find(f=>f.id===e.fieldId).feeTextOverrideId=current.overrideId;
}

export function refreshFeeTextOperations(task){
 for(const override of Object.values(task.feeTextOverrides||{}))for(const e of override.entries){
  const m=task.materials.find(m=>m.materialId===e.materialId),ids=new Set(e.changes.map(c=>c.paragraphId));m.operations=m.operations.filter(o=>!ids.has(o.paragraphId));
  for(const [i,change]of e.changes.entries())if(change.before!==change.after)m.operations.push({...change,id:'fee-text-'+e.fieldId+'-'+i,kind:'replace',source:'user_fee_text',feeTextOverrideId:override.overrideId,factIds:e.factIds,reason:'用户确认完整费率文字。'});
 }
}

export function releaseFeeText(task,id){
 const override=task.feeTextOverrides?.[id];if(!override)throw new Error('未找到本项整段修改。');delete task.feeTextOverrides[id];
 for(const saved of override.factSnapshots){const f=task.facts.find(f=>f.id===saved.id);for(const k of savedKeys)if(saved.values[k]===undefined)delete f[k];else f[k]=saved.values[k];delete f.feeTextOverrideId;}
 for(const m of task.materials){m.operations=m.operations.filter(o=>o.feeTextOverrideId!==id);for(const f of m.fields)if(f.feeTextOverrideId===id)delete f.feeTextOverrideId;}
 return override.factSnapshots.map(f=>f.id);
}
