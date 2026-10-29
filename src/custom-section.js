// Explicit whole-element wording supplied by the user; no inferred contract authority.
const labels={investment_ratios:'投资比例',investment_scope:'投资范围',investment_strategy:'投资策略',investment_restrictions:'投资限制'};
const unique=a=>[...new Set(a)];
const savedKeys=['value','status','source','reason','reasonType','evidence','userConfirmed','_confirmationBefore'];
const snapshot=f=>Object.fromEntries(savedKeys.map(k=>[k,f[k]]));
const normalize=s=>String(s??'').normalize('NFKC').replace(/[\s·•,，。;；:：、【】\[\]"“”'‘’]/g,'');
function normalizedScope(task,f){
 const scope={...f.scope};
 // Match review-card grouping: ignore only a condition that represents the
 // original wording; semantic conditions remain separate replacement scopes.
 if(scope.condition&&(task._records||[]).some(r=>r.f===f&&normalize(r.atom.before)===scope.condition))delete scope.condition;
 return Object.fromEntries(Object.entries(scope).sort(([a],[b])=>a.localeCompare(b)));
}
const scopeKey=(task,f)=>JSON.stringify(normalizedScope(task,f));
const numbering=text=>[...String(text).normalize('NFKC').matchAll(/(?:^|\n)\s*(?:\((\d+)\)|(\d+)[)、.])\s*/g)].map(m=>Number(m[1]||m[2]));
function selection(task,factIds){
 const selected=unique(factIds||[]).map(id=>task.facts.find(f=>f.id===id));
 if(!selected.length||selected.some(f=>!f)||!labels[selected[0]?.key]||selected.some(f=>f.key!==selected[0].key))throw new Error('自定义整项替换仅适用于同一项投资比例、投资范围、投资策略或投资限制。');
 const key=selected[0].key,scope=scopeKey(task,selected[0]);if(selected.some(f=>scopeKey(task,f)!==scope))throw new Error('不同份额、渠道、主体口径或适用条件请分别自定义替换。');
 // A card may start at one old clause; replacing the whole element includes all
 // sibling clauses of the same scope and every associated material.
 return {key,scope:normalizedScope(task,selected[0]),facts:task.facts.filter(f=>f.key===key&&scopeKey(task,f)===scope)};
}

export function previewCustomSection(task,factIds,text){
 const raw=String(text??'').replace(/\r\n?/g,'\n').replace(/\t/g,' ');if(!raw.trim())throw new Error('请填写本次完整替换内容；未填写时不会删除原文。');
 if(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(raw))throw new Error('粘贴内容包含无法写入文档的控制字符，请去掉异常字符后预览。');
 const {key,scope,facts}=selection(task,factIds),ids=new Set(facts.map(f=>f.id)),id='custom-section-'+key+'-'+JSON.stringify(scope),entries=[];
 for(const m of task.materials){
  const fields=m.fields.filter(f=>f.factIds.some(id=>ids.has(id))),groups=[];
  for(const field of fields){
   if(field.sectionOverrideId)throw new Error('此项已采用合同内容，请先撤销采用，再自定义替换。');
   if(field.customSectionOverrideId&&field.customSectionOverrideId!==id)throw new Error('此项已有其他口径的自定义替换，请先撤销。');
   const paragraphs=unique(field.paragraphIds).map(pid=>m.template.paragraphs.find(p=>p.id===pid));
   if(paragraphs.some(p=>!p))throw new Error('原文位置已经变化，请重新选择文件核对。');
   const pids=new Set(field.paragraphIds);
   if(field.factIds.some(fid=>!ids.has(fid))||m.fields.some(f=>f!==field&&f.paragraphIds.some(pid=>pids.has(pid))))throw new Error('此处与其他要素共用或重叠，不能整项覆盖，请分别补充原文。');
   const ps=paragraphs.filter(p=>p.text?.trim());if(!ps.length)continue;
   const first=ps[0],location=JSON.stringify([first.part,first.tableId||'',first.shapeId||'']),ordinals=numbering(field.before),last=groups.at(-1);
   const continued=last&&last.location===location&&(first.tableId||first.shapeId)&&ordinals.length&&last.lastOrdinal&&ordinals[0]===last.lastOrdinal+1;
   if(continued){last.fields.push(field);last.ps.push(...ps);last.lastOrdinal=ordinals.at(-1);}
   else groups.push({location,fields:[field],ps,lastOrdinal:ordinals.at(-1)});
  }
  for(const group of groups){
   const changes=group.ps.map((p,i)=>({paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after:i===0?raw:''}));
   entries.push({materialId:m.materialId,name:m.template.name,fieldIds:group.fields.map(f=>f.id),factIds:unique(group.fields.flatMap(f=>f.factIds)),before:group.fields.map(f=>f.after).filter(Boolean).join('\n'),after:raw,changes});
  }
 }
 if(!entries.length)throw new Error('未找到可自定义替换的完整要素位置。');
 const data={id,key,label:labels[key],text:raw,scope,factIds:[...ids],factState:facts.map(f=>({id:f.id,...snapshot(f),customSectionOverrideId:f.customSectionOverrideId})),entries};
 return {...data,signature:JSON.stringify(data)};
}

export function applyCustomSectionOverride(task,preview){
 const current=previewCustomSection(task,preview?.factIds,preview?.text);
 if(preview?.signature!==current.signature||JSON.stringify({...preview,signature:undefined})!==JSON.stringify({...current,signature:undefined}))throw new Error('本项原文或核对结果已变化，请重新预览自定义替换。');
 task.customSectionOverrides??={};const old=task.customSectionOverrides[current.id],facts=current.factIds.map(id=>task.facts.find(f=>f.id===id));
 task.customSectionOverrides[current.id]={...current,factSnapshots:old?.factSnapshots||facts.map(f=>({id:f.id,values:snapshot(f)}))};
 for(const f of facts){f.customSectionOverrideId=current.id;f.value=current.text;f.status='verified';f.source='user_custom_section';f.reason='用户确认自定义整项替换；不是合同自动提取的依据。';f.reasonType='user_confirmed';f.evidence=[{blockId:'user-custom-section',quote:current.text}];delete f.userConfirmed;delete f._confirmationBefore;}
 for(const e of current.entries){const m=task.materials.find(m=>m.materialId===e.materialId);for(const field of m.fields.filter(f=>e.fieldIds.includes(f.id)))field.customSectionOverrideId=current.id;}
 return task;
}

export function refreshCustomSectionOperations(task){
 for(const override of Object.values(task.customSectionOverrides||{}))for(const e of override.entries){
  const m=task.materials.find(m=>m.materialId===e.materialId),pids=new Set(e.changes.map(c=>c.paragraphId));m.operations=m.operations.filter(o=>!pids.has(o.paragraphId));
  for(const [i,c]of e.changes.entries())if(c.before!==c.after)m.operations.push({...c,id:'custom-section-'+e.fieldIds[0]+'-'+i,kind:'replace',source:'user_custom_section',customSectionOverrideId:override.id,factIds:e.factIds,reason:'用户确认自定义替换整个'+override.label+'；未包含的旧条目移除。'});
 }
}

export function refreshCustomSectionFields(task){
 for(const override of Object.values(task.customSectionOverrides||{}))for(const e of override.entries){
  const m=task.materials.find(m=>m.materialId===e.materialId);
  for(const [i,id]of e.fieldIds.entries()){const field=m.fields.find(f=>f.id===id);field.after=i===0?e.after:'';field.reason='用户确认自定义整项替换；不是合同自动提取的依据。';}
 }
}

export function releaseCustomSection(task,id){
 const override=task.customSectionOverrides?.[id];if(!override)throw new Error('未找到本项自定义替换。');delete task.customSectionOverrides[id];
 for(const s of override.factSnapshots){const f=task.facts.find(f=>f.id===s.id);for(const k of savedKeys)if(s.values[k]===undefined)delete f[k];else f[k]=s.values[k];delete f.customSectionOverrideId;}
 for(const m of task.materials){m.operations=m.operations.filter(o=>o.customSectionOverrideId!==id);for(const f of m.fields)if(f.customSectionOverrideId===id)delete f.customSectionOverrideId;}
 return override.factSnapshots.map(f=>f.id);
}
