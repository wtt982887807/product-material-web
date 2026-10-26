// Explicit, previewed adoption. Automatic atom matching remains unchanged.
import {getSectionReview} from './review-sections.js?v=20261026-opening-strategy';
import {resolveStrategyReplacement} from './contract-sections.js?v=20261026-opening-strategy';
const unique=a=>[...new Set(a)];
const snapshot=f=>Object.fromEntries(['value','status','source','reason','reasonType','evidence','userConfirmed','_confirmationBefore'].map(k=>[k,f[k]]));
function numberingLikeTemplate(lines,text){
 const style=String(text).match(/^\s*(?:([（(])\d+([）)])|\d+([）)、.．]))/);
 if(!style)return lines;
 return lines.map(line=>line.replace(/^\s*(?:[（(](\d+)[）)]|(\d+)[）)、.．])\s*/,(_,a,b)=>(style[1]||'')+(a||b)+(style[2]||style[3])));
}

export function previewSectionAdoption(task,factIds){
 const ids=unique(factIds||[]),review=getSectionReview(task,ids);
 if(!['investment_ratios','investment_strategy','opening_fixed_monthdays'].includes(review?.key)||!(review.replacement?.lines?.length||review.materials?.some(m=>m.replacement?.lines?.length)))throw new Error('尚无唯一、完整且适用范围明确的合同依据，不能直接采用。');
 if(task.sectionOverrides?.[review.id])throw new Error('当前已采用此项合同内容；如需重新选择，请先撤销。');
 const entries=[];
 for(const material of review.materials){
  const m=task.materials.find(m=>m.materialId===material.materialId);
  const replacement=material.replacement||review.replacement;if(!replacement?.lines?.length)continue;
  const fields=m.fields.filter(f=>material.fieldIds.includes(f.id));
  const paragraphs=new Set(fields.flatMap(f=>f.paragraphIds));
  if(m.fields.some(f=>!fields.includes(f)&&f.paragraphIds.some(id=>paragraphs.has(id)))||fields.some(f=>f.factIds.some(id=>task.facts.find(f=>f.id===id)?.key!==review.key)))throw new Error(review.label+'与其他要素共用段落，不能整段覆盖；请保留局部核对。');
  const groups=[];
  for(const field of fields){
   const first=m.template.paragraphs.find(p=>field.paragraphIds.includes(p.id)&&p.text?.trim());if(!first)continue;
   // Strategy copies in separate cells must each receive their own full title
   // list. Missing cell/shape metadata is not evidence of a shared container.
   const location=review.key==='investment_strategy'
    ?first.tableId&&Number.isInteger(first.rowIndex)&&Number.isInteger(first.cellIndex)?[first.part,first.tableId,first.rowIndex,first.cellIndex]
     :!first.tableId&&first.shapeId?[first.part,'shape',first.shapeId]:[first.part,'field',field.id]
    :[first.part,first.tableId||'',first.shapeId||''];
   const key=JSON.stringify(location),last=groups.at(-1);
   const restarts=/^\s*(?:[（(]?1[）)、.．]|一[、])/.test(field.before.normalize('NFKC'))&&last?.fields.some(f=>/^\s*(?:[（(]?1[）)、.．]|一[、])/.test(f.before.normalize('NFKC')));
   if(last?.key===key&&!restarts)last.fields.push(field);else groups.push({key,fields:[field]});
  }
  for(const group of groups){
   const fields=group.fields,ps=unique(fields.flatMap(f=>f.paragraphIds)).map(id=>m.template.paragraphs.find(p=>p.id===id)).filter(p=>p?.text?.trim());
   // Each independent table/text box receives a complete section, never a
   // fragment of a section distributed across separate copies of the field.
   const original=ps.map(p=>p.text).join('\n'),scope={...task.facts.find(f=>fields.some(field=>field.factIds.includes(f.id)))?.scope,condition:''};
   const localReplacement=review.key==='investment_strategy'?resolveStrategyReplacement(task.contract,original,scope):replacement;
   if(!localReplacement?.lines?.length)throw new Error('该处'+review.label+'尚不能安全对应合同，请核对本处原文。');
   const lines=review.key==='investment_ratios'?numberingLikeTemplate(localReplacement.lines,original):localReplacement.lines,changes=ps.map((p,i)=>({paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after:i===ps.length-1?lines.slice(i).join('\n'):(lines[i]||'')}));
   entries.push({materialId:m.materialId,name:m.template.name,fieldIds:fields.map(f=>f.id),factIds:unique(fields.flatMap(f=>f.factIds)),before:fields.map(f=>f.after).join('\n'),after:changes.map(c=>c.after).filter(Boolean).join('\n'),warnings:localReplacement.warnings||[],changes});
  }
 }
 if(!entries.length)throw new Error('未找到需要更新的可编辑字段。');
 const data={id:review.id,key:review.key,label:review.label,factIds:ids,evidence:review.evidence,entries};
 return {...data,signature:JSON.stringify(data)};
}

export function adoptSection(task,preview){
 const current=previewSectionAdoption(task,preview?.factIds);
 if(preview?.signature!==current.signature||JSON.stringify({...preview,signature:undefined})!==JSON.stringify({...current,signature:undefined}))throw new Error('核对内容已变化，请重新预览合同更新。');
 const ids=unique(current.entries.flatMap(e=>e.factIds)),facts=ids.map(id=>task.facts.find(f=>f.id===id));
 const override={...current,factSnapshots:facts.map(f=>({id:f.id,values:snapshot(f)}))};
 task.sectionOverrides??={};task.sectionOverrides[current.id]=override;
 for(const f of facts){f.sectionOverrideId=current.id;f.status='verified';f.value=null;f.source='contract_section';f.reason='已预览并采用合同'+(current.key==='investment_strategy'?'策略标题':current.label)+'。';f.reasonType='confirmed';f.evidence=current.evidence;delete f.userConfirmed;delete f._confirmationBefore;}
 for(const entry of current.entries){const m=task.materials.find(m=>m.materialId===entry.materialId);for(const field of m.fields.filter(f=>entry.fieldIds.includes(f.id)))field.sectionOverrideId=current.id;}
 return task;
}

export function refreshSectionOperations(task){
 for(const override of Object.values(task.sectionOverrides||{}))for(const entry of override.entries){
  const m=task.materials.find(m=>m.materialId===entry.materialId),paragraphIds=new Set(entry.changes.map(c=>c.paragraphId));
  m.operations=m.operations.filter(o=>!paragraphIds.has(o.paragraphId));
  for(const [i,change]of entry.changes.entries())if(change.before!==change.after)m.operations.push({...change,id:'section-'+entry.fieldIds[0]+'-'+i,kind:'replace',source:'contract_section',overrideId:override.id,factIds:entry.factIds,reason:'用户预览确认采用合同'+override.label+'。'});
 }
}

export function releaseSection(task,id){
 const override=task.sectionOverrides?.[id];if(!override)throw new Error('未找到已采用的合同更新。');
 delete task.sectionOverrides[id];
 for(const saved of override.factSnapshots){const f=task.facts.find(f=>f.id===saved.id);for(const key of Object.keys(saved.values))if(saved.values[key]===undefined)delete f[key];else f[key]=saved.values[key];delete f.sectionOverrideId;}
 for(const m of task.materials){m.operations=m.operations.filter(o=>o.overrideId!==id);for(const field of m.fields)if(field.sectionOverrideId===id)delete field.sectionOverrideId;}
 return override.factSnapshots.map(f=>f.id);
}
