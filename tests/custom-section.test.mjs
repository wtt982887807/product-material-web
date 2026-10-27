import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';

const name='晨星9号集合资产管理计划';
const old=['（1）固定收益资产：10%～100%。','（2）货币市场工具：0%～100%。','（3）公募证券投资基金：0%～80%。','（4）衍生品账户权益：0%～20%。'];
const replacement='1、固定收益类资产：20%～100%；\n2、货币市场工具：0%～100%。';
function template(id,label='投资比例',values=old,otherRows=[]){
 const paragraphs=[],part='word/document.xml';
 const rows=[['产品名称',[name]],[label,values],['产品代码',['000123']],...otherRows].map(([label,values],rowIndex)=>{
  const ids=[id+'l'+rowIndex,...values.map((_,i)=>id+'v'+rowIndex+'-'+i)];
  [label,...values].forEach((text,i)=>paragraphs.push({id:ids[i],text,part,index:paragraphs.length,page:1,tableId:'t',rowIndex,cellIndex:i?1:0}));
  return{id:id+'r'+rowIndex,label,value:values.join('\n'),part,page:1,tableId:'t',rowIndex,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:values.join('\n'),paragraphIds:ids.slice(1)}]};
 });return{id,name:id+'.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]};
}
function task(templates=[template('a'),template('b')],contractLines=['三、投资比例',...old,'四、其他事项']){
 return engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+name,...contractLines].map((text,i)=>({id:'b'+i,text}))},templates,{month:'202610'});
}
const ids=(t,key='investment_ratios')=>t.facts.filter(f=>f.key===key).map(f=>f.id);
const field=(t,i=0)=>t.materials[i].fields[1];
const state=t=>JSON.stringify(t);
const ops=t=>t.materials.map(m=>m.operations.map(({id,...o})=>o));

test('whole custom ratio replacement previews every matching material and removes all omitted clauses',()=>{
 const t=task(),before=state(t),p=engine.previewCustomSection(t,ids(t),replacement);
 assert.equal(state(t),before,'preview is read only');assert.equal(p.entries.length,2);assert.ok(p.entries.every(e=>e.after===replacement));
 assert.ok(p.entries.every(e=>e.before.includes(old[3])));engine.applyCustomSection(t,p);
 for(let i=0;i<2;i++){assert.equal(field(t,i).after,replacement);assert.equal(field(t,i).status,'verified');assert.ok(!field(t,i).after.includes('公募'));}
 assert.ok(t.facts.filter(f=>f.key==='investment_ratios').every(f=>f.source==='user_custom_section'&&f.evidence[0].blockId==='user-custom-section'));
 assert.equal(t.materials[0].operations.filter(o=>o.source==='user_custom_section'&&o.after===replacement).length,1);
 assert.ok(t.materials[0].operations.some(o=>o.source==='user_custom_section'&&o.before===old[3]&&o.after===''));
});
test('custom scope replacement works without matching contract evidence and preserves exact newlines',()=>{
 const t=task([template('a','投资范围',['现金、债券。','公募基金及其他资产。'])],[]),text='现金及银行存款。\n\n国债、政策性金融债。';
 const p=engine.previewCustomSection(t,ids(t,'investment_scope'),text);engine.applyCustomSection(t,p);assert.equal(field(t).after,text);
 assert.ok(!field(t).reason.includes('合同有明确依据'));
});
test('pasted tabs become spaces without blocking the complete replacement or removing newlines',()=>{
 const t=task(),before=state(t),input='1、固定收益类资产：\t20%～100%；\n2、货币市场工具：\t0%～100%。';
 const p=engine.previewCustomSection(t,ids(t),input);assert.equal(state(t),before);
 assert.equal(p.text,'1、固定收益类资产： 20%～100%；\n2、货币市场工具： 0%～100%。');
 engine.applyCustomSection(t,p);assert.equal(field(t).after,p.text);
});
test('independent repeated fields receive one complete copy each rather than fragments',()=>{
 const t=task([template('a','投资比例',old,[['投资比例',['固定收益资产10%～100%。']]])]),p=engine.previewCustomSection(t,ids(t),replacement);
 assert.equal(p.entries.length,2);engine.applyCustomSection(t,p);for(const f of t.materials[0].fields.filter(f=>f.label==='投资比例'))assert.equal(f.after,replacement);
});
test('numbered continuation rows form one whole replacement and numbering restarts form another copy',()=>{
 const t=task([template('a','投资比例',[old[0]],[...old.slice(1).map(line=>['投资比例',[line]]),...old.map(line=>['投资比例',[line]])])]);
 const p=engine.previewCustomSection(t,ids(t),replacement);assert.equal(p.entries.length,2);assert.deepEqual(p.entries.map(e=>e.fieldIds.length),[4,4]);
 engine.applyCustomSection(t,p);const fs=t.materials[0].fields.filter(f=>f.label==='投资比例');
 assert.deepEqual(fs.map(f=>f.after),[replacement,'','','',replacement,'','','']);
 assert.equal(t.materials[0].fields.find(f=>f.label==='产品代码').after,'000123');
});
test('selecting one clause expands to the whole same-scope element while other shares remain untouched',()=>{
 const t=task([template('a'),template('b','A类份额投资比例')]),plain=ids(t).find(id=>!t.facts.find(f=>f.id===id).scope.share);
 const p=engine.previewCustomSection(t,[plain],replacement);assert.equal(p.entries.length,1);assert.ok(p.entries[0].factIds.length>1);
 engine.applyCustomSection(t,p);assert.equal(field(t).after,replacement);assert.equal(field(t,1).after,old.join('\n'));
 assert.throws(()=>engine.previewCustomSection(task([template('a'),template('b','A类份额投资比例')]),ids(task([template('a'),template('b','A类份额投资比例')])),replacement),/口径/);
});
test('semantic conditions stay separate from original-wording groups during custom replacement and undo',()=>{
 const t=task([template('a','投资范围',['本计划投资债券。']),template('b','投资范围',['穿透后投资公募基金。'])],[]);
 const [plain,aggregate]=t.facts.filter(f=>f.key==='investment_scope');aggregate.scope.condition='aggregate_layers';
 const p=engine.previewCustomSection(t,[plain.id],'现金及银行存款。');
 assert.deepEqual(p.factIds,[plain.id]);assert.equal(p.entries.length,1);assert.equal(p.entries[0].materialId,'a');
 assert.throws(()=>engine.previewCustomSection(t,[plain.id,aggregate.id],'现金。'),/口径/);
 engine.applyCustomSection(t,p);assert.equal(field(t,1).after,'穿透后投资公募基金。');
 const other=engine.previewCustomSection(t,[aggregate.id],'穿透后仅投资债券。');
 assert.deepEqual(other.factIds,[aggregate.id]);assert.equal(other.scope.condition,'aggregate_layers');
 engine.applyCustomSection(t,other);engine.undoCustomSection(t,p.id);
 assert.equal(field(t).after,'本计划投资债券。');assert.equal(field(t,1).after,'穿透后仅投资债券。');
 engine.undoCustomSection(t,other.id);assert.equal(field(t,1).after,'穿透后投资公募基金。');
});
test('repeat custom edits and undo restore prior local edits while preserving unrelated supplements',()=>{
 const t=task(),all=ids(t),first=all[0];engine.setSupplement(t,first,'（1）固定收益资产：15%～100%。');
 const original=field(t).after,originalOps=ops(t).map(os=>os.filter(o=>o.factIds?.some(id=>all.includes(id))));
 const renders=()=>t._records.filter(r=>all.includes(r.f.id)).map(r=>r.userRenderValue),originalRenders=renders();
 const p=engine.previewCustomSection(t,all,replacement);assert.deepEqual(renders(),originalRenders,'preview preserves hidden occurrence render values');engine.applyCustomSection(t,p);engine.applyCustomSection(t,engine.previewCustomSection(t,all,'仅投资债券。'));
 assert.deepEqual(renders(),originalRenders,'custom replacements leave the original occurrence values available for undo');
 const code=t.facts.find(f=>f.key==='product_code');engine.setSupplement(t,code.id,'999999');engine.undoCustomSection(t,p.id);
 assert.equal(field(t).after,original);assert.equal(t.materials[0].fields[2].after,'999999');
 assert.deepEqual(renders(),originalRenders);
 assert.deepEqual(ops(t).map(os=>os.filter(o=>o.factIds?.some(id=>all.includes(id)))),originalOps);
});
test('stale and tampered custom previews cannot overwrite newer edits',()=>{
 const t=task(),p=engine.previewCustomSection(t,ids(t),replacement),tampered=structuredClone(p);tampered.entries[0].after='错误文字';
 assert.throws(()=>engine.applyCustomSection(t,tampered),/重新预览/);engine.setSupplement(t,ids(t)[0],'仅投资债券。');
 assert.throws(()=>engine.applyCustomSection(t,p),/重新预览/);
 const newer=engine.previewCustomSection(t,ids(t),replacement);engine.applyCustomSection(t,newer);assert.throws(()=>engine.applyCustomSection(t,newer),/重新预览/);
});
test('empty, unsupported, mixed-field and shared-paragraph requests never alter task state',()=>{
 const t=task(),before=state(t);for(const text of ['', ' \n '])assert.throws(()=>engine.previewCustomSection(t,ids(t),text),/完整|填写/);
 assert.throws(()=>engine.previewCustomSection(t,[t.facts.find(f=>f.key==='product_code').id],replacement),/投资比例|投资范围/);
 assert.throws(()=>engine.previewCustomSection(t,[...ids(t),t.facts.find(f=>f.key==='product_code').id],replacement),/投资比例|投资范围/);assert.equal(state(t),before);
 const f=field(t);t.materials[0].fields.push({...f,id:'shared',label:'其他要素'});assert.throws(()=>engine.previewCustomSection(t,ids(t),replacement),/共用|重叠/);
});
test('custom override blocks individual supplements and contract adoption until undo',()=>{
 const t=task(),p=engine.previewCustomSection(t,ids(t),replacement);engine.applyCustomSection(t,p);
 assert.throws(()=>engine.setSupplement(t,ids(t)[0],'本段新文字。'),/撤销|自定义/);
 assert.throws(()=>engine.previewContractSection(t,ids(t)),/撤销|自定义/);
 engine.undoCustomSection(t,p.id);assert.doesNotThrow(()=>engine.setSupplement(t,ids(t)[0],'本段新文字。'));
});
test('an already adopted contract section must be undone before a custom override',()=>{
 const t=task(),p=engine.previewContractSection(t,ids(t));engine.applyContractSection(t,p);
 assert.throws(()=>engine.previewCustomSection(t,ids(t),replacement),/先撤销/);engine.undoContractSection(t,p.id);
 assert.doesNotThrow(()=>engine.previewCustomSection(t,ids(t),replacement));
});
