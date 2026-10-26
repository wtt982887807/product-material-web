import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,previewContractSection,applyContractSection,undoContractSection,setSupplement,confirmFacts} from '../src/review-engine.js';
const name='晨星9号集合资产管理计划';
const ratios=['（1）固定收益类资产：占集合计划资产总值的20%～100%（不含）；','（2）权益类资产：占集合计划资产总值的0%～40%（不含）；','（3）期货和衍生品类账户权益：占集合计划资产总值的0%～20%（不含）；','（4）公募基金：占集合计划资产总值的0%～80%；','（5）境外投资：占集合计划资产总值的0%～50%（不含）。'];
const note='本集合计划投资于公募证券投资基金，大类资产配置比例应当按照穿透原则合并计算。';
function task(count=1,style=t=>t){
 const lines=['产品名称：'+name,'三、投资比例',...ratios,note,'四、投资限制'];
 const contract={name:'合同.docx',blocks:lines.map((text,i)=>({id:'c'+i,text})),warnings:[]};
 const templates=Array.from({length:count},(_,n)=>{
  const paragraphs=[],rows=[['产品名称',name],['投资比例',style(ratios.slice(0,3).join('；'))],['产品代码','000123']].map(([label,value],i)=>{
   const ids=[`m${n}-p${i*2}`,`m${n}-p${i*2+1}`];for(const [j,text]of [label,value].entries())paragraphs.push({id:ids[j],text,index:i*2+j,part:'word/document.xml',page:1,tableId:'t'});
   return{id:'r'+i,label,value,tableId:'t',rowIndex:i,page:1,part:'word/document.xml',paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
  });return{id:'m'+n,name:'一页通'+n+'.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]};
 });return analyzeMaterials(contract,templates);
}
const ids=t=>t.facts.filter(f=>f.key==='investment_ratios').map(f=>f.id);
const ops=t=>t.materials.map(m=>m.operations.map(({id,...o})=>o));
test('adoption is explicit and preview has all five clauses plus note across materials',()=>{
 const t=task(2),before=JSON.stringify(t),p=previewContractSection(t,ids(t));assert.equal(JSON.stringify(t),before);assert.equal(p.entries.length,2);
 for(const e of p.entries){assert.ok(e.after.includes(ratios[4]));assert.ok(e.after.includes(note));}
 applyContractSection(t,p);for(const m of t.materials){const f=m.fields.find(f=>f.label==='投资比例');assert.equal(f.status,'verified');assert.ok(f.after.includes(ratios[4]));assert.equal(m.operations.filter(o=>o.source==='contract_section').length,1);}
});
test('undo restores original small operations and preserves later independent supplements',()=>{
 const t=task(),original=ops(t)[0].filter(o=>!o.factIds.includes(t.facts.find(f=>f.key==='product_code').id));
 const p=previewContractSection(t,ids(t));applyContractSection(t,p);const code=t.facts.find(f=>f.key==='product_code');setSupplement(t,code.id,'999999');confirmFacts(t,ids(t));
 assert.ok(t.materials[0].fields.find(f=>f.label==='投资比例').after.includes(ratios[4]));
 undoContractSection(t,p.id);assert.equal(t.materials[0].fields.find(f=>f.label==='产品代码').after,'999999');
 assert.deepEqual(ops(t)[0].filter(o=>!o.factIds.includes(code.id)),original);
});
test('stale previews, tampered payloads and ambiguous chapters cannot be applied',()=>{
 const t=task(),p=previewContractSection(t,ids(t));const tampered=structuredClone(p);tampered.entries[0].changes[0].after='不正确';assert.throws(()=>applyContractSection(t,tampered),/重新预览/);
 t.contract.blocks.find(b=>b.text===ratios[4]).text=ratios[4].replace('50','60');assert.throws(()=>applyContractSection(t,p),/重新预览/);
 t.contract.blocks.push({id:'other',text:'五、投资比例'},{id:'other1',text:ratios[0].replace('20','30')},{id:'other2',text:'六、风险揭示'});assert.throws(()=>previewContractSection(t,ids(t)),/唯一/);
});
test('independent repeated tables each receive the complete section',()=>{
 const t=task(),m=t.materials[0],f=m.fields.find(f=>f.label==='投资比例'),p=m.template.paragraphs.find(p=>p.id===f.paragraphIds[0]);
 const copy={...p,id:p.id+'-copy',index:99,tableId:'second-table'};m.template.paragraphs.push(copy);m.fields.push({...f,id:f.id+'-copy',paragraphIds:[copy.id]});
 const preview=previewContractSection(t,ids(t));assert.equal(preview.entries.length,2);for(const e of preview.entries)assert.ok(e.after.includes(ratios[4])&&e.after.includes(ratios[0]));
});
test('shared paragraphs and mixed-field facts cannot erase unrelated changes',()=>{
 const t=task(),m=t.materials[0],f=m.fields.find(f=>f.label==='投资比例');m.fields.push({...f,id:'another',label:'产品代码',factIds:[t.facts.find(f=>f.key==='product_code').id]});
 assert.throws(()=>previewContractSection(t,ids(t)),/共用/);
});
test('adoption keeps the template numeric enumeration style without altering quoted evidence',()=>{
 const t=task(1,text=>text.replace('（1）','1）'));
 const preview=previewContractSection(t,ids(t));assert.match(preview.entries[0].after,/^1）/);assert.match(preview.entries[0].after,/\n5）境外投资/);assert.ok(preview.evidence[0].quote.includes('（5）'));
});
