import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';

const full='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';
function template(id,values,label='退出费率'){
 const paragraphs=[],part='word/document.xml';
 const rows=[['产品名称',['晨星稳健3号集合资产管理计划']],[label,Array.isArray(values)?values:[values]],['管理费',['管理费率0.6%/年']]].map(([label,values],rowIndex)=>{
  const ids=[id+'l'+rowIndex,...values.map((_,i)=>id+'v'+rowIndex+'-'+i)];
  [label,...values].forEach((text,i)=>paragraphs.push({id:ids[i],text,part,index:paragraphs.length,page:1,tableId:'t',rowIndex,cellIndex:i===0?0:1}));
  return{id:id+'r'+rowIndex,label,value:values.join('\n'),part,page:1,tableId:'t',rowIndex,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:values.join('\n'),paragraphIds:ids.slice(1)}]};
 });return{id,name:id+'.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]};
}
function task(value,others=[]){return engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'产品名称：晨星稳健9号集合资产管理计划'},{id:'f',text:full}]},[template('a',value),...others],{month:'202610'});}
const group=t=>engine.getSupplementGroups(t).find(g=>g.factIds.some(id=>t.facts.find(f=>f.id===id)?.key==='redemption_fee'));
const field=(t,i=0)=>t.materials[i].fields[1];

test('fee card accepts a number or explicitly previewed complete text even when the template has one rate',()=>{
 const t=task('退出费率0.5%。'),g=group(t);
 engine.applySupplement(t,engine.previewCardSupplement(t,g.id,'0.2%'));assert.equal(field(t).after,'退出费率0.2%。');
 const p=engine.previewCardSupplement(t,g.id,full);assert.equal(p.changes[0].before,'退出费率0.2%。');assert.equal(p.changes[0].after,full);assert.ok(p.feeText);
 assert.equal(field(t).after,'退出费率0.2%。','preview is read only');engine.applySupplement(t,p);assert.equal(field(t).after,full);
 assert.equal(field(t).status,'verified');assert.equal(t.materials[0].fields[2].after,'管理费率0.6%/年');
});
test('complete fee prose replaces a value split across holding-condition and rate paragraphs once',()=>{
 const t=task(['持有天数＜180自然日','退出费率1%','持有天数≥180自然日','退出费率0%']);
 const p=engine.previewCardSupplement(t,group(t).id,full);engine.applySupplement(t,p);
 assert.equal(field(t).after.trim(),full);assert.equal(field(t).after.split(full).length,2);assert.ok(!field(t).after.includes('1%'+full));
});
test('manual complete wording can change tiers without changing automatic contract matching',()=>{
 const t=task(full),before=field(t).after,input='持有期不足30天，退出费率1.5%；持有期满30天不足365天，退出费率0.5%；持有期满365天，退出费率0%。';
 const p=engine.previewCardSupplement(t,group(t).id,input);assert.equal(field(t).after,before);engine.applySupplement(t,p);assert.equal(field(t).after,input);
 assert.ok(t.facts.filter(f=>f.key==='redemption_fee').every(f=>f.source==='user'));
});
test('complete fee replacement preserves scope and rejects shared unrelated parameters atomically',()=>{
 const t=task('退出费率0.5%。',[template('b','退出费率0.6%。','B类份额退出费率')]);
 engine.applySupplement(t,engine.previewCardSupplement(t,group(t).id,full));assert.equal(field(t).after,full);assert.equal(field(t,1).after,'退出费率0.6%。');
 const mixed=task('退出费率0.5%；管理费率0.6%。'),old=JSON.stringify(mixed.materials);
 assert.throws(()=>engine.previewCardSupplement(mixed,group(mixed).id,full),/其他|共用|整段/);assert.equal(JSON.stringify(mixed.materials),old);
});
test('full-text previews reject stale or modified targets, allow editing again, and can restore the old result',()=>{
 const t=task('退出费率0.5%。'),g=group(t),p=engine.previewCardSupplement(t,g.id,full),tampered=structuredClone(p);tampered.changes[0].after='tampered';
 assert.throws(()=>engine.applySupplement(t,tampered),/重新预览|变化/);assert.equal(field(t).after,'退出费率0.5%。');
 engine.applySupplement(t,p);const second=full.replace('180','90').replace('≥180','≥90');engine.applySupplement(t,engine.previewCardSupplement(t,g.id,second));
 assert.throws(()=>engine.applySupplement(t,p),/重新预览|变化/);assert.equal(field(t).after,second);
 assert.throws(()=>engine.previewCardSupplement(t,g.id,'0.2%'),/多|分档|完整/);
 engine.undoFeeTextSupplement(t,g.id);assert.equal(field(t).after,'退出费率0.5%。');
 engine.applySupplement(t,engine.previewCardSupplement(t,g.id,'0.2%'));assert.equal(field(t).after,'退出费率0.2%。');
});
test('a single-rate full-text edit still accepts a later numeric edit',()=>{
 const t=task('退出费率0.5%。'),g=group(t);engine.applySupplement(t,engine.previewCardSupplement(t,g.id,'本计划退出费率为0.3%。'));
 engine.applySupplement(t,engine.previewCardSupplement(t,g.id,'0.2%'));assert.equal(field(t).after,'本计划退出费率为0.2%。');
});
test('full-text input never accepts an out-of-range fee or hides a different fee type',()=>{
 for(const text of ['持有天数不足30天，退出费率-1%。','持有天数不足30天，退出费率101%。','退出费率1%；管理费率0.6%。']){const t=task('退出费率0.5%。');assert.throws(()=>engine.previewCardSupplement(t,group(t).id,text),/费率|本项|其他/);}
});
test('one complete fee supplement updates both tiered and flat templates with the same scope',()=>{
 const t=task(full,[template('flat','退出费率0.5%。')]);
 const g=engine.getSupplementGroups(t).find(g=>g.kind==='text'&&g.factIds.some(id=>t.facts.find(f=>f.id===id)?.key==='redemption_fee'));
 const input=full.replace('1%','0.8%'),p=engine.previewCardSupplement(t,g.id,input);assert.equal(p.changes.length,2);engine.applySupplement(t,p);assert.equal(field(t).after,input);assert.equal(field(t,1).after,input);
 const g2=group(t);engine.applySupplement(t,engine.previewCardSupplement(t,g2.id,full));assert.equal(field(t,1).after,full);
 engine.undoFeeTextSupplement(t,g2.id);assert.equal(field(t).after,full);assert.equal(field(t,1).after,'退出费率0.5%。');
});
test('full fee text cannot change the explicit share or collecting party of the selected card',()=>{
 const t=engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'晨星稳健9号集合资产管理计划'}]},[template('a','退出费率0.5%。','A类份额退出费率')],{month:'202610'}),g=group(t);
 for(const input of ['B类份额退出费率为0.2%。','代销机构收取的退出费率为0.2%。'])assert.throws(()=>engine.previewCardSupplement(t,g.id,input),/份额|主体|口径/);
 assert.equal(field(t).after,'退出费率0.5%。');
});
test('specific trust fee names own their nested generic management label in full prose',()=>{
 const inputTemplate=template('a','固定信托管理费0.5%/年。','固定信托管理费');for(const p of inputTemplate.paragraphs)p.text=p.text.replace('晨星稳健3号集合资产管理计划','晨星稳健3号集合资金信托计划');inputTemplate.rows[0].value='晨星稳健3号集合资金信托计划';inputTemplate.rows[0].cells[1].text=inputTemplate.rows[0].value;
 const t=engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'晨星稳健9号集合资金信托计划'}]},[inputTemplate],{month:'202610'});
 const g=engine.getSupplementGroups(t).find(g=>g.factIds.some(id=>t.facts.find(f=>f.id===id)?.key==='trust_total_fee'));assert.ok(g);
 const input='本信托的固定信托管理费率为0.6%/年。';engine.applySupplement(t,engine.previewCardSupplement(t,g.id,input));assert.equal(field(t).after,input);
});
