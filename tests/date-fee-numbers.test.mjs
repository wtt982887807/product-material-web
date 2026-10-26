import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';
const target='晨星稳健9号集合资产管理计划';
function template(id,rows,kind='docx'){
 const paragraphs=[],part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml';
 const rr=[['产品名称','晨星稳健3号集合资产管理计划'],...rows].map(([label,value],i)=>{const ids=[id+'l'+i,id+'v'+i];[label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));return{id:id+'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};});
 return{id,name:id+'.'+kind,kind,paragraphs,rows:rr,slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素'}]:[],warnings:[]};
}
function task(templates){return engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+target,'业绩报酬基准R1、R2以公告为准。'].map((text,i)=>({id:'b'+i,text}))},templates,{month:'202610'});}
function rendered(t,id,pid){const m=t.materials.find(m=>m.materialId===id);let text=m.template.paragraphs.find(p=>p.id===pid).text;for(const o of m.operations.filter(o=>o.paragraphId===pid&&o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}

test('opening and sale dates share one group and update together without altering reservations',()=>{
 for(const input of ['本次开放日：2026/9/8至2026/9/14','本次开售日：2026/9/8至2026/9/14']){
  const t=task([template('word',[['本次开放日','2026/8/17-2026/8/21'],['预约截止日','2026/8/16']]),template('ppt',[['本次开售日','2026年8月17日至2026年8月21日']], 'pptx')]);
  assert.equal(engine.getSupplementGroups(t).filter(g=>g.kind==='date').length,2);
  const before=JSON.stringify(t.facts.filter(f=>(f.scope.condition||'').startsWith('reservation')));
  const p=engine.previewSupplementField(t,'date',input);assert.equal(p.updates.length,4);engine.applySupplement(t,p);
  assert.equal(rendered(t,'word','wordv1'),'2026/9/8-2026/9/14');assert.equal(rendered(t,'ppt','pptv1'),'2026年9月8日至2026年9月14日');assert.equal(rendered(t,'word','wordv2'),'2026/8/16');
  assert.equal(JSON.stringify(t.facts.filter(f=>(f.scope.condition||'').startsWith('reservation'))),before);
 }
});
test('bulk alias date input recomputes trading-day windows and rejects conflicting duplicate aliases',()=>{
 const original='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 const t=task([template('a',[['开放期',original],['本次开售日','2026/8/17']])]);
 const p=engine.previewSupplementText(t,'本次开售日：2026/9/8至2026/9/14');engine.applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与');assert.equal(rendered(t,'a','av2'),'2026/9/8');
 const before=JSON.stringify(t.facts);assert.throws(()=>engine.previewSupplementText(t,'本次开放日：2026/9/8；本次开售日：2026/9/9'),/同一|不一致|不同/);assert.throws(()=>engine.previewSupplementField(t,'date','本次开放日：2026/9/8；本次开售日：2026/9/9'),/同一|不一致|不同/);assert.equal(JSON.stringify(t.facts),before);
});
test('performance standard wording has a recognized threshold and numeric single-tier input keeps prose',()=>{
 const original='对年化收益率超过业绩报酬计提标准2.8%的收益部分，管理人提取60%的业绩报酬。';
 const t=task([template('word',[['业绩报酬',original]]),template('ppt',[['业绩报酬',original]],'pptx')]);
 assert.ok(t.facts.some(f=>f.key==='performance_R1'));assert.ok(!t.facts.some(f=>f.key==='custom:performance_parameter'));
 const snapshot=JSON.stringify(t.facts),p=engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'２．５０％',ratio1:'60'});assert.equal(JSON.stringify(t.facts),snapshot);
 assert.equal(p.updates.length,2);engine.applySupplement(t,p);
 for(const id of ['word','ppt'])assert.equal(rendered(t,id,id+'v1'),original.replace('2.8%','2.5%'));
});
test('four numeric tier values update prose and formula by role',()=>{
 const original='当1.0%＜年化收益率≤4.0%时计提5%；当年化收益率＞4.0%时计提50%。E=K×[(R-4.0%)×50%+(4.0%-1.0%)×5%]。';
 const t=task([template('a',[['业绩报酬',original]])]);
 engine.applySupplement(t,engine.previewPerformanceNumbers(t,{mode:'tiered',threshold1:'2',threshold2:'3',ratio1:'20',ratio2:'40'}));
 assert.equal(rendered(t,'a','av1'),'当2.0%＜年化收益率≤3.0%时计提20%；当年化收益率＞3.0%时计提40%。E=K×[(R-3.0%)×40%+(3.0%-2.0%)×20%]。');
});
test('a repeated hurdle stated after its percentage is updated with the main standard',()=>{
 const original='当年化收益率＞业绩报酬计提标准2.7%时，超额收益部分提取60%的业绩报酬（2.7%为本期业绩报酬计提标准）。';
 const t=task([template('a',[['业绩报酬',original]])]);
 engine.applySupplement(t,engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'}));
 assert.equal(rendered(t,'a','av1'),original.replaceAll('2.7%','2.5%'));
});

test('one performance card updates first-period standards in prose and parentheses across all materials',()=>{
 const original='业绩报酬计提标准由管理人在每一开放期在管理人网站上予以公告。当年化收益率＞业绩报酬计提标准2.8%时，本集合计划对超过业绩报酬计提标准的收益部分提取60%的业绩报酬（2.8%为首期业绩报酬计提标准）。业绩报酬计提日为计划分红权益登记日、投资者退出日和计划终止日。计提基准不构成收益承诺，实际收益可能低于该基准。';
 const t=task([template('word',[['业绩报酬',original]]),template('ppt',[['业绩报酬',original]],'pptx'),template('海报',[['业绩报酬',original]],'pptx')]);
 const groups=engine.getSupplementGroups(t).filter(g=>g.label.includes('业绩报酬'));
 assert.equal(groups.length,1);assert.equal(groups[0].label,'业绩报酬');
 assert.ok(!t.facts.some(f=>f.key==='custom:performance_parameter'));
 engine.applySupplement(t,engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'}));
 for(const id of ['word','ppt','海报'])assert.equal(rendered(t,id,id+'v1'),original.replaceAll('2.8%','2.5%'));
 assert.equal(engine.getSupplementGroups(t).find(g=>g.kind==='performance').status,'verified');
});

test('unclassified performance percentages share the card but cannot be overwritten or silently confirmed',()=>{
 const original='超过业绩报酬计提标准2.8%的部分提取60%作为业绩报酬。另列参数5.00%，适用口径以公告为准。';
 const t=task([template('a',[['业绩报酬',original]])]);
 const groups=engine.getSupplementGroups(t).filter(g=>g.label.includes('业绩报酬'));
 assert.equal(groups.length,1);assert.equal(groups[0].label,'业绩报酬');
 const unknown=t.facts.find(f=>f.key==='custom:performance_parameter');assert.ok(unknown);
 const preview=engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'});
 assert.match(preview.warnings.join(''),/含义|口径|未明确/);assert.ok(preview.updates.every(u=>u.factId!==unknown.id));
 engine.applySupplement(t,preview);
 assert.equal(rendered(t,'a','av1'),original.replace('2.8%','2.5%'));assert.equal(unknown.status,'pending');
 assert.equal(engine.getSupplementGroups(t).find(g=>g.kind==='performance').status,'pending');
 assert.ok(t.materials[0].operations.some(o=>o.kind==='highlight'&&o.before==='5.00%'));
 engine.confirmFacts(t,groups[0].factIds);
 assert.equal(rendered(t,'a','av1'),original.replace('2.8%','2.5%'));
 assert.ok(!t.materials[0].operations.some(o=>o.kind==='highlight'&&o.before==='5.00%'));
 engine.confirmFacts(t,groups[0].factIds,false);
 assert.equal(rendered(t,'a','av1'),original.replace('2.8%','2.5%'));assert.equal(unknown.status,'pending');
});

test('a fee card with only ambiguous percentages offers review without guessing numeric roles',()=>{
 const t=task([template('a',[['业绩报酬','参数5.00%，具体口径以公告为准。']])]);
 const before=JSON.stringify(t.facts),preview=engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'});
 assert.equal(preview.updates.length,0);assert.match(preview.warnings.join(''),/含义|口径|未明确/);assert.equal(JSON.stringify(t.facts),before);
});

test('unclassified fee values remain separated by share scope and historical values are not current standards',()=>{
 const original='超过业绩报酬计提标准2.8%的部分提取60%作为业绩报酬。（5.00%为上期业绩报酬计提标准）';
 const t=task([template('a',[['A类份额业绩报酬',original],['B类份额业绩报酬',original]])]);
 const groups=engine.getSupplementGroups(t).filter(g=>g.kind==='performance');assert.equal(groups.length,2);
 assert.deepEqual(groups.map(g=>g.scope.share).sort(),['A','B']);
 const preview=engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'});
 assert.equal(preview.updates.length,0);assert.match(preview.warnings.join(''),/多个份额/);
 const a=groups.find(g=>g.scope.share==='A');engine.applySupplement(t,engine.previewSupplement(t,a.id,'超过业绩报酬计提基准2.5%的部分提取60%作为业绩报酬'));
 assert.equal(rendered(t,'a','av1'),original.replace('2.8%','2.5%'));assert.equal(rendered(t,'a','av2'),original);
});
test('numeric performance input rejects incomplete invalid or structurally mismatched values atomically',()=>{
 const single=task([template('a',[['业绩报酬','超过业绩报酬计提标准2.8%的部分提取60%。']])]),tiered=task([template('b',[['业绩报酬','1%＜年化收益率≤4%时计提5%；年化收益率＞4%时计提50%。']])]);
 for(const values of [{mode:'single',threshold1:'',ratio1:'60'},{mode:'single',threshold1:'2.5',ratio1:'-1'},{mode:'single',threshold1:'2.5',ratio1:'101'},{mode:'tiered',threshold1:'3',threshold2:'2',ratio1:'20',ratio2:'40'}])assert.throws(()=>engine.previewPerformanceNumbers(single,values),/填写|0至100|小于/);
 const before=JSON.stringify([single.facts,tiered.facts]);
 assert.throws(()=>engine.previewPerformanceNumbers(single,{mode:'tiered',threshold1:'2',threshold2:'3',ratio1:'20',ratio2:'40'}),/单档|档位|分段/);
 assert.throws(()=>engine.previewPerformanceNumbers(tiered,{mode:'single',threshold1:'2.5',ratio1:'60'}),/单档|档位|分段/);
 assert.equal(JSON.stringify([single.facts,tiered.facts]),before);
});
test('Word poster standalone fee clauses share numeric supplements while donation and reference sections stay untouched',()=>{
 const fee='管理人对年化收益率（R）超过业绩报酬计提基准2.65%的部分提取60%作为管理人业绩报酬。';
 const donor='捐赠条款：本集合计划年化收益率达到业绩报酬计提基准后，每个收益分配权益登记日捐赠金额为持有份额×0.2%。';
 const texts=['晨星稳健3号集合资产管理计划','业绩报酬计提：',fee,donor,'底层产品要素','晨星底层2号集合资产管理计划',fee.replace('2.65%','4%'),'历史产品说明','',fee.replace('2.65%','5%')];
 const poster={id:'poster',name:'海报.docx',kind:'docx',rows:[],slides:[],warnings:[],paragraphs:texts.map((text,index)=>({id:'p'+index,text,index,part:'word/document.xml',page:1,tableId:null}))};
 const t=task([poster]);const p=engine.previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'});assert.equal(p.updates.length,2);engine.applySupplement(t,p);
 assert.equal(rendered(t,'poster','p2'),fee.replace('2.65%','2.50%'));for(const i of [3,6,9])assert.equal(rendered(t,'poster','p'+i),texts[i]);
 assert.ok(t.materials[0].fields.some(f=>f.label==='业绩报酬'));assert.ok(t.materials[0].operations.filter(o=>o.kind==='replace'&&o.paragraphId==='p2').every(o=>o.before==='2.65%'));
});
