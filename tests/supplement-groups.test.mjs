import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,getSupplementGroups,previewSupplement,previewSupplementText,applySupplement,applySupplementBatch,confirmFacts,setSupplement} from '../src/review-engine.js';
import {parsePerformanceInput,parseDateInput,formatDateLike} from '../src/supplement-parser.js';
const target='晨星稳健9号集合资产管理计划';
function template(id,rows,kind='docx'){
 const paragraphs=[],part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml';
 const rr=[['产品名称','晨星稳健3号集合资产管理计划'],...rows].map(([label,value],i)=>{const ids=[id+'l'+i,id+'v'+i];[label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));return{id:id+'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};});
 return{id,name:id+'.'+kind,kind,paragraphs,rows:rr,slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素'}]:[],warnings:[]};
}
function task(templates,extra=[]){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+target,'业绩报酬基准R1、R2以公告为准。',...extra].map((text,i)=>({id:'b'+i,text}))},templates,{month:'202610'});}
function rendered(t,id,pid){const m=t.materials.find(m=>m.materialId===id);let text=m.template.paragraphs.find(p=>p.id===pid).text;for(const o of m.operations.filter(o=>o.paragraphId===pid&&o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}

test('one opening-date group updates range, start and end while keeping each date style',()=>{
 const t=task([template('word',[['本次开放日','2026/8/12-2026/8/14']]),template('ppt',[['开放期','下一开放期的首个开放日为2026年8月12日，具体以公告为准。'],['开放截止日','2026年8月14日']], 'pptx')]);
 const gs=getSupplementGroups(t).filter(g=>g.kind==='date');assert.equal(gs.length,1);assert.equal(gs[0].factIds.length,4);
 const preview=previewSupplement(t,gs[0].id,'2026/11/3至2026/11/5');assert.equal(t.facts.find(f=>f.key==='opening_announced_date').status,'pending');applySupplement(t,preview);
 assert.equal(rendered(t,'word','wordv1'),'2026/11/3-2026/11/5');
 assert.equal(rendered(t,'ppt','pptv1'),'下一开放期的首个开放日为2026年11月3日，具体以公告为准。');
 assert.equal(rendered(t,'ppt','pptv2'),'2026年11月5日');
 assert.ok(t.materials.every(m=>!m.operations.some(o=>o.kind==='highlight')));
});

test('date groups do not merge sale, opening, reservations or share classes',()=>{
 const t=task([template('a',[['本次开放日','2026/8/12'],['募集开放日','2026/8/12'],['预约开放日','2026/8/12'],['A类份额开放日','2026/8/12'],['B类份额开放日','2026/8/12']])]);
 assert.equal(getSupplementGroups(t).filter(g=>g.kind==='date').length,5);
 const p=previewSupplementText(t,'本次开放日：2026/11/3');assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),/多个份额/);
});

test('performance text maps threshold and payout roles independently',()=>{
 assert.deepEqual(parsePerformanceInput('2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%'),{performance_R1:'2%',performance_R2:'3%',performance_ratio_1:'20%',performance_ratio_2:'40%'});
 assert.deepEqual(parsePerformanceInput('R2=3%；R＞R2，计提40%'),{performance_R2:'3%',performance_ratio_2:'40%'});
 assert.deepEqual(parsePerformanceInput('R2=3%；第2档计提40%'),{performance_R2:'3%',performance_ratio_2:'40%'});
 assert.throws(()=>parsePerformanceInput('-2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%'),/负数/);
 assert.throws(()=>parsePerformanceInput('2.0%＜年化收益率≤3.0%，计提-20%；年化收益率＞3.0%，计提40%'),/负数/);
 const t=task([template('a',[['业绩报酬','当1.0%＜年化收益率≤4.0%时计提5%；当年化收益率＞4.0%时计提50%。E=K×[(R-4.0%)×50%+(4.0%-1.0%)×5%]。']])]);
 const g=getSupplementGroups(t).find(g=>g.kind==='performance');applySupplement(t,previewSupplement(t,g.id,'2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%'));
 assert.equal(rendered(t,'a','av1'),'当2.0%＜年化收益率≤3.0%时计提20%；当年化收益率＞3.0%时计提40%。E=K×[(R-3.0%)×40%+(3.0%-2.0%)×20%]。');
 assert.throws(()=>previewSupplement(t,g.id,'年化收益率＞3.0%，计提40%'),/多个业绩报酬基准/);
});

test('opening subwindows and reservation deadline keep independent meanings within a paragraph',()=>{
 const original='2026/8/18-2026/8/24：8/18-8/19开放参与和退出，8/20-8/24仅开放参与（8/24为家族端预约截止日）';
 const t=task([template('a',[['开放期',original]])]);const groups=getSupplementGroups(t),opening=groups.find(g=>g.kind==='date'&&g.scope.event==='opening'),reservation=groups.find(g=>g.kind==='date'&&g.scope.event==='reservation');
 assert.equal(opening.factIds.length,6);assert.equal(reservation.factIds.length,1);
 const p=previewSupplementText(t,'本次开放日：2026/11/3至2026/11/9');assert.equal(p.updates.length,6);assert.equal(p.warnings.length,0);applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'2026/11/3-2026/11/9：11/3-11/4开放参与和退出，11/5-11/9仅开放参与（8/24为家族端预约截止日）');
 const detailed=previewSupplement(t,opening.id,'参与和退出：2026/11/3至2026/11/4；仅参与：2026/11/5至2026/11/9');assert.equal(detailed.updates.length,4);applySupplement(t,detailed);
 assert.equal(rendered(t,'a','av1'),'2026/11/3-2026/11/9：11/3-11/4开放参与和退出，11/5-11/9仅开放参与（8/24为家族端预约截止日）');
 applySupplement(t,previewSupplement(t,reservation.id,'2026/11/9'));assert.equal(rendered(t,'a','av1'),'2026/11/3-2026/11/9：11/3-11/4开放参与和退出，11/5-11/9仅开放参与（11/9为家族端预约截止日）');
});

test('full opening range shifts poster and PPT subwindows by stock trading days',()=>{
 const original='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与（若出现合同变更，以管理人公告为准）';
 const t=task([template('poster',[['开放期',original]],'pptx'),template('ppt',[['开放期',original]],'pptx')]);
 const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14');assert.equal(p.warnings.length,0);assert.ok(p.notes.some(n=>/沪深交易日/.test(n)));applySupplement(t,p);
 for(const id of ['poster','ppt'])assert.equal(rendered(t,id,id+'v1'),'2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与（若出现合同变更，以管理人公告为准）');
 assert.ok(t.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'));
 assert.ok(t.facts.some(f=>f.evidence.some(e=>/沪深交易日/.test(e.quote))));
});

test('explicit Excel segment rule overrides old window counts and is retained in evidence',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/19开放参与和退出，8/20-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]);const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14\n分段规则：前2个开放日开放参与和退出，后3个开放日仅开放参与');
 assert.equal(p.warnings.length,0);applySupplement(t,p);assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与');
 assert.ok(p.updates.some(u=>u.basis?.ruleSource==='explicit_rule'));
});

test('ordinary Excel opening remarks do not block template splitting, but incomplete day rules do',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 for(const remark of ['以管理人公告为准','具体开放安排以公告为准，若出现合同变更，以管理人公告为准']){
  const t=task([template('a',[['开放期',old]])]),p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14\n分段规则：'+remark);
  assert.equal(p.warnings.length,0);assert.equal(p.updates.length,6);applySupplement(t,p);
  assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与');
  assert.ok(p.updates.some(u=>u.basis?.ruleSource==='template_dates'));
 }
 for(const remark of ['前2个交易日参与和退出','后3个开放日仅参与','前几个交易日安排待公告']){
  const t=task([template('a',[['开放期',old]])]),p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14\n分段规则：'+remark);
  assert.equal(p.updates.length,2);assert.match(p.warnings.join(''),/分段规则未同时明确/);
 }
});

test('changing an already supplied overall range invalidates stale derived windows and restores template highlights',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]);applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'));
 const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/15');assert.match(p.warnings.join(''),/旧分段确认失效/);
 assert.ok(t.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'),'preview must not mutate prior confirmation');applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/15：8/17-8/18开放参与和退出，8/19-8/21仅开放参与');
 const stale=t._records.filter(r=>r.atom.dateWindow);assert.ok(stale.every(r=>r.f.status==='pending'&&!r.userRenderValue&&!r.f.userConfirmed&&r.f.value===null));
 assert.ok(stale.every(r=>!r.f.evidence.some(e=>e.basis?.kind==='trading_day_split')));
 assert.ok(stale.every(r=>t.materials[0].operations.some(o=>o.kind==='highlight'&&o.factIds.includes(r.f.id))));
 applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'));assert.ok(t.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'));
});

test('a failed explicit rule invalidates prior subwindow confirmation even when overall dates are unchanged',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]);applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'));
 applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14；分段规则：以管理人公告为准'));
 assert.ok(t.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'));
 const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14；分段规则：前2个开放日参与和退出，后4个开放日仅开放参与');
 assert.match(p.warnings.join(''),/旧分段确认失效/);applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/14：8/17-8/18开放参与和退出，8/19-8/21仅开放参与');
 assert.ok(t._records.filter(r=>r.atom.dateWindow).every(r=>r.f.status==='pending'&&!r.userRenderValue));
 assert.ok(t._records.filter(r=>r.atom.dateWindow).every(r=>t.materials[0].operations.some(o=>o.kind==='highlight'&&o.factIds.includes(r.f.id))));
});

test('a start-only edit invalidates prior overall end and manually supplied or confirmed subwindows',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 for(const method of ['manual','confirm']){
  const t=task([template('a',[['开放期',old]])]),g=getSupplementGroups(t).find(g=>g.kind==='date');
  if(method==='manual')applySupplement(t,previewSupplement(t,g.id,'本次开放日：2026/9/8至2026/9/14；参与和退出：2026/9/8至2026/9/9；仅参与：2026/9/10至2026/9/14'));
  else confirmFacts(t,g.factIds);
  const p=previewSupplementText(t,'本次开放日：2026/11/3');applySupplement(t,p);
  assert.equal(rendered(t,'a','av1'),old.replace('2026/8/17','2026/11/3'));
  assert.ok(t._records.filter(r=>r.atom.dateWindow||r.atom.dateRole==='end').every(r=>r.f.status==='pending'&&!r.f.userConfirmed));
 }
});

test('replacement subwindows in the same batch remain verified while separate reservations stay confirmed',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与（8/21为预约截止日）';
 const t=task([template('a',[['开放期',old]])]),reservation=getSupplementGroups(t).find(g=>g.scope.event==='reservation');confirmFacts(t,reservation.factIds);
 applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'));
 applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/15；参与和退出：2026/9/8至2026/9/9；仅参与：2026/9/10至2026/9/15'));
 assert.ok(t.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'));
 assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/15：9/8-9/9开放参与和退出，9/10-9/15仅开放参与（8/21为预约截止日）');
 assert.ok(t.facts.filter(f=>reservation.factIds.includes(f.id)).every(f=>f.userConfirmed));
});

test('segmented windows skip exchange holiday closures and weekend makeup workdays',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]);applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/30至2026/10/13'));
 assert.equal(rendered(t,'a','av1'),'2026/9/30-2026/10/13：9/30-10/8开放参与和退出，10/9-10/13仅开放参与');
});

test('conflicting old splits do not silently pick a template unless a full explicit rule is supplied',()=>{
 const a='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与',b=a.replace('8/17-8/18','8/17-8/19').replace('8/19-8/21','8/20-8/21');
 const t=task([template('a',[['开放期',a]]),template('b',[['开放期',b]])]);const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14');
 assert.match(p.warnings.join(''),/不同|不一致/);applySupplement(t,p);assert.equal(rendered(t,'a','av1'),a.replace('2026/8/17-2026/8/21','2026/9/8-2026/9/14'));
 const explicit=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14\n分段规则：前两个交易日参与和退出，后三个交易日仅参与');applySupplement(t,explicit);
 assert.equal(rendered(t,'a','av1'),rendered(t,'b','bv1'));
});

test('unknown calendar coverage and inconsistent counts retain only subwindows for review',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 for(const [range,warning]of [['2027/10/14至2027/10/20',/2027|交易日历/],['2026/9/8至2026/9/15',/天数|交易日/]]){
  const t=task([template('a',[['开放期',old]])]);const p=previewSupplementText(t,'本次开放日：'+range);assert.equal(p.updates.length,2);assert.match(p.warnings.join(''),warning);applySupplement(t,p);assert.equal(t.facts.filter(f=>f.key==='opening_announced_date'&&f.status==='pending').length,4);
 }
});

test('incomplete old windows and malformed explicit split cannot manufacture new subdates',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/20-8/21仅开放参与';
 for(const rule of ['', '\n分段规则：前2个交易日参与和退出，后4个交易日仅参与','\n分段规则：前2个交易日参与和退出']){
  const t=task([template('a',[['开放期',old]])]);const p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'+rule);assert.equal(p.updates.length,2);assert.ok(p.warnings.length>0);
 }
});

test('three plus two and template wording rules are preserved instead of assuming two plus three',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/19开放参与和退出，8/20-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]);applySupplement(t,previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14'));
 assert.equal(rendered(t,'a','av1'),'2026/9/8-2026/9/14：9/8-9/10开放参与和退出，9/11-9/14仅开放参与');
 const t2=task([template('a',[['开放期',old+'；前二个开放日开放参与和退出，后三个开放日仅开放参与']])]);applySupplement(t2,previewSupplementText(t2,'本次开放日：2026/9/8至2026/9/14'));
 assert.equal(rendered(t2,'a','av1'),'2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与；前二个开放日开放参与和退出，后三个开放日仅开放参与');
});

test('nontrading endpoints and overseas trading conventions stay explicit and do not shift dates silently',()=>{
 const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与';
 const t=task([template('a',[['开放期',old]])]),p=previewSupplementText(t,'本次开放日：2026/9/6至2026/9/11');assert.equal(p.updates.length,2);assert.match(p.warnings.join(''),/并非均为沪深交易日/);
 const abroad=task([template('a',[['开放期',old+'，开放日需为沪深与香港共同交易日。']])]),q=previewSupplementText(abroad,'本次开放日：2026/9/8至2026/9/14');assert.equal(q.updates.length,2);assert.match(q.warnings.join(''),/香港|境外/);
});

test('reversed old subwindow order does not prevent updating explicit overall dates',()=>{
 const old='2026/8/17-2026/8/21：8/20-8/21开放参与和退出，8/17-8/19仅开放参与';
 const t=task([template('a',[['开放期',old]])]),p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14');assert.equal(p.updates.length,2);assert.match(p.warnings.join(''),/顺序|空档|重叠/);applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),old.replace('2026/8/17-2026/8/21','2026/9/8-2026/9/14'));
});

test('mixed text supplements benchmark separately from performance hurdles and codes',()=>{
 const t=task([template('a',[['业绩比较基准','参考年化2.60%'],['业绩报酬','当1%＜年化收益率≤4%时计提5%；当年化收益率＞4%时计提50%。'],['产品代码','009999'],['开放期','每满12个月开放，每个开放期开放5个交易日。'],['本次开放日','2026年8月12日']])]);
 const p=previewSupplementText(t,'2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%；2.25% 业绩基准；产品代码：001234；本次开放日：2026/11/3；开放期：每满18个月，连续开放3个交易日');
 assert.throws(()=>previewSupplementText(t,'−2.25%业绩基准'),/负数/);
 assert.throws(()=>previewSupplement(t,getSupplementGroups(t).find(g=>g.kind==='benchmark').id,'-2.25%'),/负数/);
 assert.ok(p.updates.some(u=>u.key==='benchmark'&&u.value==='2.25%'));applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'参考年化2.25%');assert.equal(rendered(t,'a','av3'),'001234');assert.equal(rendered(t,'a','av4'),'每满18个月开放，每个开放期开放3个交易日。');assert.equal(rendered(t,'a','av5'),'2026年11月3日');
});

test('batch validation is atomic when a later value is invalid',()=>{
 const t=task([template('a',[['产品代码','009999'],['管理费','0.5%']])]);const code=t.facts.find(f=>f.key==='product_code'),fee=t.facts.find(f=>f.key==='management_fee'),before=JSON.stringify(t.materials[0].operations);
 assert.throws(()=>applySupplementBatch(t,[{updates:[{factId:code.id,value:'001234'}]},{updates:[{factId:fee.id,value:'wrong'}]}]),/数字|费率/);
 assert.equal(code.status,'pending');assert.equal(JSON.stringify(t.materials[0].operations),before);
});

test('single date updates only the start while the old end stays pending',()=>{
 const t=task([template('a',[['本次开放日','2026/8/12-2026/8/14']])]);const group=getSupplementGroups(t).find(g=>g.kind==='date');
 const p=previewSupplement(t,group.id,'2027.10.14');assert.equal(p.updates.length,1);assert.match(p.warnings.join(''),/结束日期/);applySupplement(t,p);
 assert.equal(rendered(t,'a','av1'),'2027/10/14-2026/8/14');assert.equal(t.facts.filter(f=>f.key==='opening_announced_date'&&f.status==='pending').length,1);assert.ok(t.materials[0].operations.some(o=>o.kind==='highlight'&&o.start>12));
 const t2=task([template('b',[['本次开放日','2026年8月12日']])]);applySupplement(t2,previewSupplement(t2,getSupplementGroups(t2).find(g=>g.kind==='date').id,'2027年10月14日'));assert.equal(rendered(t2,'b','bv1'),'2027年10月14日');
 assert.throws(()=>parseDateInput('2026/2/30',2026),/有效日期/);assert.throws(()=>parseDateInput('2026/11/5至2026/11/3',2026),/不能早于/);
 assert.equal(formatDateLike('2026/11/3至2026/11/5','8月12日至14日','date',2026),'11月3日至5日');
 assert.equal(formatDateLike('2026/11/3至2026/11/5','2026-08-12-2026-08-14','date',2026),'2026-11-03-2026-11-05');
 assert.throws(()=>parsePerformanceInput('2.0%＜年化收益率≤3.0%，计提20%；年化收益率≥3.0%，计提40%'),/无法对应|分段边界/);
});

test('confirmation keeps distinct originals, removes only selected annotations and preserves Word changes',()=>{
 const t=task([template('w',[['销售机构风险评级','R3（中风险）'],['管理费','0.8%'],['产品代码','009999']]),template('p',[['销售机构风险评级','R2（中低风险）'],['管理费','0.8%'],['产品代码','009999']],'pptx')],['管理费年费率为0.5%。']);
 const f=t.facts.find(f=>f.key==='risk_rating'),wordReplacements=JSON.stringify(t.materials[0].operations.filter(o=>o.kind==='replace'));
 confirmFacts(t,[f.id]);assert.equal(f.source,'user_confirmation');assert.equal(f.userConfirmed,true);assert.equal(f.status,'verified');
 for(const m of t.materials){assert.ok(!m.operations.some(o=>o.factIds.includes(f.id)));assert.ok(m.operations.some(o=>o.kind==='highlight'&&o.before==='009999'));}
 assert.equal(rendered(t,'w','wv1'),'R3（中风险）');assert.equal(rendered(t,'p','pv1'),'R2（中低风险）');assert.equal(JSON.stringify(t.materials[0].operations.filter(o=>o.kind==='replace')),wordReplacements);
 confirmFacts(t,[f.id],false);assert.equal(f.status,'pending');assert.ok(t.materials[1].operations.some(o=>o.kind==='highlight'&&o.factIds.includes(f.id)));
});

test('calendar provenance remains explicit and later supplement replaces original confirmation',()=>{
 const t=task([template('a',[['产品代码','009999']])]);const f=t.facts.find(f=>f.key==='product_code');confirmFacts(t,[f.id]);
 const g=getSupplementGroups(t).find(g=>g.kind==='code'),p=previewSupplement(t,g.id,'001234');p.source={file:'日历.xlsx',sheet:'产品',row:8,productName:target};applySupplement(t,p);
 assert.equal(f.source,'user_calendar');assert.equal(f.userConfirmed,undefined);assert.equal(f.evidence[0].source.row,8);assert.equal(rendered(t,'a','av1'),'001234');
 setSupplement(t,f.id,'001235');assert.equal(rendered(t,'a','av1'),'001235');
});
