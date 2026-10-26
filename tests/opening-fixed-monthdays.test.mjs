import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';

const name='晨星全周期QDII25号集合资产管理计划';
const rule='本集合计划自成立次日起每月1日、15日（如遇节假日顺延至下一个工作日）开放办理参与和退出业务。合同另有约定的，从其约定。除因根据合同变更或者展期等需要设立临时开放期供投资者退出外，本集合计划原则上不设立特别开放期。';
function template(text,label='开放期'){
 const paragraphs=[],part='ppt/slides/slide1.xml',rows=[['产品名称',name],[label,text]].map(([label,value],i)=>{
  const vals=Array.isArray(value)?value:[value],ids=vals.map((_,j)=>'v'+i+'-'+j),lid='l'+i;
  paragraphs.push({id:lid,text:label,part,page:1,index:paragraphs.length,tableId:'t',cellIndex:0});
  vals.forEach((text,j)=>paragraphs.push({id:ids[j],text,part,page:1,index:paragraphs.length,tableId:'t',cellIndex:1}));
  return{id:'r'+i,label,value:vals.join('\n'),part,page:1,tableId:'t',paragraphIds:[lid,...ids],cells:[{text:label,paragraphIds:[lid]},{text:vals.join('\n'),paragraphIds:ids}]};
 });return{id:'a',name:'模板.pptx',kind:'pptx',paragraphs,rows,slides:[{page:1,title:'产品要素',text:'产品要素'}],warnings:[]};
}
function task(old,clauses=[rule],label='开放期'){
 return engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+name,...clauses].map((text,i)=>({id:'b'+i,text}))},[template(old,label)],{month:'202610'});
}
const field=t=>t.materials[0].fields.find(f=>/开放期/.test(f.label));
const fact=t=>t.facts.find(f=>f.key==='opening_fixed_monthdays');

test('fixed multiple month dates update every corresponding digit and preserve 日/号 wording',()=>{
 for(const [old,next] of [
  ['本计划每月10日、25日开放办理参与和退出业务。','本计划每月1日、15日开放办理参与和退出业务。'],
  ['本计划每月10号和25号开放办理参与和退出业务。','本计划每月1号和15号开放办理参与和退出业务。'],
  ['本计划每月10、25号开放办理参与和退出业务。','本计划每月1、15号开放办理参与和退出业务。']
 ]){
  const t=task(old,['本计划每月1日、15日开放办理参与和退出业务。']);
  assert.equal(field(t).after,next);assert.equal(fact(t)?.status,'verified');
  assert.deepEqual(t.materials[0].operations.filter(o=>o.paragraphId==='v1-0'&&o.kind==='replace').map(o=>[o.before,o.after]),[['10','1'],['25','15']]);
 }
});

test('rollover and extra holding requirements require whole-rule review with complete body evidence',()=>{
 const old='本集合计划每月10号和25号开放办理参与和退出业务，非交易日顺延至下一个交易日；持有期为180个自然日。';
 const t=task(old,['释义','开放期','指本集合计划办理参与、退出业务的期间。','四、封闭期、开放期','2、开放期：'+rule]);
 assert.equal(fact(t)?.status,'pending');assert.equal(field(t).after,old);
 assert.ok(fact(t).evidence.some(e=>e.quote===rule));assert.ok(!fact(t).evidence.some(e=>/指本集合计划/.test(e.quote)));
 const review=engine.getOpeningRuleReview(t,field(t).factIds);
 assert.equal(review.key,'opening_fixed_monthdays');assert.equal(review.kind,'structure');assert.equal(review.needsReview,true);
 assert.deepEqual(review.replacement.lines,[rule]);assert.deepEqual(review.materials[0].fieldIds,[field(t).id]);
 assert.match(review.summary,/顺延|工作日/);assert.match(review.summary,/持有|180/);
});

test('identical complete multi-date rules verify the full paragraph rather than only the first date',()=>{
 const t=task(rule,['2、开放期：'+rule]);assert.equal(fact(t)?.status,'verified');assert.equal(field(t).after,rule);
 assert.equal(t.materials[0].operations.filter(o=>o.paragraphId==='v1-0').length,0);
 assert.deepEqual(engine.getOpeningRuleReview(t,field(t).factIds).replacement.lines,[rule]);
});

test('an exact continuing note already contained in the body is not duplicated into the rule',()=>{
 const note=rule.slice(rule.indexOf('除因')),t=task(rule,['2、开放期：'+rule,note]);
 assert.equal(fact(t)?.status,'verified');assert.deepEqual(engine.getOpeningRuleReview(t,field(t).factIds).replacement.lines,[rule]);
});

test('definitions, disclosure schedules, underlying products and other share schedules cannot supply a rule',()=>{
 const old='A类份额每月10日、25日开放办理参与和退出业务。';
 for(const clauses of [
  ['释义','开放期：指每月1日、15日开放办理参与和退出业务的期间。','三、集合计划概况'],
  ['管理人每月1日、15日披露净值，开放安排另行公告。'],
  ['底层基金每月1日、15日开放办理参与和退出业务。'],
  ['B类份额开放安排','每月1日、15日开放办理参与和退出业务。']
 ]){
  const t=task(old,clauses,'A类份额开放期');assert.equal(fact(t)?.status,'pending');assert.equal(field(t).after,old);
  assert.equal(engine.getOpeningRuleReview(t,field(t).factIds)?.replacement,undefined);
 }
});

test('body rules take priority over the earlier glossary and preserve complete conditional tail',()=>{
 const old=rule.replace('1日、15日','10日、25日');
 const t=task(old,['释义','开放期：指每月10日、25日办理参与和退出业务的期间。','三、集合计划概况','2、开放期：'+rule]);
 assert.equal(field(t).after,rule);assert.equal(fact(t)?.status,'verified');
 assert.ok(fact(t).evidence.every(e=>!e.quote.startsWith('指')&&!e.quote.includes('10日')));
});

test('different body date schedules in the same scope are a conflict with no automatic adoption',()=>{
 const old=rule.replace('1日、15日','10日、25日'),t=task(old,[rule,rule.replace('1日、15日','2日、16日')]);
 assert.equal(fact(t)?.status,'conflict');assert.equal(field(t).after,old);
 const review=engine.getOpeningRuleReview(t,field(t).factIds);assert.equal(review.replacement,undefined);assert.equal(review.kind,'uncertain');
});

test('a less prominent body occurrence with a contradictory holding limit cannot be ignored',()=>{
 const same='本计划每月1日、15日开放办理参与和退出业务。',old=same+'持有期为180个自然日。';
 const t=task(old,['开放期：'+old,same+'持有期为90个自然日。']);
 assert.equal(fact(t)?.status,'conflict');assert.equal(field(t).after,old);
 const review=engine.getOpeningRuleReview(t,field(t).factIds);assert.equal(review.replacement,undefined);assert.ok(review.evidence.some(e=>/90个自然日/.test(e.quote)));
});

test('a less prominent compatible body rule cannot hide an additional mandatory holding condition',()=>{
 const same='本计划每月1日、15日开放办理参与和退出业务。',note='每笔份额须持有满180个自然日后方可退出。';
 const t=task(same,['开放期：'+same,same+note]);
 assert.equal(fact(t)?.status,'pending');assert.equal(field(t).after,same);
 const review=engine.getOpeningRuleReview(t,field(t).factIds);assert.deepEqual(review.replacement.lines,[same+note]);assert.match(review.summary,/持有/);
});

test('an additional procedural cross-reference does not expand an already complete matching product rule',()=>{
 const crossReference='因根据合同变更或者展期等需要而设立临时开放期供投资者退出的程序及披露比照本部分“二、集合计划的退出”执行。';
 const t=task(rule,['开放期：'+rule,rule+crossReference]);
 assert.equal(fact(t)?.status,'verified');assert.deepEqual(engine.getOpeningRuleReview(t,field(t).factIds).replacement.lines,[rule]);
 assert.ok(fact(t).evidence.some(e=>e.quote.includes(crossReference)),'The procedural cross-reference stays visible in full evidence');
});

test('a separate mandatory holding-full-days paragraph remains part of the complete opening rule',()=>{
 const same='本计划每月1日、15日开放办理参与和退出业务。',note='每笔份额须持有满180个自然日后方可退出。';
 const t=task(same,[same,note,'五、投资范围']);
 assert.equal(fact(t)?.status,'pending');assert.equal(field(t).after,same);
 const review=engine.getOpeningRuleReview(t,field(t).factIds);assert.deepEqual(review.replacement.lines,[same,note]);assert.ok(review.evidence.some(e=>e.quote.includes(note)));
});

test('a prospective underlying fund cannot become an adoptable target opening rule',()=>{
 const same='本计划每月1日、15日开放办理参与和退出业务。';
 const t=task(same,['本计划拟投资的甲基金每月1日、15日开放办理参与和退出业务。']);
 assert.equal(fact(t)?.status,'pending');assert.equal(field(t).after,same);
 assert.equal(engine.getOpeningRuleReview(t,field(t).factIds).replacement,undefined);assert.equal(fact(t).evidence.length,0);
});

test('multi-paragraph field cannot validate only dates and ignore a separate 180-day holding paragraph',()=>{
 const old=['本计划每月10日、25日开放办理参与和退出业务。','每笔份额持有期为180个自然日。'];
 const t=task(old,['本计划每月1日、15日开放办理参与和退出业务。']);
 assert.equal(field(t).status,'pending');assert.equal(field(t).after,old.join('\n'));assert.equal(field(t).factIds.length,1);
 const review=engine.getOpeningRuleReview(t,field(t).factIds);assert.match(review.summary,/持有/);
});

test('single date, monthly ordinal weekday, and period schedules do not enter the added branch',()=>{
 for(const [old,body,key,value] of [
  ['本计划每月10日开放。','本计划每月15日开放。','opening_monthday','15'],
  ['本计划每月第二个、第四个周三开放。','本计划每月第一个、第三个周三开放。','opening_monthly_weekday','每月第1个、第3个周三'],
  ['本计划每满12个月开放。','本计划每满18个月开放。','opening_cycle','18']
 ]){
  const t=task(old,[body]);assert.equal(fact(t),undefined);assert.equal(t.facts.find(f=>f.key===key)?.value,value);
 }
});
