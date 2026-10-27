import test from 'node:test';
import assert from 'node:assert/strict';
import {buildRedemptionTierIndex,matchTemplateRedemptionTier,resolveRedemptionTier} from '../src/redemption-tiers.js';

const contract=lines=>({blocks:lines.map((text,i)=>typeof text==='string'?{id:'b'+i,text}:text)});
const atom=(text,rate,n=0)=>({key:'redemption_fee',before:rate,start:[...text.matchAll(new RegExp(rate.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'g'))][n].index});
const tiers=(text)=>[...text.matchAll(/\d+(?:\.\d+)?%/g)].map(m=>matchTemplateRedemptionTier(text,{key:'redemption_fee',start:m.index,before:m[0]}));
const result=(lines,text,rate,scope={})=>resolveRedemptionTier(buildRedemptionTierIndex(contract(lines)),matchTemplateRedemptionTier(text,atom(text,rate)),scope);
const two='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';
const table=(rows,headers=['持有天数','退出费率'],id='table-1')=>[headers,...rows].flatMap((row,rowIndex)=>row.map((text,cellIndex)=>({id:id+'-'+rowIndex+'-'+cellIndex,text,part:'word/document.xml',tableId:id,rowIndex,cellIndex})));

test('two holding-day tiers match their own atom without flattening the percentages',()=>{
 const c=buildRedemptionTierIndex(contract(['3、退出费率：'+two,'退出费100%归入集合计划资产。'])),t=tiers(two);
 assert.equal(t.length,2);assert.notEqual(t[0].condition,t[1].condition);
 assert.deepEqual(t.map(x=>resolveRedemptionTier(c,x,{}).value),['1%','0%']);
 assert.ok(t.every(x=>resolveRedemptionTier(c,x,{}).status==='verified'));
 assert.ok(resolveRedemptionTier(c,t[0],{}).evidence.some(e=>e.quote.includes('＜180自然日')&&e.quote.includes('1%')));
});

test('three tiers support chained comparisons and preserve inclusive boundaries',()=>{
 const text='持有天数D＜7自然日，退出费率1.5%；7自然日≤D＜30自然日，退出费率0.5%；D≥30自然日，退出费率0%。',c=buildRedemptionTierIndex(contract([text]));
 assert.deepEqual(tiers(text).map(t=>resolveRedemptionTier(c,t,{}).value),['1.5%','0.5%','0%']);
 const wrong='持有天数≤7自然日，退出费率2%。';assert.notEqual(result([text],wrong,'2%').status,'verified');
});

test('controlled day synonyms match, without inferring days from months or years',()=>{
 for(const [before,expected]of [['持有期不足180天，赎回费率0.5%。','1%'],['持有期不满180天，赎回费率0.5%。','1%'],['持有期不少于180天，赎回费率0.5%。','0%'],['持有期满180天，赎回费率0.5%。','0%']])assert.equal(result([two],before,'0.5%').value,expected,before);
 for(const before of ['持有期不满6个月，退出费率1%。','持有期不足1年，退出费率1%。','D<180，退出费率1%。']){const t=matchTemplateRedemptionTier(before,atom(before,'1%'));assert.ok(t?.warning);assert.notEqual(resolveRedemptionTier(buildRedemptionTierIndex(contract([two])),t,{}).status,'verified');}
 const trading='持有期不足180交易日，退出费率0.5%。';assert.notEqual(result([two],trading,'0.5%').status,'verified');
});

test('ordinary single fee, no fee and combined participation fee stay outside the tier path',()=>{
 for(const text of ['退出费率：0%。','本计划不收取退出费。','参与/退出费率0%。','退出费100%归入计划资产。']){const m=text.match(/\d+%/);assert.equal(matchTemplateRedemptionTier(text,m?{start:m.index,before:m[0],key:'redemption_fee'}:{start:0,before:text,key:'redemption_fee'}),null);}
 const c=buildRedemptionTierIndex(contract(['退出费率0%。']));assert.notEqual(resolveRedemptionTier(c,tiers(two)[0],{}).status,'verified');
});

test('structured table conditions and fee values are paired by row and deduplicate narrative copies',()=>{
 const rows=table([['D＜180自然日','1.0%'],['D≥180自然日','0%']]);const c=buildRedemptionTierIndex(contract([two,'退出费率如下表所示：',...rows,'退出费100%归入集合计划资产。']));
 for(const [i,t]of tiers(two).entries()){const r=resolveRedemptionTier(c,t,{});assert.equal(r.status,'verified');assert.equal(r.value,i?'0%':'1%');assert.ok(r.evidence.some(e=>e.blockIds?.includes('table-1-'+(i+1)+'-0')&&e.quote.includes(i?'0%':'1.0%')));}
});

test('table parsing does not zip unrelated cells or take a percentage from the following paragraph',()=>{
 const c=buildRedemptionTierIndex(contract(['退出费率如下表：',...table([['D＜180自然日','详见公告'],['D≥180自然日','0%']]),'退出费100%归入计划资产。']));
 assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');assert.equal(resolveRedemptionTier(c,tiers(two)[1],{}).value,'0%');
 const noMetadata=rows=>rows.map(({id,text})=>({id,text}));const flat=buildRedemptionTierIndex(contract(noMetadata(table([['D＜180自然日','1%'],['D≥180自然日','0%']]))));assert.notEqual(resolveRedemptionTier(flat,tiers(two)[0],{}).status,'verified');
});

test('same interval different rates and overlapping contradictory intervals remain conflicts',()=>{
 assert.equal(result([two,'持有天数＜180自然日，退出费率2%。'],two,'1%').status,'conflict');
 assert.equal(result([two,'持有天数＜200自然日，退出费率2%。'],two,'1%').status,'conflict');
 assert.equal(result([two,'持有天数≤180自然日，退出费率1%。'],two,'0%').status,'conflict');
});

test('missing interval does not rewrite template holding-day limits or borrow another tier',()=>{
 const old='持有天数＜90自然日，退出费率0.5%。';const r=result([two],old,'0.5%');assert.equal(r.status,'pending');assert.equal(r.value,null);
});

test('inline A and B share rates are separated, including subsequent clauses inheriting the local share',()=>{
 const text='A类份额：持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。B类份额：持有天数＜180自然日，退出费率2%；持有天数≥180自然日，退出费率0.5%。';
 const c=buildRedemptionTierIndex(contract([text]));assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'A'}).value,'1%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'B'}).value,'2%');assert.equal(resolveRedemptionTier(c,tiers(two)[1],{share:'B'}).value,'0.5%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');
 const template='B类份额：'+two;assert.equal(resolveRedemptionTier(c,tiers(template)[0],{}).value,'2%');
});

test('parent share and channel headings constrain child prose and table rows',()=>{
 const c=buildRedemptionTierIndex(contract(['三、A类份额','退出费用',two,'四、B类份额','退出费用',two.replace('1%','2%')]));
 assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'A'}).value,'1%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'B'}).value,'2%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');
 const channel=buildRedemptionTierIndex(contract(['渠道：建设银行',...table([['D＜180自然日','1%'],['D≥180自然日','0%']])]));assert.equal(resolveRedemptionTier(channel,tiers(two)[0],{channel:'建设银行'}).value,'1%');assert.equal(resolveRedemptionTier(channel,tiers(two)[0],{channel:'工商银行'}).status,'pending');assert.equal(resolveRedemptionTier(channel,tiers(two)[0],{}).status,'pending');
});

test('reference products, examples and temporary waivers do not become target regular rates',()=>{
 for(const lines of [['三、底层基金概况','退出费用',two],['三、计算示例',two],['三、临时费率优惠',two],['对特定投资者临时优惠：'+two]]){const r=result(lines,two,'1%');assert.equal(r.status,'pending');}
 const r=result(['三、底层基金概况',two.replace('1%','9%'),'四、本计划费用',two],two,'1%');assert.equal(r.status,'verified');assert.equal(r.value,'1%');
 assert.equal(result([two,'临时优惠期间，持有天数＜180自然日，退出费率0%。'],two,'1%').value,'1%');
});

test('ambiguous unscoped template cannot silently pick a shared rate while an applicable share exception differs',()=>{
 const c=buildRedemptionTierIndex(contract([two,'A类份额：持有天数＜180自然日，退出费率0.5%。']));assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'A'}).status,'conflict');
});

test('fee allocation percentage in the same sentence is not a rate or a reason to drop the real rate',()=>{
 const text='持有天数＜180自然日，退出费率1%，退出费100%归入集合计划资产。';
 assert.equal(matchTemplateRedemptionTier(text,atom(text,'100%')),null);
 const r=result([text],two,'1%');assert.equal(r.status,'verified');assert.equal(r.value,'1%');
});

test('table columns distinguish A and B rates and a separate reference-product table heading stays excluded',()=>{
 const c=buildRedemptionTierIndex(contract(table([['D＜180自然日','1%','2%'],['D≥180自然日','0%','0.5%']],['持有天数','A类份额退出费率','B类份额退出费率'])));
 assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'B'}).value,'2%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');
 const rows=table([['持有天数','退出费率'],['D＜180自然日','1%'],['D≥180自然日','0%']],['底层基金概况','']);
 assert.equal(result(rows,two,'1%').status,'pending');
});

test('unknown additional applicability conditions and mixed day/month versions stay pending',()=>{
 for(const text of ['违约退出时，持有天数＜180自然日，退出费率2%。','仅适用于特殊投资者，持有天数＜180自然日，退出费率2%。','通过某销售银行参与，持有天数＜180自然日，退出费率2%。'])assert.notEqual(result([text],two,'1%').status,'verified',text);
 assert.equal(result([two,'持有期限不足6个月，退出费率2%。'],two,'1%').status,'pending');
});

test('ordinary calculation examples never provide regular contractual fee tiers',()=>{
 for(const text of ['例如：'+two,'假设投资者'+two])assert.equal(result([text],two,'1%').status,'pending');
});

test('rate ceilings and discounts are not silently converted into exact contractual rates',()=>{
 for(const text of ['持有天数＜180自然日，退出费率不超过1%。','持有天数＜180自然日，按退出费标准的50%收取。','持有天数＜180自然日，退出费率为1%至2%。','持有天数＜180自然日，按退出费率1%的50%收取。']){
  assert.ok(matchTemplateRedemptionTier(text,{key:'redemption_fee',start:text.search(/\d+(?:\.\d+)?%/),before:text.match(/\d+(?:\.\d+)?%/)[0]})?.warning,text);
  assert.equal(result([text],two,'1%').status,'pending',text);
  assert.equal(result([two,text],two,'1%').status,'pending',text);
 }
});

test('a single rate accompanied by a separate minimum holding term stays on the original fee path',()=>{
 for(const text of ['退出费率为0.5%，单笔份额最短持有期180自然日。','最低持有期180天，退出费率0.5%。','持有期为180自然日，退出费率0.5%。'])assert.equal(matchTemplateRedemptionTier(text,atom(text,'0.5%')),null,text);
});

test('a non-exact table rate cannot be hidden by an older matching narrative but only affects its interval',()=>{
 for(const rate of ['不超过1%','详见公告','按原费率的50%']){
  const c=buildRedemptionTierIndex(contract([two,...table([['D＜180自然日',rate],['D≥180自然日','0%']])]));
  assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending',rate);assert.equal(resolveRedemptionTier(c,tiers(two)[1],{}).status,'verified',rate);
 }
});

test('a distributor-specific fee does not become the manager fee and open-ended starting rates stay pending',()=>{
 const text='代销机构收取的退出费率：持有期不足180天为2%。',c=buildRedemptionTierIndex(contract([text]));
 assert.equal(resolveRedemptionTier(c,tiers(two)[0],{issuer:'manager'}).status,'pending');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{issuer:'distributor'}).value,'2%');
 for(const suffix of ['起','及以上','以上','左右'])assert.equal(result(['持有期不足180天，退出费率为1%'+suffix+'。'],two,'1%').status,'pending',suffix);
});

test('an explicit no-fee tier in prose or a fee-table cell resolves to zero without treating temporary waivers as regular',()=>{
 const prose='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，不收取退出费用。';assert.equal(result([prose],two,'0%').value,'0%');
 for(const value of ['不收取','免收退出费','无'])assert.equal(result(table([['D＜180自然日','1%'],['D≥180自然日',value]]),two,'0%').value,'0%',value);
 const limited='临时优惠期间，持有天数≥180自然日，免收退出费。';assert.equal(result([limited],two,'0%').status,'pending');
});

test('explicit unconditional product policies cover old template tiers only when no applicable tier evidence exists',()=>{
 for(const [text,rate]of [['本计划不收取退出费。','0%'],['本集合计划免收赎回费用。','0%'],['本计划退出费率统一为0.5%。','0.5%']]){
  const c=buildRedemptionTierIndex(contract([text]));for(const tier of tiers(two)){const r=resolveRedemptionTier(c,tier,{});assert.equal(r.status,'verified',text);assert.equal(r.value,rate);assert.ok(r.evidence.some(e=>e.quote===text));}
 }
 const ranked=buildRedemptionTierIndex(contract(['本计划退出费率统一为0.5%。',two]));assert.deepEqual(tiers(two).map(t=>resolveRedemptionTier(ranked,t,{}).value),['1%','0%']);
 const partial=buildRedemptionTierIndex(contract(['本计划不收取退出费。','持有天数＜180自然日，退出费率1%。']));assert.equal(resolveRedemptionTier(partial,tiers(two)[1],{}).status,'pending');
 const cell=buildRedemptionTierIndex(contract(table([['赎回费安排','本信托计划不收取赎回费。']],['产品要素','内容'])));assert.equal(resolveRedemptionTier(cell,tiers(two)[0],{}).value,'0%');
});

test('uniform-policy fallback preserves scope and cannot hide conditional, conflicting or undetermined evidence',()=>{
 for(const lines of [['三、底层基金概况','本计划不收取退出费。'],['临时优惠期间，本计划不收取退出费。'],['持有期超过180天时，本计划不收取退出费。'],['本计划退出费率统一不超过0.5%。'],['本计划不收取退出费。','本计划退出费率统一为0.5%。'],['本计划不收取退出费。','持有天数＜180自然日，退出费率不超过1%。']])assert.notEqual(result(lines,two,'1%').status,'verified',lines.join('\n'));
 const c=buildRedemptionTierIndex(contract(['A类份额','本计划不收取退出费。']));assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'A'}).value,'0%');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{share:'B'}).status,'pending');assert.equal(resolveRedemptionTier(c,tiers(two)[0],{}).status,'pending');
 const issuer=buildRedemptionTierIndex(contract(['代销机构费用','本计划不收取退出费。']));assert.equal(resolveRedemptionTier(issuer,tiers(two)[0],{issuer:'manager'}).status,'pending');
});
