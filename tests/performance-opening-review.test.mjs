import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials} from '../src/review-engine.js';

const NAME='青澜信用优享18个月5号集合资产管理计划';
const contract=()=>({name:'合同.doc',warnings:[],blocks:[
 '产品名称：'+NAME,
 '本集合计划开放期：自本集合计划成立之日起至每满18个月之对日为首个开放日，每个开放期开放5个交易日。',
 '管理人按持有期年化收益率（R）分段收取业绩报酬。',
 '年化收益率（R）\t业绩报酬计提比例',
 'R≤R1\t不提取',
 'R1＜R≤R2\t5%',
 'R＞R2\t50%',
 'R1、R2为业绩报酬计提基准，具体值以管理人公告为准。',
].map((text,i)=>({id:'b-'+i,text}))});
function poster(content){
 const paragraphs=[NAME,...content].map((text,i)=>({id:'p-'+i,text,part:'ppt/slides/slide1.xml',index:i,page:1,shapeId:i?'body':'title',tableId:null}));
 return{id:'poster',name:'【海报】青澜信用优享18个月5号.pptx',kind:'pptx',warnings:[],paragraphs,rows:[],slides:[{page:1,title:NAME,text:paragraphs.map(p=>p.text).join('\n'),paragraphIds:paragraphs.map(p=>p.id)}]};
}
const ops=(task,id,kind)=>task.materials[0].operations.filter(o=>o.paragraphId===id&&(!kind||o.kind===kind));
function updated(task,id){let text=task.templates[0].paragraphs.find(p=>p.id===id).text;for(const o of ops(task,id,'replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}

test('a contract performance-table payout column is not a hurdle value',()=>{
 const task=analyzeMaterials(contract(),[poster(['业绩报酬计提：','当年化收益率≤2.25%时不提取；当2.25%＜年化收益率≤2.60%时，提取5%；超出2.60%的部分提取50%。'])]);
 const hurdles=task.facts.filter(f=>/^performance_R\d/.test(f.key));
 assert.equal(hurdles.length,2);
 assert.ok(hurdles.every(f=>f.status==='pending'&&f.value==null),'symbolic contract thresholds stay unresolved');
 assert.equal(ops(task,'p-2','replace').length,0,'the 5% payout is never copied into a hurdle');
 const payouts=task.facts.filter(f=>/^performance_ratio_/.test(f.key));
 assert.deepEqual(payouts.map(f=>f.value).sort(),['5%','50%']);
 assert.ok(payouts.every(f=>f.status==='verified'),'the payout column still supplies confirmed rates');
});

test('all repeated template hurdle values are yellow while identical payout numbers remain verified',()=>{
 const text=[
  '1）当年化收益率≤5.00%时，本集合计划不提取业绩报酬；',
  '2）当5.00%＜年化收益率≤6.00%时，对年化收益率大于5.00%且不高于6.00%的收益部分提取5%的业绩报酬；',
  '3）当年化收益率＞6.00%时，对年化收益率高于6.00%的收益部分提取50%的业绩报酬及对年化收益率大于5.00%且不高于6.00%的收益部分提取5%的业绩报酬，即E=K×[(R-6.00%)×50%+（6.00%-5.00%）×5%]。',
 ];
 const task=analyzeMaterials(contract(),[poster(['业绩报酬计提：',...text])]);
 for(const [i,line]of text.entries()){
  const id='p-'+(i+2),yellow=ops(task,id,'highlight');
  assert.equal(updated(task,id),line,'all words and all already-correct numbers stay unchanged');
  for(const m of line.matchAll(/(?:5|6)\.00%/g))assert.ok(yellow.some(o=>o.start===m.index&&o.before===m[0]),'every hurdle occurrence is yellow: '+m[0]+' at '+m.index);
  for(const m of line.matchAll(/(?<![\d.])(?:5|50)%/g))assert.ok(!yellow.some(o=>o.start<=m.index&&o.start+o.before.length>m.index),'contract-confirmed payout is not an unresolved hurdle');
 }
});

test('same old payout in two tiers is resolved by tier meaning rather than its value',()=>{
 const text='当2.25%＜年化收益率≤2.60%时，提取20%；当年化收益率＞2.60%时，对高于2.60%的收益部分提取20%，对大于2.25%且不高于2.60%的收益部分提取20%，即E=K×[(R-2.60%)×20%+（2.60%-2.25%）×20%]。';
 const task=analyzeMaterials(contract(),[poster(['业绩报酬计提：',text])]);
 assert.equal(updated(task,'p-2'),'当2.25%＜年化收益率≤2.60%时，提取5%；当年化收益率＞2.60%时，对高于2.60%的收益部分提取50%，对大于2.25%且不高于2.60%的收益部分提取5%，即E=K×[(R-2.60%)×50%+（2.60%-2.25%）×5%]。');
});

test('poster opening date paragraph does not detach the smaller cycle text below it',()=>{
 const cycle='自本集合计划成立之日起至每满12个月之对日为首个开放日，每个开放期开放5个交易日。';
 const t=poster(['开放期：','2026/8/18-2026/8/24（前2个开放日开放参与和退出，后3个开放日仅开放参与）',cycle,'历史业绩','截至2026年8月12日，过去12个月收益率为5.00%。']);
 t.paragraphs.filter(p=>p.index>=4).forEach(p=>p.shapeId='history');
 const task=analyzeMaterials(contract(),[t]);
 assert.equal(updated(task,'p-3'),cycle.replace('每满12个月','每满18个月'));
 assert.ok(ops(task,'p-2','highlight').length>0,'announcement dates still remain yellow');
 assert.equal(ops(task,'p-2','replace').length,0,'announcement dates are not invented');
 assert.deepEqual(ops(task,'p-5'),[],'historical performance dates and numbers remain untouched');
});
