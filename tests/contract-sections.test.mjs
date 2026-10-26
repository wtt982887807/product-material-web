import test from 'node:test';
import assert from 'node:assert/strict';
import {getSectionReview} from '../src/contract-sections.js';
import {analyzeMaterials} from '../src/review-engine.js';

const ratioLines=[
 '（1）固定收益类资产：占集合计划资产总值的20%～100%（不含）；',
 '（2）权益类资产：占集合计划资产总值的0%～40%（不含）；',
 '（3）期货和衍生品类账户权益：占集合计划资产总值的0%～20%（不含）；',
 '（4）公募基金：占集合计划资产总值的0%～80%；',
 '（5）境外投资：占集合计划资产总值的0%～50%（不含）。',
];
const note='本集合计划投资于公募证券投资基金，大类资产配置比例、总资产占净资产比例应当按照穿透原则合并计算。';
function taskFor(lines,materialTexts=[ratioLines.join('\n')],key='investment_ratios'){
 const labels={investment_ratios:'投资比例',investment_scope:'投资范围',investment_strategy:'投资策略',investment_restrictions:'投资限制',investment_objective:'投资目标'};
 const contract={name:'合同.docx',blocks:lines.map((text,i)=>({id:'b'+i,text,paragraphIndex:i,page:i<5?2:3})),warnings:[]};
 const facts=materialTexts.map((_,i)=>({id:'f'+i,key,label:labels[key],scope:{product:'测试产品',share:'',channel:'',condition:''},value:'unchanged',status:'verified',evidence:[],occurrences:[{materialId:'m'+i,fieldId:'field'+i,paragraphId:'p'+i}]}));
 const materials=materialTexts.map((text,i)=>({materialId:'m'+i,template:{name:'物料'+i+'.docx',paragraphs:[{id:'p'+i,text}]},operations:[],fields:[{id:'field'+i,label:labels[key],before:text,after:text,factIds:['f'+i],paragraphIds:['p'+i]}]}));
 return {contract,facts,materials,target:{name:'测试产品'}};
}
const review=t=>getSectionReview(t,t.facts.map(f=>f.id));

test('complete numbered section includes all five ratios and tail note across pages',()=>{
 const task=taskFor(['3、投资比例',...ratioLines,note,'4、投资限制','不得投资禁止类资产。']);
 const before=JSON.stringify(task);const r=review(task);
 assert.equal(r.evidence[0].itemCount,5);assert.equal(r.evidence[0].noteCount,1);
 assert.ok(r.evidence[0].quote.includes('境外投资'));assert.ok(r.evidence[0].quote.includes('穿透原则'));
 assert.ok(!r.evidence[0].quote.includes('不得投资'));assert.equal(r.materials[0].total,5);
 assert.equal(JSON.stringify(task),before,'review must never mutate the engine result');
});

test('each material gets its own missing coverage, never union across three templates',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,'投资限制'],[ratioLines.slice(0,3).join('\n'),ratioLines.join('\n'),ratioLines.slice(3).join('\n')]));
 assert.equal(r.materials[0].matched,3);assert.equal(r.materials[1].matched,5);assert.equal(r.materials[2].matched,2);
 assert.equal(r.materials[0].items.filter(i=>i.status==='missing').length,2);
 assert.match(r.summary,/公募基金|境外投资/);assert.equal(r.needsReview,true);
 assert.ok(r.materials[0].items.filter(i=>i.status==='missing').every(i=>/未体现|未呈现/.test(i.reason)));
});

test('identical attachment sections deduplicate without losing their source locations',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,note,'投资限制','其他限制。','附件一','投资比例',...ratioLines,note,'投资限制']));
 assert.equal(r.evidence.length,1);assert.equal(r.evidence[0].itemCount,5);assert.equal(r.evidence[0].duplicateCount,1);
 assert.ok(r.evidence[0].blockIds.includes('b0'));assert.ok(r.evidence[0].blockIds.includes('b10'));
});

test('separate Chinese heading number is not absorbed into prior section evidence',()=>{
 const r=review(taskFor(['二','投资比例',...ratioLines,note,'三','投资限制','附件一','二','投资比例',...ratioLines,note,'三','投资限制']));
 assert.equal(r.evidence.length,1);assert.equal(r.evidence[0].noteCount,1);assert.equal(r.evidence[0].duplicateCount,1);
 assert.ok(!r.evidence[0].quote.endsWith('三'));
});

test('look-through footnote does not cover the omitted public-fund ratio',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,note,'投资限制'],[ratioLines.slice(0,3).join('\n')+'\n'+note]));
 assert.deepEqual(r.materials[0].items.filter(i=>i.status==='missing').map(i=>i.title),['公募基金','境外投资']);
});

test('contradictory duplicate sections remain visible and uncertain',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,'投资限制','附件一','投资比例',...ratioLines.map(s=>s.replace('0%～80%','0%～60%')),'投资限制']));
 assert.equal(r.evidence.length,2);assert.equal(r.kind,'uncertain');assert.equal(r.needsReview,true);
 assert.match(r.summary,/不同|冲突|多个/);
 assert.ok(r.materials[0].items.some(i=>i.contractText.includes('0%～80%')));
 assert.ok(r.materials[0].items.some(i=>i.contractText.includes('0%～60%')));
 assert.ok(r.materials[0].items.filter(i=>i.title==='公募基金').every(i=>i.status==='uncertain'));
});

test('numeric-only drift remains local and does not create a new review blocker',()=>{
 const task=taskFor(['投资比例',ratioLines[0],'投资限制'],[ratioLines[0].replace('20%','30%')]);
 const before=JSON.stringify(task);const r=review(task);
 assert.equal(r.kind,'local');assert.equal(r.needsReview,false);assert.equal(r.materials[0].items[0].status,'local');
 assert.equal(JSON.stringify(task),before);
});

test('total versus net assets and inclusive boundaries are structural changes',()=>{
 for(const text of [ratioLines[0].replace('资产总值','资产净值'),ratioLines[0].replace('不含','含')]){
  const r=review(taskFor(['投资比例',ratioLines[0],'投资限制'],[text]));
  assert.equal(r.kind,'structure');assert.equal(r.materials[0].items[0].status,'structure');assert.equal(r.needsReview,true);
 }
});

test('semantic category matching ignores numbering but never matches solely by percentages',()=>{
 const r=review(taskFor(['投资比例',ratioLines[0],ratioLines[1],'投资限制'],['（1）权益类资产：占集合计划资产总值的0%～40%（不含）；\n（2）固定收益类资产：占集合计划资产总值的20%～100%（不含）；']));
 assert.equal(r.kind,'aligned');assert.equal(r.materials[0].matched,2);
 const wrong=review(taskFor(['投资比例',ratioLines[0],'投资限制'],['（1）房地产资产：占集合计划资产总值的20%～100%（不含）；']));
 assert.equal(wrong.materials[0].matched,0);assert.equal(wrong.materials[0].extra.length,1);assert.equal(wrong.needsReview,true);
});

test('summary wording may cover several categories without requiring verbatim paragraphs',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,'投资限制'],['固定收益类资产20%至100%（不含），权益类资产0%至40%（不含），期货和衍生品类账户权益0%至20%（不含）；公募基金0%至80%；境外投资0%至50%（不含）。']));
 assert.equal(r.materials[0].matched,5);assert.ok(r.materials[0].items.every(i=>i.status!=='missing'));
 assert.ok(!r.materials[0].items.some(i=>i.reason.includes('错误')));
});

test('nested item numbering is retained and long evidence is never clipped',()=>{
 const long='本条适用范围说明。'.repeat(220);
 const r=review(taskFor(['二、投资范围','（一）固定收益类资产','1、国债及地方政府债券。','2、政策性金融债。',long,'（二）权益类资产','1、上市公司股票。','三、投资策略','其他策略。'],['固定收益类资产、权益类资产'],'investment_scope'));
 assert.equal(r.evidence[0].itemCount,2);assert.ok(r.evidence[0].quote.length>1600);assert.ok(r.evidence[0].quote.includes('上市公司股票'));
 assert.ok(!r.evidence[0].quote.includes('其他策略'));
});

test('no reliable section heading leaves existing flow alone',()=>{
 const r=review(taskFor(['本产品的投资比例请参见公告。',ratioLines[0]],[ratioLines[0]]));assert.equal(r,null);
});

test('A and B share sections cannot be merged or silently selected without scope',()=>{
 const lines=['A类份额投资比例',ratioLines[0],'B类份额投资比例',ratioLines[0].replace('20%','30%'),'投资限制'];
 const task=taskFor(lines,[ratioLines[0]]);task.facts[0].scope.share='A';
 let r=review(task);assert.equal(r.evidence.length,1);assert.ok(!r.evidence[0].quote.includes('30%'));
 task.facts[0].scope.share='';r=review(task);assert.equal(r.kind,'uncertain');assert.equal(r.evidence.length,2);
});

test('coverage rereads current field.after after user supplementation',()=>{
 const task=taskFor(['投资比例',...ratioLines,'投资限制'],[ratioLines.slice(0,3).join('\n')]);
 assert.equal(review(task).needsReview,true);task.materials[0].fields[0].after=ratioLines.join('\n');
 assert.equal(review(task).needsReview,false);assert.equal(review(task).materials[0].matched,5);
});

test('same asset label cannot conceal prohibited-versus-permitted investment',()=>{
 const r=review(taskFor(['投资范围','公募基金：不得投资。','投资策略'],['公募基金：可以投资。'],'investment_scope'));
 assert.equal(r.kind,'structure');assert.equal(r.needsReview,true);
});

test('summary coverage never claims quantitative clauses are already checked',()=>{
 const r=review(taskFor(['投资比例',...ratioLines,'投资限制'],['固定收益类资产、权益类资产、期货和衍生品类账户权益、公募基金、境外投资']));
 assert.equal(r.materials[0].matched,5);assert.ok(r.materials[0].items.every(i=>i.status==='summary'));
 assert.match(r.summary,/摘要/);assert.ok(!/均已对应/.test(r.summary));
});

test('new conditions in the same asset category are uncertain, not harmless summaries',()=>{
 for(const text of [ratioLines[0]+'单一债券不得超过10%。',ratioLines[0].replace('；','，且单一债券不得超过10%。'),ratioLines[0].replace('20%','30%').replace('；','，仅投资境内资产。')]){
  const r=review(taskFor(['投资比例',ratioLines[0],'投资限制'],[text]));
  assert.ok(['uncertain','structure'].includes(r.kind));assert.equal(r.needsReview,true);
 }
});

test('changing contract blocks in place invalidates complete evidence cache',()=>{
 const task=taskFor(['投资比例',ratioLines[0],'投资限制'],[ratioLines[0]]);
 assert.equal(review(task).evidence[0].itemCount,1);
 task.contract.blocks.splice(2,0,{id:'added',text:ratioLines[1]});
 assert.equal(review(task).evidence[0].itemCount,2);
});

test('source evidence retains full-width glyphs and exact template wording',()=>{
 const text='（1）固定收益类资产：占集合计划资产总值的20%～100%（不含）；';
 const r=review(taskFor(['3、投资比例',text,'4、投资限制'],[text]));
 assert.ok(r.evidence[0].quote.includes(text));assert.equal(r.materials[0].items[0].templateText,text.slice(0,-1));
});

test('explicit semantic conditions cannot borrow coverage from another material scope',()=>{
 const task=taskFor(['投资比例',...ratioLines,'投资限制'],[ratioLines.slice(0,3).join('\n'),ratioLines.join('\n')]);
 task.facts[0].scope.condition='仅境内投资';task.facts[1].scope.condition='境内及境外投资';
 const r=getSectionReview(task,['f0']);assert.equal(r.materials.length,1);assert.equal(r.materials[0].matched,3);
 assert.equal(r.kind,'uncertain');
});

test('issuer and period remain part of review identity',()=>{
 const a=taskFor(['投资比例',ratioLines[0],'投资限制'],[ratioLines[0]]),b=structuredClone(a);
 a.facts[0].scope.issuer='manager';a.facts[0].scope.period='contract';
 b.facts[0].scope.issuer='distributor';b.facts[0].scope.period='announcement';
 assert.notEqual(review(a).id,review(b).id);
});

test('a correct occurrence cannot conceal a contradictory second occurrence in one material',()=>{
 const text=ratioLines[0]+'\n'+ratioLines[0].replace('资产总值','资产净值');
 const r=review(taskFor(['投资比例',ratioLines[0],'投资限制'],[text]));
 assert.equal(r.needsReview,true);assert.ok(r.materials[0].extra.some(i=>i.templateText.includes('资产净值')));
 const repeated=review(taskFor(['投资比例',ratioLines[0],'投资限制'],[ratioLines[0]+'\n'+ratioLines[0]]));
 assert.equal(repeated.needsReview,false);assert.equal(repeated.materials[0].extra.length,0);
});

test('real analyzeMaterials result uses all associated field paragraphs without changing exports',()=>{
 const name='晨星9号集合资产管理计划';
 const lines=['产品名称：'+name,'投资比例',...ratioLines,'投资限制'];
 const contract={name:'合同.docx',blocks:lines.map((text,i)=>({id:'c'+i,text})),warnings:[]};
 const paragraphs=[],rows=[['产品名称',name],...ratioLines.slice(0,3).map(text=>['投资比例',text])].map(([label,value],i)=>{
  const ids=['p'+i+'l','p'+i+'v'];paragraphs.push({id:ids[0],text:label,part:'word/document.xml',tableId:'t',rowIndex:i,cellIndex:0},{id:ids[1],text:value,part:'word/document.xml',tableId:'t',rowIndex:i,cellIndex:1});
  return {id:'r'+i,label,value,tableId:'t',rowIndex:i,page:1,part:'word/document.xml',paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });
 const task=analyzeMaterials(contract,[{id:'m',name:'一页通.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]}]);
 const ids=task.facts.filter(f=>f.key==='investment_ratios').map(f=>f.id);
 const before=JSON.stringify({facts:task.facts,operations:task.materials[0].operations});
 const r=getSectionReview(task,ids);assert.equal(r.materials[0].matched,3);assert.equal(r.materials[0].total,5);
 assert.equal(JSON.stringify({facts:task.facts,operations:task.materials[0].operations}),before);
});
