import test from 'node:test';
import assert from 'node:assert/strict';
import {getSectionReview,resolveStrategySummary,resolveStrategyReplacement} from '../src/contract-sections.js';
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
const strategyTitles=['资产配置策略','商品期货投资策略','基金策略','风险平价策略','套期保值策略','国债期货投资策略','境外结构化票据投资策略'];
const strategyList=titles=>titles.map((t,i)=>`${i+1}、${t}`).join('；');
function strategyLines(titles=strategyTitles){return ['三、投资策略','（一）决策依据','依据合同履行投资决策职责。','（二）投资决策程序','管理人按程序执行交易。','（三）投资管理的方法和标准',...titles.flatMap((title,i)=>[`${i+1}、${title}`,`${title}的具体投资方法及管理要求。`,...(title==='基金策略'?['（1）定量策略','按照数据配置基金。','（2）定性策略','结合市场研究开展投资。']:[])]),`${titles.length+1}、固定收益类资产超过计划资产总值80%的发生情形和条件，及相应投资策略、主要资产类别`,'（1）发生情形和条件：在市场变化时，固定收益资产配置可以突破80%。','（2）相应投资策略、主要资产类别：以债券投资策略为主。','四、投资限制','不得从事违法投资。'];}

test('strategy title summary resolves at the matching hierarchy, not governance headings',()=>{
 const text=strategyList(strategyTitles),task=taskFor(strategyLines(),[text],'investment_strategy');
 const before=JSON.stringify(task),result=resolveStrategySummary(task.contract,text);
 assert.equal(result.status,'verified');assert.equal(result.value,text);
 assert.equal(result.evidence[0].itemCount,7);assert.equal(result.evidence[0].noteCount,1);
 assert.ok(result.evidence[0].quote.includes('决策依据'));assert.ok(result.evidence[0].quote.includes('突破80%'));
 const r=review(task);assert.equal(r.kind,'aligned');assert.equal(r.needsReview,false);assert.equal(r.materials[0].matched,7);
 assert.deepEqual(r.materials[0].items.map(i=>i.title),strategyTitles);assert.equal(r.materials[0].extra.length,0);
 assert.equal(r.notes.length,1);assert.match(r.notes[0].contractText,/突破80%/);
 assert.equal(JSON.stringify(task),before);
});

test('strategy discovery adapts to a different count and leaves nested fund strategies nested',()=>{
 const titles=['债券投资策略','基金策略','多资产轮动策略'];
 const task=taskFor(strategyLines(titles),[strategyList(titles)],'investment_strategy');
 const r=review(task);assert.equal(r.materials[0].total,3);assert.equal(r.needsReview,false);
 assert.ok(!r.materials[0].items.some(i=>i.title==='定量策略'));assert.equal(resolveStrategySummary(task.contract,strategyList(titles)).status,'verified');
});

test('missing or extra real strategy headings stay pending and are never expanded into prose',()=>{
 const titles=['商品期货投资策略','国债期货投资策略','资产配置策略'];
 const task=taskFor(strategyLines(titles),[strategyList(titles)],'investment_strategy');
 for(const list of [titles.slice(0,2),[...titles,'信用套利策略']]){
  const result=resolveStrategySummary(task.contract,strategyList(list));assert.equal(result.status,'pending');assert.equal(result.value,null);
 }
 task.materials[0].fields[0].after=strategyList(titles.slice(0,2));
 assert.equal(review(task).materials[0].items.find(i=>i.title==='资产配置策略').status,'missing');
});

test('commodity and treasury futures strategy titles cannot substitute for one another',()=>{
 const task=taskFor(strategyLines(['商品期货投资策略','国债期货投资策略']),['商品期货投资策略；商品期货投资策略'],'investment_strategy');
 assert.equal(resolveStrategySummary(task.contract,task.materials[0].fields[0].after).status,'pending');
 const r=review(task);assert.equal(r.materials[0].matched,1);assert.equal(r.materials[0].items.find(i=>i.title==='国债期货投资策略').status,'missing');
});

test('title summaries allow numbering and separator changes while preserving exact input text',()=>{
 const task=taskFor(strategyLines(['商品期货投资策略','国债期货投资策略']));
 const text='（二）国债期货策略\n（一）商品期货投资策略';
 const r=resolveStrategySummary(task.contract,text);assert.equal(r.status,'verified');assert.equal(r.value,text);
});

test('strategy resolution refuses full prose and unheaded fuzzy matches',()=>{
 const task=taskFor(strategyLines(['资产配置策略','基金策略']));
 assert.equal(resolveStrategySummary(task.contract,'本计划采用资产配置策略，按照风险收益情况动态调整投资比例。'),null);
 assert.equal(resolveStrategySummary({blocks:[{id:'x',text:'可以采用资产配置策略；基金策略。'}]},'资产配置策略；基金策略'),null);
});

test('strategy source versions and unresolved share scope cannot be silently chosen',()=>{
 const titles=['资产配置策略','基金策略'];
 const task=taskFor([...strategyLines(titles),'附件',...strategyLines(['资产配置策略','信用套利策略'])]);
 assert.equal(resolveStrategySummary(task.contract,strategyList(titles)).status,'conflict');
 const scoped=taskFor(['A类份额投资策略','1、资产配置策略','说明。','2、基金策略','说明。','B类份额投资策略','1、资产配置策略','说明。','2、信用套利策略','说明。','投资限制']);
 assert.notEqual(resolveStrategySummary(scoped.contract,strategyList(titles)).status,'verified');
 assert.equal(resolveStrategySummary(scoped.contract,strategyList(titles),{share:'A'}).status,'verified');
});

test('missing strategy adoption supplies only titles and preserves each material own order and separators',()=>{
 const task=taskFor(strategyLines(),[strategyList(strategyTitles.slice(0,6)),strategyTitles.slice(0,6).reverse().join('\n')],'investment_strategy');
 const r=review(task),a=r.materials[0],b=r.materials[1];
 assert.deepEqual(a.fieldIds,['field0']);assert.deepEqual(a.replacement.addedTitles,['境外结构化票据投资策略']);
 assert.equal(a.replacement.text,strategyList(strategyTitles));assert.equal(a.replacement.lines.length,1);
 assert.ok(b.replacement.text.startsWith(strategyTitles[5]+'\n境外结构化票据投资策略\n'+strategyTitles[4]));
 assert.ok(!a.replacement.text.includes('具体投资方法'));assert.ok(!a.replacement.text.includes('发生情形'));
 assert.ok(!b.replacement.text.includes('；'));
});

test('inserting a middle strategy keeps correct wording and ordering and uses the original separator style',()=>{
 const values=['资产配置策略','商品期货投资策略','基金策略'],text='（1）资产配置策略； （2）基金策略';
 const r=review(taskFor(strategyLines(values),[text],'investment_strategy')).materials[0].replacement;
 assert.equal(r.text,'（1）资产配置策略； （2）商品期货投资策略； （3）基金策略');
 assert.deepEqual(r.addedTitles,['商品期货投资策略']);assert.deepEqual(r.removedTitles,[]);
});

test('obsolete strategy headings are explicitly listed in the proposed replacement, never silently removed',()=>{
 const values=['资产配置策略','基金策略'],text='基金策略；资产配置策略；信用套利策略';
 const r=review(taskFor(strategyLines(values),[text],'investment_strategy')).materials[0].replacement;
 assert.equal(r.text,'基金策略；资产配置策略');assert.deepEqual(r.removedTitles,['信用套利策略']);assert.ok(r.warnings.some(w=>w.includes('信用套利策略')));
});

test('ambiguous strategy chapter or share never offers title adoption',()=>{
 const values=['资产配置策略','基金策略'];
 const versions=taskFor([...strategyLines(values),'附件',...strategyLines(['资产配置策略','信用套利策略'])],[strategyList(values)],'investment_strategy');
 assert.ok(review(versions).materials.every(m=>!m.replacement));
 const scoped=taskFor(['A类份额投资策略','1、资产配置策略','说明。','2、基金策略','说明。','B类份额投资策略','1、资产配置策略','说明。','2、信用套利策略','说明。','投资限制'],[strategyList(values)],'investment_strategy');
 assert.ok(review(scoped).materials.every(m=>!m.replacement));
});

test('independent containers in one material receive their own strategy numbering and separators',()=>{
 const values=['资产配置策略','基金策略','境外结构化票据投资策略'],contract=taskFor(strategyLines(values)).contract;
 const containers=[{shapeId:'one',text:'1、资产配置策略；2、基金策略'},{shapeId:'two',text:'（一）基金策略\n（二）资产配置策略'}];
 const before=JSON.stringify(containers),plans=containers.map(c=>resolveStrategyReplacement(contract,c.text));
 assert.equal(plans[0].text,'1、资产配置策略；2、基金策略；3、境外结构化票据投资策略');
 assert.equal(plans[1].text,'（一）基金策略\n（二）境外结构化票据投资策略\n（三）资产配置策略');
 assert.equal(plans[0].lines.length,1);assert.equal(plans[1].lines.length,3);assert.equal(JSON.stringify(containers),before);
});

test('per-container strategy replacement keeps the same reliable-scope and short-title gates',()=>{
 const values=['资产配置策略','基金策略'],contract=taskFor([...strategyLines(values),'附件',...strategyLines(['资产配置策略','信用套利策略'])]).contract;
 assert.equal(resolveStrategyReplacement(contract,strategyList(values)),undefined);
 assert.equal(resolveStrategyReplacement(taskFor(strategyLines(values)).contract,'本计划按照市场变化调整投资。'),undefined);
 assert.equal(resolveStrategyReplacement(null,strategyList(values)),undefined);
});

test('underlying-fund strategy chapter cannot supply target-product title adoption',()=>{
 const lines=['产品名称：目标9号集合资产管理计划','三、底层基金概况','投资策略','1、资产配置策略','配置说明。','2、信用套利策略','套利说明。','四、风险揭示'];
 const task=taskFor(lines,['1、资产配置策略；2、基金策略'],'investment_strategy');
 assert.equal(resolveStrategyReplacement(task.contract,task.materials[0].fields[0].after),undefined);
 assert.equal(resolveStrategySummary(task.contract,task.materials[0].fields[0].after),null);
 assert.equal(review(task),null);
});

test('standalone share and channel headers are inherited by nested strategy headings',()=>{
 const body=['投资策略','1、资产配置策略','配置说明。','2、信用套利策略','套利说明。','四、风险揭示'],text='1、资产配置策略；2、基金策略';
 const a=taskFor(['产品名称：目标9号集合资产管理计划','A类份额',...body]);
 assert.equal(resolveStrategyReplacement(a.contract,text,{share:'B'}),undefined);
 assert.ok(resolveStrategyReplacement(a.contract,text,{share:'A'}));
 assert.equal(resolveStrategyReplacement(a.contract,text,{}),undefined);
 const channel=taskFor(['产品名称：目标9号集合资产管理计划','渠道：建设银行',...body]);
 assert.equal(resolveStrategyReplacement(channel.contract,text,{channel:'招商银行'}),undefined);
 assert.ok(resolveStrategyReplacement(channel.contract,text,{channel:'建设银行'}));
});

test('an outer reference section ends at its peer and cannot hide the later own-product strategy',()=>{
 const lines=['产品名称：目标9号集合资产管理计划','三、底层基金概况','投资策略','1、资产配置策略','说明。','2、信用套利策略','说明。','四、本计划投资安排','投资策略','1、资产配置策略','说明。','2、基金策略','说明。','五、风险揭示'];
 const task=taskFor(lines,['资产配置策略；基金策略'],'investment_strategy'),result=resolveStrategySummary(task.contract,'资产配置策略；基金策略');
 assert.equal(result.status,'verified');assert.ok(result.evidence.every(e=>!e.quote.includes('信用套利策略')));
 assert.equal(review(task).materials[0].total,2);
});

test('a following standalone share heading is processed before its own strategy chapter',()=>{
 const lines=['产品名称：目标9号集合资产管理计划','A类份额','投资策略','1、资产配置策略','说明。','2、基金策略','说明。','B类份额','投资策略','1、商品期货投资策略','说明。','2、国债期货投资策略','说明。','风险揭示'];
 const c=taskFor(lines).contract;
 assert.equal(resolveStrategySummary(c,'商品期货投资策略；国债期货投资策略',{share:'B'})?.status,'verified');
 assert.equal(resolveStrategySummary(c,'资产配置策略；基金策略',{share:'A'})?.status,'verified');
});

test('ordinary investment-target item wording is not mistaken for an underlying-product parent heading',()=>{
 const text='1、投资标的包括债券、股票及公募基金';
 const r=review(taskFor(['二、投资范围',text,'2、其他经批准的资产','三、投资策略'],[text],'investment_scope'));
 assert.ok(r);assert.ok(r.evidence[0].quote.includes(text));assert.equal(r.evidence[0].itemCount,2);
});

test('large restrictions differences prefer complete evidence without losing clauses or the tail',()=>{
 const clauses=Array.from({length:15},(_,i)=>`${i+1}、本计划不得进行限制行为${i+1}。`),tail='本计划上述投资比例根据监管规定监控。';
 const r=review(taskFor(['投资限制',...clauses,tail,'风险揭示'],['遵守合同约定的相关投资限制。'],'investment_restrictions'));
 assert.equal(r.preferCompleteEvidence,true);assert.equal(r.evidence[0].itemCount,15);assert.equal(r.evidence[0].noteCount,1);
 assert.match(r.summary,/差异较大.*完整合同依据.*15项.*1段补充说明/);assert.ok(r.evidence[0].quote.includes(clauses[14]));assert.ok(r.evidence[0].quote.includes(tail));
 const local=review(taskFor(['投资限制','1、单一资产不得超过资产净值的10%。','风险揭示'],['1、单一资产不得超过资产净值的12%。'],'investment_restrictions'));
 assert.notEqual(local.preferCompleteEvidence,true);assert.equal(local.kind,'local');
 const smallVersions=review(taskFor(['投资限制','1、单一资产不得超过资产净值的10%。','风险揭示','附件一','投资限制','1、单一资产不得超过资产净值的12%。','风险揭示'],['1、单一资产不得超过资产净值的10%。'],'investment_restrictions'));
 assert.notEqual(smallVersions.preferCompleteEvidence,true,'a small ambiguous rule still uses ordinary detailed review');
});

test('restriction attachment boundary stops at separate-numbered related-party heading and does not merge version counts',()=>{
 const clauses=Array.from({length:15},(_,i)=>`${i+1}、本计划不得进行限制行为${i+1}。`),tail='本计划上述投资比例根据监管规定监控。';
 const lines=['投资限制',...clauses,tail,'风险揭示','附件一','三','投资限制',...clauses.slice(0,14),tail,'四','关联交易','交易管控由管理人负责。','五','建仓期','建仓期6个月。'];
 const r=review(taskFor(lines,['遵守合同约定限制。'],'investment_restrictions'));
 assert.deepEqual(r.evidence.map(e=>e.itemCount),[15,14]);assert.equal(r.materials[0].total,15);assert.equal(r.preferCompleteEvidence,true);
 assert.ok(r.evidence.every(e=>!e.quote.includes('建仓期')));assert.match(r.summary,/15项/);
});

test('complete numbered section includes all five ratios and tail note across pages',()=>{
 const task=taskFor(['3、投资比例',...ratioLines,note,'4、投资限制','不得投资禁止类资产。']);
 const before=JSON.stringify(task);const r=review(task);
 assert.equal(r.evidence[0].itemCount,5);assert.equal(r.evidence[0].noteCount,1);
 assert.ok(r.evidence[0].quote.includes('境外投资'));assert.ok(r.evidence[0].quote.includes('穿透原则'));
 assert.ok(!r.evidence[0].quote.includes('不得投资'));assert.equal(r.materials[0].total,5);
 assert.equal(JSON.stringify(task),before,'review must never mutate the engine result');
});

test('ratio adoption source preserves every original content line including introductory conditions and tail notes',()=>{
 const intro='本计划在正常市场条件下遵守下列投资比例：';
 const task=taskFor(['3、投资比例',intro,...ratioLines,note,'4、投资限制'],[ratioLines.slice(0,3).join('\n')]);
 const r=review(task);assert.deepEqual(r.replacement.lines,[intro,...ratioLines,note]);
 assert.ok(!r.replacement.lines.includes('3、投资比例'));assert.equal(r.replacement.lines[1],ratioLines[0]);
});

test('ambiguous ratio versions or scope cannot expose an adoption replacement',()=>{
 const task=taskFor(['A类份额投资比例',...ratioLines,'B类份额投资比例',...ratioLines.map(s=>s.replace('20%～100%','30%～100%')),'投资限制']);
 assert.equal(review(task).replacement,undefined);
 task.facts[0].scope.share='A';assert.deepEqual(review(task).replacement.lines,ratioLines);
 task.facts[0].scope.condition='仅限特定市场条件';assert.equal(review(task).replacement,undefined);
 const versions=taskFor(['投资比例',...ratioLines,'投资限制','附件','投资比例',...ratioLines.map(s=>s.replace('20%～100%','30%～100%')),'投资限制']);
 assert.equal(review(versions).replacement,undefined);
});

test('complete duplicate ratio sections keep one safe source while non-ratio fields expose none',()=>{
 const ratio=review(taskFor(['投资比例',...ratioLines,note,'投资限制','附件','投资比例',...ratioLines,note,'投资限制']));
 assert.deepEqual(ratio.replacement.lines,[...ratioLines,note]);
 const strategy=review(taskFor(strategyLines(['资产配置策略','基金策略']),['资产配置策略；基金策略'],'investment_strategy'));
 assert.equal(strategy.replacement,undefined);
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
 const one=getSectionReview(task,[ids[0]]);assert.deepEqual(one.materials[0].fieldIds,['m-field-2','m-field-3','m-field-4']);
 assert.equal(JSON.stringify({facts:task.facts,operations:task.materials[0].operations}),before);
});
