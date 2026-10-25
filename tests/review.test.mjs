import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,setSupplement} from '../src/review-engine.js';

const contract=(lines)=>({name:'合同.docx',blocks:lines.map((text,i)=>({id:`b${i}`,text})),warnings:[]});
function template(id,rows,{kind='docx',pages=[]}={}){
 const paragraphs=[], rr=rows.map(([label,value,page=1],i)=>{const ids=[`${id}-p${i}l`,`${id}-p${i}v`]; paragraphs.push({id:ids[0],text:label,part:kind==='docx'?'word/document.xml':`ppt/slides/slide${page}.xml`,index:i*2,page,tableId:`t${page}`,rowIndex:i,cellIndex:0},{id:ids[1],text:value,part:kind==='docx'?'word/document.xml':`ppt/slides/slide${page}.xml`,index:i*2+1,page,tableId:`t${page}`,rowIndex:i,cellIndex:1});return{id:`r${i}`,part:paragraphs.at(-1).part,page,tableId:`t${page}`,rowIndex:i,label,value,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};});
 return {id,name:`${id}.${kind}`,kind,paragraphs,rows:rr,slides:pages.map(([page,title])=>({page,title,text:title,paragraphIds:paragraphs.filter(p=>p.page===page).map(p=>p.id)})),warnings:[]};
}
test('contract split cover and explicit identity beats examples; replacements keep surrounding words',async()=>{
 const c=contract(['华星稳健三年2号','集合资产管理计划','管理合同','产品名称：华星稳健三年2号集合资产管理计划','管理费按前一日净资产的0.27%年费率计提。']);
 const t=template('a',[['产品名称','华星稳健三年9号集合资产管理计划'],['管理费','按年收取管理费0.63%，每日计提。']]);
 const x=await analyzeMaterials(c,[t]); assert.equal(x.target.name,'华星稳健三年2号集合资产管理计划');
 const op=x.materials[0].operations.find(o=>o.before==='0.63%'); assert.equal(op.after,'0.27%');
 assert.ok(!x.materials[0].operations.some(o=>o.before==='按年收取管理费0.63%，每日计提。'));
});
test('same scoped atom shared once; independent missing thresholds kept and marked',async()=>{
 const c=contract(['产品名称：华星稳健2号集合资产管理计划','业绩报酬基准R1、R2以管理人公告为准。','管理费率为0.3%。']);
 const a=template('a',[['产品名称','华星稳健1号集合资产管理计划'],['业绩报酬','R1=2.6%，R2=3.9%，超过部分按45%计提。'],['管理费','0.3%/年']]);
 const b=template('b',[['产品名称','华星稳健1号集合资产管理计划'],['业绩报酬','R1=2.7%，R2=4.1%'],['管理费','0.30%/年']]);
 const x=await analyzeMaterials(c,[a,b]);
 assert.equal(x.facts.filter(f=>f.key==='management_fee').length,1);
 for(const m of x.materials){ assert.ok(m.operations.some(o=>o.kind==='highlight'&&o.before.includes('%')));assert.ok(!m.operations.some(o=>o.kind==='replace'&&o.before.includes('%')));}
 assert.ok(x.facts.some(f=>f.key.includes('R1')&&f.status==='pending'));assert.ok(x.facts.some(f=>f.key.includes('R2')&&f.status==='pending'));
});
test('new field searches contract and evidence; unknown parameter is pending',async()=>{
 const c=contract(['产品名称：星河7号集合资产管理计划','估值频率：每周四','特别回购等待期：17个交易日']);
 const a=template('a',[['产品名称','星河2号集合资产管理计划'],['特别回购等待期','9个交易日'],['未披露费率','0.8%']]);
 const x=await analyzeMaterials(c,[a]);
 assert.ok(x.materials[0].operations.some(o=>o.kind==='replace'&&o.after==='17'));
 assert.ok(x.facts.some(f=>f.label==='特别回购等待期'&&f.evidence[0]?.quote.includes('17个交易日')));
 assert.ok(x.facts.some(f=>f.label==='未披露费率'&&f.status==='pending'));
});
test('issuer risk scope must not merge management and distributor ratings',async()=>{
 const c=contract(['产品名称：晨光8号集合资产管理计划','本计划管理人风险评级为R3（中风险）。','本计划不设预警线和止损线。']);
 const a=template('a',[['产品名称','晨光3号集合资产管理计划'],['管理人风险等级','R2（中低风险）'],['销售机构风险评级','R2（中低风险）']]);
 const x=await analyzeMaterials(c,[a]);const m=x.materials[0];
 assert.ok(m.operations.some(o=>o.paragraphId==='a-p1v'&&o.kind==='replace'&&o.after.includes('R3')));
 assert.ok(m.operations.some(o=>o.paragraphId==='a-p2v'&&o.kind==='highlight'));
});
test('select own element continuation and exclude trust underlying and marketing',async()=>{
 const c=contract(['产品名称：星河8号集合资金信托计划','固定信托报酬年费率为0.41%。']);
 const a=template('a',[['产品名称','星河1号集合资金信托计划',3],['固定信托报酬','0.6%/年',3],['认购起点','100万元',3],['开放日','每周五',4],['赎回费','0.2%',4],['产品名称','底层精选1号集合资产管理计划',5],['管理费','0.6%',5]],{kind:'pptx',pages:[[1,'历史业绩'],[2,'底层精选1号产品举例'],[3,'信托产品要素'],[4,'产品要素（续）'],[5,'底层投资标的产品要素']]});
 const x=await analyzeMaterials(c,[a]);assert.deepEqual(x.materials[0].selectedPages,[3,4]);assert.ok(x.materials[0].operations.every(o=>!o.part.includes('slide5')));
});
test('user supplement resolves missing atom and retains provenance',async()=>{
 const c=contract(['产品名称：曙光2号集合资产管理计划']);const a=template('a',[['产品名称','曙光1号集合资产管理计划'],['管理费','0.6%']]);const x=await analyzeMaterials(c,[a]); const f=x.facts.find(f=>f.key==='management_fee');
 setSupplement(x,f.id,'0.2%'); assert.equal(f.status,'verified');assert.equal(f.source,'user');assert.ok(x.materials[0].operations.some(o=>o.before==='0.6%'&&o.after==='0.2%'));
});
test('long clauses never share the first paragraph value; marketing tables untouched',async()=>{
 const c=contract(['产品名称：晨星9号集合资产管理计划','投资范围','（1）投资债券。','（2）投资货币基金。']);
 const t=template('a',[['管理人实力','管理人获得8个奖项'],['公司介绍','规模500亿元']]);
 const own=template('b',[['产品名称','晨星1号集合资产管理计划'],['投资范围','（1）投资债券。'],['投资范围','（2）投资货币基金。']]);
 t.rows.forEach(r=>r.tableId='marketing');t.paragraphs.forEach(p=>p.tableId='marketing');own.rows.forEach(r=>r.tableId='own');own.paragraphs.forEach(p=>p.tableId='own');t.rows.push(...own.rows);t.paragraphs.push(...own.paragraphs);
 const x=analyzeMaterials(c,[t]);assert.ok(x.materials[0].operations.every(o=>o.paragraphId.startsWith('b-')));assert.ok(x.materials[0].operations.every(o=>!o.before.includes('投资货币')));assert.equal(x.facts.filter(f=>f.key==='investment_scope').length,2);
});
test('split cover canonical name; short display and renamed manager preserve their own style',async()=>{
 const c=contract(['晨星均衡18个月','9号集合资产管理计划管理合同','管理人：晨星旧城资产管理有限公司']);
 const t=template('a',[['产品名称','晨星均衡18个月2号集合资产管理计划'],['管理人','晨星新城资产管理有限公司']]);t.paragraphs.push({id:'cover',part:'word/document.xml',index:8,tableId:null,text:'晨星均衡18个月2号'},{id:'alias',part:'word/document.xml',index:9,tableId:null,text:'晨星旧城资产管理有限公司于2025年更名为“晨星新城资产管理有限公司”。'});
 const x=analyzeMaterials(c,[t]);assert.equal(x.target.name,'晨星均衡18个月9号集合资产管理计划');const op=x.materials[0].operations.find(o=>o.paragraphId==='cover');assert.ok(op&&!op.after.includes('集合资产管理计划'));assert.ok(!x.materials[0].operations.some(o=>o.paragraphId==='a-p1v'));
});
test('table paragraph performance tiers stay same across split paragraphs',async()=>{
 const c=contract(['产品名称：晨星8号集合资产管理计划','R1与R2以公告为准。','当R1＜R≤R2时，提取7%的业绩报酬。','当R＞R2时，提取42%的业绩报酬。']);
 const t=template('a',[['产品名称','晨星1号集合资产管理计划'],['业绩报酬','当2%＜R≤3%时，提取7%的业绩报酬。'],['业绩报酬','当R＞3%时，提取51%的业绩报酬。']]);
 const x=analyzeMaterials(c,[t]);assert.ok(x.materials[0].operations.some(o=>o.before==='51%'&&o.after==='42%'));
});
test('trust composite fees, thresholds and weekday use shifted contract values',()=>{
 const c=contract(['产品名称：云舟成长4号集合资金信托计划','固定信托管理费按净值×【0.56】%/365计提。','受托人信托报酬按净值×【0.09】%/365计提。','固定投资顾问费按净值×【0.08】%/365计提。','固定销售服务费按净值×【0.39】%/365计提。','本计划开放日为每周二。','当年化收益率F≤1.75%不计提；当1.75%＜F≤2.85%时，计提21%；当F＞2.85%时，提取39%的浮动服务费。']);
 const t=template('a',[['产品名称','云舟成长1号集合资金信托计划'],['产品费率','固定管理费0.62%，信托管理费0.10%，固定投资顾问费0.08%，固定销售服务费0.44%'],['开放安排','每周五开放'],['浮动服务费','年化2.50%-3.50%区间计提20%，超过3.50%部分计提40%。']]);
 const x=analyzeMaterials(c,[t]);const ops=x.materials[0].operations.filter(o=>o.kind==='replace');
 for(const [before,after] of [['0.62%','0.56%'],['0.10%','0.09%'],['0.44%','0.39%'],['五','二'],['2.50%','1.75%'],['3.50%','2.85%'],['20%','21%'],['40%','39%']])assert.ok(ops.some(o=>o.before===before&&o.after===after),before+' -> '+after);
});
test('preceding example table cannot set the own element identity',()=>{
 const c=contract(['产品名称：晴空9号集合资产管理计划','管理费率为0.4%。']);const t=template('a',[['产品名称','其他示例1号集合资产管理计划',2],['管理费','0.9%',2],['产品名称','晴空3号集合资产管理计划',4],['管理费','0.7%',4]],{kind:'pptx',pages:[[2,'产品举例'],[4,'产品要素']]});const x=analyzeMaterials(c,[t]);assert.deepEqual(x.materials[0].selectedPages,[4]);assert.ok(x.materials[0].operations.every(o=>o.part.includes('slide4')));
});
test('all numeric pending occurrences retain original values, fractional formatting equivalent',()=>{
 const c=contract(['产品名称：光风9号集合资产管理计划','管理费年费率0.50%。','R1与R2为业绩报酬计提基准，以公告为准。']);const a=template('a',[['产品名称','光风3号集合资产管理计划'],['管理费','0.5%'],['业绩报酬','R1为1.1%，R2为2.2%，R1=1.1%。']]);const x=analyzeMaterials(c,[a]);assert.ok(!x.materials[0].operations.some(o=>o.kind==='replace'&&o.before.includes('%')));assert.equal(x.materials[0].operations.filter(o=>o.kind==='highlight').length,3);assert.equal(x.facts.filter(f=>f.key.startsWith('performance_R')).length,2);
});
test('explicit cover product code is reviewed but codes inside marketing prose are untouched',()=>{
 const c=contract(['产品名称：新泉8号集合资产管理计划','管理费率为0.2%。']);const t=template('a',[['产品名称','新泉3号集合资产管理计划'],['管理费','0.2%']]);t.paragraphs.push({id:'cover-code',part:'word/document.xml',index:8,tableId:null,text:'产品代码（956731）'},{id:'marketing-code',part:'word/document.xml',index:9,tableId:null,text:'同策略产品代码956731，过去一年收益率8%。'});
 const x=analyzeMaterials(c,[t]);const ops=x.materials[0].operations;assert.ok(ops.some(o=>o.paragraphId==='cover-code'&&o.kind==='highlight'&&o.before==='956731'));assert.ok(ops.some(o=>o.paragraphId==='cover-code'&&o.kind==='comment'));assert.ok(!ops.some(o=>o.paragraphId==='marketing-code'));assert.equal(x.facts.find(f=>f.key==='product_code').reasonType,'needs_review');
});
test('non-numeric unresolved fields retain honest review reason and source candidates',()=>{
 const c=contract(['产品名称：新泉8号集合资产管理计划','投资目标：投资债券，兼顾权益投资。']);const t=template('a',[['产品名称','新泉3号集合资产管理计划'],['投资目标','投资债券，追求稳健增长。']]);const x=analyzeMaterials(c,[t]);const f=x.facts.find(f=>f.key==='investment_objective');assert.equal(f.status,'pending');assert.equal(f.reasonType,'needs_review');assert.ok(f.evidence.length);assert.ok(!f.reason.includes('合同未提'));assert.ok(x.materials[0].operations.some(o=>o.kind==='comment'&&o.before.includes('投资债券')));
});
test('direct adviser identity in a disclosure follows the contract; biography is untouched',()=>{
 const c=contract(['产品名称：湖光9号集合资金信托计划','投资顾问：春江新城资产管理有限公司']);const t=template('a',[['产品名称','湖光3号集合资金信托计划'],['投资顾问','春江新城资产管理有限公司']]);t.paragraphs.push({id:'disclosure',part:'word/document.xml',index:8,tableId:null,text:'本信托计划由信托公司发行。秋日证券作为投资顾问，本材料中的投资顾问信息仅供参考。'},{id:'bio',part:'word/document.xml',index:9,tableId:null,text:'李先生曾在秋日证券担任研究员，拥有十年从业经验。'});const x=analyzeMaterials(c,[t]);assert.ok(x.materials[0].operations.some(o=>o.paragraphId==='disclosure'&&o.before==='秋日证券'&&o.after==='春江新城资产管理有限公司'));assert.ok(!x.materials[0].operations.some(o=>o.paragraphId==='bio'));
});
test('supplement validates numeric meaning before changing the review',()=>{
 const c=contract(['产品名称：湖光9号集合资产管理计划']);const t=template('a',[['产品名称','湖光3号集合资产管理计划'],['管理费','0.60%']]);const x=analyzeMaterials(c,[t]),f=x.facts.find(f=>f.key==='management_fee');assert.throws(()=>setSupplement(x,f.id,'abc'),/数字|费率/);assert.equal(f.status,'pending');setSupplement(x,f.id,'0.25');assert.equal(f.value,'0.25%');assert.ok(x.materials[0].operations.some(o=>o.after==='0.25%'));
});
test('standalone other-product example name outside tables remains untouched',()=>{
 const c=contract(['产品名称：湖光9号集合资产管理计划']);const t=template('a',[['产品名称','湖光3号集合资产管理计划'],['管理费','0.60%']]);t.paragraphs.push({id:'example-title',part:'word/document.xml',index:8,tableId:null,text:'产品举例'},{id:'example-name',part:'word/document.xml',index:9,tableId:null,text:'晴川1号集合资产管理计划'});const x=analyzeMaterials(c,[t]);assert.ok(!x.materials[0].operations.some(o=>o.paragraphId==='example-name'));
});
test('Word marketing example after own table does not inherit its identity',()=>{
 const c=contract(['产品名称：湖光9号集合资产管理计划','管理费率0.2%。','托管费率0.1%。']);const own=template('a',[['产品名称','湖光3号集合资产管理计划'],['管理费','0.60%']]);own.paragraphs.push({id:'break',part:'word/document.xml',index:4,tableId:null,text:'历史代表产品费用示例'});const other=template('other',[['管理费','0.9%'],['托管费','0.8%']]);other.paragraphs.forEach(p=>{p.index+=5;p.tableId='other';});other.rows.forEach(r=>r.tableId='other');own.rows.push(...other.rows);own.paragraphs.push(...other.paragraphs);const x=analyzeMaterials(c,[own]);assert.ok(x.materials[0].operations.every(o=>!o.paragraphId.startsWith('other-')));
});
test('poster marketing heading ends the preceding parameter field',()=>{
 const c=contract(['产品名称：青山9号集合资产管理计划','管理费率0.2%。']);const t={id:'a',name:'海报.pptx',kind:'pptx',rows:[],slides:[{page:1,title:'产品名称',text:''}],warnings:[],paragraphs:['产品名称','青山1号集合资产管理计划','管理费','0.5%/年','历史业绩','过去三年年化收益9.8%'].map((text,index)=>({id:'p'+index,text,index,page:1,part:'ppt/slides/slide1.xml',tableId:null}))};const x=analyzeMaterials(c,[t]);assert.ok(x.materials[0].operations.some(o=>o.paragraphId==='p3'&&o.after==='0.2%'));assert.ok(!x.materials[0].operations.some(o=>o.paragraphId==='p5'));
});
