import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,getSectionReview} from '../src/review-engine.js';

const name='晨星全周期9号集合资产管理计划';
const titles=['资产配置策略','商品期货投资策略','基金策略','风险平价策略','套期保值策略','国债期货投资策略','境外结构化票据投资策略'];
const list=values=>values.map((s,i)=>`${i+1}、${s}`).join('；');
function contract(values=titles){
 const lines=['产品名称：'+name,'三、投资策略','（一）决策依据','根据合同及法律规定执行。','（二）投资决策程序','严格遵守投资决策流程。','（三）投资管理的方法和标准',...values.flatMap((s,i)=>[`${i+1}、${s}`,`${s}对应的详细投资方法及风险管理要求。`]),`${values.length+1}、固定收益类资产超过计划资产总值80%的发生情形和条件，及相应投资策略、主要资产类别`,'在合同约定条件下调整固定收益投资比例。','四、投资限制'];
 return {name:'合同.docx',blocks:lines.map((text,i)=>({id:'c'+i,text})),warnings:[]};
}
function template(id,strategies){
 const paragraphs=[],rows=[['产品名称',name],...strategies].map(([label,input],i)=>{
  const texts=Array.isArray(input)?input:[input],part='word/document.xml',labelId=`${id}-l${i}`,ids=texts.map((_,j)=>`${id}-v${i}-${j}`);
  paragraphs.push({id:labelId,text:label,part,index:paragraphs.length,page:1,tableId:'own',rowIndex:i,cellIndex:0});
  for(const [j,text]of texts.entries())paragraphs.push({id:ids[j],text,part,index:paragraphs.length,page:1,tableId:'own',rowIndex:i,cellIndex:1});
  return {id:'r'+i,label,value:texts.join('\n'),part,page:1,tableId:'own',rowIndex:i,paragraphIds:[labelId,...ids],cells:[{text:label,paragraphIds:[labelId]},{text:texts.join('\n'),paragraphIds:ids}]};
 });
 return {id,name:id+'.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]};
}
const strategyFacts=task=>task.facts.filter(f=>f.key==='investment_strategy');
function unchangedStrategies(task,materialId){
 const m=task.materials.find(m=>m.materialId===materialId),fields=m.fields.filter(f=>/投资策略/.test(f.label)),ids=new Set(fields.flatMap(f=>f.paragraphIds));
 assert.ok(fields.every(f=>f.before===f.after));
 assert.ok(!m.operations.some(o=>ids.has(o.paragraphId)&&['replace','highlight','comment'].includes(o.kind)));
}

test('one-paragraph seven-title strategy is verified without rewriting or highlighting',()=>{
 const text=list(titles),t=analyzeMaterials(contract(),[template('one',[['投资策略',text]])]);
 assert.equal(strategyFacts(t).length,1);assert.equal(strategyFacts(t)[0].status,'verified');assert.equal(strategyFacts(t)[0].value,text);
 unchangedStrategies(t,'one');const r=getSectionReview(t,strategyFacts(t).map(f=>f.id));assert.equal(r.needsReview,false);assert.equal(r.materials[0].matched,7);
});

test('multiple paragraphs in one strategy field retain each original paragraph value',()=>{
 const paragraphs=titles.map((s,i)=>`${i+1}、${s}`),t=analyzeMaterials(contract(),[template('paragraphs',[['投资策略',paragraphs]])]);
 assert.equal(strategyFacts(t).length,7);assert.ok(strategyFacts(t).every(f=>f.status==='verified'));
 assert.deepEqual(strategyFacts(t).map(f=>f.value),paragraphs);unchangedStrategies(t,'paragraphs');
});

test('separate same-scope strategy rows are aggregated before any row is judged missing',()=>{
 const values=titles.map((s,i)=>`${i+1}、${s}`),t=analyzeMaterials(contract(),[template('rows',values.map(s=>['投资策略',s]))]);
 assert.equal(strategyFacts(t).length,7);assert.ok(strategyFacts(t).every(f=>f.status==='verified'));
 assert.deepEqual(strategyFacts(t).map(f=>f.value),values);unchangedStrategies(t,'rows');
});

test('A and B share strategy lists are never combined to manufacture full coverage',()=>{
 const lines=['产品名称：'+name,'A类份额投资策略','1、债券投资策略','说明。','2、基金策略','说明。','B类份额投资策略','1、商品期货投资策略','说明。','2、国债期货投资策略','说明。','投资限制'];
 const c={name:'合同.docx',blocks:lines.map((text,i)=>({id:'c'+i,text})),warnings:[]};
 const t=analyzeMaterials(c,[template('shares',[['A类份额投资策略','1、债券投资策略'],['B类份额投资策略','1、商品期货投资策略；2、国债期货投资策略']])]);
 const a=strategyFacts(t).filter(f=>f.scope.share==='A'),b=strategyFacts(t).filter(f=>f.scope.share==='B');
 assert.ok(a.length&&a.every(f=>f.status==='pending'));assert.ok(b.length&&b.every(f=>f.status==='verified'));
 const m=t.materials[0],bIds=new Set(m.fields.filter(f=>f.label.startsWith('B')).flatMap(f=>f.paragraphIds));
 assert.ok(!m.operations.some(o=>bIds.has(o.paragraphId)));
});

test('missing or additional strategy titles stay pending and never receive expanded contract text',()=>{
 for(const values of [titles.slice(0,6),[...titles,'信用套利策略']]){
  const text=list(values),t=analyzeMaterials(contract(),[template('mismatch',[['投资策略',text]])]);
  assert.ok(strategyFacts(t).every(f=>f.status==='pending'));assert.ok(strategyFacts(t).every(f=>f.value===null));
  const field=t.materials[0].fields.find(f=>f.label==='投资策略');assert.equal(field.after,text);
  assert.ok(!t.materials[0].operations.some(o=>field.paragraphIds.includes(o.paragraphId)&&o.kind==='replace'));
 }
});

test('a complete strategy list in another template cannot verify this template missing one title',()=>{
 const t=analyzeMaterials(contract(),[template('complete',[['投资策略',list(titles)]]),template('partial',[['投资策略',list(titles.slice(0,6))]])]);
 const status=id=>t.materials.find(m=>m.materialId===id).fields.find(f=>f.label==='投资策略').status;
 assert.equal(status('complete'),'verified');assert.equal(status('partial'),'pending');unchangedStrategies(t,'complete');
});
