import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,setSupplement} from '../src/review-engine.js';

const OLD='青澜优选1号集合资产管理计划';
const TARGET='青澜优选9号集合资产管理计划';
const contract=()=>({name:'测试合同.docx',warnings:[],blocks:[
 {id:'contract-name',text:'产品名称：'+TARGET},
 {id:'contract-fee',text:'管理费按年费率0.2%计提。'},
 {id:'contract-custody',text:'托管费按年费率0.1%计提。'},
 {id:'contract-opening',text:'本计划开放日为每周二。'},
]});

function poster(items,{includeName=true,extraRows=[]}={}){
 const paragraphs=items.map((item,index)=>{
  const p=typeof item==='string'?{text:item}:item,page=p.page||1;
  return{id:p.id||'title-'+index,text:p.text,part:`ppt/slides/slide${page}.xml`,index,page,tableId:null,shapeId:p.shapeId||'shape-'+index};
 });
 const values=[...(includeName?[['产品名称',OLD]]:[]),['管理费','0.2%/年'],['托管费','0.1%/年'],...extraRows];
 const rows=values.map(([label,value],rowIndex)=>{
  const ids=['body-'+rowIndex+'-label','body-'+rowIndex+'-value'],part='ppt/slides/slide1.xml',tableId='body-table';
  for(const [cellIndex,text] of [label,value].entries())paragraphs.push({id:ids[cellIndex],text,part,index:paragraphs.length,page:1,tableId,rowIndex,cellIndex,shapeId:'body-table-shape'});
  return{id:'row-'+rowIndex,label,value,part,page:1,tableId,rowIndex,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });
 const slides=[...new Set(paragraphs.map(p=>p.page))].map(page=>({page,title:'产品海报',text:paragraphs.filter(p=>p.page===page).map(p=>p.text).join('\n'),paragraphIds:paragraphs.filter(p=>p.page===page).map(p=>p.id)}));
 return{id:'poster',name:'【测试海报】青澜优选1号.pptx',kind:'pptx',warnings:[],paragraphs,rows,slides};
}
const analyze=t=>analyzeMaterials(contract(),[t]);
const operations=(task,id)=>task.materials[0].operations.filter(o=>o.paragraphId===id);
function rendered(task,id){
 let text=task.templates[0].paragraphs.find(p=>p.id===id).text;
 for(const o of operations(task,id).filter(o=>o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);
 return text;
}
function assertDateMarked(task,id,date){
 const p=task.templates[0].paragraphs.find(p=>p.id===id),offset=p.text.indexOf(date);
 assert.ok(offset>=0,'test date exists in source paragraph');
 const highlights=operations(task,id).filter(o=>o.kind==='highlight');
 for(const digit of date.matchAll(/\d+/g))assert.ok(highlights.some(o=>o.start<=offset+digit.index&&o.start+o.before.length>=offset+digit.index+digit[0].length),`${id}: date number ${digit[0]} is marked`);
 assert.ok(highlights.every(o=>/^\d+(?:\.\d+)?\s*$/.test(o.before)),'only numeric date fragments are yellow');
 assert.equal(rendered(task,id),p.text,'missing date stays unchanged');
 assert.ok(task.facts.some(f=>f.key==='opening_announced_date'&&f.status==='pending'&&f.occurrences.some(o=>o.paragraphId===id)),'the date is a pending opening-date fact');
}

for(const [date,cue] of [
 ['8月12日','正式开售'],
 ['2026/8/12','正式开售'],
 ['2026年8月12日','正式开放'],
 ['8月12日至8月14日','开放申购'],
 ['8月12日至14日','开放申购'],
 ['8/12—8/14','正式开售'],
 ['8/12—14','开放申购'],
])test(`poster opening heading keeps and marks the date: ${date}`,()=>{
 const task=analyze(poster([{id:'sale-title',text:date+' '+cue}]));
 assertDateMarked(task,'sale-title',date);
});

test('poster date and sale cue in separate paragraphs of one shape are associated',()=>{
 const task=analyze(poster([{id:'date',text:'8月12日',shapeId:'sale-box'},{id:'cue',text:'正式开售',shapeId:'sale-box'}]));
 assertDateMarked(task,'date','8月12日');
 assert.deepEqual(operations(task,'cue'),[],'sale wording remains unchanged');
});

test('poster adjacent short date and sale text boxes on the same page are associated',()=>{
 const task=analyze(poster([{id:'date',text:'8月12日',shapeId:'date-box'},{id:'cue',text:'正式开售',shapeId:'sale-box'}]));
 assertDateMarked(task,'date','8月12日');
 assert.deepEqual(operations(task,'cue'),[]);
});

test('Chinese dates in an opening field remain pending alongside a verified weekly rule',()=>{
 const t=poster([],{extraRows:[['开放安排','每周三开放，下一开放期为8月12日至14日。']]}),id=t.rows.at(-1).cells[1].paragraphIds[0];
 const task=analyze(t),date='8月12日至14日',p=t.paragraphs.find(p=>p.id===id),offset=p.text.indexOf(date);
 for(const digit of date.matchAll(/\d+/g))assert.ok(operations(task,id).some(o=>o.kind==='highlight'&&o.start===offset+digit.index&&o.before.trim()===digit[0]));
 assert.equal(rendered(task,id),'每周二开放，下一开放期为8月12日至14日。');
 assert.ok(task.facts.some(f=>f.key==='opening_announced_date'&&f.status==='pending'));
});

for(const text of [
 '成立日期：2025年8月12日',
 '截至2026年8月12日，历史业绩为3.8%。',
 '历史募集期：2025年8月12日正式开售。',
 '本产品成立于2025年8月12日。',
 '8月12日',
])test(`unrelated poster dates are not opening dates: ${text}`,()=>{
 const task=analyze(poster([{id:'unrelated',text}]));
 assert.deepEqual(operations(task,'unrelated'),[]);
});

test('a sale cue on another page does not mark an isolated poster date',()=>{
 const task=analyze(poster([{id:'date',text:'8月12日',page:1},{id:'cue',text:'正式开售',page:2}]));
 assert.deepEqual(operations(task,'date'),[]);
});

for(const boundary of ['产品要素','风险等级','历史业绩'])test(`poster date context stops at a section or field boundary: ${boundary}`,()=>{
 const task=analyze(poster([{id:'date',text:'8月12日'},{id:'boundary',text:boundary},{id:'cue',text:'正式开售'}]));
 assert.deepEqual(operations(task,'date'),[]);
});

test('opening start and end date facts are separate when supplemented',()=>{
 const task=analyze(poster([{id:'start',text:'本次开放起始日为8月12日。'},{id:'end',text:'本次开放截止日为8月14日。'}]));
 assertDateMarked(task,'start','8月12日');
 assertDateMarked(task,'end','8月14日');
 const start=task.facts.find(f=>f.key==='opening_announced_date'&&f.occurrences.some(o=>o.paragraphId==='start'));
 const end=task.facts.find(f=>f.key==='opening_announced_date'&&f.occurrences.some(o=>o.paragraphId==='end'));
 assert.notEqual(start.id,end.id,'different date roles cannot share one editable fact');
 setSupplement(task,start.id,'2026/10/22');
 assert.ok(operations(task,'start').some(o=>o.kind==='replace'));
 assert.equal(rendered(task,'end'),'本次开放截止日为8月14日。');
 assert.ok(operations(task,'end').some(o=>o.kind==='highlight'));
 assert.equal(end.status,'pending');
});

test('poster product title beyond the first paragraph follows the contract',()=>{
 const task=analyze(poster([{id:'slogan',text:'稳健同行 长期相伴'},{id:'name',text:OLD}],{includeName:false}));
 assert.equal(rendered(task,'name'),TARGET);
 assert.deepEqual(operations(task,'slogan'),[]);
});

test('poster product title with a wrapped legal suffix and code keeps its layout',()=>{
 const before='青澜优选1号\n集合资产管理计划\n产品代码：（956731）';
 const task=analyze(poster([{id:'slogan',text:'稳健同行 长期相伴'},{id:'name',text:before}],{includeName:false}));
 assert.equal(rendered(task,'name'),before.replace('优选1号','优选9号'));
 assert.ok(!operations(task,'name').some(o=>o.kind==='replace'&&o.before.includes('956731')),'missing code is retained');
});

test('poster title identity does not rewrite underlying or historical product names',()=>{
 const task=analyze(poster([
  {id:'slogan',text:'稳健同行 长期相伴'},
  {id:'name',text:OLD},
  {id:'underlying-heading',text:'底层投资标的'},
  {id:'underlying-name',text:'晴川精选3号集合资产管理计划'},
  {id:'history-heading',text:'历史代表产品'},
  {id:'history-name',text:'清泉成长5号集合资产管理计划'},
 ],{includeName:false}));
 assert.equal(rendered(task,'name'),TARGET);
 assert.deepEqual(operations(task,'underlying-name'),[]);
 assert.deepEqual(operations(task,'history-name'),[]);
});

test('repeated template identity in a historical shape remains a historical product',()=>{
 const task=analyze(poster([
  {id:'name',text:OLD,shapeId:'current-title'},
  {id:'history-heading',text:'历史代表产品',shapeId:'historical-product'},
  {id:'history-name',text:OLD,shapeId:'historical-product'},
 ],{includeName:false}));
 assert.equal(rendered(task,'name'),TARGET);
 assert.deepEqual(operations(task,'history-name'),[],'a known old identity must not bypass the historical context guard');
});

test('date-role labels separate equal old opening start and end dates',()=>{
 const t=poster([],{extraRows:[['开放起始日','8月12日'],['开放截止日','8月12日']]}),startId=t.rows.at(-2).cells[1].paragraphIds[0],endId=t.rows.at(-1).cells[1].paragraphIds[0];
 const task=analyze(t);
 const start=task.facts.find(f=>f.key==='opening_announced_date'&&f.occurrences.some(o=>o.paragraphId===startId));
 const end=task.facts.find(f=>f.key==='opening_announced_date'&&f.occurrences.some(o=>o.paragraphId===endId));
 assert.ok(start&&end);
 assert.notEqual(start.id,end.id,'equal old dates with different date roles are independent facts');
 setSupplement(task,start.id,'2026/10/22');
 assert.ok(operations(task,startId).some(o=>o.kind==='replace'));
 assert.equal(rendered(task,endId),'8月12日');
 assert.ok(operations(task,endId).some(o=>o.kind==='highlight'));
});

test('a multi-line poster title preserves semantic line breaks across a different series name',()=>{
 const before='青澜资产\n福星君享年华\n12个月2号\n集合资产管理计划';
 const t=poster([{id:'name',text:before}],{includeName:false});
 const c={name:'合同.docx',warnings:[],blocks:[{id:'name',text:'产品名称：青澜资产福星君享（信用优享）18个月5号集合资产管理计划'}]};
 const task=analyzeMaterials(c,[t]);
 assert.equal(rendered(task,'name'),'青澜资产\n福星君享（信用优享）\n18个月5号\n集合资产管理计划');
 assert.ok(operations(task,'name').filter(o=>o.kind==='replace').every(o=>!/[\r\n]/.test(o.before+o.after)),'each edit stays inside its original title line');
});
