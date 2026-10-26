import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,getSupplementGroups,previewSupplement,applySupplement} from '../src/review-engine.js';
import {parseOpeningInput} from '../src/supplement-parser.js';
const target='晨星全周期QDII3号集合资产管理计划';
function template(id,value,kind='docx'){
 const part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml',paragraphs=[];
 const rows=[['产品名称','晨星全周期QDII6号集合资产管理计划'],['开放期',value]].map(([label,value],i)=>{
  const ids=[id+'l'+i,id+'v'+i];[label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));
  return{id:id+'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });return{id,name:id+'.'+kind,kind,paragraphs,rows,slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素'}]:[],warnings:[]};
}
function task(templates,clauses){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+target,...clauses].map((text,i)=>({id:'b'+i,text}))},templates,{month:'202610'});}
function rendered(t,id){const m=t.materials.find(m=>m.materialId===id);let text=m.template.paragraphs.find(p=>p.id===id+'v1').text;for(const o of m.operations.filter(o=>o.paragraphId===id+'v1'&&o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}
const clause='本集合计划自成立次日起每月第一个、第三个周三（如遇节假日顺延至下一个工作日）开放办理参与和退出业务。';
test('monthly ordinal weekdays update both materials with only the differing ordinal tokens',()=>{
 const old=clause.replace('第一个、第三个','第二个、第四个'),t=task([template('word',old),template('ppt',old,'pptx')],[clause]);
 assert.equal(rendered(t,'word'),clause);assert.equal(rendered(t,'ppt'),clause);
 const facts=t.facts.filter(f=>f.key==='opening_monthly_weekday');assert.equal(facts.length,1);assert.equal(facts[0].status,'verified');
 for(const m of t.materials)assert.deepEqual(m.operations.filter(o=>o.paragraphId===m.materialId+'v1'&&o.kind==='replace').map(o=>[o.before,o.after]),[['二','一'],['四','三']]);
});
test('manual monthly wording must not masquerade as a weekly supplement',()=>{
 assert.deepEqual(parseOpeningInput('每月第1个、第3个星期三'),{opening_monthly_weekday:'每月第1个、第3个周三'});
});
test('Chinese and Arabic enumeration styles retain their original words and only change ordinal or weekday digits',()=>{
 for(const [old,next]of [
  ['每月的第2个、第4个星期三','每月的第1个、第3个星期三'],
  ['每个自然月第二、第四个周三','每个自然月第一、第三个周三'],
  ['每月第二个星期五及第四个星期五','每月第一个星期三及第三个星期三'],
  ['每月第02个、第04个周3','每月第01个、第03个周3'],
 ]){
  const t=task([template('a','本计划'+old+'开放办理参与和退出；其他文字不变。')],[clause]);
  assert.equal(rendered(t,'a'),'本计划'+next+'开放办理参与和退出；其他文字不变。');
 }
});
test('weekly and monthly-day schedules remain independent from monthly ordinal weekdays',()=>{
 for(const [old,contract]of [['本计划每周三开放。',clause],['本计划每月3日开放。',clause]]){
  const t=task([template('a',old)],[contract]);assert.equal(rendered(t,'a'),old);assert.ok(!t.facts.some(f=>f.key==='opening_monthly_weekday'));
 }
 const t=task([template('a',clause.replace('第一个、第三个','第二个、第四个'))],['本计划每周三开放办理参与和退出。','本计划每月3日开放。']);
 assert.equal(t.facts.find(f=>f.key==='opening_monthly_weekday').status,'pending');
 assert.deepEqual(parseOpeningInput('每周四'),{opening_weekday:'四'});assert.deepEqual(parseOpeningInput('每月3日'),{opening_monthday:'3'});
});
test('different monthly candidates conflict rather than choosing the first contract mention',()=>{
 const old=clause.replace('第一个、第三个','第二个、第四个'),t=task([template('a',old)],[clause,clause.replace('第一个、第三个','第二个、第四个')]);
 assert.equal(t.facts.find(f=>f.key==='opening_monthly_weekday').status,'conflict');assert.equal(rendered(t,'a'),old);
 assert.ok(t.materials[0].operations.some(o=>o.kind==='highlight'&&o.before==='二'));
});
test('an incompatible item count in any template makes the shared schedule pending in both orders',()=>{
 const full=clause.replace('第一个、第三个','第二个、第四个'),single='本计划每月第二个周三开放。';
 for(const templates of [[template('a',full),template('b',single)],[template('b',single),template('a',full)]]){
  const t=task(templates,[clause]),f=t.facts.find(f=>f.key==='opening_monthly_weekday');assert.equal(f.status,'pending');assert.match(f.reason,/项数不同/);
  assert.equal(rendered(t,'a'),full);assert.equal(rendered(t,'b'),single);assert.ok(t.materials.every(m=>m.operations.some(o=>o.kind==='highlight')));
 }
});
test('manual complete monthly schedules update dynamic render values and reject incompatible partial updates atomically',()=>{
 const old='本计划每月第二个、第四个周三开放。',t=task([template('a',old),template('b',old,'pptx')],[clause]),g=getSupplementGroups(t).find(g=>g.kind==='opening');
 applySupplement(t,previewSupplement(t,g.id,'每月第2个、第4个周五'));
 for(const id of ['a','b'])assert.equal(rendered(t,id),'本计划每月第二个、第四个周五开放。');
 applySupplement(t,previewSupplement(t,g.id,'每月第1个、第3个星期二'));
 for(const id of ['a','b'])assert.equal(rendered(t,id),'本计划每月第一个、第三个周二开放。');
 const before=JSON.stringify(t.materials.map(m=>m.operations));assert.throws(()=>previewSupplement(t,g.id,'每月第1个周三'),/项数不同/);assert.equal(JSON.stringify(t.materials.map(m=>m.operations)),before);
});
test('unrecognized or truncated monthly clauses cannot hide behind a separately verified lockup term',()=>{
 for(const phrase of ['每月第1个周三、第3个周三、第5个','每月第二个、第四个周三、周五','每月首个、第三个周三','每月第六个周三']){
  const old='持有期为180天；本计划'+phrase+'开放。',t=task([template('a',old)],['持有期为180天。',clause]);
  assert.equal(rendered(t,'a'),old);assert.equal(t.facts.find(f=>f.key==='opening_monthly_weekday')?.status,'pending');assert.equal(t.materials[0].fields.find(f=>f.label==='开放期').status,'pending');
  assert.throws(()=>parseOpeningInput(phrase),/月内|完整|枚举/);
 }
});
test('disclosure weekdays and differently scoped share rules do not establish an opening schedule',()=>{
 const old=clause.replace('第一个、第三个','第二个、第四个');
 for(const cs of [['每月第一个、第三个周三为估值日，参与份额以该日净值确认。'],['B类份额每月第一个、第三个周三开放。']]){
  const t=task([template('a','A类份额'+old)],cs);assert.equal(t.facts.find(f=>f.key==='opening_monthly_weekday').status,'pending');assert.equal(rendered(t,'a'),'A类份额'+old);
 }
});
test('a disclosure schedule never inherits an opening context from another clause or joined block',()=>{
 const old='本计划每月第一个、第三个周三开放。';
 for(const clauses of [
  ['管理人每月第二个、第四个周三披露份额净值；开放期另行公告。'],
  ['开放安排','每月第二个、第四个周三披露份额净值。'],
  ['管理人每月第二个、第四个周三披露份额净值；本计划'+clause],
 ]){
  const t=task([template('a',old)],clauses);assert.equal(rendered(t,'a'),old);assert.ok(t.facts.find(f=>f.key==='opening_monthly_weekday').value!=='每月第2个、第4个周三');
 }
 const t=task([template('a',old)],['管理人每月第二个、第四个周三披露净值。','开放日','每月第一个、第三个周三']);assert.equal(t.facts.find(f=>f.key==='opening_monthly_weekday').status,'verified');
});
test('equivalent ordinal synonyms and weekdays preserve their exact template representation',()=>{
 const old='本计划每月第两个、第四个周天开放。',t=task([template('a',old)],['本计划每月第2个、第4个周日开放。']);
 assert.equal(rendered(t,'a'),old);assert.equal(t.materials[0].operations.filter(o=>o.paragraphId==='av1').length,0);
});
