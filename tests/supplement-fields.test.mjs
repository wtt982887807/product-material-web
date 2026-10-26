import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';

function task(){const t={month:'202610',facts:[['date','opening_announced_date'],['fee','performance_R1'],['ratio','performance_ratio_1'],['benchmark','benchmark'],['code','product_code'],['cycle','opening_cycle'],['days','opening_days']].map(([id,key])=>({id,key,label:key,value:null,status:'pending',scope:{product:'示例',...(id==='date'?{condition:'opening|date|2026/8/12'}:{})},evidence:[],occurrences:[]})),_records:[],materials:[]};t._records.push({f:t.facts[0],atom:{dateRole:'date',before:'2026/8/12'},m:{materialId:'m'},field:{id:'date'}});return t;}
test('field parser accepts bare values and only targets the clicked semantic group',()=>{
 const t=task();
 for(const [kind,text,expected]of [
  ['date','2027.10.14',{opening_announced_date:'2027/10/14'}],
  ['date','2027年10月14日',{opening_announced_date:'2027/10/14'}],
  ['performance','超过业绩报酬计提基准2.50%的部分提取60%作为管理人业绩报酬',{performance_R1:'2.5%',performance_ratio_1:'60%'}],
  ['benchmark','业绩比较基准：2.25%',{benchmark:'2.25%'}],
  ['benchmark','2.25% 业绩基准',{benchmark:'2.25%'}],
  ['code','产品代码：000123',{product_code:'000123'}],
  ['opening','开放期：每满18个月，开放5个交易日',{opening_cycle:'18',opening_days:'5'}],
 ]){const p=engine.previewSupplementField(t,kind,text);assert.deepEqual(Object.fromEntries(p.updates.map(u=>[u.key,u.value])),expected);assert.ok(t.facts.every(f=>f.status==='pending'&&f.value===null));}
});
test('blank input, mixed fields, and ambiguous scopes never silently rewrite another field',()=>{
 const t=task();assert.throws(()=>engine.previewSupplementField(t,'code',''),/填写/);
 assert.throws(()=>engine.previewSupplementField(t,'benchmark','本次开放日：2027/10/14；业绩比较基准：2.25%'),/本行|分别/);
 t.facts.push({...t.facts.find(f=>f.id==='fee'),id:'feeB',scope:{product:'示例',share:'B'}});
 const p=engine.previewSupplementField(t,'performance','超过2.5%的部分提取60%');assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),/口径|份额/);
 assert.ok(t.facts.every(f=>f.value===null));
});
test('unmatched opening prose does not block a same-scope atomic opening update',()=>{
 const t=task();t.facts.push({id:'prose',key:'opening',label:'开放期',status:'pending',value:null,scope:{product:'示例'},evidence:[],occurrences:[]});
 const p=engine.previewSupplementField(t,'opening','每满18个月，开放5个交易日');
 assert.deepEqual(p.updates.map(u=>u.key),['opening_cycle','opening_days']);
 assert.match(p.warnings.join(''),/整段/);assert.equal(t.facts.at(-1).value,null);
});
test('opening prose belonging to another share class still requires explicit scope selection',()=>{
 const t=task();t.facts.push({id:'proseB',key:'opening',label:'开放期',status:'pending',value:null,scope:{product:'示例',share:'B'},evidence:[],occurrences:[]});
 const p=engine.previewSupplementField(t,'opening','每满18个月，开放5个交易日');assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),/份额|口径/);
});
