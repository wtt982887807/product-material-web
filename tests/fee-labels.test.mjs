import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,getSupplementGroups,previewCardSupplement,applySupplement} from '../src/review-engine.js';

function template(id,rows){
 const paragraphs=[],part='word/document.xml';
 const values=[['产品名称','晨星稳健9号集合资产管理计划'],...rows];
 const table=values.map(([label,value],i)=>{
  const ids=[id+'l'+i,id+'v'+i];
  [label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));
  return{id:id+'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });return{id,name:id+'.docx',kind:'docx',paragraphs,rows:table,slides:[],warnings:[]};
}
function task(rows,clauses,others=[]){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：晨星稳健9号集合资产管理计划',...clauses].map((text,i)=>({id:'b'+i,text}))},[template('a',rows),...others],{month:'202610'});}
const after=(t,label)=>t.materials[0].fields.find(f=>f.label===label).after;

test('explicit subscription, participation and exit labels remain separate in a mixed fee row',()=>{
 const label='认购/参与/退出费率',t=task([[label,'认购费率3%；参与费率3%；退出费率3%。']],['认购费率为2%。','参与费率为1%。','退出费率为0%。']);
 assert.equal(after(t,label),'认购费率2%；参与费率1%；退出费率0%。');
 const fees=t.facts.filter(f=>f.key.endsWith('_fee'));
 assert.equal(fees.length,3);assert.ok(fees.every(f=>f.status==='verified'));
 assert.deepEqual(fees.map(f=>f.label),['认购费率','参与费率','退出费率']);
 assert.notEqual(fees.find(f=>f.value==='2%').scope.feeActions,fees.find(f=>f.value==='1%').scope.feeActions);
});

test('individual participation and subscription fields only use their matching contract action',()=>{
 const t=task([['认购费率','3%'],['参与费率','3%'],['退出费率','3%']],['认购费率2%；参与费率1%；退出费率0%。']);
 assert.equal(after(t,'认购费率'),'2%');assert.equal(after(t,'参与费率'),'1%');assert.equal(after(t,'退出费率'),'0%');
});

test('a common subscription participation rate requires evidence for both actions',()=>{
 for(const label of ['认购/参与费率','认购及参与费率','认购费率/参与费率','认申购费率']){
  const t=task([[label,'3%']],['认购费率2%。','参与费率2.0%。']);assert.equal(after(t,label),'2%');
  assert.equal(t.materials[0].fields.find(f=>f.label===label).status,'verified');
  for(const clauses of [['认购费率2%。'],['参与费率2%。'],['认购费率2%。','参与费率1%。']]){
   const missing=task([[label,'3%']],clauses);assert.equal(after(missing,label),'3%');assert.notEqual(missing.materials[0].fields.find(f=>f.label===label).status,'verified');
  }
 }
});

test('combined subscription participation and exit declarations can provide each individual fee',()=>{
 const t=task([['认购费率','3%'],['参与费率','3%'],['退出费率','3%']],['认购/参与/退出费率为2%。']);
 for(const label of ['认购费率','参与费率','退出费率'])assert.equal(after(t,label),'2%');
});

test('a single unlabeled participation exit rate needs all requested actions, not one side',()=>{
 for(const label of ['参与/退出费率','认购/参与/退出费率']){
  const actions=label.startsWith('认购')?['认购','参与','退出']:['参与','退出'];
  const t=task([[label,'3%']],actions.map(action=>action+'费率2%。'));assert.equal(after(t,label),'2%');
  const missing=task([[label,'3%']],['参与费率2%。']);assert.equal(after(missing,label),'3%');assert.notEqual(missing.materials[0].fields[1].status,'verified');
 }
});

test('common waived fees apply to each named action without borrowing another action',()=>{
 const t=task([['认购费率','3%'],['参与费率','3%']],['本计划不收取认购/参与费。']);
 assert.equal(after(t,'认购费率'),'0%');assert.equal(after(t,'参与费率'),'0%');
 const only=task([['认购费率','3%'],['参与费率','3%']],['本计划不收取参与费。']);
 assert.equal(after(only,'认购费率'),'3%');assert.equal(after(only,'参与费率'),'0%');
});

test('supplements for participation do not cross into subscription under a shared fee key',()=>{
 const t=task([['认购费率','认购费率3%。'],['参与费率','参与费率3%。']],[]);
 const target=t.facts.find(f=>f.label==='参与费率');assert.ok(target,'participation has its own fact');const g=getSupplementGroups(t).find(g=>g.factIds.includes(target.id));
 applySupplement(t,previewCardSupplement(t,g.id,'1%'));
 assert.equal(after(t,'参与费率'),'参与费率1%。');assert.equal(after(t,'认购费率'),'认购费率3%。');
 const prose='本计划参与费率为0.5%。';applySupplement(t,previewCardSupplement(t,g.id,prose));
 assert.equal(after(t,'参与费率'),prose);assert.equal(after(t,'认购费率'),'认购费率3%。');
});

test('explicit supplement action cannot change a participation card into subscription',()=>{
 const t=task([['参与费率','参与费率3%。']],[]),g=getSupplementGroups(t).find(g=>g.label==='参与费率');
 for(const text of ['认购费率1%','本计划认购费率为1%。','认购/参与费率为1%。'])assert.throws(()=>previewCardSupplement(t,g.id,text),/口径|认购|参与/);
 assert.equal(after(t,'参与费率'),'参与费率3%。');
 applySupplement(t,previewCardSupplement(t,g.id,'1%'));assert.equal(after(t,'参与费率'),'参与费率1%。');
});

test('common fee supplements must preserve both requested actions when naming a fee',()=>{
 const label='认购/参与费率',t=task([[label,'认购/参与费率3%。']],[]),g=getSupplementGroups(t).find(g=>g.label===label);
 for(const text of ['认购费率1%','本计划参与费率为1%。'])assert.throws(()=>previewCardSupplement(t,g.id,text),/口径|认购|参与/);
 applySupplement(t,previewCardSupplement(t,g.id,'1%'));assert.equal(after(t,label),'认购/参与费率1%。');
 const prose='本计划认购/参与费率为0.5%。';applySupplement(t,previewCardSupplement(t,g.id,prose));assert.equal(after(t,label),prose);
});
