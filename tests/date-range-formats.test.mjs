import test from 'node:test';
import assert from 'node:assert/strict';
import {parseDateInput,formatDateLike} from '../src/supplement-parser.js';
import {analyzeMaterials,previewSupplementField,applySupplement,getSupplementGroups} from '../src/review-engine.js';

const inputs=['2026年9月8日-2026年9月16日','2026年9月8日至2026年9月16日','2026.9.8-2026.9.16','2026.9.8至2026.9.16'];
function template(id,value,kind='docx',label='本次开放日'){
 const paragraphs=[],part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml';
 const rows=[['产品名称','晨星稳健3号集合资产管理计划'],[label,value]].map(([label,value],rowIndex)=>{
  const paragraphIds=[id+'l'+rowIndex,id+'v'+rowIndex];
  [label,value].forEach((text,cellIndex)=>paragraphs.push({id:paragraphIds[cellIndex],text,part,page:1,index:rowIndex*2+cellIndex,tableId:'t',rowIndex,cellIndex}));
  return{id:id+'r'+rowIndex,label,value,part,page:1,tableId:'t',rowIndex,paragraphIds,cells:[{text:label,paragraphIds:[paragraphIds[0]]},{text:value,paragraphIds:[paragraphIds[1]]}]};
 });
 return{id,name:id+'.'+kind,kind,paragraphs,rows,slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素'}]:[],warnings:[]};
}
function task(templates){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'name',text:'产品名称：晨星稳健9号集合资产管理计划'}]},templates,{month:'202610'});}
function rendered(t,id){const m=t.materials.find(m=>m.materialId===id);let text=m.template.paragraphs.find(p=>p.id===id+'v1').text;for(const o of m.operations.filter(o=>o.paragraphId===id+'v1'&&o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}

test('Chinese and dotted opening-date ranges accept both hyphen and zhi delimiters',()=>{
 for(const text of inputs){const parsed=parseDateInput(text,2026);assert.equal(parsed.value,'2026/9/8至2026/9/16');assert.deepEqual(parsed.end,{year:2026,month:9,day:16});}
 for(const text of ['2026-09-08-2026-09-16','2026/9/8至9/16','2026.9.8-16','2026年9月8日至16日'])assert.equal(parseDateInput(text,2026).value,'2026/9/8至2026/9/16');
 for(const text of ['2026-09-08','2026.9.8','2026年9月8日'])assert.equal(parseDateInput(text,2026).end,null);
 assert.throws(()=>parseDateInput('2026.9.16-2026.9.8',2026),/结束日期/);
 assert.throws(()=>parseDateInput('2026.9.8-2026.9.31',2026),/有效日期/);
 assert.equal(formatDateLike(inputs[2],'2026-08-12-2026-08-14','date',2026),'2026-09-08-2026-09-16');
});

test('each accepted input updates both endpoints across Chinese dotted and ISO templates',()=>{
 for(const input of inputs){
  const t=task([template('word','2026年8月12日至2026年8月14日'),template('ppt','2026.8.12-2026.8.14','pptx'),template('海报','2026-08-12-2026-08-14','pptx')]);
  const group=getSupplementGroups(t).filter(g=>g.kind==='date');assert.equal(group.length,1);assert.equal(group[0].factIds.length,6);
  const p=previewSupplementField(t,'date',input);assert.equal(p.updates.length,6);assert.equal(p.warnings.length,0);applySupplement(t,p);
  assert.equal(rendered(t,'word'),'2026年9月8日至2026年9月16日');assert.equal(rendered(t,'ppt'),'2026.9.8-2026.9.16');assert.equal(rendered(t,'海报'),'2026-09-08-2026-09-16');
 }
});

test('dotted template windows recompute by trading days and leave percentages alone',()=>{
 const old='2026.8.17-2026.8.21：8.17-8.18开放参与和退出，8.19-8.21仅开放参与；参与费率为2.50%。';
 for(const input of ['2026年9月8日-2026年9月14日','2026.9.8至2026.9.14']){
  const t=task([template('ppt',old,'pptx','开放期')]);
  assert.equal(t.facts.filter(f=>f.key==='opening_announced_date').length,6);
  const p=previewSupplementField(t,'date',input);assert.equal(p.warnings.length,0);applySupplement(t,p);
  assert.equal(rendered(t,'ppt'),'2026.9.8-2026.9.14：9.8-9.9开放参与和退出，9.10-9.14仅开放参与；参与费率为2.50%。');
 }
});
