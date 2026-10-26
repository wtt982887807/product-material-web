import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,previewSupplementText,applySupplement} from '../src/review-engine.js';
import {calendarOpeningRange,calendarSupplementText} from '../src/calendar-import.js';

const name='晨星稳健9号集合资产管理计划';
function template(id,text,rule=''){
 const rows=[['产品名称','晨星稳健3号集合资产管理计划'],['开放期',text],...(rule?[['开放规则',rule]]:[])],paragraphs=[];
 return{id,name:id+'.pptx',kind:'pptx',warnings:[],slides:[{page:1,title:'产品要素',text:'产品要素'}],paragraphs,rows:rows.map(([label,value],i)=>{
  const ids=[id+'l'+i,id+'v'+i];[label,value].forEach((text,j)=>paragraphs.push({id:ids[j],text,part:'ppt/slides/slide1.xml',page:1,index:i*2+j,tableId:'t',rowIndex:i,cellIndex:j}));
  return{id:id+'r'+i,label,value,part:'ppt/slides/slide1.xml',page:1,tableId:'t',paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 })};
}
function task(templates,clauses=[]){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：'+name,...clauses].map((text,i)=>({id:'b'+i,text}))},templates,{month:'202610'});}
const old='2026/8/17-2026/8/21：8/17-8/18开放参与和退出，8/19-8/21仅开放参与（若出现合同变更，以管理人公告为准）';
const expected='2026/9/8-2026/9/14：9/8-9/9开放参与和退出，9/10-9/14仅开放参与（若出现合同变更，以管理人公告为准）';
const after=(t,id)=>t.materials.find(m=>m.materialId===id).fields.find(f=>f.label==='开放期').after;

test('Excel end-only schedule is explicitly labeled as deadline, not opening start',()=>{
 const row={openingEnd:'2026/9/14',openingDateNote:'前2个开放日参与和退出，后3个开放日仅参与'};
 assert.deepEqual(calendarOpeningRange(row),{start:'',end:'2026/9/14',value:'',warning:''});
 assert.match(calendarSupplementText(row),/^本次开放截止日：2026\/9\/14/m);
 assert.match(calendarSupplementText(row),/分段规则/);
 assert.doesNotMatch(calendarSupplementText(row),/本次开放日：/);
});

test('end-only Excel preview derives five trading days and both subwindows before explicit application',()=>{
 const t=task([template('poster',old),template('ppt',old)]),before=JSON.stringify(t.materials.map(m=>m.operations));
 const p=previewSupplementText(t,calendarSupplementText({openingEnd:'2026/9/14'}));
 assert.equal(p.warnings.length,0);assert.equal(p.updates.length,6);
 assert.ok(p.notes.some(n=>/反推/.test(n)&&/5/.test(n)));
 assert.equal(JSON.stringify(t.materials.map(m=>m.operations)),before);
 assert.ok(p.updates.some(u=>u.basis?.kind==='trading_day_backfill'));
 applySupplement(t,p);for(const id of ['poster','ppt'])assert.equal(after(t,id),expected);
});

test('explicit contract total wins over old template duration and dates',()=>{
 const t=task([template('a','2026/8/17-2026/8/21','每个开放期开放5个交易日。')],['每个开放期为3个交易日。']);
 const p=previewSupplementText(t,'本次开放截止日：2026/9/14');assert.equal(p.warnings.length,0);
 assert.ok(p.notes.some(n=>/合同/.test(n)&&/3/.test(n)));applySupplement(t,p);
 assert.equal(after(t,'a'),'2026/9/10-2026/9/14');
});

test('a single-date template plus explicit five trading days can reconstruct the start',()=>{
 const t=task([template('a','下一开放期的首个开放日为2026年8月17日。','连续开放5个交易日。')]);
 const p=previewSupplementText(t,'本次开放截止日：2026年9月14日');assert.equal(p.warnings.length,0);applySupplement(t,p);
 assert.equal(after(t,'a'),'下一开放期的首个开放日为2026年9月8日。');
});

test('end-only backfill excludes National Day closures and weekends',()=>{
 const t=task([template('a',old)]),p=previewSupplementText(t,'本次开放截止日：2026/10/13');
 assert.equal(p.warnings.length,0);applySupplement(t,p);
 assert.equal(after(t,'a'),old.replace('2026/8/17-2026/8/21','2026/9/30-2026/10/13').replace('8/17-8/18','9/30-10/8').replace('8/19-8/21','10/9-10/13'));
});

test('conflicting totals, unknown calendars, nontrading ends and overseas rules cannot invent a start',()=>{
 for(const [templates,clauses,end,warning]of [
  [[template('a','2026/8/17-2026/8/21'),template('b','2026/8/17-2026/8/24')],[],'2026/9/14',/不一致|不同|冲突/],
  [[template('a',old)],[],'2027/10/14',/2027|日历/],
  [[template('a',old)],[],'2026/9/13',/交易日/],
  [[template('a',old+'，开放日应为沪深和香港共同交易日。')],[],'2026/9/14',/香港|境外/],
  [[template('a','2026/8/17')],[],'2026/9/14',/天数|完整/],
  [[template('a',old)],['每个开放期为3个交易日。','每个开放期为5个交易日。'],'2026/9/14',/合同|冲突|不一致/],
 ]){
  const t=task(templates,clauses),before=JSON.stringify(t.materials.map(m=>m.operations));
  const p=previewSupplementText(t,'本次开放截止日：'+end);assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),warning);
  assert.equal(JSON.stringify(t.materials.map(m=>m.operations)),before);
 }
});

test('explicit full range keeps existing behavior and an inconsistent deadline is rejected',()=>{
 const t=task([template('a',old)]),p=previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14');applySupplement(t,p);assert.equal(after(t,'a'),expected);
 assert.throws(()=>previewSupplementText(t,'本次开放日：2026/9/8至2026/9/14；本次开放截止日：2026/9/15'),/不一致|不同/);
 assert.throws(()=>previewSupplementText(t,'本次开放截止日：2026/9/14；档期结束日期：2026/9/15'),/不一致|不同/);
});

test('phase-specific opening days cannot replace or conflict with the overall duration',()=>{
 const t=task([template('a',old)],['每个开放期开放5个交易日，赎回开放2个交易日，仅参与开放3个交易日。']);
 const p=previewSupplementText(t,'本次开放截止日：2026/9/14');assert.equal(p.warnings.length,0);applySupplement(t,p);assert.equal(after(t,'a'),expected);
 const partial=task([template('a','2026/8/17')],['参与退出开放2个交易日，随后仅参与开放3个交易日。']);
 const q=previewSupplementText(partial,'本次开放截止日：2026/9/14');assert.equal(q.updates.length,0);assert.match(q.warnings.join(''),/总|天数|完整/);
});

test('separate contract share headings cannot supply an unscoped duration',()=>{
 const t=task([template('a',old)],['A类份额','每个开放期为3个交易日。','B类份额','每个开放期为5个交易日。']);
 const p=previewSupplementText(t,'本次开放截止日：2026/9/14');assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),/口径|条件|份额/);
 const b=template('b',old);b.rows[1].label=b.rows[1].cells[0].text='B类份额开放期';b.paragraphs.find(p=>p.id==='bl1').text='B类份额开放期';
 const scoped=task([b],['A类份额开放安排','每个开放期为3个交易日。','B类份额开放安排','每个开放期为5个交易日。']);
 const q=previewSupplementText(scoped,'本次开放截止日：2026/9/14');assert.equal(q.warnings.length,0,q.warnings.join('；'));applySupplement(scoped,q);
 assert.equal(scoped.materials[0].fields.find(f=>f.label==='B类份额开放期').after,expected);
});

test('underlying product duration cannot override the target template duration',()=>{
 for(const clauses of [['底层基金每个开放期为10个交易日。'],['底层基金开放安排','每个开放期为10个交易日。']]){
  const t=task([template('a',old)],clauses),p=previewSupplementText(t,'本次开放截止日：2026/9/14');
  assert.equal(p.warnings.length,0);applySupplement(t,p);assert.equal(after(t,'a'),expected);
 }
});

test('ordinary contract total-duration wording always outranks the stale three-day template',()=>{
 for(const clause of ['每个开放期为连续5个交易日。','每个开放期均为5个交易日。','每个开放期为5个连续交易日。']){
  const t=task([template('a','2026/8/17-2026/8/19')],[clause]),p=previewSupplementText(t,'本次开放截止日：2026/9/14');
  assert.equal(p.warnings.length,0,clause);applySupplement(t,p);assert.equal(after(t,'a'),'2026/9/8-2026/9/14',clause);
 }
 const t=task([template('a','2026/8/17-2026/8/19')],['每个开放期为3至5个交易日。']);
 const p=previewSupplementText(t,'本次开放截止日：2026/9/14');assert.equal(p.updates.length,0);assert.match(p.warnings.join(''),/合同|天数|口径/);
});
