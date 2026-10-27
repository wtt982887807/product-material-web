import test from 'node:test';
import assert from 'node:assert/strict';
import * as engine from '../src/review-engine.js';

const original='参与金额＜300万元，参与费率为2.0%；参与金额≥300万元，参与费率为2%。';
function task(text=original,contractFee='参与费率为2%。',label='参与费率'){
 const paragraphs=[],part='word/document.xml';
 const rows=[['产品名称','晨星稳健9号集合资产管理计划'],[label,text],['管理费','管理费率0.6%/年']].map(([label,value],i)=>{
  const ids=['l'+i,'v'+i];[label,value].forEach((text,j)=>paragraphs.push({id:ids[j],text,part,index:i*2+j,page:1,tableId:'t',rowIndex:i,cellIndex:j}));
  return {id:'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });
 return engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'产品名称：晨星稳健9号集合资产管理计划'},{id:'f',text:contractFee},{id:'m',text:'管理费率0.6%/年'}]},[{id:'a',name:'一页通.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]}],{month:'202610'});
}
const field=t=>t.materials[0].fields[1];
const group=t=>engine.getSupplementGroups(t).find(g=>g.factIds.some(id=>t.facts.find(f=>f.id===id)?.key==='subscription_fee'));

test('equal participation amount tiers use one fee despite decimal display differences',()=>{
 const t=task();assert.equal(field(t).after,'参与费率为2%。');assert.equal(field(t).before,original);
 assert.equal(t.materials[0].fields[2].after,'管理费率0.6%/年');assert.equal(field(t).status,'verified');
 assert.equal(t.materials[0].operations.filter(o=>o.source==='equal_fee_tiers').length,1);
});
test('three complete amount bands collapse after contract-supported numeric updates',()=>{
 const text='参与金额＜100万元，参与费率为1.0%；100万元≤参与金额＜300万元，参与费率为0.8%；参与金额≥300万元，参与费率为0.6%。';
 assert.equal(field(task(text)).after,'参与费率为2%。');
});
test('complete same-fee bands work with other thresholds, units, and reversed ordering',()=>{
 const text='参与金额≥5000000元，参与费率为0.25%；参与金额＜500万元，参与费率为0.250%。';
 assert.equal(field(task(text,'参与费率0.25%。')).after,'参与费率为0.25%。');
});
test('unknown and different contract fees retain tier wording',()=>{
 assert.equal(field(task(original,'未给出参与费率。')).after,original);
 const different=original.replace('2.0%','1.0%');assert.equal(field(task(different,different)).after,different);
});
test('gaps, overlaps, restricted ranges, different scopes, and additional terms never become unconditional',()=>{
 for(const text of [
  original.replace('≥300','＞300'),original.replace('＜300','≤300'),
  original.replace('参与金额＜300','100万元≤参与金额＜300'),
  original.replace('参与金额≥300','参与金额≥400'),
  'A类份额：'+original,'代销渠道：'+original,
  original+'每笔最低收取100元。',original.replace('参与金额≥300','B类份额参与金额≥300'),
  original.replace('参与费率为2%。','认购费率为2%。'),
 ])assert.equal(field(task(text)).after,text,text);
});
test('explicit full fee wording remains exactly as entered, and undo restores automatic simplification',()=>{
 const t=task(),g=group(t),p=engine.previewCardSupplement(t,g.id,original);
 assert.equal(p.changes[0].before,'参与费率为2%。');assert.equal(p.changes[0].after,original);
 engine.applySupplement(t,p);assert.equal(field(t).after,original);
 engine.undoFeeTextSupplement(t,g.id);assert.equal(field(t).after,'参与费率为2%。');
});
test('unrelated supplements rebuild the simplification without duplicate or overlapping operations',()=>{
 const t=task(),management=t.facts.find(f=>f.key==='management_fee');
 engine.setSupplement(t,management.id,'0.7%');engine.setSupplement(t,management.id,'0.8%');
 assert.equal(field(t).after,'参与费率为2%。');assert.equal(t.materials[0].operations.filter(o=>o.paragraphId==='v1').length,1);
 assert.equal(t.materials[0].fields[2].after,'管理费率0.8%/年');
});
test('equal exit holding tiers simplify while unequal rates keep their conditions',()=>{
 const text='持有天数＜90自然日，退出费率为1.0%；持有天数≥90自然日，退出费率为1%。';
 assert.equal(field(task(text,text,'退出费率')).after,'退出费率为1%。');
 const different=text.replace('为1%。','为0%。');assert.equal(field(task(different,different,'退出费率')).after,different);
});
test('mixed fee row simplifies only the complete redundant group and keeps other fee wording',()=>{
 const text='认购费率为1%；'+original+'退出费率为0%。';
 assert.equal(field(task(text,'认购费率为1%；参与费率为2%；退出费率为0%。','认购/参与/退出费率')).after,'认购费率为1%；参与费率为2%。退出费率为0%。');
});
test('common subscription participation fees simplify without losing their common label',()=>{
 const label='认购/参与费率',text=original.replaceAll('参与费率',label);
 assert.equal(field(task(text,'认购费率2%；参与费率2%。',label)).after,'认购/参与费率为2%。');
 const subscription=original.replaceAll('参与','认购');assert.equal(field(task(subscription,'认购费率2%。','认购费率')).after,'认购费率为2%。');
});
