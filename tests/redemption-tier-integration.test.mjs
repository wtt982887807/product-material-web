import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {analyzeMaterials,getSupplementGroups,previewSupplement,applySupplement} from '../src/review-engine.js';
import {readContract} from '../src/office-reader.js';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;

const original='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';
const decimalOriginal='持有天数＜180自然日，退出费率1.0%；持有天数≥180自然日，退出费率0.0%。';
function template(id,fee,label='退出费率'){
 const paragraphs=[],part='word/document.xml';
 const rows=[['产品名称','晨星稳健3号集合资产管理计划'],[label,fee],['管理费','管理费率0.6%/年']].map(([label,value],rowIndex)=>{
  const ids=[id+'l'+rowIndex,id+'v'+rowIndex];[label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,index:rowIndex*2+cellIndex,page:1,tableId:'t',rowIndex,cellIndex}));
  return{id:id+'r'+rowIndex,label,value,part,page:1,tableId:'t',rowIndex,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });return{id,name:id+'.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]};
}
function task(clauses=[original],fee=original,others=[]){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：晨星稳健9号集合资产管理计划','管理费率0.6%/年',...clauses].map((text,i)=>typeof text==='string'?{id:'b'+i,text}:text)},[template('a',fee),...others],{month:'202610'});}
const feeFacts=t=>t.facts.filter(f=>f.key==='redemption_fee');
const feeGroup=t=>getSupplementGroups(t).filter(g=>g.factIds.some(id=>feeFacts(t).some(f=>f.id===id)));
const field=t=>t.materials[0].fields.find(f=>f.label==='退出费率');

test('matching holding-day contract fees are verified automatically and appear in one review group',()=>{
 const t=task();assert.equal(feeFacts(t).length,2);assert.ok(feeFacts(t).every(f=>f.status==='verified'));assert.deepEqual(feeFacts(t).map(f=>f.value),['1%','0%']);
 assert.equal(feeGroup(t).length,1);assert.equal(field(t).after,original);assert.equal(field(t).status,'verified');
 assert.ok(!t.materials[0].operations.some(o=>o.factIds?.some(id=>feeFacts(t).some(f=>f.id===id))));
 assert.ok(feeFacts(t).every(f=>f.evidence.some(e=>e.quote.includes('180')&&e.quote.includes('%'))));
});
test('holding-day differences change only matching fee percentages and preserve template phrasing',()=>{
 const near=original.replace('1%','0.8%').replace('0%','0.2%'),t=task([original],near,[template('b','持有期不少于180天，赎回费率为0.3%；持有期少于180天，赎回费率为0.9%。','赎回费率')]);
 assert.equal(field(t).after,decimalOriginal);assert.equal(t.materials[1].fields.find(f=>f.label==='赎回费率').after,'持有期不少于180天，赎回费率为0.0%；持有期少于180天，赎回费率为1.0%。');
 assert.equal(feeGroup(t).length,1);for(const m of t.materials)for(const o of m.operations.filter(o=>o.kind==='replace'&&o.factIds?.some(id=>feeFacts(t).some(f=>f.id===id))))assert.match(o.before,/^\d+(?:\.\d+)?%$/);
});
test('only a missing or conflicting holding band remains pending',()=>{
 const t=task(['持有天数＜180自然日，退出费率1%。']);assert.equal(feeFacts(t).filter(f=>f.status==='verified').length,1);assert.equal(feeFacts(t).filter(f=>f.status!=='verified').length,1);
 const conflict=task([original,'持有期少于180天，退出费率2%。']);assert.ok(feeFacts(conflict).some(f=>f.status==='conflict'));assert.equal(feeFacts(conflict).find(f=>f.value==='0%')?.status,'verified');
});
test('different day boundaries do not silently move template holding periods',()=>{
 const t=task(['持有天数≤180自然日，退出费率1%；持有天数>180自然日，退出费率0%。']);assert.ok(feeFacts(t).every(f=>f.status!=='verified'));assert.equal(field(t).after,original);
});
test('complete manual tier text still previews and applies once after contract-based matching',()=>{
 const t=task([original],original.replace('1%','0.8%').replace('0%','0.2%')),group=feeGroup(t)[0];
 const p=previewSupplement(t,group.id,original);assert.equal(p.updates.length,2);assert.equal(p.changes[0].after,decimalOriginal);applySupplement(t,p);assert.equal(field(t).after,p.changes[0].after);
 const next=original.replace('1%','0.7%').replace('0%','0.1%');applySupplement(t,previewSupplement(t,feeGroup(t)[0].id,next));assert.equal(field(t).after,next);
 assert.throws(()=>previewSupplement(t,feeGroup(t)[0].id,'0.5%'),/分档|条件|完整/);
});
test('manual tiers from different share classes do not cross-update',()=>{
 const t=analyzeMaterials({name:'合同.docx',blocks:[{id:'n',text:'产品名称：晨星稳健9号集合资产管理计划'}],warnings:[]},[template('a',original,'A类份额退出费率'),template('b',original,'B类份额退出费率')],{month:'202610'});
 const groups=feeGroup(t);assert.equal(groups.length,2);const p=previewSupplement(t,groups[0].id,original.replace('1%','0.5%'));applySupplement(t,p);assert.ok(t.materials[0].fields[1].after.includes('0.5%'));assert.equal(t.materials[1].fields[1].after,original);
});
test('a single known holding band retains scalar supplementation while multiple bands require conditions',()=>{
 const text='持有天数＜180自然日，退出费率1%。',t=task([],text);
 const g=feeGroup(t)[0],p=previewSupplement(t,g.id,'退出费率：0.2%');assert.equal(p.updates.length,1);applySupplement(t,p);assert.equal(field(t).after,text.replace('1%','0.2%'));
});
test('flat and zero fees retain the existing processing path',()=>{
 for(const value of ['退出费率0.5%。','本计划不收取退出费。']){const t=task([value],value);assert.ok(feeFacts(t).every(f=>!f.redemptionTier));assert.equal(field(t).after,value);}
});
test('an actual DOCX fee table retains row evidence and pairs each holding condition with its rate',async()=>{
 const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',p=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
 const rows=[['持有天数','退出费率'],['D＜180自然日','1.0%'],['D≥180自然日','0%']];
 const zip=new JSZip();zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${p('产品名称：晨星稳健9号集合资产管理计划')}${p('退出费率如下表所示：')}<w:tbl>${rows.map(cells=>`<w:tr>${cells.map(t=>`<w:tc>${p(t)}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>${p('退出费100%归入集合计划资产。')}</w:body></w:document>`);
 const c=await readContract({name:'合同.docx',bytes:await zip.generateAsync({type:'uint8array'})}),t=analyzeMaterials(c,[template('a',original)],{month:'202610'});
 assert.ok(feeFacts(t).every(f=>f.status==='verified'));assert.deepEqual(feeFacts(t).map(f=>f.value),['1%','0%']);
 assert.ok(feeFacts(t).every(f=>f.evidence.some(e=>e.quote.includes('180')&&!e.quote.includes('100%'))));
});
