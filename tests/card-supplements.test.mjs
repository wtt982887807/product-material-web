import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials,getSupplementGroups,previewSupplement,applySupplement,confirmFacts} from '../src/review-engine.js';

function template(id,rows){
 const paragraphs=[],part='word/document.xml';
 const items=[['产品名称','晨星稳健3号集合资产管理计划'],...rows].map(([label,value],i)=>{
  const ids=[id+'l'+i,id+'v'+i];[label,value].forEach((text,cellIndex)=>paragraphs.push({id:ids[cellIndex],text,part,page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));
  return{id:id+'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });return{id,name:id+'.docx',kind:'docx',paragraphs,rows:items,slides:[],warnings:[]};
}
function task(rows,extra=[],other=[]){return analyzeMaterials({name:'合同.docx',warnings:[],blocks:['产品名称：晨星稳健9号集合资产管理计划',...extra].map((text,i)=>({id:'b'+i,text}))},[template('a',rows),...other],{month:'202610'});}
const group=(t,key)=>getSupplementGroups(t).find(g=>g.factIds.some(id=>t.facts.find(f=>f.id===id)?.key===key));
const after=(t,label,index=0)=>t.materials[index].fields.find(f=>f.label===label).after;
const original='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';

test('single fee card accepts its explicit label and annual unit, without inserting prose into a number',()=>{
 for(const [label,key,before,input,expected]of [
  ['退出费率','redemption_fee','退出费率0.5%。','退出费率为0.2%。','退出费率0.2%。'],
  ['管理费','management_fee','管理费率0.6%/年。','管理费率：0.4%/年','管理费率0.4%/年。'],
  ['托管费','custody_fee','0.03%/年','0.02%/年','0.02%/年'],
  ['认购费','subscription_fee','认购费率0.5%','认购费率0.2%','认购费率0.2%'],
  ['退出费率','redemption_fee','退出费率0.5%。','免收退出费','退出费率0.0%。']
 ]){const t=task([[label,before]]),p=previewSupplement(t,group(t,key).id,input);assert.ok(p.updates.length);applySupplement(t,p);assert.equal(after(t,label),expected);}
});

test('one scalar cannot flatten multiple exit fee tiers, even when original percentages agree',()=>{
 for(const text of [original,original.replace('0%。','1%。')]){const t=task([['退出费率',text]]);assert.throws(()=>previewSupplement(t,group(t,'redemption_fee').id,'0.5%'),/分档|条件|多个/);assert.equal(after(t,'退出费率'),text);}
});

test('full exit fee clauses map by conditions across templates and produce real paragraph previews',()=>{
 const other='持有期少于180天，赎回费率为1%；持有期不少于180天，赎回费率为0%。';
 const t=task([['退出费率',original]],[],[template('b',[['赎回费率',other]])]);
 const p=previewSupplement(t,group(t,'redemption_fee').id,'持有天数≥180自然日，退出费率0.2%；持有天数＜180自然日，退出费率0.8%。');
 assert.equal(p.changes.length,2);assert.ok(p.changes.every(c=>c.before&&c.after&&c.materialName));
 applySupplement(t,p);assert.equal(after(t,'退出费率'),'持有天数＜180自然日，退出费率0.8%；持有天数≥180自然日，退出费率0.2%。');
 assert.equal(after(t,'赎回费率',1),'持有期少于180天，赎回费率为0.8%；持有期不少于180天，赎回费率为0.2%。');
 const q=previewSupplement(t,group(t,'redemption_fee').id,'持有天数＜180自然日，退出费率0.7%；持有天数≥180自然日，退出费率0.1%。');applySupplement(t,q);
 assert.match(after(t,'退出费率'),/0.7%.*0.1%/);
 const fee=t.facts.find(f=>f.key==='redemption_fee');confirmFacts(t,[fee.id]);confirmFacts(t,[fee.id],false);assert.match(after(t,'退出费率'),/0.7%.*0.1%/);
});

test('fee full-text matching rejects missing, changed and ambiguous conditions and tampered anchors atomically',()=>{
 const t=task([['退出费率',original]]),g=group(t,'redemption_fee'),snapshot=JSON.stringify(t.materials[0].operations);
 for(const input of ['退出费率0.8%；退出费率0.2%。','持有天数＜90自然日，退出费率0.8%；持有天数≥180自然日，退出费率0.2%。','持有天数＜180自然日，退出费率0.8%。'])assert.throws(()=>previewSupplement(t,g.id,input),/条件|对应|完整|分档/);
 assert.equal(JSON.stringify(t.materials[0].operations),snapshot);
 const p=previewSupplement(t,g.id,original.replace('1%','0.8%').replace('0%','0.2%'));p.updates[0].occurrenceValues[0].before='tampered';
 assert.throws(()=>applySupplement(t,p),/变化|对应|定位|原文/);assert.equal(JSON.stringify(t.materials[0].operations),snapshot);
});

test('risk descriptions share validation with codes, including full-width labeled input',()=>{
 for(const before of ['R3（中风险）','中风险']){const t=task([['风险等级',before]]),g=group(t,before.startsWith('R')?'risk_rating':'risk_description');
  assert.throws(()=>previewSupplement(t,g.id,'R9'),/R1|风险/);assert.throws(()=>previewSupplement(t,g.id,'R2（高风险）'),/不一致|冲突|风险/);
  applySupplement(t,previewSupplement(t,g.id,'风险等级：Ｒ２（中低风险）'));assert.equal(after(t,'风险等级'),before.startsWith('R')?'R2（中低风险）':'中低风险');
 }
});

test('typed number cards parse their own labels and reject incompatible units',()=>{
 for(const [label,key,before,input,wrong,expected]of [
  ['锁定期','lockup_days','持有期为180天','持有期为360天','12个月','持有期为360天'],
  ['存续期限','duration_years','存续期限为5年','存续期限为10年','18个月','存续期限为10年'],
  ['认购起点','minimum_amount','首次认购100万元','首次认购50万元','50年','首次认购50万元']
 ]){const t=task([[label,before]]),g=group(t,key);assert.throws(()=>previewSupplement(t,g.id,wrong),/单位|数字|填写/);applySupplement(t,previewSupplement(t,g.id,input));assert.equal(after(t,label),expected);}
});

test('whole prose supplements stay whole and never replace an entire lockup sentence with a bare number',()=>{
 const t=task([['锁定期','锁定期为12个月']]),g=group(t,'lockup');assert.throws(()=>previewSupplement(t,g.id,'18'),/完整|原文|表述/);
 applySupplement(t,previewSupplement(t,g.id,'锁定期为18个月'));assert.equal(after(t,'锁定期'),'锁定期为18个月');
 const o=task([['开放期','每满12个月进行一次开放，具体以公告为准。']]),og=group(o,'opening');
 const p=previewSupplement(o,og.id,'每满18个月进行一次开放，具体以公告为准。');assert.equal(p.updates.length,1);applySupplement(o,p);assert.equal(after(o,'开放期'),'每满18个月进行一次开放，具体以公告为准。');
});

test('fixed-monthday full prose supplementation produces applied content rather than a zero-update preview',()=>{
 const old='本计划每月10日、25日开放办理参与和退出业务。',t=task([['开放期',old]],['本计划每月1日、15日开放办理参与和退出业务。']),g=group(t,'opening_fixed_monthdays');
 const input='本计划每月2日、16日开放办理参与和退出业务。',p=previewSupplement(t,g.id,input);assert.ok(p.updates.length);applySupplement(t,p);assert.equal(after(t,'开放期'),input);
});

test('code and benchmark cards accept their own label and never pick one number from a range or negative value',()=>{
 const t=task([['产品代码','000999'],['业绩比较基准','2.5%']]),code=group(t,'product_code'),benchmark=group(t,'benchmark');
 applySupplement(t,previewSupplement(t,code.id,'产品代码：000123'));assert.equal(after(t,'产品代码'),'000123');
 for(const text of ['2%至3%','-2.25%','−2.25%','２％至３％','2.25%，另一个为3%'])assert.throws(()=>previewSupplement(t,benchmark.id,text),/单个|负数|基准|百分比/);
 applySupplement(t,previewSupplement(t,benchmark.id,'业绩比较基准：２.２５％'));assert.equal(after(t,'业绩比较基准'),'2.25%');
});

test('a single numeric supplement does not flatten separate holding-day conditions',()=>{
 const t=task([['锁定期','持有期为180天，特殊条件持有期为360天']]),g=group(t,'lockup_days');assert.throws(()=>previewSupplement(t,g.id,'90天'),/多个|条件|不同/);
 const d=task([['存续期限','存续期为5年，可展期5年。']]);assert.throws(()=>previewSupplement(d,group(d,'duration_years').id,'10年'),/多个|条件|不同/);
});

test('whole text line wrapping is normalized before preview and application',()=>{
 const t=task([['投资范围','投资于债券和基金。']]),g=group(t,'investment_scope'),p=previewSupplement(t,g.id,'投资于债券、\n股票及\t基金。');
 assert.ok(p.notes.some(n=>/换行|空格/.test(n)));assert.equal(p.changes[0].after,'投资于债券、 股票及 基金。');applySupplement(t,p);assert.equal(after(t,'投资范围'),p.changes[0].after);
 assert.ok(t.materials[0].operations.filter(o=>o.kind==='replace').every(o=>!/[\r\n\t]/.test(o.after)));
});

test('old tier previews fail after a newer supplement and retain newer per-occurrence values',()=>{
 const t=task([['退出费率',original],['管理费','0.6%']]),g=group(t,'redemption_fee');
 const old=previewSupplement(t,g.id,original.replace('1%','0.8%').replace('0%','0.2%'));
 applySupplement(t,previewSupplement(t,g.id,original.replace('1%','0.7%').replace('0%','0.1%')));
 assert.throws(()=>applySupplement(t,old),/变化|重新预览/);
 applySupplement(t,previewSupplement(t,group(t,'management_fee').id,'0.4%'));assert.match(after(t,'退出费率'),/0.7%.*0.1%/);
});

test('card changes retain current replacements of unrelated parameters in the same paragraph',()=>{
 const t=task([['产品费率','管理费0.6%；托管费0.03%']]);applySupplement(t,previewSupplement(t,group(t,'management_fee').id,'0.4%'));
 const p=previewSupplement(t,group(t,'custody_fee').id,'0.02%');assert.equal(p.changes[0].before,'管理费0.4%；托管费0.03%');assert.equal(p.changes[0].after,'管理费0.4%；托管费0.02%');applySupplement(t,p);assert.equal(after(t,'产品费率'),p.changes[0].after);
});

test('unknown performance percent remains an unresolved number rather than accepting an entire fee sentence',()=>{
 const t=task([['业绩报酬','本计划的业绩报酬为2.5%。']]),g=group(t,'custom:performance_parameter'),p=previewSupplement(t,g.id,'超过业绩报酬计提基准2.50%的部分提取60%作为管理人业绩报酬');
 assert.equal(p.updates.length,0);assert.equal(after(t,'业绩报酬'),'本计划的业绩报酬为2.5%。');
});

test('insufficient and full holding-period wording is recognized without discarding its tier conditions',()=>{
 const before='持有期不足180天，退出费率1%；持有期满180天，退出费率0%。',t=task([['退出费率',before]]);
 const input=before.replace('1%','0.8%').replace('0%','0.2%');applySupplement(t,previewSupplement(t,group(t,'redemption_fee').id,input));assert.equal(after(t,'退出费率'),input);
});

test('a whole-key fact narrowed by automatic local updates accepts matching prose without inserting prose into its digit',()=>{
 for(const [label,key,before,contract,input,expected]of [
  ['开放期','opening','每满12个月进行一次开放，具体以公告为准。','开放期：每满18个月进行一次开放，具体以公告为准。','每满24个月进行一次开放，具体以公告为准。','每满24个月进行一次开放，具体以公告为准。'],
  ['锁定期','lockup','锁定期为12个月','锁定期：锁定期为18个月','锁定期为24个月','锁定期为24个月'],
  ['投资比例','investment_ratios','固定收益资产占资产总值20%。','投资比例：固定收益资产占资产总值30%。','固定收益资产占资产总值40%。','固定收益资产占资产总值40%。']
 ]){const t=task([[label,before]],[contract]),g=group(t,key);assert.ok(t._records.some(r=>r.f.id===g.factIds[0]&&/^\d+$/.test(r.atom.before)),label);
  const p=previewSupplement(t,g.id,input);assert.equal(p.changes[0].after,expected);applySupplement(t,p);assert.equal(after(t,label),expected);
  assert.throws(()=>previewSupplement(t,g.id,input+'另增加其他条件。'),/局部|对应|完整|条件/);
 }
});
