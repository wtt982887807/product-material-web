import test from 'node:test';
import assert from 'node:assert/strict';
import {simplifyEqualFeeTiers} from '../src/equal-fee-tiers.js';
import * as engine from '../src/review-engine.js';

test('different amount subjects cannot supply each other missing interval',()=>{
 const original='认购金额＜300万元，参与费率为2%；参与金额≥300万元，参与费率为2%。';
 assert.equal(simplifyEqualFeeTiers(original),original);
});
test('holding units do not silently convert to complete a fee interval',()=>{
 const original='持有天数＜180自然日，退出费率为1%；持有期限≥180交易日，退出费率为1%。';
 assert.equal(simplifyEqualFeeTiers(original),original);
});
test('distinct participation and exit labels retain their individual fee values',()=>{
 const original='参与金额＜300万元，参与费率为2%；参与金额≥300万元，参与费率为2%；持有天数＜180自然日，退出费率为1%；持有天数≥180自然日，退出费率为0%。';
 assert.equal(simplifyEqualFeeTiers(original),'参与费率为2%；持有天数＜180自然日，退出费率为1%；持有天数≥180自然日，退出费率为0%。');
});

test('numeric preview on a compacted mixed fee row preserves other current contract values and matches apply',()=>{
 const label='认购/参与/退出费率',text='认购费率0.5%；参与金额＜300万元，参与费率为1.5%；参与金额≥300万元，参与费率为1.5%；退出费率0.2%。';
 const paragraphs=[],part='word/document.xml';
 const rows=[['产品名称','晨星稳健9号集合资产管理计划'],[label,text]].map(([label,value],i)=>{
  const ids=['l'+i,'v'+i];[label,value].forEach((text,j)=>paragraphs.push({id:ids[j],text,part,index:i*2+j,page:1,tableId:'t',rowIndex:i,cellIndex:j}));
  return{id:'r'+i,label,value,part,page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });
 const t=engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'产品名称：晨星稳健9号集合资产管理计划'},{id:'f',text:'认购费率2%；参与费率3%；退出费率0%。'}]},[{id:'a',name:'一页通.docx',kind:'docx',paragraphs,rows,slides:[],warnings:[]}],{month:'202610'});
 const group=engine.getSupplementGroups(t).find(g=>g.scope.feeActions==='认购');
 const before='认购费率2.0%；参与费率为3%；退出费率0.0%。',after='认购费率0.7%；参与费率为3%；退出费率0.0%。';
 assert.equal(t.materials[0].fields[1].after,before);
 const preview=engine.previewCardSupplement(t,group.id,'0.7%');
 assert.equal(preview.changes[0].before,before);assert.equal(preview.changes[0].after,after);
 assert.equal(t.materials[0].fields[1].after,before,'preview leaves all current facts and wording unchanged');
 engine.applySupplement(t,preview);assert.equal(t.materials[0].fields[1].after,after);
});
