import test from 'node:test';
import assert from 'node:assert/strict';
import {currentMonth,deliveryName,resolveProductionMonth,monthInputValue} from '../src/naming.js';

test('month is captured using Beijing time at a month boundary',()=>{
 assert.equal(currentMonth(new Date('2026-09-30T16:01:00Z')),'202610');
 assert.equal(currentMonth(new Date('2026-09-30T15:59:00Z')),'202609');
});
test('replaces product number and filename month only; preserves shortened poster product',()=>{
 const result=deliveryName('【甲行 海报2609】示例稳健18个月5号(1).pptx',{targetName:'示例稳健（优选配置）18个月1号集合资产管理计划',templateProductName:'示例稳健（优选配置）18个月5号集合资产管理计划',kind:'poster',month:'202610'});
 assert.equal(result,'【甲行 海报2610】示例稳健18个月1号.pptx');
});
test('keeps bracket style of product, title prefix and exactly one element suffix',()=>{
 const result=deliveryName('【甲行新发 ppt202607】示例远帆组合低波3号产品要素(2).pptx',{targetName:'示例远帆组合（低波）5号集合资产管理计划',templateProductName:'示例远帆组合（低波）3号集合资产管理计划',kind:'ppt',month:'202610'});
 assert.equal(result,'【甲行新发 ppt202610】示例远帆组合低波5号产品要素.pptx');
});
test('adds missing month without overwriting target name already in filename',()=>{
 assert.equal(deliveryName('【甲行一页通】示例双债1号.docx',{targetName:'示例双债1号集合资金信托计划',templateProductName:'示例多元1号集合资金信托计划',kind:'word',month:'202610'}),'【甲行一页通 202610】示例双债1号.docx');
});
test('does not mistake 18-month term for a publication date',()=>{
 assert.equal(deliveryName('示例18个月7号.pptx',{targetName:'示例18个月7号集合资产管理计划',templateProductName:'示例18个月7号',kind:'ppt',month:'202610'}),'【202610】示例18个月7号产品要素.pptx');
});
test('full filename date is not silently transformed to a month',()=>{
 assert.equal(deliveryName('【一页通20260630】示例1号.docx',{targetName:'示例1号集合资产管理计划',templateProductName:'示例1号',kind:'word',month:'202610'}),'【一页通20260630 202610】示例1号.docx');
});
test('rejects invalid month',()=>assert.throws(()=>deliveryName('a.docx',{targetName:'示例1号',kind:'word',month:'202613'})));

test('manual production month survives a later task start and names both date conventions',()=>{
 const month=resolveProductionMonth({manualValue:'2026-08',now:new Date('2026-10-25T00:00:00Z')});
 assert.equal(month,'202608');
 assert.equal(deliveryName('【海报2609】示例1号.pptx',{targetName:'示例1号',kind:'poster',month}),'【海报2608】示例1号.pptx');
 assert.equal(deliveryName('【一页通202609】示例1号.docx',{targetName:'示例1号',kind:'word',month}),'【一页通202608】示例1号.docx');
 assert.equal(monthInputValue(month),'2026-08');
});
test('automatic production month refreshes at Beijing month boundary; empty and impossible manual values fail',()=>{
 assert.equal(resolveProductionMonth({now:new Date('2026-10-31T16:00:00Z')}),'202611');
 for(const value of ['', '2026-13','2026-00','2026-1','abcd','2026-10-25'])assert.throws(()=>resolveProductionMonth({manualValue:value}));
});

test('a filename may omit the confirmed product issuer without blocking a Word export',()=>{
 const name='国泰海通君享湘安21号集合资产管理计划';
 assert.equal(deliveryName('【长沙银行一页通202608】君享湘安21号集合资产管理计划.docx',{targetName:name,templateProductName:name,kind:'word',month:'202610'}),'【长沙银行一页通202610】君享湘安21号集合资产管理计划.docx');
});

test('an issuer-free filename follows the contract number and preserves its legal suffix',()=>{
 assert.equal(deliveryName('【长沙银行一页通202608】君享湘安21号集合资产管理计划.docx',{targetName:'国泰海通君享湘安22号集合资产管理计划',templateProductName:'国泰海通君享湘安21号集合资产管理计划',kind:'word',month:'202610'}),'【长沙银行一页通202610】君享湘安22号集合资产管理计划.docx');
});

test('a two-character series abbreviation must also match the confirmed full product name',()=>{
 assert.equal(deliveryName('【一页通202608】鑫源25号.docx',{targetName:'示例机构鑫源26号集合资产管理计划',templateProductName:'示例机构鑫源25号集合资产管理计划',kind:'word',month:'202610'}),'【一页通202610】鑫源26号.docx');
});

test('a changed issuer preserves a proven series abbreviation and PPT suffix',()=>{
 assert.equal(deliveryName('【甲行 PPT202609】远航稳健21号产品要素.pptx',{targetName:'新品牌远航稳健22号集合资产管理计划',templateProductName:'旧品牌远航稳健21号集合资产管理计划',kind:'ppt',month:'202610'}),'【甲行 PPT202610】远航稳健22号产品要素.pptx');
});

test('an issuer name extension cannot leak into an unchanged product series abbreviation',()=>{
 assert.equal(deliveryName('【一页通202608】君享湘安21号.docx',{targetName:'国泰海通资管君享湘安22号集合资产管理计划',templateProductName:'国泰海通君享湘安21号集合资产管理计划',kind:'word',month:'202610'}),'【一页通202610】君享湘安22号.docx');
});

test('a known omitted prefix applies to a new series, while unprovable abbreviations use the full target',()=>{
 const input='【一页通2609】远航稳健21号.docx';
 const args={templateProductName:'示例机构远航稳健21号集合资产管理计划',kind:'word',month:'202610'};
 assert.equal(deliveryName(input,{...args,targetName:'示例机构星河多元3号集合资产管理计划'}),'【一页通2610】星河多元3号.docx');
 assert.equal(deliveryName(input,{...args,targetName:'不同机构星河多元3号集合资产管理计划'}),'【一页通2610】不同机构星河多元3号.docx');
});

test('a proven abbreviation also retains omitted parenthetical qualifiers and distinct month terms',()=>{
 assert.equal(deliveryName('【甲行 海报2609】远航稳健18个月5号.pptx',{targetName:'示例机构远航稳健（优选配置）18个月1号集合资产管理计划',templateProductName:'示例机构远航稳健（优选配置）18个月5号集合资产管理计划',kind:'poster',month:'202610'}),'【甲行 海报2610】远航稳健18个月1号.pptx');
});

test('a series abbreviation cannot be inferred from a bare number, another series or marketing text',()=>{
 const args={targetName:'示例机构远航稳健22号集合资产管理计划',templateProductName:'示例机构远航稳健21号集合资产管理计划',kind:'word',month:'202610'};
 for(const stem of ['21号','星河多元21号','代表产品21号','历史业绩 远航稳健21号'])assert.throws(()=>deliveryName('【一页通202608】'+stem+'.docx',args),/无法对应/);
});
