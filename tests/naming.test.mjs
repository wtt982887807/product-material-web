import test from 'node:test';
import assert from 'node:assert/strict';
import {currentMonth,deliveryName} from '../src/naming.js';

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
