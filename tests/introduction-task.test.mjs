import test from 'node:test';
import assert from 'node:assert/strict';
import {createIntroductionTask,finalizeIntroductionTask} from '../src/introduction-task.js';
import {deliveryName,materialDeliveryName} from '../src/naming.js';
const template=(id,kind='docx',name=`【一页通202609】${id}.docx`)=>({id,name,kind,paragraphs:[],slides:kind==='pptx'?[1,2,3,4].map(page=>({page})):[],warnings:[]});
const field=(source,id)=>({id,source,paragraphIds:[id],factIds:[id],status:'verified'});
test('introduction task has no contract or fabricated shared product identity',()=>{
 const task=createIntroductionTask([template('甲产品'),template('乙产品')],{month:'202610',productionDate:'2026-10-29'});
 assert.equal(task.updateMode,'introductions');assert.equal(task.contract,undefined);assert.equal(task.target.name,'公司和经理介绍更新');assert.equal(task.facts.length,0);
 assert.equal(materialDeliveryName(task.materials[0],{month:task.month}),'【一页通202610】甲产品.docx');
 assert.equal(materialDeliveryName(task.materials[1],{month:task.month}),'【一页通202610】乙产品.docx');
});
test('introduction-only keeps recognized unchanged pages and skips missing sections',()=>{
 const task=createIntroductionTask([template('a','pptx','【PPT202609】甲产品产品要素.pptx'),template('empty')],{month:'202610'}),m=task.materials[0];
 m.fields.push(field('manager','bio'));m.profilePages={company:[4],manager:[2,4]};
 finalizeIntroductionTask(task);assert.deepEqual(m.selectedPages,[2,4]);assert.equal(m.operations.length,0);assert.equal(task.materials.length,1);assert.equal(task.inputIssues[0].reason,'no_introductions');
 assert.equal(materialDeliveryName(m,{month:'202610'}),'【PPT202610】甲产品公司及经理介绍.pptx');
});
test('full mode keeps original product fields and page subset, adds only profile pages',()=>{
 const m={kind:'ppt',template:template('a','pptx','【PPT202609】甲产品.pptx'),selectedPages:[2,3],fields:[field('contract','risk')],operations:[{id:'risk',source:'contract'}],profilePages:{company:[1],manager:[4]}};
 const task={facts:[],materials:[m],stats:{}};const fields=JSON.stringify(m.fields),ops=JSON.stringify(m.operations);
 finalizeIntroductionTask(task);assert.deepEqual(m.selectedPages,[1,2,3,4]);assert.deepEqual(m.productSelectedPages,[2,3]);assert.equal(JSON.stringify(m.fields),fields);assert.equal(JSON.stringify(m.operations),ops);
 m.profilePages.manager=[];finalizeIntroductionTask(task);assert.deepEqual(m.selectedPages,[1,2,3]);
 assert.equal(materialDeliveryName(m,{targetName:'乙产品',month:'202610'}),'【PPT202610】乙产品产品要素及介绍.pptx');
});
test('old naming stays byte-identical if no introduction pages are added',()=>{
 const m={kind:'ppt',template:template('a','pptx','【PPT202609】甲产品.pptx'),templateProductName:'甲产品'};
 const options={targetName:'乙产品',month:'202610'};assert.equal(materialDeliveryName(m,options),deliveryName(m.template.name,{...options,templateProductName:'甲产品',kind:'ppt'}));
});
test('introduction-only rejects accidental product edits instead of disguising them',()=>{
 const task=createIntroductionTask([template('a')],{month:'202610'});task.materials[0].fields.push(field('manager','bio'));task.materials[0].operations.push({source:'contract'});
 assert.throws(()=>finalizeIntroductionTask(task),/介绍/);
});
test('one-page manager PPT is a PPT, while an explicitly named poster is skipped',()=>{
 const t=template('a','pptx','经理介绍.pptx');t.slides=[{page:1}];const task=createIntroductionTask([t,template('p','pptx','海报.pptx')],{month:'202610'});
 assert.equal(task.materials[0].kind,'ppt');assert.equal(task.materials.length,1);assert.equal(task.inputIssues[0].reason,'poster');
});
