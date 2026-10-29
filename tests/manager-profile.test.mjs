import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture,p,table,profile,task,normalize} from './manager-fixture.mjs';
import {applyManagerProfile} from '../src/manager-profile.js';
import {analyzeMaterials,setSupplement} from '../src/review-engine.js';

test('manager authority updates the full biography and excludes other sections and PPT',async()=>{
 const t=await fixture(p('投资经理简介')+p('张三（照片）：')+p('9年证券从业经验。金融硕士。')+p('现任旧公司投资经理。')+p('旧投资理念。')+p('公司实力雄厚')+p('注册资本99亿元。'));
 const x=task(t);x.materials.push({...x.materials[0],materialId:'ppt',kind:'ppt',operations:[],fields:[]});
 const biography='10年证券从业经验。金融博士。\n现任新公司投资经理，历任研究员。\n擅长量化投资。';
 applyManagerProfile(x,profile([{biography}]),{productionDate:'2026-10-29'});
 assert.equal(x.facts.length,1);assert.equal(x.facts[0].source,'manager');assert.equal(x.facts[0].value,biography);
 assert.equal(x.materials[0].fields[0].after,biography);assert.equal(x.materials[1].operations.length,0);
 assert.ok(x.materials[0].operations.every(o=>!['word/document.xml#p-1','word/document.xml#p-5','word/document.xml#p-6'].includes(o.paragraphId)));
 assert.match(x.facts[0].evidence[0].quote,/经理简介.*第2行/);assert.equal(x.warnings.length,0);
});

test('same-name source rows select the closest biography deterministically and remain idempotent',async()=>{
 const t=await fixture(p('基金经理')+p('张三')+p('复旦大学金融学硕士，9年证券从业经验。')+p('现任甲公司基金经理。')+p('产品要素'));
 const x=task(t),data=profile([{biography:'北京大学医学博士，3年从业经验。\n现任乙公司。'},{biography:'复旦大学金融学硕士，10年证券从业经验。\n现任甲公司基金经理。'}]);
 applyManagerProfile(x,data);const first=JSON.stringify(x.materials[0].operations);applyManagerProfile(x,data);
 assert.equal(JSON.stringify(x.materials[0].operations),first);assert.equal(x.facts.length,1);assert.equal(x.facts[0].scope.row,3);assert.equal(x.warnings.length,0);
});

test('manager regions span table cells and text boxes but stop at another manager or section',async()=>{
 const t=await fixture(table(p('张三'),p('旧学历。')+p('现任甲经理。'))+p('基金经理简介')+p('',`<w:r><w:txbxContent>${p('李四：')}${p('旧学历。')}${p('旧投资理念。')}</w:txbxContent></w:r>`)+p('风险揭示')+p('张三负责本产品的历史业绩统计。'));
 const x=applyManagerProfile(task(t),profile([{name:'张三',biography:'金融硕士。\n现任乙经理。'},{name:'李四',biography:'经济学博士。\n投资风格稳健。'}]));
 assert.deepEqual(x.facts.map(f=>f.scope.manager),['张三','李四']);assert.equal(x.facts.length,2);
 assert.equal(x.materials[0].operations.some(o=>o.before.includes('历史业绩')),false);
});

test('date reapplication retains product and company operations and manual supplements',async()=>{
 const t=await fixture(p('投资经理')+p('张三')+p('9年证券从业经验。')+p('公司实力雄厚'));
 const x=task(t),operation={id:'product',kind:'replace',source:'user',paragraphId:'other',start:0,before:'1',after:'2'},fact={id:'user-fact',source:'user',status:'verified',value:'manual'};
 x.facts.push(fact);x.materials[0].operations.push(operation);x.materials[0].fields.push({id:'company',source:'company'});
 const data=profile([{biography:'9年证券从业经验。',workStart:'2016-10-30'}]);
 applyManagerProfile(x,data,{productionDate:'2026-10-29'});assert.match(x.facts.find(f=>f.source==='manager').value,/9年/);
 applyManagerProfile(x,data,{productionDate:'2026-10-30'});assert.match(x.facts.find(f=>f.source==='manager').value,/10年/);
 assert.ok(x.facts.includes(fact));assert.ok(x.materials[0].operations.includes(operation));assert.equal(x.materials[0].fields.filter(f=>f.source==='company').length,1);
});

test('decorated team headings and manager names in product rows stay outside biography',async()=>{
 const team='-------千亿固收舰队 20年资产管理经验行业领跑-------';
 const t=await fixture(p('基金经理')+p('张三（配照片）：')+p('金融硕士。')+p('现任基金经理。')+p(team)+p('团队现有16人，经验丰富。')+table(p('投资经理'),p('张三')));
 const x=applyManagerProfile(task(t),profile([{biography:'金融博士。\n现任投资经理。'}]));
 assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>[2,3].includes(o.index)));
 assert.doesNotMatch(x.materials[0].fields[0].before,/配照片|舰队|团队/);
});

test('product supplement finalization preserves manager full preview and exact operations',async()=>{
 const productTable=table(p('产品名称'),p('示例稳健1号集合资产管理计划')).replace('</w:tbl>','<w:tr><w:tc>'+p('管理费')+'</w:tc><w:tc>'+p('0.8%')+'</w:tc></w:tr></w:tbl>');
 const t=await fixture(productTable+p('投资经理简介')+p('张三')+p('旧学历。')+p('现任甲公司经理。'));
 const contract={name:'合同.docx',blocks:[{id:'b0',text:'产品名称：示例稳健2号集合资产管理计划'}],warnings:[]};
 const x=applyManagerProfile(analyzeMaterials(contract,[t]),profile([{biography:'金融博士。\n现任乙公司经理。\n擅长量化投资。'}]));
 const field=x.materials[0].fields.find(f=>f.source==='manager'),ops=JSON.stringify(x.materials[0].operations.filter(o=>o.source==='manager')),preview=JSON.stringify(field),fee=x.facts.find(f=>f.key==='management_fee');
 assert.ok(fee);setSupplement(x,fee.id,'0.2%');assert.equal(JSON.stringify(field),preview);assert.equal(JSON.stringify(x.materials[0].operations.filter(o=>o.source==='manager')),ops);
});

test('a manager quotation in market commentary is not a biography anchor',async()=>{
 const t=await fixture(p('市场观点')+p('张三 表示，当前市场估值合理，建议保持稳健投资策略。')+p('风险揭示'));
 const x=applyManagerProfile(task(t),profile([{biography:'金融硕士。\n现任投资经理。'}]));
 assert.equal(x.facts.length,0);assert.equal(x.materials[0].operations.length,0);
});
