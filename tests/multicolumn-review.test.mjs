import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials} from '../src/review-engine.js';

const contract=lines=>({name:'合同.docx',blocks:lines.map((text,i)=>({id:`b${i}`,text})),warnings:[]});
function gridTemplate(pages,{kind='pptx'}={}){
 const t={id:'grid',name:`双栏要素.${kind}`,kind,paragraphs:[],rows:[],slides:[],warnings:[]};
 for(const [page,title,rows]of pages){const part=kind==='docx'?'word/document.xml':`ppt/slides/slide${page}.xml`,tableId=`${part}#table-${page}`;
  for(const [rowIndex,values]of rows.entries()){
   const cells=values.map((text,cellIndex)=>{const index=t.paragraphs.filter(p=>p.part===part).length,id=`${part}#p-${index}`;t.paragraphs.push({id,part,page,index,text,tableId,rowIndex,cellIndex});return{text,paragraphIds:[id]};});
   t.rows.push({id:`${tableId}#row-${rowIndex}`,part,page,tableId,rowIndex,label:cells[0].text,value:cells.slice(1).map(c=>c.text).join('\n'),paragraphIds:cells.flatMap(c=>c.paragraphIds),cells});
  }
  t.slides.push({page,title,text:title,paragraphIds:t.paragraphs.filter(p=>p.page===page).map(p=>p.id)});
 }
 return t;
}
const ownRows=[
 ['产品名称','星河全景X期集合资金信托计划','','','','',''],
 ['受托人','旧受托信托有限公司','','','代销机构','','旧代销银行股份有限公司'],
 ['产品期限','3年，可提前结束或延期','','','保管行','','旧保管银行股份有限公司'],
 ['代销机构风险评级','R5（高风险）','','','产品认购起点','','40万元'],
 ['开放日','成立日所在周不开放；每周五为开放日。','','','赎回锁定期','','第6个赎回开放日前不得赎回。'],
 ['费用信息','认/申购费','0.5%，内扣法','赎回费','','无',''],
 ['','固定管理费','1.55%/年','保管费','','0.06%/年',''],
 ['','浮动管理费','计提10%，提取后年化收益率不低于5%。','','','','']
];
test('a merged multicolumn trust element table is not mistaken for performance data',()=>{
 const t=gridTemplate([[2,'产品要素',ownRows],[3,'底层产品要素',[['产品名称','底层精选1号集合资产管理计划'],['管理费','0.9%']]]]);
 const task=analyzeMaterials(contract(['产品名称：星河全景1期集合资金信托计划','受托人：新受托信托有限公司','代销机构：新代销银行股份有限公司','保管行：新保管银行股份有限公司','管理人风险等级为R2。']),[t]);
 const m=task.materials[0];assert.deepEqual(m.selectedPages,[2]);
 const fields=new Map(m.fields.map(f=>[f.label,f]));
 for(const label of ['受托人','代销机构','保管行','产品认购起点','赎回锁定期','认/申购费','赎回费','固定管理费','保管费','浮动管理费'])assert.ok(fields.has(label),label);
 assert.equal(fields.get('受托人').before,'旧受托信托有限公司');assert.equal(fields.get('代销机构').before,'旧代销银行股份有限公司');
 assert.equal(fields.get('受托人').after,'新受托信托有限公司');assert.equal(fields.get('代销机构').after,'新代销银行股份有限公司');assert.equal(fields.get('保管行').after,'新保管银行股份有限公司');
 assert.equal(fields.get('代销机构风险评级').after,'R5（高风险）');assert.equal(fields.get('代销机构风险评级').status,'pending');
 assert.ok(m.operations.every(o=>o.part==='ppt/slides/slide2.xml'));
 const valueIds=new Set(m.fields.flatMap(f=>f.paragraphIds));
 assert.ok(m.operations.every(o=>valueIds.has(o.paragraphId)),'field labels cannot be edited as their neighbor value');
});
test('performance columns stay excluded even when the first physical paragraph is a disclaimer',()=>{
 const historical=[['产品名称','成立时间','成立以来表现','近一月收益率'],['旧产品1号集合资金信托计划','2023-01-01','40.5%','1.1%']];
 const t=gridTemplate([[1,'本材料仅供合格投资者使用',historical],[2,'产品要素',ownRows]]);
 const task=analyzeMaterials(contract(['产品名称：星河全景1期集合资金信托计划']),[t]);
 assert.deepEqual(task.materials[0].selectedPages,[2]);assert.ok(task.materials[0].operations.every(o=>o.part!=='ppt/slides/slide1.xml'));
});
test('Word double label-value rows retain separate field scopes and source paragraphs',()=>{
 const t=gridTemplate([[1,'产品要素',[['产品名称','晨星2号集合资产管理计划','',''],['管理费','0.5%','托管费','0.2%'],['开放日','每周五开放，成立当周不开放','产品期限','3年']]]],{kind:'docx'});
 const task=analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划','管理费率0.3%。','托管费率0.1%。']),[t]);
 const fields=new Map(task.materials[0].fields.map(f=>[f.label,f]));
 assert.equal(fields.get('管理费').after,'0.3%');assert.equal(fields.get('托管费').after,'0.1%');
 assert.ok(!task.materials[0].operations.some(o=>o.before==='托管费'));
});
test('separate trust fee labels retain total-manager meaning and percentage punctuation',()=>{
 const t=gridTemplate([[2,'产品要素',[['产品名称','星河1期集合资金信托计划','',''],['固定管理费','【1.55】%/年','保管费','【0.06】%/年']]]]);
 const task=analyzeMaterials(contract(['产品名称：星河2期集合资金信托计划','固定信托管理费年费率0.6%。','受托人信托报酬年费率0.09%。','保管费年费率0.01%。']),[t]);
 const fields=new Map(task.materials[0].fields.map(f=>[f.label,f]));
 assert.equal(fields.get('固定管理费').after,'【0.60】%/年');assert.equal(fields.get('保管费').after,'【0.01】%/年');
});
test('a two-layer aggregate fee never borrows a single-layer contract rate or shared fact',()=>{
 const t=gridTemplate([[2,'产品要素',[['产品名称','星河1期集合资金信托计划','',''],['保管费','【0.06】%/年（两层合计）','固定管理费','【1.55】%/年（两层合计）']]]]);
 const single=gridTemplate([[2,'产品要素',[['产品名称','星河1期集合资金信托计划','',''],['保管费','0.06%/年','固定管理费','1.55%/年']]]]);single.id='single';
 const task=analyzeMaterials(contract(['产品名称：星河2期集合资金信托计划','保管费年费率0.01%。','固定信托管理费年费率0.6%。']),[t,single]);
 const fields=new Map(task.materials[0].fields.map(f=>[f.label,f]));
 for(const label of ['保管费','固定管理费']){assert.equal(fields.get(label).after,fields.get(label).before);assert.equal(fields.get(label).status,'pending');}
 assert.ok(task.materials[0].operations.some(o=>o.kind==='highlight'&&o.before.includes('0.06')));
 assert.equal(task.materials[1].fields.find(f=>f.label==='保管费').after,'0.01%/年');
 assert.equal(task.facts.filter(f=>f.key==='custody_fee').length,2);
});
test('materials without any product elements are explicitly marked unlocated',()=>{
 const t=gridTemplate([[1,'公司介绍',[['公司介绍','管理规模100亿元'],['团队简介','专业投资团队']]]],{kind:'docx'});
 const task=analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划']),[t]);
 assert.equal(task.materials[0].productElementsLocated,false);
 assert.ok(task.materials[0].warnings.some(w=>/未定位|无法定位|尚未识别/.test(w)));
 const recognized=gridTemplate([[1,'产品要素',[['产品名称','晨星2号集合资产管理计划'],['管理费','0.5%']]]],{kind:'docx'});
 assert.equal(analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划']),[recognized]).materials[0].productElementsLocated,true);
});
test('risk notes about an underlying product and exposure formulas are not own product facts',()=>{
 const t=gridTemplate([[2,'产品要素',[['产品名称','星河1期集合资金信托计划','',''],['代销机构风险评级','R5（高风险）','产品类型','资产管理信托，混合类']]]]);
 const task=analyzeMaterials(contract(['产品名称：星河2期集合资金信托计划','管理人\t示例资产管理有限公司\n风险等级\t管理人将本集合计划的风险评级评价为中高风险R4。\n注：本集合计划为信托投资的标的产品，标的产品风险等级可能与本信托计划风险等级、代销机构风险评级存在差异。','权益类风险敞口=|（权益类现货多头市值-权益类现货空头市值）/产品净资产|。']),[t]);
 for(const f of task.materials[0].fields.filter(f=>['代销机构风险评级','产品类型'].includes(f.label))){assert.equal(f.after,f.before);assert.equal(f.status,'pending');}
 const explicit=analyzeMaterials(contract(['产品名称：星河2期集合资金信托计划','代销机构风险评级为R4（中高风险）。','产品类型：权益类。']),[t]);
 assert.equal(explicit.materials[0].fields.find(f=>f.label==='代销机构风险评级').after,'R4（中高风险）');
 assert.equal(explicit.materials[0].fields.find(f=>f.label==='产品类型').after,'资产管理信托，权益类');
});
test('a value that resembles a label cannot merge neighboring labels into a product type',()=>{
 const t=gridTemplate([[1,'产品要素',[['产品名称','晨星2号集合资产管理计划'],['产品类型','开放式','托管人','旧银行股份有限公司']]]],{kind:'docx'});
 const task=analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划','产品类型：固定收益类。','托管人：新银行股份有限公司']),[t]);
 const fields=new Map(task.materials[0].fields.map(f=>[f.label,f]));
 assert.equal(fields.get('产品类型').before,'开放式');assert.equal(fields.get('产品类型').after,'开放式');assert.equal(fields.get('产品类型').status,'pending');
 assert.equal(fields.get('托管人').after,'新银行股份有限公司');assert.ok(task.materials[0].operations.every(o=>o.before!=='托管人'));
});
test('ambiguous multicolumn rows retain each cell instead of applying the first label to all cells',()=>{
 const t=gridTemplate([[1,'产品要素',[['产品名称','晨星2号集合资产管理计划'],['托管人','旧银行股份有限公司','缴费说明','待定','备注']]]],{kind:'docx'});
 const task=analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划','托管人：新银行股份有限公司']),[t]);
 const field=task.materials[0].fields.find(f=>f.label==='托管人');assert.equal(field.after,field.before);assert.equal(field.status,'pending');
 assert.ok(!task.materials[0].operations.some(o=>o.kind==='replace'&&field.paragraphIds.includes(o.paragraphId)));
 assert.ok(task.facts.some(f=>/多栏/.test(f.reason)));
});
test('only an existing asset-category token changes inside a compound product type',()=>{
 const t=gridTemplate([[1,'产品要素',[['产品名称','晨星2号集合资产管理计划'],['产品类型','开放式，混合类']]]],{kind:'docx'});
 const task=analyzeMaterials(contract(['产品名称：晨星3号集合资产管理计划','产品类型：固定收益类。']),[t]);
 assert.equal(task.materials[0].fields.find(f=>f.label==='产品类型').after,'开放式，固定收益类');
});
