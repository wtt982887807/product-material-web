import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser} from '@xmldom/xmldom';
import {analyzeMaterials} from '../src/review-engine.js';

const TARGET='青澜福星君享（信用优享）18个月5号集合资产管理计划';
const TABLE_NAME='青澜福星君享（信用优享）18个月1号集合资产管理计划';
const c={name:'合同.docx',warnings:[],blocks:[{id:'name',text:'产品名称：'+TARGET},{id:'fee',text:'管理费年费率为0.3%。'}]};
function template(cover,tableName=TABLE_NAME,tail=[]){
 const part='word/document.xml',paragraphs=cover.map((text,index)=>({id:'cover-'+index,text,part,index,tableId:null}));
 const rows=[['产品名称',tableName],['管理费','0.3%']].map(([label,value],rowIndex)=>{
  const ids=['table-'+rowIndex+'-label','table-'+rowIndex+'-value'];
  for(const [cellIndex,text]of[label,value].entries())paragraphs.push({id:ids[cellIndex],text,part,index:paragraphs.length,tableId:'elements',rowIndex,cellIndex});
  return{id:'row-'+rowIndex,label,value,part,tableId:'elements',rowIndex,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};
 });
 for(const [i,text]of tail.entries())paragraphs.push({id:'tail-'+i,text,part,index:paragraphs.length,tableId:null});
 return{id:'word',name:'【一页通】青澜福星君享.docx',kind:'docx',warnings:[],paragraphs,rows,slides:[]};
}
function after(task,id){let text=task.templates[0].paragraphs.find(p=>p.id===id).text;for(const o of task.materials[0].operations.filter(o=>o.paragraphId===id&&o.kind==='replace').sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}

test('Word cover identity follows contract even when the element name omits a series alias',()=>{
 const t=template(['青澜福星君享年华（信用优享）12个月2号集合资产管理计划']);
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-0'),TARGET);
 assert.equal(after(task,'table-0-value'),TARGET);
});

test('Word cover identity does not inherit malformed repeated wording from the element table',()=>{
 const t=template(['青澜福星君享年华12个月定开债2号集合资产管理计划'],'青澜福星君享君享年华18个月定开债1号集合资产管理计划');
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-0'),TARGET);
 assert.equal(after(task,'table-0-value'),TARGET);
});

test('Word cover identity updates when the element table already has the target name',()=>{
 const t=template(['青澜福星君享（信用优享）12个月2号集合资产管理计划'],TARGET);
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-0'),TARGET);
 assert.equal(after(task,'table-0-value'),TARGET);
});

test('Word cover accepts leading empty and branding paragraphs',()=>{
 const t=template(['','青澜资产管理','青澜福星君享年华（信用优享）12个月2号集合资产管理计划']);
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-2'),TARGET);
 assert.equal(after(task,'cover-1'),'青澜资产管理');
});

test('Word cover split across short name and legal suffix keeps both paragraphs',()=>{
 const t=template(['青澜福星君享（信用优享）12个月2号','集合资产管理计划']);
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-0'),'青澜福星君享（信用优享）18个月5号');
 assert.equal(after(task,'cover-1'),'集合资产管理计划');
 assert.ok(task.materials[0].operations.filter(o=>o.kind==='replace').every(o=>!/[\n\r]/.test(o.after)),'replacement stays inside each existing paragraph');
});

test('independent Word cover recognition does not rewrite later historical or underlying names',()=>{
 const cover='青澜福星君享年华（信用优享）12个月2号集合资产管理计划';
 const t=template([cover],TABLE_NAME,['历史代表产品',cover,'底层投资标的',cover,'历史业绩：'+cover+'成立以来收益率3.0%。']);
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-0'),TARGET);
 for(let i=0;i<5;i++)assert.equal(after(task,'tail-'+i),t.paragraphs.find(p=>p.id==='tail-'+i).text);
 assert.ok(!task.materials[0].operations.some(o=>o.paragraphId.startsWith('tail-')));
});

for(const [boundary,firstXml,nextXml]of[
 ['explicit page break','<w:p><w:r><w:t>青澜资产管理</w:t><w:br w:type="page"/></w:r></w:p>',null],
 ['page break before',null,'<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>青澜历史策略2号集合资产管理计划</w:t></w:r></w:p>'],
 ['section break','<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>青澜资产管理</w:t></w:r></w:p>',null],
])test(`Word cover recognition stops before a later page at ${boundary}`,()=>{
 const later='青澜旧策略2号集合资产管理计划',t=template(['青澜资产管理',later]);
 const parse=xml=>new DOMParser().parseFromString(xml.replace('<w:p>','<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'),'text/xml').documentElement;
 if(firstXml)t.paragraphs[0].node=parse(firstXml);
 if(nextXml)t.paragraphs[1].node=parse(nextXml.replace('青澜历史策略2号集合资产管理计划',later));
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-1'),later,'a full legal name after an explicit page boundary is not the main title');
 assert.ok(!task.materials[0].operations.some(o=>o.paragraphId==='cover-1'));
 assert.equal(after(task,'table-0-value'),TARGET,'the actual element name still updates');
});

test('Word cover recognition retains empty paragraphs when checking page boundaries',()=>{
 const later='青澜旧策略2号集合资产管理计划',t=template(['',later]);
 t.paragraphs[0].node=new DOMParser().parseFromString('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:br w:type="page"/></w:r></w:p>','text/xml').documentElement;
 const task=analyzeMaterials(c,[t]);
 assert.equal(after(task,'cover-1'),later);
 assert.ok(!task.materials[0].operations.some(o=>o.paragraphId==='cover-1'));
});
