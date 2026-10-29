import test from 'node:test';
import assert from 'node:assert/strict';
import {analyzeMaterials} from '../src/review-engine.js';

function analyze(lines,{value='本计划为R2（中低风险）产品，适合风险承受能力为C3及以上的投资者。',kind='docx'}={}){
 const contract={name:'合同.docx',blocks:['产品名称：晨星稳健9号集合资产管理计划',...lines].map((text,i)=>({id:'b'+i,text})),warnings:[]};
 const rows=[['产品名称','晨星稳健2号集合资产管理计划'],['风险等级',value],['代销机构风险评级','R5（高风险）']],paragraphs=[];
 const part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml';
 const rr=rows.map(([label,value],i)=>{const ids=[`p${i}l`,`p${i}v`];paragraphs.push(...[label,value].map((text,j)=>({id:ids[j],text,part,index:i*2+j,page:1,tableId:'t1',rowIndex:i,cellIndex:j})));return{id:'r'+i,part,page:1,tableId:'t1',rowIndex:i,label,value,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]};});
 return analyzeMaterials(contract,[{id:'risk',name:`风险模板.${kind}`,kind,paragraphs,rows:rr,slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素',paragraphIds:paragraphs.map(p=>p.id)}]:[],warnings:[]}]);
}
const fact=(task,issuer='manager')=>task.facts.find(f=>f.key==='risk_rating'&&f.scope.issuer===issuer);
const field=task=>task.materials[0].fields.find(f=>f.label==='风险等级');

for(const [word,code] of [['低风险','R1'],['中低风险','R2'],['中风险','R3'],['中高风险','R4'],['高风险','R5']]){
 test(`Chinese product rating ${word} maps to ${code}, preserving wording and investor codes`,()=>{
  for(const kind of ['docx','pptx']){
   const task=analyze([`本计划属于${word}投资品种，适合风险承受能力为C3及以上的投资者。`],{kind});
   assert.equal(fact(task).value,code);
   assert.equal(field(task).after,`本计划为${code}（${word}）产品，适合风险承受能力为C3及以上的投资者。`);
   assert.equal(fact(task,'distributor').status,'pending');
  }
 });
}
for(const lines of [
 ['本产品为中风险产品。'],['管理人将本集合计划评定为中风险。'],
 ['管理人根据相关制度，将本集合计划的风险评级评价为中风险。'],
 ['本计划风险等级根据管理人内部相关风险评价制度评定为中风险。'],
 ['产品评级：中风险'],['风险等级','中风险'],['风险等级\n中风险'],['本计划属于','中风险投资品种。'],
])test(`recognizes product assertion: ${lines.join(' / ')}`,()=>assert.equal(fact(analyze(lines)).value,'R3'));

for(const line of [
 '投资者风险等级为中风险。','本计划适合风险承受能力为C3（中风险）的投资者。',
 '标的产品风险等级为中风险。','本计划投资的标的产品属于中风险产品。',
 '风险等级分为低风险、中风险和高风险。','本计划风险等级并非中风险。',
 '若本计划风险等级为中风险，则应适用相关规定。','本计划原为中风险产品。',
 '本计划可能为中风险产品。','招商银行将本计划评定为中风险。',
])test(`does not infer product rating from: ${line}`,()=>assert.equal(fact(analyze([line])).status,'pending'));

test('a split heading does not skip unrelated definitions',()=>assert.equal(fact(analyze(['风险等级','等级说明','中风险'])).status,'pending'));
test('explicit R code retains precedence over verbal mapping',()=>{
 const task=analyze(['本计划风险等级为R4（中高风险）。','本计划属于中风险投资品种。']);
 assert.equal(fact(task).value,'R4');assert.match(field(task).after,/R4（中高风险）/);
});
test('manager and distributor verbal ratings remain separate',()=>{
 const task=analyze(['管理人风险评级为中风险，代销机构风险评级为高风险。']);
 assert.equal(fact(task).value,'R3');assert.equal(fact(task,'distributor').value,'R5');
 for(const lines of [['代销机构对本计划的风险评级为中风险。'],['代销机构将本计划评定为中风险。'],['代销机构风险评级','中风险']]){
  const t=analyze(lines);assert.equal(fact(t).status,'pending');assert.equal(fact(t,'distributor').value,'R3');
 }
});
test('conflicting actual product verbal ratings require review',()=>assert.equal(fact(analyze(['本计划属于中风险产品。','本计划风险等级为中低风险。'])).status,'conflict'));
test('a Chinese-only template retains its presentation',()=>{
 const task=analyze(['本计划属于中风险产品。'],{value:'本产品为中低风险，适合C3及以上投资者。'});
 assert.equal(field(task).after,'本产品为中风险，适合C3及以上投资者。');
});
test('direct classification keeps adjacent explicit R code precedence',()=>{
 for(const line of ['本计划属于中高风险R3投资品种。','本计划属于中风险（R4）投资品种。','本产品属于R4（中高风险）产品。']){
  assert.equal(fact(analyze([line])).value,line.includes('R3')?'R3':'R4');
 }
});
test('a distributor assertion retains its subject before a comma',()=>{
 const task=analyze(['根据代销机构评估，本计划属于中风险产品。']);
 assert.equal(fact(task).status,'pending');assert.equal(fact(task,'distributor').value,'R3');
});
test('split headings retain the underlying product subject',()=>{
 for(const lines of [['标的产品\n风险等级\n中风险'],['标的产品','风险等级','中风险']])assert.equal(fact(analyze(lines)).status,'pending');
});
test('alternative verbal ratings do not become a confirmed grade',()=>assert.equal(fact(analyze(['本计划属于中风险或者高风险产品。'])).status,'pending'));
