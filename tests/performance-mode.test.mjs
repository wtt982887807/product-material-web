import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {analyzeMaterials,getSupplementGroups,previewPerformanceNumbers} from '../src/review-engine.js';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const target='青澜稳健9号集合资产管理计划';
const charged='管理人对年化收益率超过5.00%的收益部分提取20%的业绩报酬。';
const none='本集合计划不提取业绩报酬。';
function template(id,lines,label='业绩报酬',kind='pptx'){
 const paragraphs=[],part=kind==='docx'?'word/document.xml':'ppt/slides/slide1.xml';
 const rows=[['产品名称',[target]], [label,lines]].map(([label,values],i)=>{
  const lid=id+'l'+i,ids=values.map((_,j)=>id+'v'+i+'-'+j);
  paragraphs.push({id:lid,text:label,part,page:1,index:paragraphs.length,tableId:'t',rowIndex:i,cellIndex:0});
  values.forEach((text,j)=>paragraphs.push({id:ids[j],text,part,page:1,index:paragraphs.length,tableId:'t',rowIndex:i,cellIndex:1}));
  return{id:id+'r'+i,label,value:values.join('\n'),part,page:1,tableId:'t',rowIndex:i,paragraphIds:[lid,...ids],cells:[{text:label,paragraphIds:[lid]},{text:values.join('\n'),paragraphIds:ids}]};
 });
 return{id,name:id+'.'+kind,kind,paragraphs,rows,warnings:[],slides:kind==='pptx'?[{page:1,title:'产品要素',text:'产品要素'}]:[]};
}
const contract=lines=>({name:'合同.docx',warnings:[],blocks:['产品名称：'+target,...lines].map((text,i)=>({id:'b'+i,text}))});
const task=(lines,templates=[template('ppt',[charged])])=>analyzeMaterials(contract(lines),templates,{month:'202610'});
const policy=t=>t.facts.find(f=>f.key==='performance_policy');
const value=t=>t.materials[0].fields[1].after;

test('unconditional no-performance-fee contract replaces obsolete fee wording with one explicit conclusion',()=>{
 const t=task([none],[template('word',[charged,'业绩报酬计提日为分红权益登记日，计算公式E=K×(R-5.00%)×20%。'],'业绩报酬','docx'),template('ppt',[charged]),template('海报',[charged])]);
 assert.equal(policy(t)?.policyMode,'none');assert.equal(policy(t)?.status,'verified');assert.equal(policy(t)?.value,none);
 assert.equal(policy(t)?.source,'contract');assert.ok(policy(t)?.evidence.some(e=>e.quote.includes(none)));
 assert.equal(t.facts.filter(f=>f.key==='performance_policy').length,1);
 assert.ok(!t.facts.some(f=>/^performance_(?:R|ratio)/.test(f.key)||f.key==='custom:performance_parameter'));
 for(const m of t.materials){assert.equal(m.fields[1].after.trim(),none);assert.ok(!m.operations.some(o=>['highlight','comment'].includes(o.kind)));}
 assert.equal(getSupplementGroups(t).find(g=>g.factIds.includes(policy(t).id)).kind,'performance_policy');
 assert.equal(previewPerformanceNumbers(t,{mode:'single',threshold1:'2.5',ratio1:'60'}).updates.length,0);
});

test('an already no-fee template has a verified policy fact and is not an empty unverified field',()=>{
 const t=task([none],[template('word',['不提取业绩报酬。'],'业绩报酬','docx')]);
 assert.equal(policy(t)?.status,'verified');assert.equal(policy(t)?.policyMode,'none');assert.ok(t.materials[0].fields[1].factIds.length);
 assert.equal(value(t),none);
});

test('independent fee fields each retain the no-fee conclusion while continuation paragraphs are cleared only within that field',()=>{
 const a=template('word',[charged,'业绩报酬计提日为分红日，计算公式E=K×(R-5.00%)×20%。'],'业绩报酬','docx');
 const other=template('other',[charged],'业绩报酬','docx'),row=other.rows[1];
 const offset=a.paragraphs.length;
 for(const [i,p]of other.paragraphs.filter(p=>row.paragraphIds.includes(p.id)).entries())a.paragraphs.push({...p,index:offset+i,rowIndex:2});
 a.rows.push({...row,rowIndex:2});
 const t=task([none],[a]),fields=t.materials[0].fields.filter(f=>f.label==='业绩报酬');
 assert.equal(fields.length,2);
 for(const field of fields)assert.equal(field.after.trim(),none);
 assert.equal(fields[0].after.split(none).length-1,1);
 assert.equal(fields[1].after.split(none).length-1,1);
});

test('charging contract and charging template retain their existing numeric-only edits and precision',()=>{
 const t=task(['管理人对年化收益率超过2.5%的收益部分提取60%的业绩报酬。']);
 assert.equal(policy(t),undefined);assert.equal(value(t),charged.replace('5.00%','2.50%').replace('20%','60%'));
 const replacements=t.materials[0].operations.filter(o=>o.kind==='replace');assert.deepEqual(replacements.map(o=>[o.before,o.after]),[['5.00%','2.50%'],['20%','60%']]);
});

test('no-fee template cannot be verified when the contract charges or leaves the mode unknown',()=>{
 for(const lines of [['管理人按收益表现提取业绩报酬，具体基准以公告为准。'],[]]){
  const t=task(lines,[template('a',[none])]);assert.equal(policy(t)?.status,'pending');assert.equal(value(t),none);
  assert.ok(t.materials[0].fields[1].factIds.length);if(lines.length)assert.ok(policy(t).evidence.length);
 }
});

test('conditional waivers, temporary periods, examples and other fee types are not unconditional no-fee policies',()=>{
 for(const text of ['当年化收益率不超过5%时，本集合计划不提取业绩报酬。','本集合计划暂不提取业绩报酬。','本集合计划首期不收取业绩报酬。','本次开放期免收业绩报酬。','例如本集合计划不提取业绩报酬。','底层产品不收取业绩报酬。','本集合计划不收取销售佣金。']){
  const t=task([text]);assert.notEqual(policy(t)?.policyMode,'none',text);assert.equal(value(t),charged,text);
 }
});

test('explicit charging and no-fee clauses in the same scope remain a conflict',()=>{
 const t=task([none,'本计划管理人收取业绩报酬，超过5%的部分提取20%。']);
 assert.equal(policy(t)?.status,'conflict');assert.equal(policy(t)?.policyMode,'conflict');assert.equal(value(t),charged);
 assert.ok(policy(t).evidence.length>=2);
});

test('share and charging-subject scope prevent a different product fee policy from replacing the target clause',()=>{
 const t=task(['A类份额','本集合计划不提取业绩报酬。','B类份额','管理人对年化收益率超过2.5%的收益部分提取60%的业绩报酬。'],[template('b',[charged],'B类份额业绩报酬')]);
 assert.notEqual(policy(t)?.policyMode,'none');
 const unscoped=task(['A类份额','本集合计划不提取业绩报酬。']);assert.notEqual(policy(unscoped)?.policyMode,'none');
 const advisor=task(['投资顾问不收取业绩报酬。']);assert.notEqual(policy(advisor)?.policyMode,'none');assert.equal(value(advisor),charged);
 const underlying=task(['底层基金的业绩报酬','不提取业绩报酬。']);assert.notEqual(policy(underlying)?.policyMode,'none');
});

test('split heading and explicit no-fee value are supported without manufacturing percentages',()=>{
 const t=task(['业绩报酬','不提取']);assert.equal(policy(t)?.policyMode,'none');assert.match(value(t),/不提取业绩报酬/);assert.doesNotMatch(value(t),/%/);
});

test('generic fee headings and if-any liquidation boilerplate do not override an explicit no-fee decision',()=>{
 const t=task(['四、管理人提取业绩报酬',none,'清算后如有余额，分别扣除应计提业绩报酬（如有）。']);
 assert.equal(policy(t)?.policyMode,'none');assert.equal(value(t),none);
});

test('an immediately preceding waiver period or return condition prevents a global no-fee replacement',()=>{
 for(const condition of ['首期优惠安排','当年化收益率为0%时','当年化收益率低于3%时']){
  const t=task([condition,none]);assert.notEqual(policy(t)?.policyMode,'none',condition);assert.equal(value(t),charged,condition);
  assert.equal(policy(t)?.status,'pending',condition);
 }
});

test('a new unqualified performance-fee heading ends a previous local condition',()=>{
 for(const condition of ['首期优惠安排','当年化收益率为0%时']){
  const t=task([condition,'投资者可参与。','业绩报酬',none]);
  assert.equal(policy(t)?.policyMode,'none');assert.equal(value(t),none);
 }
});

test('a no-charge cell below a hurdle belongs to that band, not the entire charging plan',()=>{
 const t=task(['5、业绩报酬','管理人根据持有期年化收益率（R）的不同，分段收取管理人业绩报酬。','R≤R1','不提取','R1＜R≤R2','5%','R＞R2','40%','R1、R2以公告为准。']);
 assert.equal(policy(t),undefined);assert.ok(t.facts.some(f=>f.key==='performance_R1'));
});

test('no-performance-fee replacement does not erase other fee sentences in a compound cell',()=>{
 const t=task([none],[template('a',['管理费0.50%；'+charged+'托管费0.02%。'])]);
 assert.equal(value(t),'管理费0.50%；'+none+'托管费0.02%。');
});

test('Word export retains real deletion/insertion revisions while accepted content contains only the new fee conclusion',async()=>{
 const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/package/2006/relationships',O='http://schemas.openxmlformats.org/officeDocument/2006/relationships',C='http://schemas.openxmlformats.org/package/2006/content-types';
 const esc=s=>s.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;'),p=s=>`<w:p><w:r><w:t>${esc(s)}</w:t></w:r></w:p>`,cell=lines=>`<w:tc>${lines.map(p).join('')}</w:tc>`;
 const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body><w:tbl><w:tr>${cell(['产品名称'])}${cell([target])}</w:tr><w:tr>${cell(['业绩报酬'])}${cell([charged,'业绩报酬计提日为分红日，计算公式E=K×(R-5.00%)×20%。'])}</w:tr></w:tbl><w:sectPr/></w:body></w:document>`);
 z.file('[Content_Types].xml',`<Types xmlns="${C}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
 z.file('_rels/.rels',`<Relationships xmlns="${R}"><Relationship Id="rId1" Type="${O}/officeDocument" Target="word/document.xml"/></Relationships>`);
 const {readTemplate}=await import('../src/office-reader.js'),{writeTask}=await import('../src/office-writer.js');
 const input=await readTemplate({name:'【一页通202608】青澜稳健9号.docx',bytes:await z.generateAsync({type:'uint8array'})});
 const t=task([none],[input]),out=await writeTask(t);assert.equal(out.failures.length,0);assert.equal(out.files.length,1);assert.equal(out.files[0].report.structuralValidation,true);
 const zip=await JSZip.loadAsync(out.files[0].bytes),xml=await zip.file('word/document.xml').async('string');
 assert.match(xml,/<w:del\b/);assert.match(xml,/<w:ins\b/);assert.match(xml,/5\.00%/);
 const accepted=await readTemplate({name:out.files[0].name,bytes:out.files[0].bytes}),text=accepted.paragraphs.map(p=>p.text).join('\n');
 assert.match(text,/本集合计划不提取业绩报酬/);assert.doesNotMatch(text,/5\.00%|20%|E=K|业绩报酬计提日/);
});
