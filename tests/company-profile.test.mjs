import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import {analyzeMaterials,setSupplement} from '../src/review-engine.js';
import {writeMaterial} from '../src/office-writer.js';
import {parseCompanyProfile,applyCompanyProfile} from '../src/company-profile.js';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR='http://schemas.openxmlformats.org/package/2006/relationships';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
const run=s=>`<w:r><w:rPr><w:b/></w:rPr><w:t>${s}</w:t></w:r>`;
const p=(s,split=false)=>`<w:p>${split?[...s].map(run).join(''):run(s)}</w:p>`;
const row=cells=>`<w:tr>${cells.map(s=>`<w:tc>${s}</w:tc>`).join('')}</w:tr>`;
const table=rows=>`<w:tbl>${rows.join('')}</w:tbl>`;
async function doc(body,name='公司介绍2032.04版本.docx',extra={}){
 const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
 z.file('_rels/.rels',`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
 z.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
 for(const [name,content]of Object.entries(extra))z.file(name,content);
 return{name,bytes:await z.generateAsync({type:'uint8array'})};
}
const sourceFoot='数据来源：星河资管。金牛奖颁奖机构为示例证券报，统计时间为2012年-2032年；近两年管理人类获奖情况为2032.3-五年期示例奖、2031.5-三年期示例奖。';
const bullets=(old=false,split=false)=>p(`注册资本${old?'9':'12'}亿元，首批券商资产管理子公司`,split)+p(`全能型资产管理机构，${old?'14':'16'}多年资产管理经验`,split)+p('专业的投研团队，全面的产品体系',split)+p(`荣获${old?'31':'43'}座业内权威奖项“金牛奖”`,split)+p(`截至${old?'2030':'2031'}年6月末，公司资产管理规模超${old?'5100':'6200'}亿元*`,split);
function narrative(){return table([row([p('雄厚公司实力支持')]),row([p('星河资产管理'),p('股东星河证券')]),row([bullets(),p('星河证券股份有限公司是综合金融服务商。')+p('公司分支机构遍布全国29个省市自治区。')]),row([p(sourceFoot)])]);}
function kpi(){return p('公司实力雄厚')+table([row([p('6300亿元'),p(''),p('荣获43座'),p(''),p('16年')]),row([p('资管规模'),p(''),p('金牛奖'),p(''),p('资管经验')])])+p('数据来源：星河资管。资管规模数据时间截至2032.3.31。金牛奖颁奖机构为示例证券报，统计时间为2012年-2032年；近两年管理人类获奖情况为2032.3-五年期示例奖、2031.5-三年期示例奖。');}
async function profile(){return parseCompanyProfile(await doc(narrative()+kpi()));}
function task(t){return{facts:[],materials:[{materialId:t.id,kind:'word',template:t,fields:[],operations:[],warnings:[]}],warnings:[],stats:{}};}
function revised(m,id){let s=m.template.paragraphs.find(p=>p.id===id).text;for(const o of m.operations.filter(o=>o.paragraphId===id&&o.kind==='replace').sort((a,b)=>b.start-a.start))s=s.slice(0,o.start)+o.after+s.slice(o.start+o.before.length);return s;}

test('structured Word conversion warnings survive company review without blocking it',async()=>{
 const t=await readTemplate(await doc(narrative()));const x=task(t),warning={code:'DOC_SHADING_FALLBACK',message:'请核对转换后的表格底色'};
 x.warnings.push(warning);x.materials[0].warnings.push(warning);
 applyCompanyProfile(x,await profile());applyCompanyProfile(x,await profile());
 assert.ok(x.warnings.includes(warning));assert.ok(x.materials[0].warnings.includes(warning));
});

test('profile is serializable, reads both layouts independently and leaves source bytes intact',async()=>{
 const file=await doc(narrative()+kpi());const original=file.bytes.slice();const x=await parseCompanyProfile(file);
 assert.equal(x.version,'2032.04');assert.match(x.sha256,/^[0-9a-f]{64}$/);assert.deepEqual(x,JSON.parse(JSON.stringify(x)));assert.deepEqual(file.bytes,original);
 assert.deepEqual(x.sections.map(s=>s.layout),['narrative','kpi']);assert.ok(JSON.stringify(x.sections[0]).includes('2031年6月末'));assert.ok(JSON.stringify(x.sections[1]).includes('2032.3.31'));
});
test('trust/advisor columns, compound runs, joint source and optional content stay correctly scoped',async()=>{
 const body=table([row([p('受托人：')+p('晨光信托有限公司'),p('投资顾问：')+p('星河资产管理')]),row([p('晨光信托资产管理规模8万亿元，经验30年。'),bullets(true,true)+p('特色业务由专门团队提供。')]),row([p('数据来源：晨光信托、星河资管。星河资管获得金牛奖颁奖机构为示例证券报，统计时间为2012年-2030年；近两年管理人类获奖情况为2031.5-三年期示例奖。')+p('星河旧公司更名为星河资产管理。')])]);
 const t=await readTemplate(await doc(body,'信托一页通.docx'));const x=applyCompanyProfile(task(t),await profile());const m=x.materials[0];
 const trust=t.paragraphs.find(p=>p.text.startsWith('晨光信托资产'));assert.equal(revised(m,trust.id),trust.text);
 assert.ok(m.operations.length>5);for(const op of m.operations){assert.equal(op.source,'company');assert.equal(t.paragraphs.find(p=>p.id===op.paragraphId).text.slice(op.start,op.start+op.before.length),op.before);}
 const all=t.paragraphs.map(p=>revised(m,p.id)).join('\n');assert.match(all,/注册资本12亿元/);assert.match(all,/截至2031年6月末，公司资产管理规模超6200亿元/);assert.match(all,/数据来源：晨光信托、星河资管/);assert.match(all,/特色业务由专门团队提供/);assert.match(all,/旧公司更名/);assert.ok(!all.includes('股东星河证券'));
 assert.ok(x.facts.every(f=>f.source==='company'&&f.evidence[0].blockId.startsWith('company-profile-')));
 const beforeCount=m.operations.length;applyCompanyProfile(x,await profile());assert.equal(m.operations.length,beforeCount);
 const out=await writeMaterial(m,{month:'203204',targetName:'星河示例集合资产管理计划'});const z=await JSZip.loadAsync(out.bytes);const xml=await z.file('word/document.xml').async('string');assert.match(xml,/<w:ins/);assert.match(xml,/<w:del/);assert.deepEqual(t.bytes,(await readTemplate({name:t.name,bytes:t.bytes})).bytes);
});
test('KPI mapping follows labels even with spacer columns and preserves distinct source date',async()=>{
 const body=p('公司实力雄厚')+table([row([p('31座'),p('5000亿元'),p('14年')]),row([p('金牛奖'),p('资管规模'),p('资管经验')])])+p('数据来源：星河资管，截至2030.12.31。金牛奖颁奖机构为示例证券报，统计时间为2012年-2030年；近两年管理人类获奖情况为2031.5-三年期示例奖；近两年团队获奖情况为2031.1-示例团队奖。');
 const t=await readTemplate(await doc(body));const x=applyCompanyProfile(task(t),await profile());const all=t.paragraphs.map(p=>revised(x.materials[0],p.id)).join('\n');
 assert.match(all,/荣获43座/);assert.match(all,/6300亿元/);assert.match(all,/资管规模数据时间截至2032.3.31/);assert.match(all,/近两年团队获奖情况为2031.1-示例团队奖/);assert.ok(!all.includes('2031年6月末'));
});
test('unsupported subject and PPT cannot acquire company replacements',async()=>{
 const t=await readTemplate(await doc(narrative().replaceAll('星河','远峰')));const x=applyCompanyProfile(task(t),await profile());assert.equal(x.materials[0].operations.length,0);assert.ok(x.materials[0].warnings.length);
 const ppt=task({...t,kind:'pptx'});ppt.materials[0].kind='poster';applyCompanyProfile(ppt,await profile());assert.equal(ppt.materials[0].operations.length,0);
});
test('contract supplements retain company changes without duplicate or missing operations',async()=>{
 const t=await readTemplate(await doc(narrative().replace('6200','5100')+table([row([p('产品名称'),p('星河稳健1号集合资产管理计划')]),row([p('管理费'),p('0.8%')])])));
 const c={name:'示例合同.docx',blocks:[{id:'b0',text:'产品名称：星河稳健2号集合资产管理计划'}],warnings:[]};const x=applyCompanyProfile(analyzeMaterials(c,[t]),await profile());const snapshot=JSON.stringify(x.materials[0].operations.filter(o=>o.source==='company'));const fee=x.facts.find(f=>f.key==='management_fee');assert.ok(fee);setSupplement(x,fee.id,'0.2%');assert.equal(JSON.stringify(x.materials[0].operations.filter(o=>o.source==='company')),snapshot);
});
test('supported old revisions are accepted only in the in-memory company source view',async()=>{
 const body=narrative().replace(run('荣获43座业内权威奖项“金牛奖”'),`<w:del w:id="1"><w:r><w:delText>旧奖项文字</w:delText></w:r></w:del><w:ins w:id="2">${run('荣获43座业内权威奖项“金牛奖”')}</w:ins>`);const f=await doc(body);const before=f.bytes.slice();const x=await parseCompanyProfile(f);assert.ok(x.warnings.some(s=>s.includes('旧修订')));assert.ok(!JSON.stringify(x).includes('旧奖项文字'));assert.deepEqual(f.bytes,before);
});

test('generic manager/shareholder headings use local company source and fill adjacent missing source clauses without new paragraphs',async()=>{
 const body=table([row([p('管理人实力'),p('股东实力')]),row([p('首批券商资产管理子公司')+p('至今14年资产管理经验，专业的投研团队，覆盖面广的产品体系')+p('近三年多次荣获业内权威奖项“金牛奖”')+p('截至2030年6月末，公司资产管理规模达5100亿元*'),p('星河证券股份有限公司是金融服务商。')+p('公司分支机构遍布全国27个省市自治区。')]),row([p('数据来源：星河资管'),p('数据来源：星河证券2030半年度报告')]),row([p('注：金牛奖颁奖机构为示例证券报，统计时间为2012年-2030年；近两年管理人类获奖情况为2031.5-三年期示例奖')+p('星河旧公司更名为星河资产管理。')])]);
 const t=await readTemplate(await doc(body));const x=applyCompanyProfile(task(t),await profile());const all=t.paragraphs.map(p=>revised(x.materials[0],p.id)).join('\n');
 assert.match(all,/注册资本12亿元，首批券商资产管理子公司/);assert.match(all,/全能型资产管理机构，16多年资产管理经验，专业的投研团队，全面的产品体系/);assert.match(all,/荣获43座业内权威奖项“金牛奖”/);assert.match(all,/数据来源：星河证券2030半年度报告/);
 const out=await writeMaterial(x.materials[0],{month:'203204',targetName:'星河示例集合资产管理计划'});const again=await readTemplate({name:out.name,bytes:out.bytes});assert.equal(again.paragraphs.length,t.paragraphs.length);assert.equal(again.paragraphs.map(p=>p.text).join('\n'),all);
});

test('a longer unrelated brand ending in the authority company name is not the same subject',async()=>{
 const t=await readTemplate(await doc(narrative().replaceAll('星河','远峰星河').replace('6200','5100')));const x=applyCompanyProfile(task(t),await profile());assert.equal(x.materials[0].operations.length,0);assert.ok(x.materials[0].warnings.length);
});
test('separate shareholder award footnotes and unlocatable source paragraphs are retained with an explicit warning',async()=>{
 const body=table([row([p('星河资产管理'),p('股东星河证券')]),row([p('注册资本9亿元')+p('14多年资产管理经验')+p('荣获31座业内权威奖项“金牛奖”'),p('星河证券股份有限公司是综合金融服务商。')]),row([p('数据来源：星河资管。统计时间为2012年-2030年'),p('统计时间为2001年-2030年；近两年管理人类获奖情况为2030.2-股东专属奖')])]);
 const t=await readTemplate(await doc(body));const x=applyCompanyProfile(task(t),await profile());const p0=t.paragraphs.find(p=>p.text.includes('股东专属奖'));assert.equal(revised(x.materials[0],p0.id),p0.text);assert.ok(x.materials[0].warnings.some(w=>w.includes('未找到')||w.includes('缺少')));
});
test('KPI award label is reconciled and a separate general date is not overwritten as the scale date',async()=>{
 const body=p('公司实力雄厚')+table([row([p('31座'),p('5000亿元'),p('14年')]),row([p('荣获金牛奖'),p('资管规模'),p('资管经验')])])+p('数据来源：星河资管，截至2030.12.31；规模数据截至2031.1.31。金牛奖颁奖机构为示例证券报，统计时间为2012年-2030年。');
 const t=await readTemplate(await doc(body));const x=applyCompanyProfile(task(t),await profile());const all=t.paragraphs.map(p=>revised(x.materials[0],p.id)).join('\n');assert.ok(!all.includes('荣获金牛奖'));assert.match(all,/截至2030.12.31/);assert.match(all,/资管规模数据时间截至2032.3.31/);assert.ok(x.materials[0].warnings.some(w=>w.includes('日期')));
});

test('referenced Word footnotes are updated but unrelated footnotes are not company evidence',async()=>{
 const body=narrative().replace(p(sourceFoot),'').replace(run('荣获43座业内权威奖项“金牛奖”'),run('荣获31座业内权威奖项“金牛奖”')+'<w:r><w:footnoteReference w:id="3"/></w:r>');
 const notes=`<w:footnotes xmlns:w="${W}"><w:footnote w:id="3">${p('数据来源：星河资管。统计时间为2012年-2030年；近两年管理人类获奖情况为2031.5-三年期示例奖。')}</w:footnote><w:footnote w:id="4">${p('数据来源：另一家公司。统计时间为2001年-2030年。')}</w:footnote></w:footnotes>`;
 const t=await readTemplate(await doc(body,'示例一页通.docx',{'word/footnotes.xml':notes}));const x=applyCompanyProfile(task(t),await profile());const first=t.paragraphs.find(p=>p.part==='word/footnotes.xml'&&p.text.includes('星河资管')),other=t.paragraphs.find(p=>p.part==='word/footnotes.xml'&&p.text.includes('另一家'));
 assert.match(revised(x.materials[0],first.id),/统计时间为2012年-2032年/);assert.equal(revised(x.materials[0],other.id),other.text);
});
