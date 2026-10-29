import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import JSZip from 'jszip';
import {readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {applyCompanyProfile,parseCompanyProfile} from '../src/company-profile.js';
import {readTemplate} from '../src/office-reader.js';

const module = await import('../src/default-company-profile.js').catch(error => {
  if(error.code === 'ERR_MODULE_NOT_FOUND') return {};
  throw error;
});
const load = (variant='general') => {
  assert.equal(typeof module.loadDefaultCompanyProfile,'function','built-in company profile loader is implemented');
  return module.loadDefaultCompanyProfile(variant);
};
globalThis.DOMParser=DOMParser;
globalThis.XMLSerializer=XMLSerializer;
globalThis.JSZip=JSZip;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/package/2006/relationships';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
const paragraph=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const row=cells=>'<w:tr>'+cells.map(text=>'<w:tc>'+text+'</w:tc>').join('')+'</w:tr>';
const table=rows=>'<w:tbl>'+rows.join('')+'</w:tbl>';
async function document(body){
  const zip=new JSZip();
  zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  zip.file('_rels/.rels',`<Relationships xmlns="${R}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  return{name:'合成旧版一页通.docx',bytes:await zip.generateAsync({type:'uint8array'})};
}

test('built-in profile includes both complete company layouts with the corrected June 2026 scale date',()=>{
  const profile=load();
  assert.equal(profile.version,'2026.10');
  assert.equal(profile.name,'公司介绍2026.10（通用版）.docx');
  assert.match(profile.sha256,/^[a-f\d]{64}$/);
  assert.deepEqual(profile.sections.map(section=>section.layout),['narrative','kpi']);
  const [narrative,kpi]=profile.sections;
  assert.deepEqual(narrative.columns.find(column=>column.role==='company').entries.map(entry=>[entry.key,entry.text]),[
    ['capital','注册资本20亿元'],['qualification','首批券商资产管理子公司'],
    ['institution','全能型资产管理机构'],['experience','20多年资产管理经验'],
    ['team','专业的投研团队'],['product_system','全面的产品体系'],
    ['awards','荣获71座业内权威奖项“金牛奖”'],
    ['aum','截至2026年6月末，公司资产管理规模超8800亿元'],
  ]);
  assert.deepEqual(narrative.columns.find(column=>column.role==='shareholder').entries.map(entry=>entry.text),[
    '国泰海通证券股份有限公司由国泰君安与海通证券合并而成，是中国资本市场长期、持续、全面领先的综合金融服务商。',
    '公司分支机构遍布全国31个省市自治区，服务范围覆盖17个国家和地区。资本实力位居行业最前列。',
    '（管理人与股东之间实行业务隔离制度，股东情况仅供参考，股东不直接参与产品运作管理）',
  ]);
  assert.deepEqual(kpi.columns[0].entries.map(entry=>entry.text),['8800亿元','荣获71座','20年','资管规模','金牛奖','资管经验']);
  assert.equal(kpi.notes[0].entries.find(entry=>entry.key==='aum_date').text,'资管规模数据时间截至2026.6.30');
  for(const section of profile.sections){
    const notes=section.notes.map(note=>note.text).join('\n');
    assert.match(notes,/统计时间为2010年-2026年/);
    assert.match(notes,/2026\.9-五年期金牛券商集合资产管理人、2025\.11-三年期金牛券商集合资产管理人、2024\.11-五年期金牛券商集合资产管理人/);
    for(const entry of [...section.columns.flatMap(column=>column.entries),...section.notes.flatMap(note=>note.entries)]){
      const source=profile.sourceTexts.find(text=>text.paragraphId===entry.paragraphId);
      assert.ok(source,'every field retains its source paragraph');
      assert.equal(source.text.slice(entry.start,entry.start+entry.text.length),entry.text);
    }
  }
});

test('generic is the default while CCB has independent wording and two complete layouts',()=>{
  const general=load(),ccb=load('ccb');
  assert.deepEqual(module.loadDefaultCompanyProfile(),general);
  assert.equal(ccb.name,'公司介绍2026.10（建行版）.docx');
  assert.deepEqual(ccb.sections.map(section=>section.layout),['narrative','kpi']);
  const shareholder=profile=>profile.sections[0].columns.find(column=>column.role==='shareholder').entries.map(e=>e.text).join('\n');
  assert.match(shareholder(general),/全面领先/);assert.match(shareholder(general),/资本实力位居行业最前列/);
  assert.doesNotMatch(shareholder(ccb),/全面领先|资本实力位居行业最前列/);
  assert.match(shareholder(ccb),/长期、持续的综合金融服务商/);
  const kpiValues=profile=>profile.sections[1].columns.map(c=>({role:c.role,subject:c.subject,entries:c.entries.map(({key,text})=>({key,text}))}));
  assert.deepEqual(kpiValues(ccb),kpiValues(general));
  ccb.sections[0].columns[0].entries[0].text='changed';assert.notEqual(load('ccb').sections[0].columns[0].entries[0].text,'changed');
  assert.throws(()=>load('unknown'),/版本/);
});

test('each downloadable Word is exactly the matching built-in reference with metadata removed',async()=>{
  assert.equal(typeof module.getDefaultCompanyTemplate,'function');
  for(const variant of ['general','ccb']){
    const descriptor=module.getDefaultCompanyTemplate(variant),bytes=await readFile(new URL(descriptor.url)),profile=load(variant);
    assert.equal(descriptor.name,profile.name);assert.equal(descriptor.sha256,profile.sha256);
    assert.equal(createHash('sha256').update(bytes).digest('hex'),profile.sha256);
    assert.deepEqual(await parseCompanyProfile({name:descriptor.name,bytes}),profile);
    const zip=await JSZip.loadAsync(bytes);
    assert.ok(!Object.keys(zip.files).some(name=>name.startsWith('docProps/')||name.startsWith('customXml/')));
    assert.doesNotMatch(await zip.file('_rels/.rels').async('string'),/docProps/);
    assert.doesNotMatch(await zip.file('[Content_Types].xml').async('string'),/docProps/);
    const body=await zip.file('word/document.xml').async('string');
    assert.ok(body.includes(variant==='general'?'通用版本':'建行版本'));
    assert.ok(!body.includes(variant==='general'?'建行版本':'通用版本'));
    const template=await readTemplate({name:descriptor.name,bytes});assert.equal(new Set(template.rows.map(row=>row.tableId)).size,2);
  }
  assert.throws(()=>module.getDefaultCompanyTemplate('unknown'),/版本/);
});

test('built-in profile is a fresh plain object and cannot be corrupted by a previous task',()=>{
  const original=load();
  assert.deepEqual(original,JSON.parse(JSON.stringify(original)));
  const changed=load();
  changed.sections[0].columns[0].entries[0].text='changed';
  changed.sections[1].notes[0].entries.length=0;
  changed.sourceTexts[0].text='changed';
  changed.warnings.push('local task warning');
  assert.deepEqual(load(),original);
});

test('default contains review evidence only, without raw document bytes or private file metadata',()=>{
  const profile=load();
  assert.deepEqual(Object.keys(profile).sort(),['name','schemaVersion','sections','sha256','sourceTexts','version','warnings']);
  assert.match(profile.sha256,/^[a-f\d]{64}$/);
  const visit=value=>{
    if(!value||typeof value!=='object')return;
    assert.ok(Array.isArray(value)||Object.getPrototypeOf(value)===Object.prototype);
    for(const [key,child]of Object.entries(value)){
      assert.ok(!['bytes','author','creator','lastModifiedBy','path','absolutePath','filePath'].includes(key));
      if(typeof child==='string')assert.doesNotMatch(child,/[A-Z]:\\|\\\\|docProps\/|file:\/\//i);
      visit(child);
    }
  };
  visit(profile);
});

test('built-in evidence updates both supported layouts in the existing company review engine',async()=>{
  const narrative=table([
    row([paragraph('国泰海通资产管理')]),
    row([paragraph('注册资本9亿元，首批券商资产管理子公司')+paragraph('全能型资产管理机构，12多年资产管理经验')+paragraph('荣获12座业内权威奖项“金牛奖”')+paragraph('截至2024年6月末，公司资产管理规模超5000亿元*')]),
    row([paragraph('数据来源：国泰海通资管。金牛奖颁奖机构为中国证券报，统计时间为2010年-2024年。')]),
  ]);
  const kpi=paragraph('公司实力雄厚')+table([
    row([paragraph('5000亿元'),paragraph('12座'),paragraph('12年')]),
    row([paragraph('资管规模'),paragraph('金牛奖'),paragraph('资管经验')]),
  ])+paragraph('数据来源：国泰海通资管。资管规模数据时间截至2024.6.30。');
  const template=await readTemplate(await document(narrative+kpi));
  const task={facts:[],materials:[{materialId:template.id,kind:'word',template,fields:[],operations:[],warnings:[]}],warnings:[],stats:{}};
  applyCompanyProfile(task,load());
  const material=task.materials[0];
  const text=template.paragraphs.map(p=>{
    let value=p.text;
    for(const operation of material.operations.filter(op=>op.kind==='replace'&&op.paragraphId===p.id).sort((a,b)=>b.start-a.start))value=value.slice(0,operation.start)+operation.after+value.slice(operation.start+operation.before.length);
    return value;
  }).join('\n');
  assert.match(text,/注册资本20亿元/);
  assert.match(text,/荣获71座业内权威奖项“金牛奖”/);
  assert.match(text,/截至2026年6月末，公司资产管理规模超8800亿元/);
  assert.match(text,/资管规模数据时间截至2026\.6\.30/);
  assert.ok(material.operations.length>0);
  assert.ok(task.facts.some(fact=>fact.source==='company'));
});
