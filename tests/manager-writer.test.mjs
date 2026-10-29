import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {fixture,p,run,table,profile,task,normalize} from './manager-fixture.mjs';
import {applyManagerProfile} from '../src/manager-profile.js';
import {writeMaterial} from '../src/office-writer.js';
import {readTemplate} from '../src/office-reader.js';
import {NS,parseXml,elements,serializeXml} from '../src/office-xml.js';
async function output(t,biography){const x=applyManagerProfile(task(t),profile([{biography}])),file=await writeMaterial(x.materials[0],{month:'202610',targetName:'示例产品'}),zip=await JSZip.loadAsync(file.bytes),xml=await zip.file('word/document.xml').async('string');return{x,file,zip,doc:parseXml(xml),accepted:await readTemplate({name:'output.docx',bytes:file.bytes})};}

test('Word tracks only changed characters, including pure insertion without unchanged anchors',async()=>{
 const t=await fixture(p('投资经理')+p('张三')+p('9年证券从业经验。金融硕士。')+p('现任甲公司投资经理。')+p('产品要素'));
 const biography='10年证券从业经验。金融硕士。\n现任甲公司高级投资经理。',o=await output(t,biography);
 assert.equal(elements(o.doc,NS.w,'delText').map(n=>n.textContent).join(''),'9');
 assert.equal(elements(o.doc,NS.w,'ins').map(n=>n.textContent).join(''),'10高级');
 assert.equal(normalize(o.accepted.paragraphs.slice(2,4).map(p=>p.text).join('\n')),normalize(biography));
 assert.equal(o.accepted.paragraphs.at(-1).text,'产品要素');
});

test('Word adds missing biography paragraphs and deletes obsolete content while preserving names and photo nodes',async()=>{
 const t=await fixture(p('投资经理')+p('张三：','<w:r><w:drawing/></w:r>')+p('金融硕士。')+p('现任甲公司经理。')+p('已不再使用的长段落。')+p('产品要素'));
 const nameXml=serializeXml(t.paragraphs[1].node),biography='金融硕士。\n现任甲公司经理。\n擅长量化投资。\n重视风险管理。',o=await output(t,biography);
 assert.equal(serializeXml(o.accepted.paragraphs[1].node),nameXml);
 assert.equal(normalize(o.accepted.paragraphs.slice(2,-1).map(p=>p.text).join('\n')),normalize(biography));
 assert.equal(o.accepted.paragraphs.at(-1).text,'产品要素');
 assert.equal(elements(o.doc,NS.w,'drawing').length,1);assert.ok(elements(o.doc,NS.w,'ins').length);
 assert.ok(o.accepted.paragraphs.some(p=>p.text==='重视风险管理。'));
});

test('Word retains inline name and line breaks and matches the full canonical biography',async()=>{
 const t=await fixture(table(p('投资经理'),`<w:p>${run('张三')}<w:r><w:br/></w:r>${run('9年证券从业经验。')}<w:r><w:br/></w:r>${run('现任旧公司经理。')}</w:p>`)+p('产品要素'));
 const biography='10年证券从业经验。\n现任新公司经理。\n投资风格稳健。',o=await output(t,biography);
 assert.equal(normalize(o.accepted.paragraphs.find(p=>p.text.startsWith('张三')).text.slice(2)),normalize(biography));
 assert.equal(elements(o.doc,NS.w,'t').some(t=>t.textContent.includes('\n')),false);
});

test('Word removes a surplus paragraph without rewriting unchanged biography text',async()=>{
 const t=await fixture(p('张三')+p('金融硕士。')+p('现任甲公司经理。')+p('过时的投资理念和奖项。')+p('产品要素'));
 const o=await output(t,'金融硕士。\n现任甲公司经理。');
 assert.equal(elements(o.doc,NS.w,'delText').map(n=>n.textContent).join(''),'过时的投资理念和奖项。');
 assert.equal(elements(o.doc,NS.w,'ins').length,0);
 assert.deepEqual(o.accepted.paragraphs.map(p=>p.text),['张三','金融硕士。','现任甲公司经理。','产品要素']);
});

test('removed biography bullet disappears and inserted paragraphs retain native bullet formatting',async()=>{
 const t=await fixture(p('张三')+p('金融硕士。')+p('现任甲公司经理。')+p('过时的投资理念。')+p('产品要素'));
 const o=await output(t,'金融硕士。\n现任甲公司经理。');
 assert.deepEqual(o.accepted.paragraphs.map(p=>p.text),['张三','金融硕士。','现任甲公司经理。','产品要素']);
 assert.match(elements(o.doc,NS.w,'delText').map(n=>n.textContent).join(''),/过时的投资理念。/);
});

test('paragraph insertion inherits text style when native tab-stop properties precede the text',async()=>{
 const styled=text=>`<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="7938"/></w:tabs></w:pPr>${run(text)}</w:p>`;
 const t=await fixture(p('投资经理')+p('张三')+styled('金融硕士。')+styled('投资风格稳健。')+p('产品要素'));
 const biography='金融硕士。\n现任甲公司经理。\n投资风格稳健。',o=await output(t,biography);
 assert.equal(normalize(o.accepted.paragraphs.slice(2,-1).map(p=>p.text).join('\n')),normalize(biography));
});

test('a missing first biography paragraph is inserted after an inline manager name',async()=>{
 const t=await fixture(p('投资经理简介')+p('张三：现任甲公司投资经理。')+p('产品要素'));
 const biography='10年证券从业经验。金融硕士。\n现任甲公司投资经理。\n投资风格稳健。',o=await output(t,biography);
 assert.ok(o.accepted.paragraphs[1].text.startsWith('张三：'));
 assert.equal(normalize(o.accepted.paragraphs.slice(1,-1).map(p=>p.text).join('\n').replace(/^张三：/,'')),normalize(biography));
});

test('biography internal headings are updated together with all their following content',async()=>{
 const t=await fixture(p('投资经理简介')+p('张三')+p('金融硕士。')+p('投资特点')+p('坚持旧投资策略。')+p('产品要素'));
 const biography='金融博士。\n投资特点\n坚持新投资策略。',o=await output(t,biography);
 assert.equal(normalize(o.accepted.paragraphs.slice(2,-1).map(p=>p.text).join('\n')),normalize(biography));
});
