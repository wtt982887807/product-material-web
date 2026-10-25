import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import JSZip from 'jszip';
import fixture from './doc-fixture.cjs';
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
globalThis.JSZip = JSZip;
globalThis.self = globalThis;
await import('../vendor/docToText.js');
const reader = await import('../src/office-reader.js');
const xml = await import('../src/office-xml.js');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const p = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
async function word(body, extra={}) {
  const zip = new JSZip();
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  for (const [part, value] of Object.entries(extra)) zip.file(part,value);
  return {name:'synthetic.docx', bytes:await zip.generateAsync({type:'uint8array'})};
}

test('reader exposes usable entry points',()=>{
  assert.equal(typeof reader.readTemplate,'function');
  assert.equal(typeof reader.readContract,'function');
});

test('Word accepts old inline revisions and preserves comments and source bytes',async()=>{
  assert.equal(typeof reader.readTemplate,'function');
  const file=await word('<w:p><w:commentRangeStart w:id="7"/><w:del w:id="1"><w:r><w:delText>旧值</w:delText></w:r></w:del><w:ins w:id="2"><w:r><w:t>新值</w:t></w:r></w:ins><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>',{'word/comments.xml':`<w:comments xmlns:w="${W}"><w:comment w:id="7">${p('原批注')}</w:comment></w:comments>`});
  const original=file.bytes.slice(); const doc=await reader.readTemplate(file);
  assert.equal(doc.paragraphs[0].text,'新值');
  assert.equal(xml.elements(doc.parts.get('word/document.xml'),W,'ins').length,0);
  assert.equal(xml.elements(doc.parts.get('word/document.xml'),W,'commentRangeStart').length,1);
  assert.equal(await doc.zip.file('word/comments.xml').async('string'), await JSZip.loadAsync(original).then(z=>z.file('word/comments.xml').async('string')));
  assert.deepEqual(file.bytes,original); assert.notEqual(doc.bytes,file.bytes);
  assert.ok(doc.preparation.changedParts.includes('word/document.xml'));
});

test('Word safely accepts a deleted paragraph mark by joining adjacent paragraphs',async()=>{
  assert.equal(typeof reader.readTemplate,'function');
  const doc=await reader.readTemplate(await word(`<w:p><w:pPr><w:rPr><w:del w:id="4"/></w:rPr></w:pPr><w:r><w:t>前</w:t></w:r></w:p>${p('后')}`));
  assert.deepEqual(doc.paragraphs.map(p=>p.text),['前后']);
});

test('old formatting snapshots cannot make accepted visible revisions fail',async()=>{
  const doc=await reader.readTemplate(await word('<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4"><w:rPr><w:ins w:id="2"/></w:rPr></w:rPrChange></w:rPr><w:t>当前值</w:t></w:r></w:p>'));
  assert.equal(doc.paragraphs[0].text,'当前值');
  assert.equal(xml.elements(doc.parts.get('word/document.xml'),W,'rPrChange').length,0);
  assert.equal(xml.elements(doc.parts.get('word/document.xml'),W,'b').length,1);
});

test('accepting deleted text retains existing comment anchors as an empty span',async()=>{
  const doc=await reader.readTemplate(await word('<w:p><w:del w:id="1"><w:commentRangeStart w:id="9"/><w:r><w:delText>删除文字</w:delText></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:del><w:r><w:t>保留文字</w:t></w:r></w:p>'));
  const part=doc.parts.get('word/document.xml');
  assert.equal(doc.paragraphs[0].text,'保留文字');
  assert.equal(xml.elements(part,W,'commentRangeStart').length,1);
  assert.equal(xml.elements(part,W,'commentRangeEnd').length,1);
  assert.equal(xml.elements(part,W,'commentReference').length,1);
  assert.equal(xml.elements(part,W,'commentReference')[0].parentNode.localName,'r');
});

test('Word textboxes choose a single compatibility branch and retain paragraph indexes',async()=>{
  const doc=await reader.readTemplate(await word('<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:r><w:t>标题</w:t></w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:txbxContent>'+p('文本框')+'</w:txbxContent></mc:Choice><mc:Fallback><w:txbxContent>'+p('文本框')+'</w:txbxContent></mc:Fallback></mc:AlternateContent></w:p>'+p('末尾')));
  assert.deepEqual(doc.paragraphs.map(p=>[p.index,p.text]),[[0,'标题'],[1,'文本框'],[3,'末尾']]);
});

test('Word rejects structural cell and unsupported cross-container paragraph revisions',async()=>{
  assert.equal(typeof reader.readTemplate,'function');
  await assert.rejects(reader.readTemplate(await word('<w:tbl><w:tr><w:tc><w:tcPr><w:cellMerge w:id="2"/></w:tcPr>'+p('值')+'</w:tc></w:tr></w:tbl>')), /修订|revision/i);
  await assert.rejects(reader.readTemplate(await word('<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr></w:p><w:tbl/>')), /修订|revision/i);
});

test('Word table cells preserve paragraph boundaries and textbox text is not duplicated',async()=>{
  assert.equal(typeof reader.readTemplate,'function');
  const doc=await reader.readTemplate(await word('<w:tbl><w:tr><w:tc>'+p('管理费')+'</w:tc><w:tc>'+p('A类：0.30%')+p('B类：0.40%')+'</w:tc></w:tr></w:tbl><w:p><w:r><w:t>外层</w:t><w:drawing><w:txbxContent>'+p('文本框')+'</w:txbxContent></w:drawing></w:r></w:p>'));
  assert.equal(doc.rows[0].label,'管理费');assert.equal(doc.rows[0].value,'A类：0.30%\nB类：0.40%');
  assert.deepEqual(doc.paragraphs.map(p=>p.text),['管理费','A类：0.30%','B类：0.40%','外层','文本框']);
  assert.equal(new Set(doc.paragraphs.map(p=>p.id)).size,5);
  assert.equal(doc.rows[0].cells[1].paragraphIds.length,2);
});

test('PPT follows presentation relationships rather than lexical slide filenames',async()=>{
  assert.equal(typeof reader.readTemplate,'function');
  const zip=new JSZip();
  zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="rId9"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels',`<Relationships xmlns="${REL}"><Relationship Id="rId9" Type="${R}/slide" Target="slides/slide9.xml"/><Relationship Id="rId2" Type="${R}/slide" Target="slides/slide2.xml"/></Relationships>`);
  for (const [n,t] of [[9,'第一页'],[2,'第二页']]) zip.file(`ppt/slides/slide${n}.xml`,`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="4" name="文字"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>${t}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  const doc=await reader.readTemplate({name:'synthetic.pptx',bytes:await zip.generateAsync({type:'uint8array'})});
  assert.deepEqual(doc.slides.map(s=>s.part),['ppt/slides/slide9.xml','ppt/slides/slide2.xml']);
  assert.deepEqual(doc.paragraphs.map(p=>[p.text,p.page,p.shapeId]),[['第一页',1,'4'],['第二页',2,'4']]);
});

test('legacy DOC parser reads a binary CFB contract locally and rejects corrupt input',async()=>{
  assert.equal(typeof reader.readContract,'function');
  const c=await reader.readContract({name:'synthetic.doc',bytes:fixture.buildDoc()});
  assert.equal(c.blocks.map(b=>b.text).join('\n').trim(),fixture.EXPECTED.trim());
  await assert.rejects(reader.readContract({name:'broken.doc',bytes:new Uint8Array([1,2,3])}),/DOC|读取/i);
});

test('XML rejects malformed or externally declared documents',()=>{
  assert.equal(typeof xml.parseXml,'function');
  assert.throws(()=>xml.parseXml('<root><x></root>'),/XML/i);
  assert.throws(()=>xml.parseXml('<!DOCTYPE root SYSTEM "https://example.invalid/private"><root/>'),/XML/i);
});

test('XML element traversal supports namespace and local-name wildcards for package edits',()=>{
  const doc=xml.parseXml(`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:extLst/><a:p/></p:sld>`);
  assert.equal(xml.elements(doc,'*','*').length,3);
  assert.equal(xml.elements(doc,'*','p').length,1);
  assert.equal(xml.children(doc.documentElement,'*','*').length,2);
});
