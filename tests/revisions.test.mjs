import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import {writeMaterial} from '../src/office-writer.js';
import {NS,elements,children,serializeXml} from '../src/office-xml.js';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const W=NS.w,p=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,r=text=>`<w:r><w:t>${text}</w:t></w:r>`,cell=(text,pr='')=>`<w:tc><w:tcPr>${pr}</w:tcPr>${p(text)}</w:tc>`,row=(cells,pr='')=>`<w:tr><w:trPr>${pr}</w:trPr>${cells}</w:tr>`;
async function fixture(body,extra={}){const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);for(const [name,s]of Object.entries(extra))z.file(name,s);z.file('[Content_Types].xml',`<Types xmlns="${NS.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);z.file('_rels/.rels',`<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`);return{name:'修订样例.docx',bytes:await z.generateAsync({type:'uint8array'})};}
const root=t=>t.parts.get('word/document.xml'),text=t=>t.paragraphs.map(p=>p.text).join('|');

test('tblPrExChange is an obsolete snapshot; current table formatting and immutable source survive',async()=>{
 const file=await fixture(`<w:tbl>${row(cell('当前值')) .replace('<w:trPr>','<w:tblPrEx><w:tblW w:w="7200" w:type="dxa"/><w:tblPrExChange w:id="8"><w:tblPrEx><w:tblW w:w="4000" w:type="dxa"/></w:tblPrEx></w:tblPrExChange></w:tblPrEx><w:trPr>')}</w:tbl>`);const original=file.bytes.slice();const t=await readTemplate(file);
 assert.equal(text(t),'当前值');assert.equal(elements(root(t),W,'tblPrExChange').length,0);assert.equal(elements(root(t),W,'tblW')[0].getAttributeNS(W,'w'),'7200');assert.deepEqual(file.bytes,original);assert.equal(t.preparation.acceptedRevisions,1);
});
test('deleted table rows and cells disappear; inserted rows and cells keep current content and grid',async()=>{
 const t=await readTemplate(await fixture(`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>${row(cell('删除整行'),'<w:del w:id="1"/>')}${row(cell('新增行A','<w:cellIns w:id="3"/>')+cell('新增行B'),'<w:ins w:id="2"/>')}${row(cell('保留单元格')+cell('删除单元格','<w:cellDel w:id="4"/>'))}</w:tbl>`));
 assert.equal(text(t),'新增行A|新增行B|保留单元格');assert.equal(t.rows.length,2);assert.equal(t.rows[1].cells.length,1);assert.equal(elements(root(t),W,'gridCol').length,2);assert.ok(!/w:(?:del|ins|cellDel|cellIns)\b/.test(serializeXml(root(t))));
});
test('a wholly deleted table is removed instead of retaining an invalid zero-cell shell',async()=>{
 const t=await readTemplate(await fixture(`<w:tbl>${row(cell('删除','<w:cellDel w:id="1"/>'))}</w:tbl>${p('保留正文')}`));assert.equal(text(t),'保留正文');assert.equal(elements(root(t),W,'tbl').length,0);
});
test('move revisions remove source runs/paragraphs and preserve destination properties and paired range markers',async()=>{
 const moved='<w:moveFromRangeStart w:id="1" w:name="m"/><w:p><w:pPr><w:rPr><w:moveFrom w:id="2"/></w:rPr></w:pPr>'+r('旧段落位置')+'</w:p><w:moveFromRangeEnd w:id="1"/>';
 const dest='<w:moveToRangeStart w:id="3" w:name="m"/><w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:moveTo w:id="4"/></w:rPr></w:pPr><w:moveTo w:id="5">'+r('新段落位置')+'</w:moveTo></w:p><w:moveToRangeEnd w:id="3"/>';
 const t=await readTemplate(await fixture(moved+'<w:p>'+r('甲')+'<w:moveFrom w:id="6">'+r('旧片段')+'</w:moveFrom><w:moveTo w:id="7">'+r('新片段')+'</w:moveTo>'+r('乙')+'</w:p>'+dest));assert.equal(text(t),'甲新片段乙|新段落位置');assert.equal(elements(root(t),W,'jc')[0].getAttributeNS(W,'val'),'right');assert.ok(!/w:move/.test(serializeXml(root(t))));
});
test('custom XML deletion ranges unwrap deleted markup without deleting unmarked text',async()=>{
 const t=await readTemplate(await fixture('<w:p><w:customXmlDelRangeStart w:id="5"/><w:customXml w:element="outer"><w:customXmlPr/><w:customXml w:element="inner">'+r('仍保留的文字')+'</w:customXml><w:customXmlDelRangeEnd w:id="5"/></w:customXml>'+r('后文')+'</w:p>'));
 assert.equal(text(t),'仍保留的文字后文');assert.equal(elements(root(t),W,'customXml').length,0);assert.ok(!serializeXml(root(t)).includes('customXmlDelRange'));
});
test('custom XML insertion ranges keep markup and malformed move ranges remain an explicit error',async()=>{
 const t=await readTemplate(await fixture('<w:p><w:customXmlInsRangeStart w:id="1"/><w:customXml w:element="new">'+r('新节点')+'</w:customXml><w:customXmlInsRangeEnd w:id="1"/></w:p>'));assert.equal(elements(root(t),W,'customXml').length,1);assert.ok(!serializeXml(root(t)).includes('customXmlInsRange'));
 await assert.rejects(readTemplate(await fixture('<w:p><w:moveFromRangeStart w:id="1"/>'+r('不完整范围')+'</w:p>')),/范围|修订/);
});
test('accepting a deleted row preserves its old comment at a surviving paragraph without invalid table children',async()=>{
 const anchors='<w:commentRangeStart w:id="8"/><w:commentRangeEnd w:id="8"/><w:r><w:commentReference w:id="8"/></w:r>';
 const t=await readTemplate(await fixture(`<w:tbl>${row('<w:tc><w:p>'+anchors+r('待删')+'</w:p></w:tc>','<w:del w:id="1"/>')}${row(cell('保留'))}</w:tbl>`,{'word/comments.xml':`<w:comments xmlns:w="${W}"><w:comment w:id="8">${p('原批注')}</w:comment></w:comments>`}));
 assert.equal(text(t),'保留');for(const name of ['commentRangeStart','commentRangeEnd','commentReference']){const nodes=elements(root(t),W,name);assert.equal(nodes.length,1);assert.equal(nodes[0].parentNode.localName,name==='commentReference'?'r':'p');}
});
test('paragraph mark deletion chains keep successor paragraph formatting and numbered current state',async()=>{
 const old=s=>'<w:p><w:pPr><w:jc w:val="left"/><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>'+r(s)+'</w:p>';
 const final='<w:p><w:pPr><w:jc w:val="right"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/><w:numberingChange w:id="6" w:original="old"/></w:numPr></w:pPr>'+r('丙')+'</w:p>';
 const t=await readTemplate(await fixture(old('甲')+old('乙')+final));assert.equal(text(t),'甲乙丙');assert.equal(elements(root(t),W,'jc')[0].getAttributeNS(W,'val'),'right');assert.equal(elements(root(t),W,'numberingChange').length,0);assert.equal(elements(root(t),W,'numId')[0].getAttributeNS(W,'val'),'9');
});
test('style, numbering, comment and other Word XML revisions are accepted without turning them into contract text',async()=>{
 const change='<w:rPr><w:b/><w:rPrChange w:id="4"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>';
 const t=await readTemplate(await fixture(p('正文'),{'word/styles.xml':`<w:styles xmlns:w="${W}"><w:style w:styleId="N">${change}</w:style></w:styles>`,'word/numbering.xml':`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">${change}</w:lvl></w:abstractNum></w:numbering>`,'word/comments.xml':`<w:comments xmlns:w="${W}"><w:comment w:id="1"><w:p><w:ins w:id="8">${r('批注保留')}</w:ins></w:p></w:comment></w:comments>`}));
 assert.equal(text(t),'正文');for(const part of ['word/styles.xml','word/numbering.xml','word/comments.xml']){assert.ok(t.preparation.changedParts.includes(part));assert.ok(!/w:(?:rPrChange|ins)\b/.test(serializeXml(t.parts.get(part))));}assert.equal(elements(t.parts.get('word/styles.xml'),W,'b').length,1);
});
test('a cell merge with explicit consistent current merge geometry accepts history without losing content',async()=>{
 const t=await readTemplate(await fixture(`<w:tbl>${row(cell('合并头','<w:vMerge w:val="restart"/>'))}${row(cell('','<w:vMerge w:val="continue"/><w:cellMerge w:id="2" w:vMerge="cont"/>'))}</w:tbl>`));assert.equal(elements(root(t),W,'cellMerge').length,0);assert.deepEqual(elements(root(t),W,'vMerge').map(n=>n.getAttributeNS(W,'val')),['restart','continue']);
 await assert.rejects(readTemplate(await fixture(`<w:tbl>${row(cell('未指明新状态','<w:cellMerge w:id="2"/>'))}</w:tbl>`)),/修订/);
});

test('the output package contains accepted old revisions and only this round of new tracked edits',async()=>{
 const file=await fixture('<w:tbl>'+row(cell('旧行'),'<w:del w:id="3"/>')+row(cell('更新数值7'))+'</w:tbl>',{'word/styles.xml':`<w:styles xmlns:w="${W}"><w:style w:styleId="N"><w:rPr><w:b/><w:rPrChange w:id="4"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr></w:style></w:styles>`});
 const t=await readTemplate(file),p0=t.paragraphs[0];const m={kind:'word',template:t,fields:[],operations:[{id:'new-op',kind:'replace',paragraphId:p0.id,part:p0.part,index:p0.index,start:p0.text.indexOf('7'),before:'7',after:'9',factIds:[]}]};
 const out=await writeMaterial(m,{month:'203204',targetName:'示例集合资产管理计划'}),z=await JSZip.loadAsync(out.bytes),body=await z.file('word/document.xml').async('string'),styles=await z.file('word/styles.xml').async('string');assert.ok(!body.includes('旧行'));assert.equal((body.match(/<w:del\b/g)||[]).length,1);assert.equal((body.match(/<w:ins\b/g)||[]).length,1);assert.ok(!styles.includes('rPrChange'));const reread=await readTemplate({name:out.name,bytes:out.bytes});assert.equal(text(reread),'更新数值9');assert.equal(t.paragraphs[0].text,'更新数值7');
});
test('normalizing one text revision does not delete unrelated pre-existing empty structures',async()=>{
 const t=await readTemplate(await fixture('<w:tbl/><w:p><w:ins w:id="1">'+r('保留新文字')+'</w:ins></w:p>'));assert.equal(elements(root(t),W,'tbl').length,1);
});

test('a deleted trailing row collapses a comment after its surviving start instead of reversing the range',async()=>{
 const body='<w:tbl>'+row('<w:tc><w:p><w:commentRangeStart w:id="5"/>'+r('保留')+'</w:p></w:tc>')+row('<w:tc><w:p>'+r('删除')+'<w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r></w:p></w:tc>','<w:del w:id="8"/>')+'</w:tbl>';
 const t=await readTemplate(await fixture(body)),s=serializeXml(root(t));assert.equal(text(t),'保留');assert.ok(s.indexOf('commentRangeStart')<s.indexOf('commentRangeEnd'));
});
test('complex supported-boundary fallback and damaged revision ranges have different machine-readable codes',async()=>{
 await assert.rejects(readTemplate(await fixture('<w:p><w:pPr><w:rPr><w:del w:id="1"/></w:rPr></w:pPr>'+r('跨表合并')+'</w:p><w:tbl/>')),e=>e.code==='WORD_REVISION_COMPLEX');
 await assert.rejects(readTemplate(await fixture('<w:p><w:moveToRangeStart w:id="1"/>'+r('缺结束标记')+'</w:p>')),e=>e.code==='WORD_REVISION_INVALID');
});
test('deleting through an existing vertical merge is delegated instead of leaving dangling merge continuations',async()=>{
 const body='<w:tbl>'+row(cell('上','<w:vMerge w:val="restart"/>'),'<w:del w:id="1"/>')+row(cell('','<w:vMerge w:val="continue"/>'))+'</w:tbl>';
 await assert.rejects(readTemplate(await fixture(body)),e=>e.code==='WORD_REVISION_COMPLEX');
});
test('deleting every row of a merged group is safe when the next surviving row starts independent cells',async()=>{
 const body='<w:tbl>'+row(cell('合并上','<w:vMerge w:val="restart"/>'),'<w:del w:id="1"/>')+row(cell('','<w:vMerge w:val="continue"/>'),'<w:del w:id="2"/>')+row(cell('独立行'))+'</w:tbl>';
 const t=await readTemplate(await fixture(body));assert.equal(text(t),'独立行');assert.equal(t.rows.length,1);assert.equal(elements(root(t),W,'vMerge').length,0);
});
