import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import JSZip from 'jszip';
import {NS,parseXml,elements,serializeXml} from '../src/office-xml.js';

globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;globalThis.JSZip=JSZip;
const module=await import('../src/company-template.js').catch(error=>{if(error.code==='ERR_MODULE_NOT_FOUND')return{};throw error;});
const prepare=(file,options)=>{assert.equal(typeof module.prepareCompanyTemplate,'function');return module.prepareCompanyTemplate(file,options);};
const p=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table=text=>`<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tr><w:tc><w:tcPr><w:shd w:fill="AABBCC"/></w:tcPr>${p(text)}</w:tc></w:tr></w:tbl>`;
async function fixture(body){
 const zip=new JSZip();
 zip.file('word/document.xml',`<w:document xmlns:w="${NS.w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);
 zip.file('_rels/.rels',`<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${NS.pr}/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${NS.r}/custom-properties" Target="docProps/custom.xml"/></Relationships>`);
 zip.file('word/_rels/document.xml.rels',`<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${NS.r}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${NS.r}/customXml" Target="../customXml/item1.xml"/></Relationships>`);
 zip.file('[Content_Types].xml',`<Types xmlns="${NS.ct}"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/><Override PartName="/customXml/item1.xml" ContentType="application/xml"/></Types>`);
 zip.file('word/styles.xml','<styles>retained style definitions</styles>');
 zip.file('docProps/core.xml','<creator>Private author</creator>');zip.file('docProps/custom.xml','<property>Private WPS ID</property>');zip.file('customXml/item1.xml','<private>Private source payload</private>');
 return{name:'公司介绍2026.10 .docx',bytes:await zip.generateAsync({type:'uint8array'})};
}
test('unmarked files retain their original bytes and name with a defensive copy',async()=>{
 const input=await fixture(table('通用版本只是表格中的普通内容'));
 const before=input.bytes.slice();const result=await prepare(input,{variant:'ccb'});
 assert.equal(result.name,input.name);assert.deepEqual(result.bytes,before);assert.notEqual(result.bytes,input.bytes);
 result.bytes[0]=0;assert.deepEqual(input.bytes,before);
});
test('explicit version headings split the selected body, retain layout and remove private metadata',async()=>{
 const input=await fixture(p('通用版本')+p('general heading')+table('general contents')+p('建行版本')+p('ccb heading')+table('ccb contents'));
 const before=input.bytes.slice(),source=await JSZip.loadAsync(input.bytes),sourceDoc=parseXml(await source.file('word/document.xml').async('string'));
 for(const variant of ['general','ccb']){
  const result=await prepare(input,{variant}),zip=await JSZip.loadAsync(result.bytes),doc=parseXml(await zip.file('word/document.xml').async('string'));
  assert.equal(result.name,`公司介绍2026.10（${variant==='general'?'通用版':'建行版'}）.docx`);
  assert.match(doc.documentElement.textContent,new RegExp(`${variant} contents`));
  assert.doesNotMatch(doc.documentElement.textContent,new RegExp(`${variant==='general'?'ccb':'general'} contents`));
  assert.equal(elements(doc,NS.w,'tbl').length,1);
  assert.equal(serializeXml(elements(doc,NS.w,'tbl')[0]),serializeXml(elements(sourceDoc,NS.w,'tbl')[variant==='general'?0:1]));
  assert.equal(elements(doc,NS.w,'sectPr').length,1);assert.equal(elements(doc,NS.w,'pgSz')[0].getAttributeNS(NS.w,'w'),'11906');
  assert.equal(await zip.file('word/styles.xml').async('string'),await source.file('word/styles.xml').async('string'));
  assert.ok(!Object.keys(zip.files).some(path=>/^(?:docProps|customXml)\//.test(path)));
  for(const part of ['_rels/.rels','word/_rels/document.xml.rels','[Content_Types].xml'])assert.doesNotMatch(await zip.file(part).async('string'),/docProps|customXml/);
 }
 assert.deepEqual(input.bytes,before);
 assert.deepEqual(await prepare(input),await prepare(input,{variant:'general'}));
});
test('unknown or absent explicit variants produce a clear error instead of the wrong template',async()=>{
 const input=await fixture(p('通用版')+table('general only'));
 await assert.rejects(()=>prepare(input,{variant:'unknown'}),/版本/);
 await assert.rejects(()=>prepare(input,{variant:'ccb'}),/建行版本/);
 await assert.rejects(()=>prepare({...input,name:'not-word.pptx'}),/DOCX/);
});
test('version downloads do not retain notes belonging only to the other version',async()=>{
 const withNote=(text,id)=>p(text).replace('</w:p>',`<w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`);
 const input=await fixture(p('通用版本')+withNote('general body',1)+p('建行版本')+withNote('ccb body',2)),zip=await JSZip.loadAsync(input.bytes);
 zip.file('word/footnotes.xml',`<w:footnotes xmlns:w="${NS.w}"><w:footnote w:type="separator" w:id="-1">${p('')}</w:footnote><w:footnote w:id="1">${p('general source note')}</w:footnote><w:footnote w:id="2">${p('ccb source note')}</w:footnote></w:footnotes>`);
 input.bytes=await zip.generateAsync({type:'uint8array'});
 const output=await JSZip.loadAsync((await prepare(input,{variant:'ccb'})).bytes),xml=await output.file('word/footnotes.xml').async('string');
 assert.match(xml,/ccb source note/);assert.doesNotMatch(xml,/general source note/);assert.match(xml,/w:type="separator"/);
});
