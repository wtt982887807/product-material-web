import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {readWordTemplate}=await import('../src/word-input.js');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function docx(){const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>产品名称：示例产品</w:t></w:r></w:p></w:body></w:document>`);z.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');return z.generateAsync({type:'uint8array'});}
const oldDoc=new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1,0,1]);
test('normal DOCX avoids the conversion engine and leaves input bytes unchanged',async()=>{const bytes=await docx(),original=bytes.slice();let called=false;const result=await readWordTemplate({name:'模板.docx',bytes},{convert:async()=>{called=true;}});assert.equal(called,false);assert.equal(result.name,'模板.docx');assert.equal(result.paragraphs[0].text,'产品名称：示例产品');assert.deepEqual(bytes,original);});
test('a DOC filename containing an actual DOCX package is recognized without conversion',async()=>{let called=false;const result=await readWordTemplate({name:'旧名.doc',bytes:await docx()},{convert:async()=>{called=true;}});assert.equal(called,false);assert.equal(result.name,'旧名.docx');assert.match(result.warnings.join(''),/实际.*DOCX|识别.*DOCX/);});
test('a binary DOC is converted once and the source digest stays available for duplicate checks',async()=>{let calls=0;const output=await docx();const result=await readWordTemplate({name:'老模板.doc',bytes:oldDoc},{convert:async input=>{calls++;assert.deepEqual(input.bytes,oldDoc);return output;}});assert.equal(calls,1);assert.equal(result.name,'老模板.docx');assert.ok(result.sourceSha256);assert.notEqual(result.sourceSha256,result.sha256);assert.match(result.warnings.join(''),/转换.*DOCX/);});
test('complex DOCX revisions automatically use the full engine then re-enter the normal reader',async()=>{const z=await JSZip.loadAsync(await docx());const xml=await z.file('word/document.xml').async('string');z.file('word/document.xml',xml.replace('<w:p>','<w:p><w:pPr><w:rPr><w:del w:id="8"/></w:rPr></w:pPr>'));let calls=0;const result=await readWordTemplate({name:'复杂.docx',bytes:await z.generateAsync({type:'uint8array'})},{convert:async()=>{calls++;return docx();}});assert.equal(calls,1);assert.equal(result.paragraphs[0].text,'产品名称：示例产品');assert.match(result.warnings.join(''),/旧修订/);});
test('arbitrary damaged inputs do not launch the full engine',async()=>{let calls=0;await assert.rejects(readWordTemplate({name:'坏文件.docx',bytes:new Uint8Array([1,2,3])},{convert:async()=>{calls++;return docx();}}),/识别|损坏|读取/);assert.equal(calls,0);});
test('a failed conversion remains an actionable failure and never masquerades as a DOCX',async()=>{await assert.rejects(readWordTemplate({name:'旧模板.doc',bytes:oldDoc},{convert:async()=>oldDoc}),/DOCX|转换/);});
test('a macro-enabled Word package cannot be relabeled as ordinary DOCX without conversion',async()=>{
 const zip=await JSZip.loadAsync(await docx());zip.file('[Content_Types].xml',(await zip.file('[Content_Types].xml').async('string')).replace('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml','application/vnd.ms-word.document.macroEnabled.main+xml'));zip.file('word/vbaProject.bin',new Uint8Array([1,2]));let calls=0;
 const result=await readWordTemplate({name:'带宏包.doc',bytes:await zip.generateAsync({type:'uint8array'})},{convert:async()=>{calls++;return docx();}});
 assert.equal(calls,1);assert.equal(result.zip.file('word/vbaProject.bin'),null);
});
