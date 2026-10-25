import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
Object.assign(globalThis,{JSZip,DOMParser,XMLSerializer});
import {readTemplate} from '../src/office-reader.js';
import {writeTask} from '../src/office-writer.js';

async function material(name,text){
 const z=new JSZip();
 z.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
 z.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
 z.file('word/document.xml',`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`);
 return {template:await readTemplate({name,bytes:await z.generateAsync({type:'uint8array'})}),kind:'word',operations:[],fields:[]};
}
test('one invalid material is isolated; other deliveries keep preview names and remain readable',async()=>{
 const first=await material('【一页通202608】旧2号.docx','不应交付的第一份'),second=await material('【一页通202608】旧3号.docx','应当交付的第二份');
 first.template.sha256='changed-after-review';
 const result=await writeTask({target:{name:'示例8号集合资产管理计划'},month:'202610',materials:[first,second]});
 assert.equal(result.files.length,1);assert.equal(result.files[0].name,'【一页通202610】示例8号（2）.docx');
 assert.equal(result.failures[0].sourceName,first.template.name);assert.match(result.failures[0].message,/发生变化/);
 assert.equal(result.report.complete,false);assert.equal(result.report.requestedCount,2);
 const zip=await JSZip.loadAsync(result.files[0].bytes);assert.match(await zip.file('word/document.xml').async('string'),/应当交付的第二份/);
});
test('when every material is invalid no successful empty package is reported',async()=>{
 const m=await material('示例.docx','原文');m.template.sha256='invalid';
 await assert.rejects(()=>writeTask({target:{name:'示例8号'},month:'202610',materials:[m]}),/没有物料成功生成.*示例\.docx.*发生变化/s);
});
test('unlocated product elements cannot become a renamed successful document',async()=>{
 const bad=await material('旧产品宣传.docx','旧产品2号的广告原文'),good=await material('有效模板.docx','有效要素');bad.productElementsLocated=false;
 const result=await writeTask({target:{name:'示例8号'},month:'202610',materials:[bad,good]});
 assert.equal(result.files.length,1);assert.match(result.failures[0].message,/未识别.*产品要素/);assert.equal(result.report.businessVerified,false);
});
