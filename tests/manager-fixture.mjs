import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import {NS} from '../src/office-xml.js';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
export const run=text=>`<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r>`;
export const p=(text,extra='')=>`<w:p><w:pPr><w:spacing w:after="0"/><w:ind w:left="240"/></w:pPr>${run(text)}${extra}</w:p>`;
export const table=(...cells)=>`<w:tbl><w:tr>${cells.map(c=>`<w:tc>${c}</w:tc>`).join('')}</w:tr></w:tbl>`;
export async function fixture(body){
 const z=new JSZip();z.file('word/document.xml',`<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
 z.file('_rels/.rels',`<Relationships xmlns="${NS.pr}"><Relationship Id="r1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`);
 z.file('[Content_Types].xml',`<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
 return readTemplate({name:'【一页通202609】示例产品.docx',bytes:await z.generateAsync({type:'uint8array'})});
}
export const profile=(records)=>({schemaVersion:1,name:'经理资料.xlsx',sha256:'a'.repeat(64),records:records.map((r,i)=>({id:`row-${i+2}`,name:'张三',biography:'',workStart:null,investmentStart:null,sheet:'经理简介',row:i+2,...r}))});
export const task=t=>({facts:[],materials:[{materialId:t.id,kind:'word',template:t,operations:[],fields:[],warnings:[]}],warnings:[],stats:{}});
export const normalize=text=>text.replace(/[\s•·●◆\uf0d8]/g,'');
