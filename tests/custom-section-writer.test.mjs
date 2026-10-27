import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import * as engine from '../src/review-engine.js';
import {writeTask} from '../src/office-writer.js';
import {NS,elements,serializeXml} from '../src/office-xml.js';

globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {w:W,a:A,p:P,r:R,pr:PR,ct:CT}=NS;
const name='晨星稳健9号集合资产管理计划',note='其他介绍内容和排版保持原样。';
const old=['（1）固定收益类资产：占资产总值的20%至80%。','（2）权益类资产：占资产总值的0%至40%。','（3）公募基金：占资产总值的0%至80%。','（4）衍生品账户权益：占资产总值的0%至20%。'];
const replacement='（1）固定收益类资产：占资产总值的80%至100%。\n（2）货币市场工具：占资产总值的0%至20%。';
const parse=s=>new DOMParser().parseFromString(s,'application/xml');
const rels=items=>`<Relationships xmlns="${PR}">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

async function fixture(kind,label){
 const word=kind==='docx',zip=new JSZip(),part=word?'word/document.xml':'ppt/slides/slide1.xml';
 const run=text=>word?`<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r>`:`<a:r><a:rPr lang="zh-CN" sz="1400"><a:ea typeface="宋体"/></a:rPr><a:t>${text}</a:t></a:r>`;
 const paragraph=content=>word?`<w:p><w:pPr><w:spacing w:after="0"/><w:ind w:left="120"/></w:pPr>${content}</w:p>`:`<a:p><a:pPr marL="120" indent="0"/>${content}<a:endParaRPr lang="zh-CN"/></a:p>`;
 const cell=paragraphs=>word?`<w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${paragraphs}</w:tc>`:`<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr/></a:tc>`;
 const content=paragraph(run(old[0])+(word?'<w:r><w:br/></w:r>':'<a:br/>')+run(old[1]))+paragraph(run(old[2]))+paragraph(run(old[3]));
 const rows=[['产品名称',paragraph(run(name))],[label,content],['管理费',paragraph(run('管理费率0.6%/年'))]].map(([label,value])=>word?`<w:tr>${cell(paragraph(run(label)))}${cell(value)}</w:tr>`:`<a:tr h="914400">${cell(paragraph(run(label)))}${cell(value)}</a:tr>`).join('');
 const table=word?`<w:tbl><w:tblPr><w:tblW w:w="10000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>${rows}</w:tbl>`:`<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="7315200"/></a:tblGrid>${rows}</a:tbl>`;
 const xml=word?`<w:document xmlns:w="${W}"><w:body>${table}${paragraph(run(note))}<w:sectPr/></w:body></w:document>`:`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="标题"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraph(run('产品要素'))}</p:txBody></p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="产品要素表"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="200000"/><a:ext cx="8229600" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table">${table}</a:graphicData></a:graphic></p:graphicFrame><p:sp><p:nvSpPr><p:cNvPr id="3" name="说明"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraph(run(note))}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
 zip.file(part,xml);const overrides=[[part,word?'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']];
 if(word)zip.file('_rels/.rels',rels([['r1','officeDocument',part]]));
 else{zip.file('_rels/.rels',rels([['r1','officeDocument','ppt/presentation.xml']]));zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>`);zip.file('ppt/_rels/presentation.xml.rels',rels([['r1','slide','slides/slide1.xml']]));overrides.push(['ppt/presentation.xml','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml']);}
 zip.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([p,t])=>`<Override PartName="/${p}" ContentType="${t}"/>`).join('')}</Types>`);
 return{name:'【'+(word?'一页通':'PPT')+'202608】晨星稳健9号.'+kind,bytes:await zip.generateAsync({type:'uint8array'})};
}

for(const kind of ['docx','pptx'])for(const [label,key]of [['投资比例','investment_ratios'],['投资范围','investment_scope']])test(`${kind} custom ${key} replaces all old paragraphs, including inline breaks, with native reviewable text`,async()=>{
 const file=await fixture(kind,label),original=file.bytes.slice(),template=await readTemplate(file),ns=kind==='docx'?W:A;
 const row=template.rows.find(r=>r.label===label),ps=row.cells[1].paragraphIds.map(id=>template.paragraphs.find(p=>p.id===id));
 assert.equal(ps.length,3);assert.equal(ps[0].text,old[0]+'\n'+old[1]);
 const props=ps.map(p=>serializeXml(elements(p.node,ns,'pPr')[0]));
 const task=engine.analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'产品名称：'+name},{id:'m',text:'管理费率0.6%/年'}]},[template],{month:'202610'}),ids=task.facts.filter(f=>f.key===key).map(f=>f.id);
 assert.ok(ids.length>1,'multiple source clauses share one complete element');
 assert.equal(typeof engine.previewCustomSection,'function','custom whole replacement API must be available');
 const p=engine.previewCustomSection(task,ids,replacement);assert.equal(p.entries.length,1);assert.equal(p.entries[0].after,replacement);
 assert.equal(task.materials[0].fields.find(f=>f.label===label).after,old.join('\n'),'preview is read only');
 engine.applyCustomSection(task,p);
 const output=(await writeTask(task)).files[0];assert.equal(output.report.structuralValidation,true);
 const zip=await JSZip.loadAsync(output.bytes),doc=parse(await zip.file(ps[0].part).async('string'));
 if(kind==='docx'){
  assert.equal(elements(doc,W,'delText').map(t=>t.textContent).join(''),old.join(''),'all old clauses have real deletion revisions');
  assert.ok(elements(doc,W,'ins').length>0);assert.ok((await zip.file('word/settings.xml').async('string')).includes('trackRevisions'));
  assert.equal(elements(doc,W,'commentRangeStart').length,0);assert.equal(elements(doc,W,'highlight').length,0);
 }else{assert.ok(!Object.keys(zip.files).some(p=>/^ppt\/comments\//.test(p)));assert.equal(elements(doc,A,'highlight').length,0);}
 const reread=await readTemplate({name:output.name,bytes:output.bytes}),updated=reread.rows.find(r=>r.label===label),currentPs=updated.cells[1].paragraphIds.map(id=>reread.paragraphs.find(p=>p.id===id));
 assert.equal(updated.value.trim(),replacement);assert.ok(old.slice(2).every(t=>!updated.value.includes(t)),'omitted old third and fourth clauses are gone');
 assert.deepEqual(currentPs.map(p=>serializeXml(elements(p.node,ns,'pPr')[0])),props);
 assert.equal(reread.rows.find(r=>r.label==='产品名称').value,name);assert.equal(reread.rows.find(r=>r.label==='管理费').value,'管理费率0.6%/年');assert.ok(reread.paragraphs.some(p=>p.text===note));assert.equal(reread.rows.length,template.rows.length);assert.deepEqual(file.bytes,original);
});
