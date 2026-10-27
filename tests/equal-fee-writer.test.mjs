import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import {analyzeMaterials} from '../src/review-engine.js';
import {writeTask} from '../src/office-writer.js';
import {NS,elements,serializeXml} from '../src/office-xml.js';

globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {w:W,a:A,p:P,r:R,pr:PR,ct:CT}=NS;
const name='晨星稳健9号集合资产管理计划';
const first='参与金额＜300万元，参与费率为2.0%；';
const second='参与金额≥300万元，参与费率为2%。';
const replacement='参与费率为2%。';
const note='其他介绍内容和排版保持原样。';
const parse=s=>new DOMParser().parseFromString(s,'application/xml');
const rels=items=>`<Relationships xmlns="${PR}">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;

async function fixture(kind,{feeLabel='参与费率',feeFirst=first,feeSecond=second}={}){
 const word=kind==='docx',zip=new JSZip(),part=word?'word/document.xml':'ppt/slides/slide1.xml';
 const run=text=>word?`<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="20"/></w:rPr><w:t>${text}</w:t></w:r>`:`<a:r><a:rPr lang="zh-CN" sz="1400"><a:ea typeface="宋体"/></a:rPr><a:t>${text}</a:t></a:r>`;
 const paragraph=content=>word?`<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${content}</w:p>`:`<a:p><a:pPr marL="0" indent="0"/>${content}<a:endParaRPr lang="zh-CN"/></a:p>`;
 const cell=content=>word?`<w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${paragraph(content)}</w:tc>`:`<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph(content)}</a:txBody><a:tcPr/></a:tc>`;
 // Different runs and an explicit Office line break reproduce real template structure.
 const fee=run(feeFirst)+(feeSecond?(word?'<w:r><w:br/></w:r>':'<a:br/>')+run(feeSecond):'');
 const rows=[['产品名称',run(name)],[feeLabel,fee],['管理费',run('管理费率0.6%/年')]].map(([label,value])=>word?`<w:tr>${cell(run(label))}${cell(value)}</w:tr>`:`<a:tr h="914400">${cell(run(label))}${cell(value)}</a:tr>`).join('');
 const table=word?`<w:tbl><w:tblPr><w:tblW w:w="10000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/><w:gridCol w:w="5000"/></w:tblGrid>${rows}</w:tbl>`:`<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="914400"/><a:gridCol w="7315200"/></a:tblGrid>${rows}</a:tbl>`;
 const xml=word?`<w:document xmlns:w="${W}"><w:body>${table}${paragraph(run(note))}<w:sectPr/></w:body></w:document>`:`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="标题"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraph(run('产品要素'))}</p:txBody></p:sp><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="产品要素表"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="200000"/><a:ext cx="8229600" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table">${table}</a:graphicData></a:graphic></p:graphicFrame><p:sp><p:nvSpPr><p:cNvPr id="3" name="说明"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraph(run(note))}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
 zip.file(part,xml);
 const overrides=[[part,word?'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']];
 if(word)zip.file('_rels/.rels',rels([['r1','officeDocument',part]]));
 else{
  zip.file('_rels/.rels',rels([['r1','officeDocument','ppt/presentation.xml']]));
  zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels',rels([['r1','slide','slides/slide1.xml']]));
  overrides.push(['ppt/presentation.xml','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml']);
 }
 zip.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([p,t])=>`<Override PartName="/${p}" ContentType="${t}"/>`).join('')}</Types>`);
 return{name:'【'+(word?'一页通':'PPT')+'202608】晨星稳健9号.'+kind,bytes:await zip.generateAsync({type:'uint8array'})};
}

for(const kind of ['docx','pptx'])test(`${kind} equal participation tiers export one rate across inline breaks with formatting intact`,async()=>{
 const file=await fixture(kind),original=file.bytes.slice(),template=await readTemplate(file),ns=kind==='docx'?W:A;
 const feeRow=template.rows.find(r=>r.label==='参与费率'),feeParagraph=template.paragraphs.find(p=>p.id===feeRow.cells[1].paragraphIds[0]);
 assert.equal(feeRow.value,first+'\n'+second);
 const paragraphProperties=serializeXml(elements(feeParagraph.node,ns,'pPr')[0]);
 const task=analyzeMaterials({name:'合同.docx',warnings:[],blocks:[{id:'n',text:'产品名称：'+name},{id:'f',text:'参与费率为2%。'},{id:'m',text:'管理费率0.6%/年'}]},[template],{month:'202610'});
 assert.equal(task.materials[0].fields.find(f=>f.label==='参与费率').after,replacement,'the automatic review collapses equal fee tiers without manual input');
 const result=await writeTask(task);assert.equal(result.files.length,1);assert.equal(result.files[0].report.structuralValidation,true);
 const output=result.files[0],zip=await JSZip.loadAsync(output.bytes),doc=parse(await zip.file(feeParagraph.part).async('string'));
 if(kind==='docx'){
  assert.ok(elements(doc,W,'del').length>0,'removed tiers have real deletion revisions');
  assert.ok(elements(doc,W,'ins').length>0,'the collapsed fee has real insertion revisions');
  assert.equal(elements(doc,W,'delText').map(t=>t.textContent).join(''),first+second);
  assert.equal(elements(doc,W,'ins').flatMap(n=>elements(n,W,'t')).map(t=>t.textContent).join(''),replacement);
  assert.ok((await zip.file('word/settings.xml').async('string')).includes('trackRevisions'));
 }else assert.ok(!Object.keys(zip.files).some(p=>/^ppt\/comments\//.test(p)),'verified equivalent tiers do not introduce pending comments');
 const reread=await readTemplate({name:output.name,bytes:output.bytes}),updated=reread.rows.find(r=>r.label==='参与费率');
 assert.equal(updated.value,replacement);
 assert.equal(reread.rows.find(r=>r.label==='产品名称').value,name);
 assert.equal(reread.rows.find(r=>r.label==='管理费').value,'管理费率0.6%/年');
 assert.ok(reread.paragraphs.some(p=>p.text===note));
 const newParagraph=reread.paragraphs.find(p=>p.id===updated.cells[1].paragraphIds[0]);
 assert.equal(serializeXml(elements(newParagraph.node,ns,'pPr')[0]),paragraphProperties);
 assert.equal(reread.rows.length,template.rows.length);
 assert.deepEqual(file.bytes,original,'input files remain immutable');
});

for(const kind of ['docx','pptx'])test(`${kind} mixed fee labels export separately matched rates without collapsing distinct actions`,async()=>{
 const feeLabel='认购/参与/退出费率',file=await fixture(kind,{feeLabel,feeFirst:'认购费率3%；参与费率3%；退出费率3%。',feeSecond:''}),original=file.bytes.slice(),template=await readTemplate(file);
 const contract={name:'合同.docx',warnings:[],blocks:['产品名称：'+name,'认购费率2%；参与费率1%；退出费率0%。','管理费率0.6%/年'].map((text,i)=>({id:'b'+i,text}))};
 const task=analyzeMaterials(contract,[template],{month:'202610'}),expected='认购费率2%；参与费率1%；退出费率0%。';
 assert.equal(task.materials[0].fields.find(f=>f.label===feeLabel).after,expected);
 const result=await writeTask(task),output=result.files[0],reread=await readTemplate({name:output.name,bytes:output.bytes});
 assert.equal(reread.rows.find(r=>r.label===feeLabel).value,expected);
 assert.equal(reread.rows.find(r=>r.label==='管理费').value,'管理费率0.6%/年');
 assert.ok(reread.paragraphs.some(p=>p.text===note));assert.deepEqual(file.bytes,original);
 if(kind==='docx'){
  const zip=await JSZip.loadAsync(output.bytes),doc=parse(await zip.file('word/document.xml').async('string'));
  assert.ok(elements(doc,W,'del').length>=3);assert.ok(elements(doc,W,'ins').length>=3);
 }
});
