import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {readTemplate} from '../src/office-reader.js';
import {writeMaterial,writeTask} from '../src/office-writer.js';
import {NS,elements,children,parseXml,serializeXml,textOf} from '../src/office-xml.js';
import {fixture as wordFixture,p as wordParagraph} from './manager-fixture.mjs';

globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {a:A,p:P,r:R,pr:PR,ct:CT}=NS;
const properties='<a:pPr marL="120" indent="0" algn="just"><a:lnSpc><a:spcPct val="110000"/></a:lnSpc></a:pPr>';
const run=(text,bold=false)=>`<a:r><a:rPr lang="zh-CN" sz="1600" b="${bold?'1':'0'}"><a:ea typeface="宋体"/></a:rPr><a:t>${text}</a:t></a:r>`;
const paragraph=(text,content)=>`<a:p>${properties}${content??run(text)}<a:endParaRPr lang="zh-CN" sz="1600"/></a:p>`;
const shape=(id,content)=>`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="介绍${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="8000000" cy="3000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${content}</p:txBody></p:sp>`;
const rels=items=>`<Relationships xmlns="${PR}">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
async function fixture(contents=[shape(1,paragraph('公司：旧公司。',run('公司：',true)+run('旧公司。'))+paragraph('旧介绍补充。'))+shape(2,paragraph('张三：投资经理。'))]){
 const zip=new JSZip(),overrides=[['ppt/presentation.xml','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml']];
 zip.file('_rels/.rels',rels([['r1','officeDocument','ppt/presentation.xml']]));
 zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst>${contents.map((_,i)=>`<p:sldId id="${256+i}" r:id="r${i+1}"/>`).join('')}</p:sldIdLst></p:presentation>`);
 zip.file('ppt/_rels/presentation.xml.rels',rels(contents.map((_,i)=>[`r${i+1}`,'slide',`slides/slide${i+1}.xml`])));
 for(const [i,content] of contents.entries()){
  const part=`ppt/slides/slide${i+1}.xml`;
  const photo='<p:pic><p:nvPicPr><p:cNvPr id="20" name="经理照片"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="photo"/></p:blipFill><p:spPr/></p:pic>';
  zip.file(part,`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>${content}${photo}</p:spTree></p:cSld></p:sld>`);
  zip.file(`ppt/slides/_rels/slide${i+1}.xml.rels`,rels([['photo','image','../media/photo.png'],['layout','slideLayout','../slideLayouts/layout1.xml']]));
  overrides.push([part,'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']);
 }
 zip.file('ppt/media/photo.png',new Uint8Array([137,80,78,71,13,10,26,10]));
 zip.file('ppt/slideLayouts/layout1.xml',`<p:sldLayout xmlns:p="${P}"/>`);
 zip.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.map(([part,type])=>`<Override PartName="/${part}" ContentType="${type}"/>`).join('')}</Types>`);
 return readTemplate({name:'【PPT202609】示例甲产品.pptx',bytes:await zip.generateAsync({type:'uint8array'})});
}
function operation(p,after,source='company',extra={}){return{id:`${source}-${p.id}`,kind:'replace',source,paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after,profilePpt:true,...(after.includes('\n')?{profileParagraphs:true}:{}),factIds:[`fact-${source}`],...extra};}
function review(template,operations=[],extra={}){
 const fields=['company','manager'].map(source=>({id:`field-${source}`,source,status:'verified',factIds:[`fact-${source}`],paragraphIds:operations.filter(o=>o.source===source).map(o=>o.paragraphId)})).filter(f=>f.paragraphIds.length);
 return{template,materialId:template.id,kind:'ppt',templateProductName:'示例甲产品',productElementsLocated:true,selectedPages:[1],profilePages:{company:[1],manager:[1]},fields,operations,...extra};
}
const options={month:'202610',targetName:'新产品',now:new Date('2026-10-29T00:00:00Z')};
async function output(material,opts=options){const file=await writeMaterial(material,opts),zip=await JSZip.loadAsync(file.bytes);return{file,zip,doc:parseXml(await zip.file('ppt/slides/slide1.xml').async('string'))};}

test('PPT profile multiline replacements create native paragraphs with inherited style and preserve photos and relationships',async()=>{
 const t=await fixture(),before=t.bytes.slice(),p=t.paragraphs[0],last=t.paragraphs[2];
 const oldPhoto=serializeXml(elements(t.parts.get(p.part),P,'pic')[0]),oldShape=serializeXml(elements(t.parts.get(p.part),P,'spPr')[0]),oldProperties=serializeXml(children(p.node,A,'pPr')[0]);
 const o=await output(review(t,[operation(p,'公司：新公司。\n公司：完整新介绍。'),operation(t.paragraphs[1],'','company',{profileDeleteParagraph:true}),operation(last,'张三：高级投资经理。','manager')]));
 assert.deepEqual(elements(o.doc,A,'p').map(p=>textOf(p,'pptx')),['公司：新公司。','公司：完整新介绍。','张三：高级投资经理。']);
 assert.equal(elements(o.doc,A,'br').length,0);assert.ok(elements(o.doc,A,'t').every(n=>!n.textContent.includes('\n')));
 for(const p of elements(o.doc,A,'p')){assert.equal(serializeXml(children(p,A,'pPr')[0]),oldProperties);assert.equal(children(p,A,'endParaRPr')[0].getAttribute('sz'),'1600');}
 const inserted=elements(o.doc,A,'p')[1];assert.equal(children(children(inserted,A,'r')[0],A,'rPr')[0].getAttribute('b'),'1');assert.equal(children(children(inserted,A,'r').at(-1),A,'rPr')[0].getAttribute('b'),'0');
 assert.equal(serializeXml(elements(o.doc,P,'pic')[0]),oldPhoto);assert.equal(serializeXml(elements(o.doc,P,'spPr')[0]),oldShape);
 for(const part of ['ppt/slides/_rels/slide1.xml.rels','ppt/media/photo.png','ppt/slideLayouts/layout1.xml'])assert.deepEqual(await o.zip.file(part).async('uint8array'),await t.zip.file(part).async('uint8array'));
 assert.deepEqual(t.bytes,before);assert.equal(o.file.report.structuralValidation,true);
});

test('introduction-only PPT accepts no product facts or target name and retains only recognized pages',async()=>{
 const t=await fixture([shape(1,paragraph('公司旧介绍。')),shape(2,paragraph('无关团队获奖页。'))]),op=operation(t.paragraphs[0],'公司新介绍。');
 const o=await output(review(t,[op],{updateMode:'introductions',productElementsLocated:false,profilePages:{company:[1],manager:[]}}),{month:'202610'});
 assert.match(o.file.name,/示例甲产品/);assert.match(o.file.name,/202610/);assert.equal(o.zip.file('ppt/slides/slide2.xml'),null);assert.deepEqual(o.file.report.slideMapping.map(p=>p.sourcePage),[1]);
 assert.equal(textOf(elements(o.doc,A,'p')[0],'pptx'),'公司新介绍。');
});

test('matched but unchanged introductions deliver and batch preview names keep each template identity',async()=>{
 const t=await fixture(),m=review(t,[],{updateMode:'introductions',productElementsLocated:false,fields:[{id:'matched',source:'company',status:'verified',factIds:['fact-company'],paragraphIds:[t.paragraphs[0].id]}],profilePages:{company:[1],manager:[]}});
 const t2={...t,id:'other',name:'【PPT202609】示例乙产品.pptx'},result=await writeTask({month:'202610',materials:[m,{...m,template:t2},m]});
 assert.equal(result.files.length,3);assert.match(result.files[0].name,/示例甲产品/);assert.match(result.files[1].name,/示例乙产品/);assert.notEqual(result.files[0].name,result.files[2].name);assert.ok(result.files.every(f=>f.report.operations.length===0));
});

test('explicit profile paragraph deletion removes empty and obsolete paragraphs but retains the final valid txBody paragraph',async()=>{
 const t=await fixture([shape(1,paragraph('保留。')+paragraph('', '')+paragraph('删除。'))+shape(2,paragraph('删除最后段。'))]);
 const ops=[operation(t.paragraphs[1],'','company',{profileDeleteParagraph:true}),operation(t.paragraphs[2],'','company',{profileDeleteParagraph:true}),operation(t.paragraphs[3],'','manager',{profileDeleteParagraph:true})];
 const o=await output(review(t,ops)),bodies=elements(o.doc,P,'txBody');
 assert.deepEqual(bodies.map(b=>children(b,A,'p').length),[1,1]);assert.deepEqual(bodies.map(b=>children(b,A,'p').map(p=>textOf(p,'pptx'))),[['保留。'],['']]);
 assert.equal(serializeXml(children(children(bodies[1],A,'p')[0],A,'pPr')[0]),serializeXml(children(t.paragraphs[3].node,A,'pPr')[0]));
});

test('an unmarked empty replacement keeps its original paragraph instead of deleting layout',async()=>{
 const t=await fixture(),o=await output(review(t,[operation(t.paragraphs[1],'')]));
 assert.equal(elements(o.doc,A,'p').length,3);assert.equal(textOf(elements(o.doc,A,'p')[1],'pptx'),'');
});

test('company text inside a native PPT table becomes separate paragraphs without changing cell formatting',async()=>{
 const table=`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="1" name="公司信息表"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100" y="200"/><a:ext cx="8000000" cy="3000000"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="8000000"/></a:tblGrid><a:tr h="3000000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraph('公司旧介绍。')}</a:txBody><a:tcPr marL="12700" marR="12700" anchor="ctr"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
 const t=await fixture([table]),op=operation(t.paragraphs[0],'公司新介绍。\n公司补充介绍。'),o=await output(review(t,[op]));
 assert.deepEqual(children(elements(o.doc,A,'txBody')[0],A,'p').map(p=>textOf(p,'pptx')),['公司新介绍。','公司补充介绍。']);
 assert.equal(serializeXml(elements(o.doc,A,'tcPr')[0]),serializeXml(elements(t.parts.get(op.part),A,'tcPr')[0]));
});

test('a PPT profile marker never allows fields or special paragraph structures to be overwritten',async()=>{
 const t=await fixture([shape(1,paragraph('',`<a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:rPr/><a:t>1</a:t></a:fld>`))]);
 await assert.rejects(writeMaterial(review(t,[operation(t.paragraphs[0],'公司新介绍。')]),options),/特殊.*结构|普通文字|非普通文本/);
});

test('introduction-only guards reject product operations, unrelated paragraphs and mismatched evidence',async()=>{
 const t=await fixture(),op=operation(t.paragraphs[0],'新公司。'),valid=review(t,[op],{updateMode:'introductions',productElementsLocated:false});
 for(const patch of [
  {operations:[{...op,source:'contract'}]},
  {operations:[{...op,paragraphId:t.paragraphs[2].id,part:t.paragraphs[2].part,index:t.paragraphs[2].index,before:t.paragraphs[2].text}]},
  {operations:[{...op,factIds:['product-fact']}]},
  {operations:[{...op,factIds:[]}]},
  {fields:[...valid.fields,{id:'product',source:'contract',paragraphIds:[t.paragraphs[2].id]}]},
  {fields:[]},
  {profilePages:{}},
  {profilePages:{manager:[1],company:[]}}
 ])await assert.rejects(writeMaterial({...valid,...patch},{month:'202610'}),/介绍|来源|证据|匹配|识别|页面/);
});

test('profile native paragraph flags cannot authorize another source, partial structural edits or malformed deletion',async()=>{
 const t=await fixture(),p=t.paragraphs[0],op=operation(p,'新公司。');
 for(const patch of [{source:'contract'},{source:'user_custom_section'},{kind:'comment'},{profileParagraphs:true,start:1,before:p.text.slice(1)},{profileParagraphs:true,before:'',after:'新增'},{profileDeleteParagraph:true,after:'非空'},{after:'甲\t乙'},{after:'甲\r乙'},{after:'甲\n乙'}])await assert.rejects(writeMaterial(review(t,[{...op,...patch}]),options),/介绍|整段|来源|位置|分段|制表|删除/);
 const w=await wordFixture(wordParagraph('公司旧介绍。')),wp=w.paragraphs[0];
 await assert.rejects(writeMaterial({...review(w,[operation(wp,'公司新介绍。')]),kind:'word'},options),/介绍|PPT/);
});

test('a profile whole paragraph operation cannot mix with another operation on that paragraph',async()=>{
 const t=await fixture(),p=t.paragraphs[0],op=operation(p,'新公司。\n新增内容。','company',{profileParagraphs:true}),comment={...op,id:'comment',profilePpt:undefined,profileParagraphs:undefined,kind:'comment',source:'contract',after:'附加批注'};
 await assert.rejects(writeMaterial(review(t,[op,comment]),options),/同段|混用|介绍/);
});

test('PPT profile span differences preserve untouched runs and paragraph formatting byte-for-byte',async()=>{
 const t=await fixture([shape(1,paragraph('',run('14年证券从业经验。',true)+run('现任投资经理。')+run('说明：保持原有格式。',true)))]),p=t.paragraphs[0],unchanged=serializeXml(children(p.node,A,'r')[2]),props=serializeXml(children(p.node,A,'pPr')[0]);
 const ops=[operation(p,'15','manager',{id:'years',start:0,before:'14'}),operation(p,'高级','manager',{id:'insert-title',start:p.text.indexOf('投资经理'),before:''})];
 const o=await output(review(t,ops)),after=elements(o.doc,A,'p')[0];
 assert.equal(textOf(after,'pptx'),'15年证券从业经验。现任高级投资经理。说明：保持原有格式。');
 assert.equal(serializeXml(children(after,A,'pPr')[0]),props);assert.equal(serializeXml(children(after,A,'r').at(-1)),unchanged);
 assert.equal(elements(o.doc,A,'p').length,1);assert.equal(elements(o.doc,A,'br').length,0);
 const inserted=children(after,A,'r').find(r=>r.textContent.includes('高级'));assert.equal(children(inserted,A,'rPr')[0].getAttribute('b'),'0');
});

test('profile insertion at replacement boundaries uses original offsets and preserves later styled text',async()=>{
 const t=await fixture([shape(1,paragraph('',run('旧值')+run('原样保留',true)))]),p=t.paragraphs[0];
 const ops=[operation(p,'前','company',{id:'first-insert',start:0,before:''}),operation(p,'新值','company',{id:'replace-old',start:0,before:'旧值'}),operation(p,'后','company',{id:'end-insert',start:2,before:''})];
 const o=await output(review(t,ops));assert.equal(textOf(elements(o.doc,A,'p')[0],'pptx'),'前新值后原样保留');
 const oldTail=serializeXml(children(p.node,A,'r')[1]);assert.equal(serializeXml(children(elements(o.doc,A,'p')[0],A,'r').at(-1)),oldTail);
});

test('profile insertion rejects duplicate anchors, empty text frames and field anchors without changing other sources',async()=>{
 const t=await fixture(),p=t.paragraphs[0],insert=operation(p,'新增','manager',{before:'',start:0});
 await assert.rejects(writeMaterial(review(t,[insert,{...insert,id:'duplicate'}]),options),/同一位置|重叠|合并/);
 await assert.rejects(writeMaterial(review(t,[{...insert,source:'contract',profilePpt:undefined}]),options),/原文位置/);
 const empty=await fixture([shape(1,paragraph('', ''))]);await assert.rejects(writeMaterial(review(empty,[operation(empty.paragraphs[0],'新增','manager',{before:'',start:0})]),options),/样式|文字|空/);
 const field=await fixture([shape(1,paragraph('',`<a:fld id="{11111111-1111-1111-1111-111111111111}" type="slidenum"><a:rPr/><a:t>1</a:t></a:fld>`))]);
 await assert.rejects(writeMaterial(review(field,[operation(field.paragraphs[0],'新增','manager',{before:'',start:0})]),options),/样式|域|结构/);
});

test('profile changes cannot target pages outside delivery selection in full or introduction modes',async()=>{
 const t=await fixture([shape(1,paragraph('公司旧介绍。')),shape(2,paragraph('第二页旧介绍。'))]),op=operation(t.paragraphs[1],'第二页新介绍。');
 for(const updateMode of [undefined,'introductions'])await assert.rejects(writeMaterial(review(t,[op],{updateMode,profilePages:{company:[2]},selectedPages:[1]}),options),/非交付|页面|介绍/);
});

test('introduction-only selection is also respected for a poster material',async()=>{
 const t=await fixture([shape(1,paragraph('公司旧介绍。')),shape(2,paragraph('无关产品内容。'))]),op=operation(t.paragraphs[0],'公司新介绍。');
 const o=await output(review(t,[op],{kind:'poster',updateMode:'introductions',profilePages:{company:[1],manager:[]}}),{month:'202610'});
 assert.equal(o.zip.file('ppt/slides/slide2.xml'),null);assert.equal(o.file.report.outputSlideCount,1);
});

test('full mode still refuses unlocated product templates in Word and PPT',async()=>{
 const ppt=await fixture(),word=await wordFixture(wordParagraph('公司简介。'));
 for(const t of [ppt,word])await assert.rejects(writeMaterial({...review(t),kind:t.kind==='pptx'?'ppt':'word',productElementsLocated:false},options),/产品要素/);
});

test('introduction-only Word retains native manager revisions without adopting PPT flags',async()=>{
 const t=await wordFixture(wordParagraph('张三：9年证券从业经验。')),p=t.paragraphs[0],op={id:'manager-word',kind:'replace',source:'manager',paragraphId:p.id,part:p.part,index:p.index,start:3,before:'9',after:'10',factIds:['fact-manager']};
 const m=review(t,[op],{kind:'word',updateMode:'introductions',productElementsLocated:false}),file=await writeMaterial(m,{month:'202610'}),zip=await JSZip.loadAsync(file.bytes),doc=parseXml(await zip.file('word/document.xml').async('string'));
 assert.equal(elements(doc,NS.w,'delText').map(n=>n.textContent).join(''),'9');assert.equal(elements(doc,NS.w,'ins').map(n=>n.textContent).join(''),'10');assert.match(file.name,/示例产品/);
});
