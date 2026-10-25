import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser, XMLSerializer} from '@xmldom/xmldom';
globalThis.JSZip = JSZip;
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const parse = xml => new DOMParser().parseFromString(xml, 'application/xml');
const els = (doc, ns, local) => Array.from(doc.getElementsByTagNameNS(ns, local));
const rels = items => `<Relationships xmlns="${PR}">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const types = overrides => `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([part,type])=>`<Override PartName="/${part}" ContentType="${type}"/>`).join('')}</Types>`;
const wr = (text,font='宋体') => `<w:r><w:rPr><w:rFonts w:eastAsia="${font}"/><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r>`;
const ar = text => `<a:r><a:rPr lang="zh-CN" sz="1600"><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r>`;
async function documentFixture({paragraphs=[wr('原产品，收益')+wr('3.','仿宋')+wr('80%')+wr('；评级R3；费率1%。')],oldComment=false}={}) {
 const zip=new JSZip();
 const part='word/document.xml';
 const xml=`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraphs.map(p=>`<w:p>${p}</w:p>`).join('')}<w:sectPr/></w:body></w:document>`;
 zip.file(part,xml); zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Normal"/></w:styles>`);
 zip.file('_rels/.rels',rels([['rId1','officeDocument',part]]));
 const relations=[['rId1','styles','styles.xml']];
 const overrides=[[part,'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml']];
 if(oldComment){zip.file('word/comments.xml',`<w:comments xmlns:w="${W}"><w:comment w:id="9" w:author="原作者"><w:p>${wr('旧批注')}</w:p></w:comment></w:comments>`);relations.push(['rId2','comments','comments.xml']);overrides.push(['word/comments.xml','application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml']);}
 zip.file('word/_rels/document.xml.rels',rels(relations));zip.file('[Content_Types].xml',types(overrides));
 const bytes=await zip.generateAsync({type:'uint8array'});const doc=parse(xml);
 const pars=els(doc,W,'p').map((node,index)=>({id:`p${index}`,part,index,node,text:els(node,W,'t').map(n=>n.textContent).join('')}));
 return {id:'test',name:'旧产品介绍202001.docx',kind:'docx',bytes,zip,parts:new Map([[part,doc]]),paragraphs:pars,slides:[],preparation:{changedParts:[],acceptedRevisions:0}};
}
test('unknown template names and duplicate output names never block valid Word files',async()=>{
 const {writeTask}=await import('../src/office-writer.js');
 const templates=await Promise.all([documentFixture({paragraphs:[wr('第一份模板内容')]}),documentFixture({paragraphs:[wr('第二份模板内容')]})]);
 templates[0].name='【一页通202608】未知2号X期.docx';templates[1].name='【一页通202608】别名3号.docx';
 const result=await writeTask({target:{name:'示例稳健8号集合资产管理计划'},month:'202610',materials:templates.map(template=>({template,kind:'word',templateProductName:'另一旧产品5号',operations:[],fields:[]}))});
 assert.deepEqual(result.files.map(f=>f.name),['【一页通202610】示例稳健8号.docx','【一页通202610】示例稳健8号（2）.docx']);
 assert.deepEqual(result.report.files.map(f=>f.name),result.files.map(f=>f.name));
 for(const [i,file]of result.files.entries()){
  const zip=await JSZip.loadAsync(file.bytes),xml=await zip.file('word/document.xml').async('string');
  assert.ok(xml.includes(i===0?'第一份模板内容':'第二份模板内容'));assert.equal(file.report.structuralValidation,true);
 }
});
async function presentationFixture({duplicateAuthors=false,usedDuplicate=false,linkExcluded=false}={}) {
 const zip=new JSZip(),parts=new Map(),paragraphs=[],slides=[];
 const presentation=`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst>${[1,2,3].map(i=>`<p:sldId id="${255+i}" r:id="rId${i}"/>`).join('')}</p:sldIdLst><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId4"/></p:sldMasterIdLst><p:custShowLst><p:custShow name="旧放映" id="1"><p:sldLst><p:sld r:id="rId1"/></p:sldLst></p:custShow></p:custShowLst></p:presentation>`;
 zip.file('ppt/presentation.xml',presentation);zip.file('_rels/.rels',rels([['rId1','officeDocument','ppt/presentation.xml']]));
 const presentationRels=[...[1,2,3].map(i=>[`rId${i}`,'slide',`slides/slide${i}.xml`]),['rId4','slideMaster','slideMasters/slideMaster1.xml']];
 const overrides=[['ppt/presentation.xml','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml']];
 for(let i=1;i<=3;i++){
   const part=`ppt/slides/slide${i}.xml`;
   const xml=`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="${i}" name="要素">${i===2&&linkExcluded?'<a:hlinkClick r:id="rId2" action="ppaction://hlinksldjump"/>':''}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="158750" y="317500"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${ar(`第${i}页；原产品；收益`)}${ar('3.80%')}${ar('；评级R3。')}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
   zip.file(part,xml);const doc=parse(xml);parts.set(part,doc);
   const node=els(doc,A,'p')[0];paragraphs.push({id:`p${i}`,part,index:0,node,page:i,text:els(node,A,'t').map(t=>t.textContent).join('')});slides.push({page:i,part,title:`第${i}页`,paragraphIds:[`p${i}`]});
   const rr=[['rId1','slideLayout','../slideLayouts/slideLayout1.xml']];
   if(i===2&&linkExcluded)rr.push(['rId2','slide','slide1.xml']);
   if(i===2&&usedDuplicate)rr.push(['rId3','comments','../comments/comment1.xml']);
   zip.file(`ppt/slides/_rels/slide${i}.xml.rels`,rels(rr));
   overrides.push([part,'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']);
 }
 zip.file('ppt/slideLayouts/slideLayout1.xml',`<p:sldLayout xmlns:p="${P}"/>`);zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels',rels([['rId1','slideMaster','../slideMasters/slideMaster1.xml']]));
 zip.file('ppt/slideMasters/slideMaster1.xml',`<p:sldMaster xmlns:p="${P}"/>`);zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels',rels([['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','image','../media/image1.png']]));zip.file('ppt/media/image1.png',new Uint8Array([1,2,3,4]));
 if(duplicateAuthors){zip.file('ppt/commentAuthors.xml',`<p:cmAuthorLst xmlns:p="${P}"><p:cmAuthor id="0" name="甲" initials="甲" lastIdx="1" clrIdx="0"/><p:cmAuthor id="0" name="乙" initials="乙" lastIdx="0" clrIdx="1"/></p:cmAuthorLst>`);presentationRels.push(['rId5','commentAuthors','commentAuthors.xml']);overrides.push(['ppt/commentAuthors.xml','application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml']);}
 if(usedDuplicate){zip.file('ppt/comments/comment1.xml',`<p:cmLst xmlns:p="${P}"><p:cm authorId="0" idx="1" dt="2020-01-01T00:00:00Z"><p:pos x="1" y="2"/><p:text>旧批注</p:text></p:cm></p:cmLst>`);overrides.push(['ppt/comments/comment1.xml','application/vnd.openxmlformats-officedocument.presentationml.comments+xml']);}
 zip.file('ppt/_rels/presentation.xml.rels',rels(presentationRels));zip.file('[Content_Types].xml',types(overrides));
 const bytes=await zip.generateAsync({type:'uint8array'});
 return {id:'test',name:'旧产品介绍202001.pptx',kind:'pptx',bytes,zip,parts,paragraphs,slides,preparation:{changedParts:[],acceptedRevisions:0}};
}
const operation=(p,kind,before,after,id=kind)=>({id,paragraphId:p.id,part:p.part,index:p.index,kind,start:p.text.indexOf(before),before,after,reason:'合同未列明，保留原值待核',factIds:[]});
const review=(template,operations,selectedPages=[])=>({materialId:template.id,template,kind:template.kind==='docx'?'word':'ppt',templateProductName:'旧产品',selectedPages,operations,fields:[],warnings:[]});
const options={month:'202610',targetName:'新产品',now:new Date('2026-10-25T00:00:00Z')};
async function writer(){return (await import('../src/office-writer.js')).writeMaterial;}

test('Word splits runs precisely, retains fonts, old comments and immutable input; offsets survive length changes',async()=>{
 const template=await documentFixture({oldComment:true});const original=template.bytes.slice();const p=template.paragraphs[0];
 const ops=[operation(p,'replace','原产品','目标产品甲','r1'),operation(p,'replace','3.80%','4.20%','r2'),operation(p,'highlight','R3',undefined,'h'),operation(p,'comment','R3','合同未载明评级，原值R3保留待核。','c')];
 const out=await(await writer())(review(template,ops),options);const zip=await JSZip.loadAsync(out.bytes);const doc=parse(await zip.file(p.part).async('string'));
 const visible=els(doc,W,'t').map(n=>n.textContent).join('');assert.equal(visible,'目标产品甲，收益4.20%；评级R3；费率1%。');
 assert.equal(els(doc,W,'delText').map(n=>n.textContent).join(''),'原产品3.80%');assert.ok(els(doc,W,'ins').length>=2);
 const hi=els(doc,W,'highlight');assert.equal(hi.length,1);assert.equal(hi[0].parentNode.parentNode.textContent,'R3');
 assert.equal(els(doc,W,'commentRangeStart').length,1);assert.equal(els(doc,W,'commentRangeEnd').length,1);assert.equal(els(doc,W,'commentReference').length,1);
 const comments=parse(await zip.file('word/comments.xml').async('string'));assert.equal(els(comments,W,'comment').length,2);assert.match(comments.documentElement.textContent,/旧批注/);
 const newRateRun=els(doc,W,'t').find(n=>n.textContent==='4.');assert.equal(els(newRateRun.parentNode,W,'rFonts')[0].getAttributeNS(W,'eastAsia'),'仿宋');
 assert.equal(await zip.file('word/styles.xml').async('string'),await template.zip.file('word/styles.xml').async('string'));
 assert.deepEqual(template.bytes,original);assert.equal(template.paragraphs[0].text,p.text);assert.match(out.name,/202610/);assert.equal(out.report.structuralValidation,true);
});

test('source-mismatch, overlapping edits and unsafe field text fail visibly',async()=>{
 const write=await writer();const t=await documentFixture(),p=t.paragraphs[0];
 await assert.rejects(write(review(t,[{...operation(p,'replace','原产品','新产品'),start:1}]),options),/原文|位置|匹配/);
 await assert.rejects(write(review(t,[operation(p,'replace','原产品','新产品','r1'),operation(p,'replace','产品','计划','r2')]),options),/重叠/);
 const f=await documentFixture({paragraphs:['<w:fldSimple w:instr="TEST">'+wr('原产品')+'</w:fldSimple>']});
 await assert.rejects(write(review(f,[operation(f.paragraphs[0],'replace','原产品','新产品')]),options),/域|结构/);
});

test('PPT emits real positioned comments and exact highlight; selected source pages remain ordered and dependencies resolve',async()=>{
 const t=await presentationFixture({duplicateAuthors:true});const p=t.paragraphs[1],source=t.bytes.slice();
 const ops=[operation(p,'replace','原产品','目标产品甲'),operation(p,'highlight','R3'),operation(p,'comment','R3','风险等级保留待核')];
 const out=await(await writer())(review(t,ops,[2,3]),options);const zip=await JSZip.loadAsync(out.bytes);
 assert.equal(zip.file('ppt/slides/slide1.xml'),null);assert.ok(zip.file('ppt/slides/slide2.xml'));assert.ok(zip.file('ppt/slides/slide3.xml'));
 const pres=parse(await zip.file('ppt/presentation.xml').async('string'));assert.deepEqual(els(pres,P,'sldId').map(n=>n.getAttributeNS(R,'id')),['rId2','rId3']);assert.equal(els(pres,P,'custShowLst').length,0);
 const doc=parse(await zip.file(p.part).async('string'));assert.match(els(doc,A,'t').map(n=>n.textContent).join(''),/目标产品甲/);
 const mark=els(doc,A,'highlight')[0];assert.equal(mark.parentNode.parentNode.textContent,'R3');assert.equal(els(mark,A,'srgbClr')[0].getAttribute('val'),'FFFF00');
 const commentFiles=Object.keys(zip.files).filter(n=>/^ppt\/comments\/.*\.xml$/.test(n));assert.equal(commentFiles.length,1);
 const comments=parse(await zip.file(commentFiles[0]).async('string'));assert.match(comments.documentElement.textContent,/风险等级保留待核/);assert.equal(els(comments,P,'pos')[0].getAttribute('x'),'100');
 const authors=parse(await zip.file('ppt/commentAuthors.xml').async('string'));const ids=els(authors,P,'cmAuthor').map(n=>n.getAttribute('id'));assert.equal(new Set(ids).size,ids.length);
 assert.deepEqual(await zip.file('ppt/media/image1.png').async('uint8array'),new Uint8Array([1,2,3,4]));assert.deepEqual(t.bytes,source);
 assert.equal(out.report.structuralValidation,true);assert.deepEqual(out.report.slideMapping.map(x=>x.sourcePage),[2,3]);assert.equal(out.report.repairs[0].removedCount,1);
});

test('PPT refuses ambiguous used duplicate authors',async()=>{
 const write=await writer();const t=await presentationFixture({duplicateAuthors:true,usedDuplicate:true});await assert.rejects(write(review(t,[],[2,3]),options),/批注.*作者|作者.*重复/);
});

test('PPT removes only navigation into excluded pages and preserves visible content and valid links',async()=>{
 const t=await presentationFixture({linkExcluded:true}),part='ppt/slides/slide2.xml',doc=t.parts.get(part);
 const hover=doc.createElementNS(A,'a:hlinkMouseOver');hover.setAttributeNS(R,'r:id','rId3');hover.setAttribute('action','ppaction://hlinksldjump');els(doc,P,'cNvPr')[0].appendChild(hover);
 t.zip.file('ppt/slides/_rels/slide2.xml.rels',rels([['rId1','slideLayout','../slideLayouts/slideLayout1.xml'],['rId2','slide','slide1.xml'],['rId3','slide','slide3.xml']]));
 t.zip.file(part,new XMLSerializer().serializeToString(doc));t.bytes=await t.zip.generateAsync({type:'uint8array'});
 const originalText=els(doc,A,'t').map(n=>n.textContent).join(''),out=await(await writer())(review(t,[],[2,3]),options),zip=await JSZip.loadAsync(out.bytes),output=parse(await zip.file(part).async('string'));
 assert.equal(els(output,A,'t').map(n=>n.textContent).join(''),originalText);assert.equal(els(output,A,'hlinkClick').length,0);assert.equal(els(output,A,'hlinkMouseOver')[0].getAttributeNS(R,'id'),'rId3');
 const outputRels=parse(await zip.file('ppt/slides/_rels/slide2.xml.rels').async('string'));assert.deepEqual(els(outputRels,PR,'Relationship').map(n=>n.getAttribute('Id')),['rId1','rId3']);
 assert.deepEqual(await zip.file('ppt/media/image1.png').async('uint8array'),new Uint8Array([1,2,3,4]));assert.equal(zip.file('ppt/slides/slide1.xml'),null);
 assert.ok(out.report.repairs.some(r=>r.kind==='removedSlideNavigation'&&r.part===part&&r.removedCount===1&&r.message));
});

test('PPT relative navigation is removed only when pages are trimmed; poster navigation remains',async()=>{
 const t=await presentationFixture(),part='ppt/slides/slide2.xml',doc=t.parts.get(part);
 for(const kind of ['hlinkClick','hlinkMouseOver']){const node=doc.createElementNS(A,'a:'+kind);node.setAttribute('action','ppaction://hlinkshowjump?jump=nextslide');els(doc,P,'cNvPr')[0].appendChild(node);}
 t.zip.file(part,new XMLSerializer().serializeToString(doc));t.bytes=await t.zip.generateAsync({type:'uint8array'});
 const write=await writer(),out=await write(review(t,[],[2,3]),options),zip=await JSZip.loadAsync(out.bytes),output=parse(await zip.file(part).async('string'));
 assert.equal(els(output,A,'hlinkClick').length,0);assert.equal(els(output,A,'hlinkMouseOver').length,0);assert.ok(out.report.repairs.some(r=>r.kind==='removedRelativeNavigation'&&r.removedCount===2&&r.message));
 const poster=await write({...review(t,[],[2]),kind:'poster'},options),posterZip=await JSZip.loadAsync(poster.bytes),posterDoc=parse(await posterZip.file(part).async('string'));
 assert.equal(els(posterDoc,A,'hlinkClick').length,1);assert.equal(els(posterDoc,A,'hlinkMouseOver').length,1);assert.ok(!poster.report.repairs.some(r=>r.kind==='removedRelativeNavigation'));
});

test('PPT refuses non-navigation dependencies on excluded pages',async()=>{
 const t=await presentationFixture({linkExcluded:true}),part='ppt/slides/slide2.xml',doc=t.parts.get(part);
 els(doc,P,'cNvPr')[0].setAttributeNS(R,'r:embed','rId2');t.zip.file(part,new XMLSerializer().serializeToString(doc));t.bytes=await t.zip.generateAsync({type:'uint8array'});
 await assert.rejects((await writer())(review(t,[],[2,3]),options),/排除|导航|依赖/);
});

test('poster exports every page without source-page trimming',async()=>{
 const t=await presentationFixture();const r={...review(t,[],[2]),kind:'poster'};const out=await(await writer())(r,options);const zip=await JSZip.loadAsync(out.bytes);assert.ok(zip.file('ppt/slides/slide1.xml'));assert.equal(out.report.outputSlideCount,3);
});

test('Word exact comments split a multi-text run at node boundaries and preserve existing anchors',async()=>{
 const t=await documentFixture({paragraphs:[`<w:commentRangeStart w:id="9"/><w:r><w:rPr><w:b/></w:rPr><w:t>甲</w:t><w:t>乙</w:t><w:t>丙</w:t></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r>`],oldComment:true});
 const p=t.paragraphs[0],out=await(await writer())(review(t,[operation(p,'comment','乙','仅乙待核')]),options),zip=await JSZip.loadAsync(out.bytes),doc=parse(await zip.file(p.part).async('string'));
 const start=els(doc,W,'commentRangeStart').find(n=>n.getAttributeNS(W,'id')!=='9'),end=els(doc,W,'commentRangeEnd').find(n=>n.getAttributeNS(W,'id')!=='9');
 let text='';for(let n=start.nextSibling;n&&n!==end;n=n.nextSibling)text+=n.textContent;assert.equal(text,'乙');
 assert.equal(els(doc,W,'commentRangeStart').filter(n=>n.getAttributeNS(W,'id')==='9').length,1);assert.equal(els(doc,W,'commentReference').filter(n=>n.getAttributeNS(W,'id')==='9').length,1);
});

test('Word comments covering a whole replacement use original offsets; later spans survive a deletion',async()=>{
 const t=await documentFixture({paragraphs:[wr('甲旧值乙留值丙删除丁待核。')]}),p=t.paragraphs[0];
 const ops=[operation(p,'replace','旧值','新的长值','r1'),operation(p,'comment','旧值','已按合同更新','c1'),operation(p,'replace','删除','','r2'),operation(p,'comment','删除','删除原文','c2'),operation(p,'highlight','待核',undefined,'h'),operation(p,'comment','待核','待核说明','c3')];
 const out=await(await writer())(review(t,ops),options),zip=await JSZip.loadAsync(out.bytes),doc=parse(await zip.file(p.part).async('string'));
 assert.equal(els(doc,W,'t').map(n=>n.textContent).join(''),'甲新的长值乙留值丙丁待核。');assert.equal(els(doc,W,'commentReference').length,3);
 assert.equal(els(doc,W,'highlight')[0].parentNode.parentNode.textContent,'待核');
});

test('Word editable hyperlink keeps its relationship and font after native revisions',async()=>{
 const t=await documentFixture({paragraphs:[wr('说明：')+`<w:hyperlink r:id="rId8">${wr('原产品','仿宋')}</w:hyperlink>`]});
 t.zip.file('word/_rels/document.xml.rels',`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId8" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`);t.bytes=await t.zip.generateAsync({type:'uint8array'});
 const out=await(await writer())(review(t,[operation(t.paragraphs[0],'replace','原产品','新产品')]),options),zip=await JSZip.loadAsync(out.bytes),doc=parse(await zip.file('word/document.xml').async('string'));
 const link=els(doc,W,'hyperlink')[0];assert.equal(link.getAttributeNS(R,'id'),'rId8');assert.equal(els(link,W,'t').map(n=>n.textContent).join(''),'新产品');assert.equal(els(link,W,'rFonts')[0].getAttributeNS(W,'eastAsia'),'仿宋');
});

test('Word a field in an unrelated part of the paragraph remains untouched',async()=>{
 const field='<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'+wr('1')+'<w:r><w:fldChar w:fldCharType="end"/></w:r>';
 const t=await documentFixture({paragraphs:[field+wr('；原产品')]}),p=t.paragraphs[0];
 const out=await(await writer())(review(t,[operation(p,'replace','原产品','新产品')]),options),zip=await JSZip.loadAsync(out.bytes),doc=parse(await zip.file(p.part).async('string'));
 assert.equal(els(doc,W,'fldChar').length,3);assert.equal(els(doc,W,'instrText')[0].textContent,'PAGE');assert.equal(els(doc,W,'t').map(n=>n.textContent).join(''),'1；新产品');
 await assert.rejects((await writer())(review(t,[operation(p,'replace','1','2')]),options),/域/);
});

test('PPT rejects reversed selected pages and dangling internal resources',async()=>{
 const t=await presentationFixture();const write=await writer();await assert.rejects(write(review(t,[],[3,2]),options),/顺序/);
 t.zip.remove('ppt/media/image1.png');t.bytes=await t.zip.generateAsync({type:'uint8array'});
 await assert.rejects(write(review(t,[],[2,3]),options),/缺失|失效/);
});

test('PPT duplicate author safety checks actual comments in producer-specific unlinked parts',async()=>{
 const t=await presentationFixture({duplicateAuthors:true});
 t.zip.file('ppt/auxiliary/exporter.xml',`<p:cmLst xmlns:p="${P}"><p:cm authorId="0" idx="1" dt="2020-01-01T00:00:00Z"><p:pos x="1" y="2"/><p:text>原批注</p:text></p:cm></p:cmLst>`);t.bytes=await t.zip.generateAsync({type:'uint8array'});
 await assert.rejects((await writer())(review(t,[],[2,3]),options),/批注.*作者|作者.*重复/);
});
