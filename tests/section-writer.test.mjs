import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
globalThis.JSZip=JSZip;globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const {NS,elements,textOf,serializeXml}=await import('../src/office-xml.js');
const {writeMaterial}=await import('../src/office-writer.js');
const {w:W,a:A,p:P,r:R,pr:PR,ct:CT}=NS;
const parse=s=>new DOMParser().parseFromString(s,'application/xml');
const rels=items=>`<Relationships xmlns="${PR}">${items.map(([id,type,target])=>`<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const wr=(text,bold=false,font='宋体')=>`<w:r><w:rPr><w:rFonts w:eastAsia="${font}"/><w:sz w:val="20"/>${bold?'<w:b/>':''}</w:rPr><w:t>${text}</w:t></w:r>`;
const ar=(text,bold=false,font='宋体')=>`<a:r><a:rPr sz="1400" b="${bold?'1':'0'}"><a:ea typeface="${font}"/></a:rPr><a:t>${text}</a:t></a:r>`;
async function fixture(kind,{content,second=false}={}){
 const zip=new JSZip(),word=kind==='docx',part=word?'word/document.xml':'ppt/slides/slide1.xml';
 const source=content??(word?wr('（1）',true,'黑体')+wr('固定收益：',true)+wr('原值80%。')+'<w:r><w:rPr/><w:br/></w:r>'+wr('（2）',true,'黑体')+wr('权益：',true)+wr('原值20%。'):ar('（1）',true,'黑体')+ar('固定收益：',true)+ar('原值80%。')+'<a:br><a:rPr sz="1400"/></a:br>'+ar('（2）',true,'黑体')+ar('权益：',true)+ar('原值20%。'));
 const ppr=word?'<w:pPr><w:spacing w:after="0"/><w:ind w:left="120"/></w:pPr>':'<a:pPr marL="120" indent="0"><a:lnSpc><a:spcPct val="110000"/></a:lnSpc></a:pPr>';
 const ps=`<${word?'w':'a'}:p>${ppr}${source}${word?'':'<a:endParaRPr lang="zh-CN"/>'}</${word?'w':'a'}:p>`+(second?`<${word?'w':'a'}:p>${ppr}${word?wr('旧续段'):ar('旧续段')}</${word?'w':'a'}:p>`:'');
 const xml=word?`<w:document xmlns:w="${W}"><w:body>${ps}<w:sectPr/></w:body></w:document>`:`<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="1" name="投资比例"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="4000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${ps}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
 const overrides=[[part,word?'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']];
 zip.file(part,xml);
 if(word)zip.file('_rels/.rels',rels([['r1','officeDocument',part]]));
 else{
  zip.file('_rels/.rels',rels([['r1','officeDocument','ppt/presentation.xml']]));
  zip.file('ppt/presentation.xml',`<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="r1"/></p:sldIdLst></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels',rels([['r1','slide','slides/slide1.xml']]));
  overrides.push(['ppt/presentation.xml','application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml']);
 }
 zip.file('[Content_Types].xml',`<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${overrides.map(([p,t])=>`<Override PartName="/${p}" ContentType="${t}"/>`).join('')}</Types>`);
 const bytes=await zip.generateAsync({type:'uint8array'}),doc=parse(xml),paragraphs=elements(doc,word?W:A,'p').map((node,index)=>({id:'p'+index,part,index,node,page:word?null:1,text:textOf(node,kind)}));
 return{id:'section-template',name:'旧产品202608.'+kind,kind,bytes,zip,parts:new Map([[part,doc]]),paragraphs,slides:word?[]:[{page:1,part,title:'产品要素'}],preparation:{changedParts:[],acceptedRevisions:0}};
}
const op=(p,after,id='section-op')=>({id,kind:'replace',source:'contract_section',paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after,factIds:[]});
async function write(t,ops){const out=await writeMaterial({template:t,kind:t.kind==='docx'?'word':'ppt',operations:ops,fields:[],selectedPages:[1]},{month:'202610',targetName:'新产品'}),zip=await JSZip.loadAsync(out.bytes);return{out,zip,doc:parse(await zip.file(t.paragraphs[0].part).async('string'))};}
const updated='（1）固定收益资产：占总资产的80%至100%。\n（2）权益类资产：占总资产的0%至20%。\n（3）商品资产：不得投资。';
for(const kind of ['docx','pptx'])test(`${kind} section replacement preserves paragraph and title styles and uses native line breaks`,async()=>{
 const t=await fixture(kind),original=t.bytes.slice(),ns=kind==='docx'?W:A,ppr=serializeXml(elements(t.paragraphs[0].node,ns,'pPr')[0]);
 const {doc,zip}=await write(t,[op(t.paragraphs[0],updated)]),p=elements(doc,ns,'p')[0];
 assert.equal(textOf(p,kind),updated);assert.equal(serializeXml(elements(p,ns,'pPr')[0]),ppr);assert.deepEqual(t.bytes,original);
 assert.ok(elements(p,ns,'t').every(t=>!t.textContent.includes('\n')));
 const scope=kind==='docx'?elements(p,W,'ins')[0]:p;
 assert.equal(elements(scope,ns,'br').length,2);
 const titles=elements(scope,ns,'t').filter(t=>t.textContent.includes('固定收益资产')||t.textContent.includes('权益类资产'));
 assert.equal(titles.length,2);assert.ok(titles.every(t=>kind==='docx'?elements(t.parentNode,W,'b').length:elements(t.parentNode,A,'rPr')[0]?.getAttribute('b')==='1'));
 const bodies=elements(scope,ns,'t').filter(t=>t.textContent.includes('占总资产'));assert.equal(bodies.length,2);assert.ok(bodies.every(t=>kind==='docx'?!elements(t.parentNode,W,'b').length:elements(t.parentNode,A,'rPr')[0]?.getAttribute('b')!=='1'));
 if(kind==='docx'){
  assert.equal(elements(p,W,'del').length,1);assert.equal(elements(p,W,'ins').length,1);
  const deleted=elements(p,W,'del')[0].cloneNode(true);for(const n of elements(deleted,W,'delText')){const visible=doc.createElementNS(W,'w:t');visible.textContent=n.textContent;n.parentNode.replaceChild(visible,n);}const holder=doc.createElementNS(W,'w:p');while(deleted.firstChild)holder.appendChild(deleted.firstChild);assert.equal(textOf(holder),t.paragraphs[0].text);
  assert.ok((await zip.file('word/settings.xml').async('string')).includes('trackRevisions'));
  const {readTemplate}=await import('../src/office-reader.js'),accepted=await readTemplate({name:'更新.docx',bytes:await zip.generateAsync({type:'uint8array'})});assert.equal(accepted.paragraphs[0].text,updated);assert.equal(elements(accepted.paragraphs[0].node,W,'del').length,0);
 }else assert.equal(elements(doc,P,'sp').length,1);
});
for(const kind of ['docx','pptx'])test(`${kind} section clearing keeps the original continuation paragraph and its properties`,async()=>{
 const t=await fixture(kind,{second:true}),ns=kind==='docx'?W:A,{doc}=await write(t,[op(t.paragraphs[0],updated),op(t.paragraphs[1],'','clear')]),ps=elements(doc,ns,'p');
 assert.equal(ps.length,2);assert.equal(textOf(ps[0],kind),updated);assert.equal(textOf(ps[1],kind),'');assert.equal(elements(ps[1],ns,'pPr').length,1);
 if(kind==='docx')assert.equal(elements(ps[1],W,'delText')[0].textContent,'旧续段');
});
test('section edits reject partial ranges, stale originals, mixed edits and non-text structures',async()=>{
 const t=await fixture('docx'),p=t.paragraphs[0],base=op(p,updated);
 await assert.rejects(write(t,[{...base,start:1,before:p.text.slice(1)}]),/整段|原文|位置/);
 await assert.rejects(write(t,[{...base,before:p.text+'旧'}]),/整段|原文|匹配/);
 for(const kind of ['replace','comment','highlight'])await assert.rejects(write(t,[base,{...base,id:'other',source:undefined,kind,start:p.text.indexOf('80%'),before:'80%',after:kind==='replace'?'90%':'核对'}]),/整段|混用|重叠/);
 for(const text of ['新值\t附注','新值\r附注'])await assert.rejects(write(t,[op(p,text)]),/分段|制表|换行|结构/);
 const unsafe=await fixture('docx',{content:'<w:fldSimple w:instr="TEST">'+wr('旧值')+'</w:fldSimple>'});await assert.rejects(write(unsafe,[op(unsafe.paragraphs[0],'新值\n其他')]),/结构|域/);
});
test('ordinary replacement still refuses multiline content and leaves ordinary scalar writes unchanged',async()=>{
 const t=await fixture('docx',{content:wr('旧值80%。')}),p=t.paragraphs[0];
 await assert.rejects(write(t,[{...op(p,'新值\n附注'),source:undefined}]),/分段|制表/);
 const {doc}=await write(t,[{...op(p,'90%'),source:undefined,start:2,before:'80%'}]);assert.equal(textOf(elements(doc,W,'p')[0]),'旧值90%。');
});
test('PPT paragraph fields are rejected while explicit blank lines remain native breaks',async()=>{
 const t=await fixture('pptx'),value='\n（1）新投资比例：80%。\n\n',result=await write(t,[op(t.paragraphs[0],value)]);
 assert.equal(textOf(elements(result.doc,A,'p')[0],'pptx'),value);assert.equal(elements(result.doc,A,'br').length,3);
 const field=await fixture('pptx',{content:'<a:fld id="field-id" type="slidenum"><a:rPr/><a:t>旧值</a:t></a:fld>'});await assert.rejects(write(field,[op(field.paragraphs[0],'新值\n续项')]),/结构|文字/);
});
