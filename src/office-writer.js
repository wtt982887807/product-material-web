import {NS,parseXml,serializeXml,elements,children,textOf,sha256,isVisibleBranch} from './office-xml.js';
import {deliveryName,currentMonth,uniqueDeliveryNames} from './naming.js?v=20261026-drop';
import {OfficePackage,fail,tag,create,setW,directElements,allocateWordIds,wordCommentPart,enableWordTracking,repairPptCommentAuthors,createPptCommentState,addPptComment,subsetPresentation,presentationOrder,validatePackage} from './office-package.js';

const AUTHOR='产品要素核对';
const W=NS.w,A=NS.a;
const XML=NS.xml||'http://www.w3.org/XML/1998/namespace';
const MIME={docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation'};
const between=(x,start,end)=>x>start&&x<end;
const overlaps=(a,b)=>a.start<b.end&&b.start<a.end;

function location(entry) {
  return `${entry.op.part}，段落 ${entry.op.index+1}，原文「${entry.op.before.slice(0,45)}」`;
}
function failAt(entry,message){fail(`${message}（${location(entry)}）`);}
function textRecords(paragraph,kind) {
  const records=[];let cursor=0;
  const walk=node=>{
    if(!isVisibleBranch(node))return;
    if(node!==paragraph&&((kind==='docx'&&tag(node,W,'p'))||(kind==='pptx'&&tag(node,A,'p'))))return;
    if(kind==='docx'&&(tag(node,W,'del')||tag(node,W,'moveFrom')))return;
    let value;
    if(kind==='docx'){
      if(tag(node,W,'t'))value=node.textContent||'';
      else if(tag(node,W,'tab'))value='\t';
      else if(tag(node,W,'br')||tag(node,W,'cr'))value='\n';
      else if(tag(node,W,'noBreakHyphen'))value='\u2011';
      else if(tag(node,W,'softHyphen'))value='\u00ad';
    } else if(tag(node,A,'t'))value=node.textContent||'';
    else if(tag(node,A,'br'))value='\n';
    if(value!==undefined){records.push({node,start:cursor,end:cursor+value.length,text:value});cursor+=value.length;return;}
    for(const child of Array.from(node.childNodes||[]))if(child.nodeType===1)walk(child);
  };
  walk(paragraph);return records;
}
function isPlainRun(run,paragraph,kind,{allowInsertion=false}={}) {
  const ns=kind==='docx'?W:A;
  if(!tag(run,ns,'r')||directElements(run).some(n=>n.namespaceURI!==ns||!['rPr','t'].includes(n.localName)))return false;
  if(kind==='pptx')return run.parentNode===paragraph&&children(run,A,'t').length===1;
  let parent=run.parentNode;
  // Hyperlink text can be revised without modifying the link target or styling.
  if(tag(parent,W,'hyperlink'))parent=parent.parentNode;
  if(allowInsertion&&tag(parent,W,'ins'))parent=parent.parentNode;
  return parent===paragraph;
}
function wordFieldDepthBefore(node) {
  let depth=0,found=false;
  const visit=current=>{
    if(current===node){found=true;return;}
    if(tag(current,W,'fldChar')){
      const type=current.getAttributeNS(W,'fldCharType');if(type==='begin')depth++;else if(type==='end')depth=Math.max(0,depth-1);
    }
    for(const child of Array.from(current.childNodes||[])){if(found)break;if(child.nodeType===1)visit(child);}
  };
  visit(node.ownerDocument.documentElement);return depth;
}
function runSegments(entry,kind,{current=false}={}) {
  const {paragraph,start,end}=entry;
  const records=textRecords(paragraph,kind),touched=records.filter(r=>r.end>start&&r.start<end);
  if(!touched.length)failAt(entry,'未找到可处理的文字');
  const runs=new Set();
  for(const record of touched){
    const run=record.node.parentNode,ns=kind==='docx'?W:A;
    if(!tag(record.node,ns,'t')||!isPlainRun(run,paragraph,kind))failAt(entry,'此位置包含域、特殊文字结构或非普通文本，无法安全自动修改');
    if(kind==='docx'){
      for(let ancestor=run.parentNode;ancestor&&ancestor!==paragraph;ancestor=ancestor.parentNode)if(tag(ancestor,W,'fldSimple')||tag(ancestor,W,'sdtContent'))failAt(entry,'此位置位于域或内容控件内，需人工核对');
      if(wordFieldDepthBefore(run))failAt(entry,'此位置位于 Word 域内，需人工核对');
    }
    runs.add(run);
  }
  return [...runs].map(run=>{
    const rr=records.filter(r=>r.node.parentNode===run);
    return {run,start:rr[0].start,end:rr[rr.length-1].end,text:rr.map(r=>r.text).join('')};
  });
}
function runPiece(run,text,kind,{deleted=false}={}) {
  const clone=run.cloneNode(true),ns=kind==='docx'?W:A,prefix=kind==='docx'?'w':'a';
  for(const node of directElements(clone))if(!tag(node,ns,'rPr'))clone.removeChild(node);
  const t=create(clone.ownerDocument,ns,`${prefix}:${deleted?'delText':'t'}`);t.textContent=text;
  if(/^\s|\s$/u.test(text))t.setAttributeNS(XML,'xml:space','preserve');
  clone.appendChild(t);return clone;
}
function replaceNodeWith(node,pieces){const parent=node.parentNode;for(const piece of pieces)parent.insertBefore(piece,node);parent.removeChild(node);}
function revision(doc,type,id,timestamp){const node=create(doc,W,`w:${type}`);setW(node,'id',id);setW(node,'author',AUTHOR);setW(node,'date',timestamp);return node;}

const sectionReplacement=op=>op.kind==='replace'&&['contract_section','user_fee_text','equal_fee_tiers','user_custom_section'].includes(op.source);
function sectionContent(entry,kind){
  const {paragraph,op}=entry,ns=kind==='docx'?W:A;
  if(op.start!==0||op.before!==textOf(paragraph,kind))failAt(entry,'整段文字更新必须准确对应整段原文');
  const all=directElements(paragraph),content=all.filter(n=>tag(n,ns,'r')||kind==='pptx'&&tag(n,A,'br'));
  const first=all.indexOf(content[0]),last=all.indexOf(content.at(-1));
  for(const [index,node]of all.entries()){
    if(tag(node,ns,'pPr')||kind==='pptx'&&tag(node,A,'endParaRPr'))continue;
    if(kind==='docx'&&node.namespaceURI===W&&['bookmarkStart','bookmarkEnd','commentRangeStart','commentRangeEnd','proofErr','permStart','permEnd'].includes(node.localName)&&!(index>first&&index<last))continue;
    if(tag(node,ns,'r')){
      const allowed=kind==='docx'?['rPr','t','br','cr']:['rPr','t'];
      if(directElements(node).some(n=>n.namespaceURI!==ns||!allowed.includes(n.localName)))failAt(entry,'整段包含域、制表或非普通文字结构，无法安全更新');
      if(kind==='docx'&&wordFieldDepthBefore(node))failAt(entry,'整段位于 Word 域内，无法安全更新');
    }else if(!(kind==='pptx'&&tag(node,A,'br')&&directElements(node).every(n=>tag(n,A,'rPr'))))failAt(entry,'整段包含特殊结构或内部批注边界，无法安全更新');
  }
  if(!content.length)failAt(entry,'整段没有可处理的文字');
  return content;
}
function sectionStyleLines(paragraph,kind){
  const lines=[{text:'',spans:[]}];
  for(const record of textRecords(paragraph,kind)){
    const run=tag(record.node.parentNode,kind==='docx'?W:A,'r')?record.node.parentNode:null;
    const pieces=record.text.split('\n');
    pieces.forEach((text,i)=>{if(i)lines.push({text:'',spans:[]});const line=lines.at(-1),start=line.text.length;line.text+=text;if(text&&run)line.spans.push({start,end:line.text.length,run});});
  }
  return lines;
}
function sectionRegions(text){
  const prefix=text.match(/^\s*(?:(?:[（(][\d一二三四五六七八九十百]+[）)]|[\d一二三四五六七八九十百]+[.．、)])\s*)/)?.[0].length||0;
  const colon=text.slice(prefix).search(/[:：]/),title=colon>=0&&colon<=80?prefix+colon+1:prefix;
  return [[0,prefix],[prefix,title],[title,text.length]];
}
function sectionStyledRuns(line,source,fallback,kind){
  const target=sectionRegions(line),origin=sectionRegions(source.text),runs=[];
  for(let i=0;i<target.length;i++){
    let remaining=line.slice(...target[i]);if(!remaining)continue;
    const [start,end]=origin[i],spans=source.spans.filter(s=>s.end>start&&s.start<end);
    if(!spans.length){const seed=(i===2?source.spans.at(-1):source.spans[0])?.run||fallback;runs.push(runPiece(seed,remaining,kind));continue;}
    for(let j=0;j<spans.length&&remaining;j++){
      const span=spans[j],length=j===spans.length-1?remaining.length:Math.min(remaining.length,Math.min(end,span.end)-Math.max(start,span.start));
      if(length){runs.push(runPiece(span.run,remaining.slice(0,length),kind));remaining=remaining.slice(length);}
    }
  }
  return runs;
}
function sectionBreak(run,kind){
  const doc=run.ownerDocument;
  if(kind==='docx'){const clone=run.cloneNode(true);for(const n of directElements(clone))if(!tag(n,W,'rPr'))clone.removeChild(n);clone.appendChild(create(doc,W,'w:br'));return clone;}
  const br=create(doc,A,'a:br'),properties=children(run,A,'rPr')[0];if(properties)br.appendChild(properties.cloneNode(true));return br;
}
function replaceSection(entry,kind,nextId,timestamp){
  const {paragraph,op}=entry,content=sectionContent(entry,kind),lines=sectionStyleLines(paragraph,kind);
  const fallback=lines.flatMap(line=>line.spans)[0]?.run||content.find(n=>tag(n,kind==='docx'?W:A,'r'));
  if(!fallback)failAt(entry,'整段缺少可继承的文字样式');
  const next=[];
  if(op.after)op.after.split('\n').forEach((line,index)=>{
    if(index)next.push(sectionBreak(next.filter(n=>tag(n,kind==='docx'?W:A,'r')).at(-1)||fallback,kind));
    next.push(...sectionStyledRuns(line,lines[Math.min(index,lines.length-1)],fallback,kind));
  });
  const nodes=[],ids=[];
  if(kind==='docx'){
    const del=revision(paragraph.ownerDocument,'del',nextId(),timestamp);ids.push(del.getAttributeNS(W,'id'));
    for(const old of content){const clone=old.cloneNode(true);for(const t of children(clone,W,'t')){const d=create(paragraph.ownerDocument,W,'w:delText');for(const attr of Array.from(t.attributes||[]))d.setAttributeNS(attr.namespaceURI,attr.name,attr.value);d.textContent=t.textContent;clone.replaceChild(d,t);}del.appendChild(clone);}nodes.push(del);
    if(next.length){const ins=revision(paragraph.ownerDocument,'ins',nextId(),timestamp);ids.push(ins.getAttributeNS(W,'id'));for(const node of next)ins.appendChild(node);nodes.push(ins);}
  }else nodes.push(...next);
  for(const node of nodes)paragraph.insertBefore(node,content[0]);for(const node of content)paragraph.removeChild(node);
  return kind==='docx'?{revisionIds:ids,section:true}:{replaced:true,section:true};
}

function replace(entry,kind,nextId,timestamp) {
  if(sectionReplacement(entry.op))return replaceSection(entry,kind,nextId,timestamp);
  const segments=runSegments(entry,kind),ids=[];let remaining=entry.op.after;
  for(let i=0;i<segments.length;i++){
    const s=segments[i],left=Math.max(0,entry.start-s.start),right=Math.min(s.text.length,entry.end-s.start);
    const old=s.text.slice(left,right),length=i===segments.length-1?remaining.length:Math.min(old.length,remaining.length),value=remaining.slice(0,length);remaining=remaining.slice(length);
    const pieces=[];
    if(s.text.slice(0,left))pieces.push(runPiece(s.run,s.text.slice(0,left),kind));
    if(kind==='docx'){
      const deletion=revision(s.run.ownerDocument,'del',nextId(),timestamp);deletion.appendChild(runPiece(s.run,old,kind,{deleted:true}));pieces.push(deletion);ids.push(deletion.getAttributeNS(W,'id'));
      if(value){const insertion=revision(s.run.ownerDocument,'ins',nextId(),timestamp);insertion.appendChild(runPiece(s.run,value,kind));pieces.push(insertion);ids.push(insertion.getAttributeNS(W,'id'));}
    } else if(value)pieces.push(runPiece(s.run,value,kind));
    if(s.text.slice(right))pieces.push(runPiece(s.run,s.text.slice(right),kind));
    replaceNodeWith(s.run,pieces);
  }
  return kind==='docx'?{revisionIds:ids}:{replaced:true};
}
function translated(position,replacements){return position+replacements.filter(r=>r.end<=position).reduce((sum,r)=>sum+r.op.after.length-(r.end-r.start),0);}
function setHighlight(run,kind) {
  const ns=kind==='docx'?W:A,prefix=kind==='docx'?'w':'a',doc=run.ownerDocument;
  let properties=children(run,ns,'rPr')[0];if(!properties){properties=create(doc,ns,`${prefix}:rPr`);run.insertBefore(properties,run.firstChild);}
  const mark=create(doc,ns,`${prefix}:highlight`);
  if(kind==='docx')setW(mark,'val','yellow');else mark.appendChild(create(doc,A,'a:srgbClr',{val:'FFFF00'}));
  const old=children(properties,ns,'highlight');
  const later=new Set(kind==='docx'?['u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange']:['uLnTx','uLn','uFillTx','uFill','latin','ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst']);
  const reference=old[0]||directElements(properties).find(n=>later.has(n.localName))||null;
  properties.insertBefore(mark,reference);for(const node of old)properties.removeChild(node);
}
function highlight(entry,kind,replacements) {
  const start=translated(entry.start,replacements),end=translated(entry.end,replacements);
  const current={...entry,start,end};
  for(const s of runSegments(current,kind,{current:true})){
    const left=Math.max(0,start-s.start),right=Math.min(s.text.length,end-s.start),pieces=[];
    if(left)pieces.push(runPiece(s.run,s.text.slice(0,left),kind));
    const marked=runPiece(s.run,s.text.slice(left,right),kind);setHighlight(marked,kind);pieces.push(marked);
    if(right<s.text.length)pieces.push(runPiece(s.run,s.text.slice(right),kind));
    replaceNodeWith(s.run,pieces);
  }
  return {color:'yellow',appliedStart:start,appliedEnd:end,highlightText:entry.op.before};
}
function splitWordBoundary(entry,position) {
  const records=textRecords(entry.paragraph,'docx'),runs=new Map();
  for(const record of records){const run=record.node.parentNode;if(!runs.has(run))runs.set(run,[]);runs.get(run).push(record);}
  const selected=[...runs.entries()].find(([,rr])=>between(position,rr[0].start,rr[rr.length-1].end));
  if(!selected)return;
  const [run,rr]=selected;
  if(!isPlainRun(run,entry.paragraph,'docx'))failAt(entry,'批注边界落在修订或特殊文本内部，无法安全定位');
  const text=rr.map(r=>r.text).join(''),offset=position-rr[0].start;
  replaceNodeWith(run,[runPiece(run,text.slice(0,offset),'docx'),runPiece(run,text.slice(offset),'docx')]);
}
function wordAnchorNode(record,paragraph) {
  let node=record.node.parentNode;
  if(tag(node.parentNode,W,'ins'))node=node.parentNode;
  if(node.parentNode===paragraph||tag(node.parentNode,W,'hyperlink'))return node;
  fail('Word 批注位置含不支持的嵌套结构。');
}
async function commentWord(pkg,entry,comments,nextId,timestamp,replacements) {
  const start=translated(entry.start,replacements),end=translated(entry.end,replacements),id=nextId(),doc=entry.paragraph.ownerDocument;
  splitWordBoundary(entry,end);splitWordBoundary(entry,start);
  const covered=textRecords(entry.paragraph,'docx').filter(r=>r.end>start&&r.start<end);
  const begin=create(doc,W,'w:commentRangeStart'),finish=create(doc,W,'w:commentRangeEnd');setW(begin,'id',id);setW(finish,'id',id);
  const ref=create(doc,W,'w:r'),pr=create(doc,W,'w:rPr'),style=create(doc,W,'w:rStyle');setW(style,'val','CommentReference');pr.appendChild(style);ref.appendChild(pr);
  const reference=create(doc,W,'w:commentReference');setW(reference,'id',id);ref.appendChild(reference);
  if(covered.length){
    const first=wordAnchorNode(covered[0],entry.paragraph),last=wordAnchorNode(covered[covered.length-1],entry.paragraph);
    first.parentNode.insertBefore(begin,first);last.parentNode.insertBefore(finish,last.nextSibling);finish.parentNode.insertBefore(ref,finish.nextSibling);
  } else {
    // A replacement may delete the whole source span; retain a point comment.
    const after=textRecords(entry.paragraph,'docx').find(r=>r.start>=start),next=after?wordAnchorNode(after,entry.paragraph):null;
    const parent=next?.parentNode||entry.paragraph;parent.insertBefore(begin,next);parent.insertBefore(finish,next);parent.insertBefore(ref,next);
  }
  const c=create(comments.doc,W,'w:comment');setW(c,'id',id);setW(c,'author',AUTHOR);setW(c,'date',timestamp);
  for(const line of (entry.op.after||entry.op.reason||'请核实此项。').replace(/\r\n?/g,'\n').split('\n')){
    const p=create(comments.doc,W,'w:p'),r=create(comments.doc,W,'w:r');
    line.split('\t').forEach((fragment,index)=>{if(index)r.appendChild(create(comments.doc,W,'w:tab'));const t=create(comments.doc,W,'w:t');t.textContent=fragment;if(/^\s|\s$/u.test(fragment))t.setAttributeNS(XML,'xml:space','preserve');r.appendChild(t);});
    p.appendChild(r);c.appendChild(p);
  }
  comments.doc.documentElement.appendChild(c);pkg.mark(comments.part);
  return {commentId:id,commentPart:comments.part,appliedStart:start,appliedEnd:end};
}

async function prepare(pkg,review) {
  const template=review.template,kind=template.kind,groups=new Map(),ids=new Set();
  const paragraphsById=new Map(template.paragraphs.map(p=>[p.id,p]));
  for(const [part,doc] of template.parts){
    // Never share live reader DOMs with the writer.
    pkg.docs.set(part,parseXml(serializeXml(doc)));
  }
  for(const part of template.preparation?.changedParts||[])pkg.mark(part);
  for(const raw of review.operations||[]){
    const op={...raw};
    if(!op.id||ids.has(op.id))fail('审阅操作编号缺失或重复。');ids.add(op.id);
    if(!['replace','highlight','comment'].includes(op.kind))fail(`不支持的审阅操作：${op.kind}`);
    if(!Number.isInteger(op.start)||op.start<0||typeof op.before!=='string'||!op.before.length)fail('审阅操作缺少有效原文位置。');
    if(op.kind==='replace'&&(typeof op.after!=='string'||(sectionReplacement(op)?/[\r\t]/:/[\r\n\t]/).test(op.after)))fail('替换内容包含不支持的分段或制表结构，需人工核对。');
    const descriptor=paragraphsById.get(op.paragraphId);
    if(!descriptor||descriptor.part!==op.part||descriptor.index!==op.index)fail(`审阅操作无法唯一定位原段落：${op.id}`);
    if(!Number.isInteger(op.index)||op.index<0)fail('审阅段落编号无效。');
    const doc=await pkg.xml(op.part),paragraph=elements(doc,kind==='docx'?W:A,'p')[op.index];
    if(!paragraph)fail(`模板段落不存在：${op.part}，${op.index+1}`);
    const source=textOf(paragraph,kind);
    const entry={op,paragraph,start:op.start,end:op.start+op.before.length};
    if(source!==descriptor.text)failAt(entry,'模板内容在审阅后发生变化');
    if(source.slice(entry.start,entry.end)!==op.before)failAt(entry,'审阅原文与指定位置不匹配');
    if(kind==='docx'&&op.kind==='comment'&&op.part!=='word/document.xml')failAt(entry,'页眉、页脚或脚注中的批注暂不支持，请人工核对');
    if(sectionReplacement(op))sectionContent(entry,kind);else runSegments(entry,kind);
    const key=`${op.part}\u0000${op.index}`;
    if(!groups.has(key))groups.set(key,{paragraph,source,entries:[]});groups.get(key).entries.push(entry);
  }
  for(const group of groups.values()){
    if(group.entries.some(e=>sectionReplacement(e.op))&&group.entries.length!==1)failAt(group.entries[0],'合同整段更新不能与同段其他操作混用');
    const replacements=group.entries.filter(e=>e.op.kind==='replace').sort((a,b)=>a.start-b.start);
    for(let i=1;i<replacements.length;i++)if(overlaps(replacements[i-1],replacements[i]))failAt(replacements[i],'替换范围重叠，需先合并审阅操作');
    for(const e of group.entries)if(e.op.kind!=='replace')for(const r of replacements){
      if(e.op.kind==='highlight'&&overlaps(e,r))failAt(e,'保留待核的标黄范围与替换范围重叠');
      if(e.op.kind==='comment'&&(between(e.start,r.start,r.end)||between(e.end,r.start,r.end)))failAt(e,'批注边界不能落在替换范围内部');
    }
  }
  return groups;
}

function validateWordComments(pkg,comments,ids) {
  if(!comments)return;
  const known=new Set(children(comments.doc.documentElement,W,'comment').map(n=>n.getAttributeNS(W,'id')));
  for(const id of ids){
    if(!known.has(id))fail('生成的 Word 批注没有对应内容。');
    let starts=0,ends=0,refs=0;
    for(const [name,doc] of pkg.docs)if(name.startsWith('word/')&&name!==comments.part){
      starts+=elements(doc,W,'commentRangeStart').filter(n=>n.getAttributeNS(W,'id')===id).length;
      ends+=elements(doc,W,'commentRangeEnd').filter(n=>n.getAttributeNS(W,'id')===id).length;
      refs+=elements(doc,W,'commentReference').filter(n=>n.getAttributeNS(W,'id')===id).length;
    }
    if(starts!==1||ends!==1||refs!==1)fail('生成的 Word 批注锚点不完整或重复。');
  }
}

export async function writeMaterial(materialReview,{month,targetName,now=new Date()}={}) {
  const template=materialReview?.template;
  if(!template||!['docx','pptx'].includes(template.kind))fail('请选择有效的 Word 或 PPTX 模板。');
  if(materialReview.productElementsLocated===false)fail('未识别到可定位的产品要素，未生成此模板，避免只改文件名而保留旧产品正文。');
  const sourceHash=await sha256(template.bytes);
  if(template.sha256&&template.sha256!==sourceHash)fail('上传的模板在审阅后发生变化，请重新上传。');
  const pkg=await OfficePackage.load(template.bytes),groups=await prepare(pkg,materialReview),kind=template.kind;
  const timestamp=new Date(now).toISOString().replace(/\.\d{3}Z$/,'Z');
  const entries=[...groups.values()].flatMap(g=>g.entries),hasComments=entries.some(e=>e.op.kind==='comment'),hasReplacements=entries.some(e=>e.op.kind==='replace');
  const nextId=kind==='docx'?await allocateWordIds(pkg):null;
  if(kind==='pptx')await repairPptCommentAuthors(pkg);
  const comments=kind==='docx'&&hasComments?await wordCommentPart(pkg):null;
  const pptState=kind==='pptx'&&hasComments?await createPptCommentState(pkg,AUTHOR):null;
  const results=[],commentIds=[];
  for(const group of groups.values()){
    const replacements=group.entries.filter(e=>e.op.kind==='replace').sort((a,b)=>b.start-a.start);
    for(const entry of replacements)results.push({id:entry.op.id,kind:'replace',...replace(entry,kind,nextId,timestamp)});
    for(const entry of group.entries.filter(e=>e.op.kind==='highlight'))results.push({id:entry.op.id,kind:'highlight',...highlight(entry,kind,replacements)});
    for(const entry of group.entries.filter(e=>e.op.kind==='comment')){
      const detail=kind==='docx'?await commentWord(pkg,entry,comments,nextId,timestamp,replacements):await addPptComment(pkg,entry,pptState,timestamp);
      results.push({id:entry.op.id,kind:'comment',...detail});if(kind==='docx')commentIds.push(detail.commentId);
    }
    let expected=group.source;
    for(const r of replacements)expected=expected.slice(0,r.start)+r.op.after+expected.slice(r.end);
    if(textOf(group.paragraph,kind)!==expected)fail('输出文字与核对结果不一致，已停止生成。');
    pkg.mark(group.entries[0].op.part);
  }
  if(kind==='docx'&&hasReplacements)await enableWordTracking(pkg);
  let slideReport={};
  if(kind==='pptx'){
    const order=await presentationOrder(pkg);
    const pages=materialReview.kind==='poster'?order.slides.map((s,i)=>i+1):materialReview.selectedPages;
    const retained=new Set(pages?.map(n=>order.slides[n-1]?.part));
    if(entries.some(e=>!retained.has(e.op.part)))fail('存在指向非交付页面的修改，请先完成目标产品页选择。');
    slideReport=await subsetPresentation(pkg,pages);
  }
  validateWordComments(pkg,comments,commentIds);
  await validatePackage(pkg);
  const changedParts=[...pkg.changed].filter(p=>!pkg.removed.has(p)),removedParts=[...pkg.removed];
  const bytes=await pkg.serialize();
  if(await sha256(template.bytes)!==sourceHash)fail('生成期间原始文件发生变化，已停止下载。');
  const outputHash=await sha256(bytes);
  const name=deliveryName(template.name,{targetName,templateProductName:materialReview.templateProductName,kind:materialReview.kind,month:month||currentMonth(now)});
  return {name,bytes,mime:MIME[kind],report:{sourceName:template.name,sourceSha256:sourceHash,outputSha256:outputHash,structuralValidation:true,nativeOfficeValidation:'notPerformedInBrowser',businessVerified:!(materialReview.fields||[]).some(f=>['pending','conflict','retained_pending','待核'].includes(f.status)),acceptedRevisions:template.preparation?.acceptedRevisions||0,changedParts,removedParts,repairs:pkg.repairs,operations:results,...slideReport}};
}

export async function writeTask(task) {
  if(!task||!Array.isArray(task.materials)||!task.materials.length)fail('没有可生成的物料。');
  const files=[],failures=[],now=task.now||new Date(),month=task.month||currentMonth(now),targetName=task.target?.name;
  // Reserve all preview names, including failed items, so surviving deliveries
  // do not unexpectedly take another template's name after a partial failure.
  const names=uniqueDeliveryNames(task.materials.map((m,i)=>{try{return deliveryName(m.template?.name,{targetName,templateProductName:m.templateProductName,kind:m.kind,month});}catch{return `未命名物料-${i+1}`;}}));
  for(const [index,material]of task.materials.entries()){
    try{const file=await writeMaterial(material,{month,targetName,now});file.name=names[index];files.push(file);}
    catch(error){failures.push({sourceName:material.template?.originalName||material.template?.name||`物料${index+1}`,stage:'generate',message:error.message||String(error)});}
  }
  if(!files.length)fail('没有物料成功生成。'+failures.map(f=>`${f.sourceName}：${f.message}`).join('\n'));
  return {files,failures,report:{fileCount:files.length,requestedCount:task.materials.length,complete:!failures.length,failures,structuralValidation:true,businessVerified:!failures.length&&files.every(f=>f.report.businessVerified),files:files.map(f=>({name:f.name,...f.report}))}};
}
