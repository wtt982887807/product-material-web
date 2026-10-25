import {NS, parseXml, serializeXml, elements, children} from './office-xml.js';

const PR = NS.pr || NS.rel;
const R = NS.r;
const CT = NS.ct;
const P = NS.p;
const W = NS.w;
const EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
export const fail = message => { throw new Error(message); };
export const tag = (node, ns, name) => node?.nodeType === 1 && node.namespaceURI === ns && node.localName === name;
export const directElements = node => Array.from(node.childNodes || []).filter(n => n.nodeType === 1);
export const create = (doc, ns, name, attrs={}) => {
  const node = doc.createElementNS(ns, name);
  for (const [key,value] of Object.entries(attrs)) node.setAttribute(key,String(value));
  return node;
};
export function setW(node, name, value) { node.setAttributeNS(W,`w:${name}`,String(value)); }
export function relsName(part='') {
  const slash = part.lastIndexOf('/');
  return part ? `${part.slice(0,slash+1)}_rels/${part.slice(slash+1)}.rels` : '_rels/.rels';
}
export function resolveTarget(owner, target) {
  if (!target || /[\\?#]/.test(target) || /^[a-z][a-z\d+.-]*:/i.test(target)) fail(`无法安全读取文件内部链接：${target}`);
  let decoded;
  try { decoded=decodeURIComponent(target); } catch { fail('文件内部链接编码无效。'); }
  if (/[\\?#]/.test(decoded)) fail('文件内部链接含不支持的特殊字符。');
  const segments=(decoded.startsWith('/') ? decoded.slice(1) : `${owner.slice(0,owner.lastIndexOf('/')+1)}${decoded}`).split('/');
  const resolved=[];
  for (const segment of segments) {
    if (!segment || segment === '.') continue;
    if (segment === '..') { if (!resolved.length) fail('文件内部链接超出压缩包根目录。'); resolved.pop(); }
    else resolved.push(segment);
  }
  return resolved.join('/');
}
function relativeTarget(owner, target) {
  const from=owner.split('/').slice(0,-1), to=target.split('/');
  while(from.length && to.length && from[0]===to[0]) {from.shift();to.shift();}
  return [...from.map(()=> '..'), ...to].join('/');
}

/** A private working copy. Uploaded bytes and reader DOMs are never mutated. */
export class OfficePackage {
  constructor(zip) { this.zip=zip; this.docs=new Map(); this.changed=new Set(); this.removed=new Set(); this.repairs=[]; }
  static async load(bytes) {
    if (!globalThis.JSZip) fail('本地压缩组件尚未加载，请刷新页面。');
    const zip=await globalThis.JSZip.loadAsync(bytes,{checkCRC32:true});
    const pkg=new OfficePackage(zip);
    for(const [name,file] of Object.entries(zip.files)) {
      if(file.dir)continue;
      if(name.startsWith('/') || name.includes('\\') || name.split('/').some(s=>s==='.'||s==='..') || (file.unsafeOriginalName && file.unsafeOriginalName!==name)) fail('压缩包含不安全的文件路径，无法处理。');
      if(name.startsWith('_xmlsignatures/')) fail('带数字签名的文件需先另存为无签名模板。');
    }
    if(!pkg.has('[Content_Types].xml')) fail('文件缺少 Office 内容类型表。');
    return pkg;
  }
  has(name) { return !this.removed.has(name) && (this.docs.has(name) || !!this.zip.file(name)); }
  names() { return [...new Set([...Object.keys(this.zip.files).filter(n=>!this.zip.files[n].dir),...this.docs.keys()])].filter(n=>!this.removed.has(n)); }
  async xml(name, {optional=false,root,ns}={}) {
    if(this.docs.has(name))return this.docs.get(name);
    const file=this.zip.file(name);
    if(!file) {
      if(!optional)fail(`文件缺少必要内容：${name}`);
      if(!root)return null;
      const doc=parseXml(`<${root} xmlns${root.includes(':')?`:${root.split(':')[0]}`:''}="${ns}"/>`);
      this.docs.set(name,doc);this.changed.add(name);return doc;
    }
    const doc=parseXml(await file.async('string'));this.docs.set(name,doc);return doc;
  }
  setXml(name, doc) {this.docs.set(name,doc);this.changed.add(name);this.removed.delete(name);}
  mark(name) {this.changed.add(name);}
  remove(name) {this.removed.add(name);this.changed.delete(name);}
  async relationships(owner, createMissing=false) {
    const name=relsName(owner);
    const doc=await this.xml(name,{optional:true,...(createMissing?{root:'Relationships',ns:PR}:{})});
    if(!doc)return {name,doc:null,nodes:[],byId:new Map()};
    if(!tag(doc.documentElement,PR,'Relationships'))fail(`文件关系表格式不受支持：${name}`);
    const nodes=children(doc.documentElement,PR,'Relationship'),byId=new Map();
    for(const node of nodes){const id=node.getAttribute('Id');if(!id||byId.has(id))fail(`文件关系编号缺失或重复：${name}`);byId.set(id,node);}
    return {name,doc,nodes,byId};
  }
  async relation(owner, type, defaultPart) {
    const rels=await this.relationships(owner,true);
    const found=rels.nodes.filter(n=>n.getAttribute('Type')===`${R}/${type}`);
    if(found.length>1)fail(`文件存在多个 ${type} 关系，无法安全确定目标。`);
    if(found.length){
      if(found[0].getAttribute('TargetMode')==='External')fail('批注或设置不能指向外部文件。');
      const part=resolveTarget(owner,found[0].getAttribute('Target'));
      if(!this.has(part))fail(`批注或设置文件不存在：${part}`);
      return part;
    }
    let part=defaultPart,count=1;
    while(this.has(part))part=defaultPart.replace(/(\.[^./]+)$/u,`-${count++}$1`);
    let id=1;while(rels.byId.has(`rId${id}`))id++;
    rels.doc.documentElement.appendChild(create(rels.doc,PR,'Relationship',{Id:`rId${id}`,Type:`${R}/${type}`,Target:relativeTarget(owner,part)}));
    this.mark(rels.name);return part;
  }
  async contentType(part, contentType) {
    const doc=await this.xml('[Content_Types].xml');
    if(!tag(doc.documentElement,CT,'Types'))fail('文件内容类型表格式不受支持。');
    const matches=children(doc.documentElement,CT,'Override').filter(n=>n.getAttribute('PartName')===`/${part}`);
    if(matches.length>1)fail(`文件内容类型表重复：${part}`);
    if(matches.length && matches[0].getAttribute('ContentType')!==contentType)fail(`文件内容类型冲突：${part}`);
    if(!matches.length){doc.documentElement.appendChild(create(doc,CT,'Override',{PartName:`/${part}`,ContentType:contentType}));this.mark('[Content_Types].xml');}
  }
  async serialize() {
    for(const name of this.changed)if(!this.removed.has(name))this.zip.file(name,serializeXml(this.docs.get(name)));
    for(const name of this.removed)this.zip.remove(name);
    return this.zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  }
}

export async function allocateWordIds(pkg) {
  let maximum=-1;
  for(const name of pkg.names())if(name.startsWith('word/')&&name.endsWith('.xml')){
    const doc=await pkg.xml(name);
    for(const node of [doc.documentElement,...elements(doc.documentElement,'*','*')]){
      const id=node.getAttributeNS?.(W,'id');if(/^\d+$/.test(id||'')){const n=Number(id);if(!Number.isSafeInteger(n))fail('Word 修订或批注编号超出可安全处理的范围。');maximum=Math.max(maximum,n);}
    }
  }
  return ()=>String(++maximum);
}

export async function wordCommentPart(pkg) {
  const part=await pkg.relation('word/document.xml','comments','word/comments.xml');
  const doc=await pkg.xml(part,{optional:true,root:'w:comments',ns:W});
  if(!tag(doc.documentElement,W,'comments'))fail('Word 批注格式不受支持。');
  const ids=children(doc.documentElement,W,'comment').map(n=>n.getAttributeNS(W,'id'));
  if(new Set(ids).size!==ids.length)fail('原有 Word 批注编号重复，需先修复模板。');
  await pkg.contentType(part,'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');
  pkg.mark(part);return {part,doc};
}

export async function enableWordTracking(pkg) {
  const part=await pkg.relation('word/document.xml','settings','word/settings.xml');
  const doc=await pkg.xml(part,{optional:true,root:'w:settings',ns:W});
  if(!tag(doc.documentElement,W,'settings'))fail('Word 设置格式不受支持。');
  let track=children(doc.documentElement,W,'trackRevisions')[0];
  if(!track){
    track=create(doc,W,'w:trackRevisions');
    const early=new Set(['writeProtection','view','zoom','removePersonalInformation','removeDateAndTime','doNotDisplayPageBoundaries','displayBackgroundShape','printPostScriptOverText','printFractionalCharacterWidth','printFormsData','embedTrueTypeFonts','embedSystemFonts','saveSubsetFonts','saveFormsData','mirrorMargins','alignBordersAndEdges','bordersDoNotSurroundHeader','bordersDoNotSurroundFooter','gutterAtTop','hideSpellingErrors','hideGrammaticalErrors','activeWritingStyle','proofState','formsDesign','attachedTemplate','linkStyles','stylePaneFormatFilter','stylePaneSortMethod','documentType','mailMerge','revisionView']);
    doc.documentElement.insertBefore(track,directElements(doc.documentElement).find(n=>!early.has(n.localName))||null);
  }
  setW(track,'val','true');pkg.mark(part);
  await pkg.contentType(part,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
}

export async function presentationOrder(pkg) {
  const root=await pkg.relationships('');
  const links=root.nodes.filter(n=>n.getAttribute('Type')===`${R}/officeDocument`);
  if(links.length!==1||links[0].getAttribute('TargetMode')==='External')fail('PPT 必须包含一个有效的演示文稿入口。');
  const part=resolveTarget('',links[0].getAttribute('Target'));
  const doc=await pkg.xml(part);
  if(!tag(doc.documentElement,P,'presentation'))fail('仅支持标准 PPTX 演示文稿。');
  const rels=await pkg.relationships(part),slides=[];
  const list=children(doc.documentElement,P,'sldIdLst')[0];
  if(!list)fail('PPT 没有可交付的页面。');
  for(const node of children(list,P,'sldId')){
    const rid=node.getAttributeNS(R,'id'),rel=rels.byId.get(rid);
    if(!rel||rel.getAttribute('Type')!==`${R}/slide`||rel.getAttribute('TargetMode')==='External')fail('PPT 页面关系无效。');
    const target=resolveTarget(part,rel.getAttribute('Target'));
    if(!pkg.has(target))fail(`PPT 页面不存在：${target}`);
    slides.push({part:target,rid,slideId:node.getAttribute('id'),node});
  }
  if(new Set(slides.map(s=>s.part)).size!==slides.length||new Set(slides.map(s=>s.slideId)).size!==slides.length)fail('PPT 页面编号重复。');
  return {part,doc,rels,list,slides};
}

/** Legacy comment IDs may be duplicated by exporters even with no comments. */
export async function repairPptCommentAuthors(pkg) {
  const references=new Set(),authorParts=new Set();
  for(const name of pkg.names())if(name.endsWith('.rels')){
    const owner=name==='_rels/.rels'?'':name.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
    const rels=await pkg.relationships(owner);
    for(const rel of rels.nodes)if(rel.getAttribute('Type')===`${R}/commentAuthors`){
      if(rel.getAttribute('TargetMode')==='External')fail('PPT 批注作者表不能为外部链接。');
      authorParts.add(resolveTarget(owner,rel.getAttribute('Target')));
    }
  }
  for(const name of pkg.names())if(name.startsWith('ppt/')&&name.endsWith('.xml')){
    const doc=await pkg.xml(name);
    if(tag(doc.documentElement,P,'cmAuthorLst'))authorParts.add(name);
    for(const cm of elements(doc,P,'cm'))references.add(cm.getAttribute('authorId'));
  }
  // Resolve actual comment relationships, including producer-specific part names.
  for(const name of pkg.names())if(name.endsWith('.rels')){
    const owner=name==='_rels/.rels'?'':name.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
    for(const rel of (await pkg.relationships(owner)).nodes)if(rel.getAttribute('Type')===`${R}/comments`&&rel.getAttribute('TargetMode')!=='External'){
      const doc=await pkg.xml(resolveTarget(owner,rel.getAttribute('Target')));
      for(const cm of elements(doc,P,'cm'))references.add(cm.getAttribute('authorId'));
    }
  }
  for(const part of authorParts){
    const doc=await pkg.xml(part);if(!tag(doc.documentElement,P,'cmAuthorLst'))fail('PPT 批注作者表结构不受支持。');
    const seen=new Set();let removedCount=0;
    for(const node of directElements(doc.documentElement)){
      const id=node.getAttribute('id');if(!tag(node,P,'cmAuthor')||!/^\d+$/.test(id))fail('PPT 批注作者编号无效。');
      if(seen.has(id)){
        if(references.has(id))fail('PPT 原有批注引用了重复作者编号，作者身份不明确，请先修复模板。');
        doc.documentElement.removeChild(node);removedCount++;
      } else seen.add(id);
    }
    if(removedCount){pkg.mark(part);pkg.repairs.push({kind:'unusedDuplicateCommentAuthors',part,removedCount});}
  }
}

export async function createPptCommentState(pkg, author) {
  for(const name of pkg.names())if(name.endsWith('.rels')){
    const doc=await pkg.xml(name);
    for(const rel of elements(doc,PR,'Relationship')){
      const type=rel.getAttribute('Type');
      if(/comment|person/i.test(type)&&![`${R}/comments`,`${R}/commentAuthors`].includes(type))fail('此 PPT 使用暂不支持的新式批注，请另存为兼容版后处理。');
    }
  }
  const order=await presentationOrder(pkg);
  const part=await pkg.relation(order.part,'commentAuthors','ppt/commentAuthors.xml');
  const doc=await pkg.xml(part,{optional:true,root:'p:cmAuthorLst',ns:P});
  if(!tag(doc.documentElement,P,'cmAuthorLst'))fail('PPT 批注作者格式不受支持。');
  const authors=children(doc.documentElement,P,'cmAuthor');
  if(authors.some(n=>!Number.isSafeInteger(Number(n.getAttribute('id')))))fail('PPT 批注作者编号超出可安全处理的范围。');
  const id=String(Math.max(-1,...authors.map(n=>Number(n.getAttribute('id'))))+1);
  if(!/^\d+$/.test(id))fail('PPT 批注作者编号无效。');
  const record=create(doc,P,'p:cmAuthor',{id,name:author,initials:'核对',lastIdx:'0',clrIdx:String(authors.length)});
  doc.documentElement.appendChild(record);pkg.mark(part);
  await pkg.contentType(part,'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
  return {id,record,part,index:0};
}

export async function addPptComment(pkg, entry, state, timestamp) {
  const {paragraph,op}=entry;
  const part=await pkg.relation(op.part,'comments',`ppt/comments/${op.part.split('/').pop().replace(/\.xml$/,'')}-review.xml`);
  const doc=await pkg.xml(part,{optional:true,root:'p:cmLst',ns:P});
  if(!tag(doc.documentElement,P,'cmLst')||directElements(doc.documentElement).some(n=>!tag(n,P,'cm')))fail('PPT 批注内容格式不受支持。');
  const idx=String(++state.index);state.record.setAttribute('lastIdx',idx);pkg.mark(state.part);
  let x=0,y=0;
  for(let ancestor=paragraph.parentNode;ancestor;ancestor=ancestor.parentNode)if(tag(ancestor,P,'sp')||tag(ancestor,P,'graphicFrame')){
    const properties=children(ancestor,P,tag(ancestor,P,'sp')?'spPr':'xfrm')[0];
    const xfrm=tag(properties,P,'xfrm')?properties:properties&&children(properties,NS.a,'xfrm')[0];
    const off=xfrm&&children(xfrm,NS.a,'off')[0];
    if(off){x=Math.round(Number(off.getAttribute('x'))/1587.5);y=Math.round(Number(off.getAttribute('y'))/1587.5);if(!Number.isFinite(x)||!Number.isFinite(y))fail('PPT 批注位置无效。');}
    break;
  }
  const comment=create(doc,P,'p:cm',{authorId:state.id,dt:timestamp,idx});
  comment.appendChild(create(doc,P,'p:pos',{x,y}));
  const text=create(doc,P,'p:text');text.textContent=`定位原文：${op.before}\n${op.after||op.reason||'请核实此项。'}`;comment.appendChild(text);
  doc.documentElement.appendChild(comment);pkg.mark(part);
  await pkg.contentType(part,'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
  return {commentId:idx,commentPart:part,authorId:state.id,anchorKind:'slidePosition',anchorText:op.before};
}

async function trimSlideNavigation(pkg, owner, rels, excluded) {
  const removed=new Set();
  if(!owner.endsWith('.xml'))return removed;
  const doc=await pkg.xml(owner);
  const isNavigation=node=>tag(node,NS.a,'hlinkClick')||tag(node,NS.a,'hlinkMouseOver');
  const relative=elements(doc,NS.a,'*').filter(node=>isNavigation(node)&&(node.getAttribute('action')||'').startsWith('ppaction://hlinkshowjump'));
  for(const node of relative)node.parentNode.removeChild(node);
  if(relative.length){
    pkg.mark(owner);
    pkg.repairs.push({kind:'removedRelativeNavigation',part:owner,removedCount:relative.length,message:`仅输出产品要素页，已移除 ${relative.length} 个可能因删页改变目标的相对翻页动作；原文字和图形保留。`});
  }
  for(const rel of rels.nodes){
    if(rel.getAttribute('TargetMode')==='External')continue;
    const target=resolveTarget(owner,rel.getAttribute('Target'));
    if(!excluded.has(target))continue;
    const id=rel.getAttribute('Id'),references=[];
    for(const node of [doc.documentElement,...elements(doc,'*','*')])for(const attr of Array.from(node.attributes||[]))if(attr.namespaceURI===R&&attr.value===id)references.push({node,attr});
    // Only navigation is optional. A resource/object reference must still fail
    // if removing a slide would make it incomplete.
    if(![`${R}/slide`,`${R}/hyperlink`].includes(rel.getAttribute('Type'))||references.some(({node,attr})=>!isNavigation(node)||attr.localName!=='id'))continue;
    const actions=new Set(references.map(r=>r.node));
    for(const node of actions)node.parentNode.removeChild(node);
    rels.doc.documentElement.removeChild(rel);removed.add(id);pkg.mark(rels.name);
    if(actions.size)pkg.mark(owner);
    pkg.repairs.push({kind:'removedSlideNavigation',part:owner,target,removedCount:actions.size,relationshipId:id,message:actions.size?`已移除 ${actions.size} 个指向未交付页面的点击／悬停跳转；原文字和图形保留。`:'已清理指向未交付页面的未使用跳转关系。'});
  }
  return removed;
}

export async function subsetPresentation(pkg, selectedPages) {
  const order=await presentationOrder(pkg),count=order.slides.length;
  if(!Array.isArray(selectedPages)||!selectedPages.length||selectedPages.some((n,i)=>!Number.isInteger(n)||n<1||n>count||(i&&n<=selectedPages[i-1])))fail('PPT 交付页必须是原文件中按顺序排列、无重复的有效页码。');
  const selected=selectedPages.map(n=>order.slides[n-1]),keep=new Set(selected.map(s=>s.part));
  const excluded=new Set([...order.slides.map(s=>s.part),...pkg.names().filter(n=>/^ppt\/slides\/[^/]+\.xml$/.test(n))].filter(n=>!keep.has(n)));
  const subset=selected.length!==count;
  if(subset){
    for(const s of order.slides)if(!keep.has(s.part))order.list.removeChild(s.node);
    for(const n of elements(order.doc,P,'custShowLst'))n.parentNode.removeChild(n);
    for(const n of elements(order.doc,P,'showPr'))n.parentNode.removeChild(n);
    for(const n of elements(order.doc,'*','sectionLst'))n.parentNode.removeChild(n);
    pkg.mark(order.part);
    const rids=new Set(selected.map(s=>s.rid));
    for(const rel of order.rels.nodes)if(rel.getAttribute('Type')===`${R}/slide`&&!rids.has(rel.getAttribute('Id')))order.rels.doc.documentElement.removeChild(rel);
    pkg.mark(order.rels.name);
  }
  const reachable=new Set(),pending=[''],visited=new Set();
  while(pending.length){
    const owner=pending.pop();if(visited.has(owner))continue;visited.add(owner);
    if(owner){if(excluded.has(owner))fail(`内部链接指向已排除页面：${owner}`);if(!pkg.has(owner))fail(`内部链接目标缺失：${owner}`);reachable.add(owner);}
    const rels=await pkg.relationships(owner),removedNavigation=subset?await trimSlideNavigation(pkg,owner,rels,excluded):new Set();
    if(!rels.doc)continue;reachable.add(rels.name);
    for(const rel of rels.nodes){
      if(removedNavigation.has(rel.getAttribute('Id')))continue;
      if(subset&&rel.getAttribute('Type').endsWith('/thumbnail')){rels.doc.documentElement.removeChild(rel);pkg.mark(rels.name);continue;}
      if(rel.getAttribute('TargetMode')==='External')continue;
      const target=resolveTarget(owner,rel.getAttribute('Target'));
      if(excluded.has(target))fail(`保留页面或资源含指向已排除页面的链接：${owner} → ${target}`);
      if(!pkg.has(target))fail(`文件内部链接目标缺失：${owner} → ${target}`);
      pending.push(target);
    }
  }
  const types=await pkg.xml('[Content_Types].xml');
  for(const node of children(types.documentElement,CT,'Override'))if(!reachable.has(decodeURIComponent(node.getAttribute('PartName')).replace(/^\//,''))){types.documentElement.removeChild(node);pkg.mark('[Content_Types].xml');}
  reachable.add('[Content_Types].xml');
  for(const name of pkg.names())if(!reachable.has(name))pkg.remove(name);
  if(subset)for(const name of reachable)if(name.endsWith('.xml')){
    const doc=await pkg.xml(name);if(!tag(doc.documentElement,EP,'Properties'))continue;
    const notes=[...reachable].filter(n=>/\/notesSlides\/[^/]+\.xml$/.test(n)).length;
    let hidden=0;for(const s of selected)if(['0','false'].includes((await pkg.xml(s.part)).documentElement.getAttribute('show')))hidden++;
    for(const [key,value] of Object.entries({Slides:selected.length,Notes:notes,HiddenSlides:hidden})){const node=children(doc.documentElement,EP,key)[0];if(node)node.textContent=String(value);}
    for(const key of ['HeadingPairs','TitlesOfParts','Words','Paragraphs','TotalTime'])for(const node of children(doc.documentElement,EP,key))node.parentNode.removeChild(node);
    pkg.mark(name);
  }
  return {sourceSlideCount:count,outputSlideCount:selected.length,slideMapping:selected.map((s,i)=>({outputPage:i+1,sourcePage:selectedPages[i],part:s.part})),removedSlideCount:count-selected.length};
}

export async function validatePackage(pkg) {
  const names=pkg.names();
  for(const name of names)if(name.endsWith('.rels')){
    const owner=name==='_rels/.rels'?'':name.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
    if(owner&&!pkg.has(owner))fail(`文件关系表没有所属文件：${name}`);
    for(const rel of (await pkg.relationships(owner)).nodes)if(rel.getAttribute('TargetMode')!=='External'){
      const target=resolveTarget(owner,rel.getAttribute('Target'));if(!pkg.has(target))fail(`输出包含失效链接：${owner} → ${target}`);
    }
  }
  for(const name of names)if(/\.(xml|vml)$/.test(name)&&name!=='[Content_Types].xml'){
    const doc=await pkg.xml(name),rels=await pkg.relationships(name);
    for(const node of [doc.documentElement,...elements(doc.documentElement,'*','*')])for(const attr of Array.from(node.attributes||[]))if(attr.namespaceURI===R&&attr.value&&!rels.byId.has(attr.value))fail(`输出中的资源编号不存在：${name}（${attr.value}）`);
  }
  const types=await pkg.xml('[Content_Types].xml');const overrideParts=new Set();
  for(const override of children(types.documentElement,CT,'Override')){
    const part=decodeURIComponent(override.getAttribute('PartName')).replace(/^\//,'');
    if(overrideParts.has(part))fail(`输出内容类型重复：${part}`);overrideParts.add(part);if(!pkg.has(part))fail(`输出内容类型指向已移除文件：${part}`);
  }
  return true;
}
