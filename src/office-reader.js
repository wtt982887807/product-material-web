import {NS,parseXml,elements,children,ancestor,isVisibleBranch,textOf,sha256} from './office-xml.js';

function copyBytes(input) {
  if (input instanceof ArrayBuffer) return new Uint8Array(input.slice(0));
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer,input.byteOffset,input.byteLength).slice();
  throw new Error('未取得文件内容，请重新选择文件。');
}

function revisionError(part,type) { const error=new Error(`“${part}”含需要完整 Word 排版引擎处理的旧修订（${type}）。`);error.code='WORD_REVISION_COMPLEX';return error; }
function invalidRevision(part,type) { const error=new Error(`“${part}”的旧修订结构不完整或无效（${type}）。`);error.code='WORD_REVISION_INVALID';return error; }

const propertyChanges=new Set(['rPrChange','pPrChange','tblPrChange','tblPrExChange','tblGridChange','trPrChange','tcPrChange','sectPrChange','numberingChange']);
const isW=(n,name)=>n?.namespaceURI===NS.w&&n.localName===name;
const isRevision=n=>n.namespaceURI===NS.w&&(/Change$/.test(n.localName)||/^(ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|customXml.*Range(Start|End)|move.*Range(Start|End))$/.test(n.localName));
const attached=(n,doc)=>{for(let p=n;p;p=p.parentNode)if(p===doc)return true;return false;};
const within=(n,parent)=>{for(let p=n;p;p=p.parentNode)if(p===parent)return true;return false;};
function unwrap(n,{customXml=false}={}){const parent=n.parentNode;if(!parent)return;for(const child of Array.from(n.childNodes))if(!(customXml&&isW(child,'customXmlPr')))parent.insertBefore(child,n);parent.removeChild(n);}

// ISO/IEC 29500 17.13.5: property changes contain OLD snapshots; deletions,
// row/cell markers and moved paragraph marks have different structural meanings.
function removeRevisedContent(n,doc,part){
  const anchors=elements(n,NS.w).filter(e=>['commentRangeStart','commentRangeEnd','commentReference','bookmarkStart','bookmarkEnd','permStart','permEnd'].includes(e.localName));
  if(anchors.length){
    let parent=n.parentNode,before=n;
    if(['tr','tc','p','tbl'].some(name=>isW(n,name))||!ancestor(n,NS.w,'p')){
      const ordered=elements(doc,NS.w,'p'),inside=ordered.filter(p=>within(p,n));
      const first=inside.length?ordered.indexOf(inside[0]):0,last=inside.length?ordered.indexOf(inside.at(-1)):first;
      const following=ordered.slice(last+1).find(p=>!within(p,n));
      parent=following||ordered.slice(0,first).reverse().find(p=>!within(p,n));
      if(!parent){
        let container=n.parentNode;while(container&&!['body','hdr','ftr','footnote','endnote','comment','tc','txbxContent'].some(name=>isW(container,name)))container=container.parentNode;
        if(!container||within(container,n))throw revisionError(part,'删除内容的批注无法安全锚定');
        parent=doc.createElementNS(NS.w,'w:p');container.insertBefore(parent,children(container,NS.w,'sectPr')[0]||null);
      }
      before=following?children(parent).find(e=>!isW(e,'pPr'))||null:null;
    }
    for(const anchor of anchors){
      if(isW(anchor,'commentReference')){const run=doc.createElementNS(NS.w,'w:r'),pr=children(anchor.parentNode,NS.w,'rPr')[0];if(pr)run.appendChild(pr.cloneNode(true));run.appendChild(anchor);parent.insertBefore(run,before);}
      else parent.insertBefore(anchor,before);
    }
  }
  n.parentNode?.removeChild(n);
}
function acceptRevisionRanges(doc,part){
  const bounds=new Map(),positions=new Map();let cursor=0;
  const walk=n=>{if(n.nodeType!==1)return;const start=cursor++;for(const c of children(n))walk(c);positions.set(n,{start,end:cursor++});};walk(doc.documentElement);
  const ranges=elements(doc,NS.w).filter(n=>/^(?:move(?:From|To)|customXml(?:Ins|Del|MoveFrom|MoveTo))Range(?:Start|End)$/.test(n.localName));
  for(const n of ranges){const type=n.localName.replace(/(?:Start|End)$/,''),id=n.getAttributeNS(NS.w,'id'),key=type+'|'+id,bound=bounds.get(key)||{type};const side=n.localName.endsWith('Start')?'start':'end';if(!id||bound[side])throw invalidRevision(part,'重复或缺失的修订范围编号');bound[side]=n;bounds.set(key,bound);}
  const removeMarkup=new Set();
  for(const b of bounds.values()){
    if(!b.start||!b.end||positions.get(b.start).start>=positions.get(b.end).start)throw invalidRevision(part,'不完整或倒置的修订范围');
    if(/^customXml(?:Del|MoveFrom)Range$/.test(b.type)){
      const low=positions.get(b.start).end,high=positions.get(b.end).start;
      const affected=elements(doc,NS.w,'customXml').filter(n=>{const pos=positions.get(n);return pos.start>low&&pos.start<high||pos.end>low&&pos.end<high;});
      if(!affected.length)throw invalidRevision(part,'自定义 XML 删除范围没有对应标记');
      affected.forEach(n=>removeMarkup.add(n));
    }
  }
  for(const n of removeMarkup)unwrap(n,{customXml:true});
  for(const n of ranges)n.parentNode?.removeChild(n);
}
function acceptCellMerge(n,part){
  if(!isW(n.parentNode,'tcPr'))throw invalidRevision(part,'cellMerge 位置无效');
  const value=n.getAttributeNS(NS.w,'vMerge'),original=n.getAttributeNS(NS.w,'vMergeOrig'),current=children(n.parentNode,NS.w,'vMerge');
  if(!value&&original){for(const v of current)v.parentNode.removeChild(v);n.parentNode.removeChild(n);return;}
  const target={cont:'continue',rest:'restart'}[value];
  // A tracked merge alone can require redistributing cell contents and geometry.
  // Accept the safe case where the producer already wrote the current layout.
  if(!target||current.length!==1||(current[0].getAttributeNS(NS.w,'val')||'continue')!==target)throw revisionError(part,'单元格合并缺少一致的当前合并结构');
  n.parentNode.removeChild(n);
}
function rowCells(row){
  const pr=children(row,NS.w,'trPr')[0];let start=Number(children(pr,NS.w,'gridBefore')[0]?.getAttributeNS(NS.w,'val')||0);
  return children(row,NS.w,'tc').map(cell=>{const properties=children(cell,NS.w,'tcPr')[0],span=Number(children(properties,NS.w,'gridSpan')[0]?.getAttributeNS(NS.w,'val')||1),merge=children(properties,NS.w,'vMerge')[0];const result={start,span,merge:merge?(merge.getAttributeNS(NS.w,'val')||'continue'):null,pending:children(properties,NS.w,'cellMerge').length>0};start+=span;return result;});
}
function rowDeletionNeedsLayout(row,deletedRows){
  const cells=rowCells(row);if(cells.some(c=>c.pending))return true;
  const rows=children(row.parentNode,NS.w,'tr'),after=rows.slice(rows.indexOf(row)+1).find(r=>!deletedRows.has(r));
  if(!after)return false;const next=rowCells(after);
  return cells.some(c=>c.merge==='restart'&&next.some(n=>n.start===c.start&&n.merge==='continue'));
}
function acceptRevisions(doc,part) {
  const original=elements(doc).filter(isRevision),count=original.length;
  for(const n of original.filter(n=>propertyChanges.has(n.localName)))if(attached(n,doc))n.parentNode.removeChild(n);
  acceptRevisionRanges(doc,part);
  const paragraphMarks=[],affectedRows=new Set(),affectedTables=new Set();
  const deletedRows=new Set(elements(doc,NS.w,'del').filter(n=>isW(n.parentNode,'trPr')).map(n=>n.parentNode.parentNode));
  // Remove whole deleted structures before visiting nested revisions in them.
  for(const n of elements(doc,NS.w).filter(n=>['del','moveFrom','cellDel'].includes(n.localName))){
    if(!attached(n,doc))continue;const parent=n.parentNode,mark=isW(parent,'rPr')&&isW(parent.parentNode,'pPr');
    if(mark&&n.localName==='del'){paragraphMarks.push(parent.parentNode.parentNode);parent.removeChild(n);continue;}
    if(mark&&n.localName==='moveFrom'){removeRevisedContent(parent.parentNode.parentNode,doc,part);continue;}
    if(isW(parent,'trPr')&&n.localName==='del'){
      if(rowDeletionNeedsLayout(parent.parentNode,deletedRows))throw revisionError(part,'删除行涉及纵向合并单元格');
      affectedTables.add(ancestor(parent,NS.w,'tbl'));removeRevisedContent(parent.parentNode,doc,part);continue;
    }
    if(isW(parent,'tcPr')&&n.localName==='cellDel'){
      const table=ancestor(parent,NS.w,'tbl');if(elements(table,NS.w).some(e=>['vMerge','hMerge','cellMerge'].includes(e.localName)))throw revisionError(part,'删除单元格涉及合并表格结构');
      affectedRows.add(ancestor(parent,NS.w,'tr'));affectedTables.add(table);removeRevisedContent(parent.parentNode,doc,part);continue;
    }
    if(n.localName==='cellDel'||['trPr','tcPr','rPr','numPr'].some(name=>isW(parent,name))||parent?.namespaceURI!==NS.w)throw revisionError(part,n.localName);
    removeRevisedContent(n,doc,part);
  }
  for(const n of elements(doc,NS.w).filter(n=>['ins','moveTo','cellIns'].includes(n.localName))){
    if(!attached(n,doc))continue;const parent=n.parentNode,mark=isW(parent,'rPr')&&isW(parent.parentNode,'pPr');
    if(mark||isW(parent,'trPr')&&n.localName==='ins'||isW(parent,'tcPr')&&n.localName==='cellIns'||isW(parent,'numPr')&&n.localName==='ins'){parent.removeChild(n);continue;}
    if(n.localName==='cellIns'||['trPr','tcPr','rPr'].some(name=>isW(parent,name))||parent?.namespaceURI!==NS.w)throw revisionError(part,n.localName);
    unwrap(n);
  }
  // Empty rows/tables created by accepted cell deletion must not survive as
  // invalid shells. Existing untouched empty structures are not rewritten.
  for(const row of affectedRows)if(row&&attached(row,doc)&&!elements(row,NS.w,'tc').length)removeRevisedContent(row,doc,part);
  for(const table of [...affectedTables].reverse())if(table&&attached(table,doc)&&!elements(table,NS.w,'tr').length)removeRevisedContent(table,doc,part);
  for(const p of paragraphMarks){
    if(!attached(p,doc))continue;
    let next=p.nextSibling;while(next&&next.nodeType!==1)next=next.nextSibling;
    if(!isW(next,'p')||elements(p,NS.w,'sectPr').length)throw revisionError(part,'删除段落标记跨越表格、容器或分节边界');
    const anchor=children(next).find(n=>!isW(n,'pPr'))||null;
    for(const child of Array.from(p.childNodes))if(!isW(child,'pPr'))next.insertBefore(child,anchor);
    p.parentNode.removeChild(p);
  }
  for(const n of elements(doc,NS.w,'cellMerge'))acceptCellMerge(n,part);
  const unresolved=elements(doc).find(n=>isRevision(n)||/^conflict(?:Ins|Del)$|^customXmlConflict/.test(n.localName));
  if(unresolved)throw revisionError(part,unresolved.localName);
  return count;
}

async function xmlPart(zip,name) {
  const entry=zip.file(name); if (!entry) throw new Error(`Office 文件缺少必要部分：${name}`);
  return parseXml(await entry.async('string'));
}
function resolvePart(base,target) {
  if (!target || /^[a-z][a-z0-9+.-]*:|^\/\//i.test(target)) throw new Error('Office 内部关系地址无效。');
  const stack=target.startsWith('/')?[]:base.split('/').slice(0,-1);
  for (const p of target.split('/')) { if (!p || p==='.') continue; if(p==='..') {if(!stack.length) throw new Error('Office 内部关系越界。');stack.pop();} else stack.push(p); }
  return stack.join('/');
}

function collectPart(result,part,doc,page) {
  const ns=result.kind==='docx'?NS.w:NS.a;
  const allP=elements(doc,ns,'p');
  const tableNodes=elements(doc,ns,'tbl').filter(isVisibleBranch);
  const tableIds=new Map(tableNodes.map((t,i)=>[t,`${part}#table-${i}`]));
  const paragraphs=allP.map((node,index)=>({node,index})).filter(p=>isVisibleBranch(p.node)).map(({node,index})=>{
    const table=ancestor(node,ns,'tbl');const row=ancestor(node,ns,'tr');const cell=ancestor(node,ns,'tc');
    const shape=result.kind==='pptx' ? ancestor(node,NS.p,'sp') || ancestor(node,NS.p,'graphicFrame') : null;
    return {id:`${part}#p-${index}`,part,index,text:textOf(node,result.kind),node,page,tableId:tableIds.get(table)||null,rowIndex:row&&table?children(table,ns,'tr').indexOf(row):null,cellIndex:cell&&row?children(row,ns,'tc').indexOf(cell):null,shapeId:shape?elements(shape,NS.p,'cNvPr')[0]?.getAttribute('id') || null:null};
  });
  result.paragraphs.push(...paragraphs);
  for (const table of tableNodes) {
    const tableId=tableIds.get(table);
    children(table,ns,'tr').forEach((row,rowIndex)=>{
      const cells=children(row,ns,'tc').map((cell,cellIndex)=>{
        const ps=paragraphs.filter(p=>p.tableId===tableId && p.rowIndex===rowIndex && p.cellIndex===cellIndex);
        return {text:ps.map(p=>p.text).join('\n'),paragraphIds:ps.map(p=>p.id)};
      });
      result.rows.push({id:`${tableId}#row-${rowIndex}`,part,page,tableId,rowIndex,label:cells[0]?.text||'',value:cells.slice(1).map(c=>c.text).join('\n'),paragraphIds:cells.flatMap(c=>c.paragraphIds),cells});
    });
  }
  return paragraphs;
}

export async function readOfficeFile(file,{role='template'}={}) {
  const name=String(file?.name||'');const kind=name.split('.').pop().toLowerCase();
  if (!['docx','pptx'].includes(kind)) throw new Error('模板请使用 DOCX 或 PPTX 文件。旧 DOC 仅支持作为合同。');
  if (!globalThis.JSZip) throw new Error('本地文件读取组件未加载，请刷新网页。');
  const bytes=copyBytes(file.bytes);const hash=await sha256(bytes);
  let zip;try {zip=await globalThis.JSZip.loadAsync(bytes);} catch {throw new Error('Office 文件无法读取：文件可能已损坏、加密或扩展名不正确。');}
  const result={id:`${kind}-${hash.slice(0,20)}`,name,kind,sha256:hash,bytes,zip,parts:new Map(),paragraphs:[],rows:[],slides:[],warnings:[],preparation:{acceptedRevisions:0,changedParts:[]}};
  if(kind==='docx') {
    if(!zip.file('word/document.xml')) throw new Error('此文件不包含可读取的 Word 正文。');
    const stories=['word/document.xml',...Object.keys(zip.files).filter(n=>/^word\/(header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(n)).sort()];
    // Revision history can also live in styles, numbering, comments or glossary
    // parts. Normalize those parts for output without treating them as evidence.
    const allParts=[...stories,...Object.keys(zip.files).filter(n=>/^word\/.*\.xml$/.test(n)&&!stories.includes(n)).sort()];
    for(const part of allParts) {
      const doc=await xmlPart(zip,part);
      const count=acceptRevisions(doc,part);
      if(count) {result.preparation.acceptedRevisions+=count;result.preparation.changedParts.push(part);}
      result.parts.set(part,doc);if(stories.includes(part))collectPart(result,part,doc,null);
    }
    if (result.preparation.acceptedRevisions) result.warnings.push(`已按当前视图接受 ${result.preparation.acceptedRevisions} 项旧修订；原文件不变，旧批注保留。`);
  } else {
    const presentation=await xmlPart(zip,'ppt/presentation.xml');const relationships=await xmlPart(zip,'ppt/_rels/presentation.xml.rels');
    const relMap=new Map(elements(relationships,NS.pr,'Relationship').map(r=>[r.getAttribute('Id'),r]));
    const ids=elements(presentation,NS.p,'sldId');
    if(!ids.length) throw new Error('演示文稿没有可读取的幻灯片。');
    for(const [i,id] of ids.entries()) {
      const rel=relMap.get(id.getAttributeNS(NS.r,'id'));
      if(!rel || rel.getAttribute('TargetMode')==='External' || !rel.getAttribute('Type').endsWith('/slide')) throw new Error('演示文稿的页面关系不完整。');
      const part=resolvePart('ppt/presentation.xml',rel.getAttribute('Target'));const doc=await xmlPart(zip,part);
      result.parts.set(part,doc);const paragraphs=collectPart(result,part,doc,i+1);
      result.slides.push({page:i+1,part,title:paragraphs.find(p=>p.text.trim())?.text||'',text:paragraphs.map(p=>p.text).join('\n'),paragraphIds:paragraphs.map(p=>p.id)});
    }
  }
  if(!result.paragraphs.length) throw new Error('没有找到可读取的正文；暂不支持图片扫描件或此 Office 结构。');
  return result;
}

export const readTemplate=file=>readOfficeFile(file,{role:'template'});

export async function readContract(file) {
  const ext=String(file?.name||'').split('.').pop().toLowerCase();
  if(ext==='docx') {
    const doc=await readOfficeFile(file,{role:'contract'});
    return {name:doc.name,sha256:doc.sha256,blocks:doc.paragraphs.filter(p=>p.text.trim()).map(p=>({id:p.id,text:p.text,part:p.part,paragraphIndex:p.index,page:p.page,...(p.tableId?{tableId:p.tableId,rowIndex:p.rowIndex,cellIndex:p.cellIndex}:{})})),warnings:doc.warnings};
  }
  if(ext!=='doc') throw new Error('合同请使用 DOCX 或旧 DOC 文件。');
  const bytes=copyBytes(file.bytes);
  if(typeof globalThis.docToText!=='function') throw new Error('旧 DOC 读取组件未加载，请刷新网页。');
  const stories=globalThis.docToText.sections(bytes);
  if(!stories || !stories.body?.trim()) throw new Error('旧 DOC 读取失败：暂不支持加密、损坏或 Word 97 之前的格式；请另存为 DOCX。');
  const blocks=[];
  // Annotation stories are comments, not contract evidence.
  for(const part of ['body','footnotes','endnotes','headers','textboxes','headerTextboxes']) {
    const text=stories[part]; if(typeof text!=='string') continue;
    text.split(/\r\n?|\n/).forEach((text,index)=>{if(text.trim()) blocks.push({id:`doc-${part}#p-${index}`,text,part,paragraphIndex:index,page:null});});
  }
  return {name:file.name,sha256:await sha256(bytes),blocks,warnings:['旧 DOC 已在本机提取当前视图文本；页码和原版式不用于合同核对。']};
}
