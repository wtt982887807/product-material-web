import {NS,parseXml,elements,children,ancestor,isVisibleBranch,textOf,sha256} from './office-xml.js';

function copyBytes(input) {
  if (input instanceof ArrayBuffer) return new Uint8Array(input.slice(0));
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer,input.byteOffset,input.byteLength).slice();
  throw new Error('未取得文件内容，请重新选择文件。');
}

function revisionError(part,type) { return new Error(`“${part}”含暂不支持的结构修订（${type}）。请在 Word 中接受旧修订后另存为 DOCX，再重新选择。`); }

/** Accept only revisions whose current view can be preserved without rebuilding layout. */
function acceptRevisions(doc,part) {
  const propertyChanges=new Set(['rPrChange','pPrChange','tblPrChange','tblGridChange','trPrChange','tcPrChange','sectPrChange']);
  const revisionNodes=elements(doc,NS.w).filter(n=> /Change$/.test(n.localName) || /^(ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|numberingChange|customXml.*Range(Start|End)|move.*Range(Start|End))$/.test(n.localName)).filter(n=> {
    for(let parent=n.parentNode;parent;parent=parent.parentNode) if(parent.namespaceURI===NS.w && propertyChanges.has(parent.localName)) return false;
    return true;
  });
  if (!revisionNodes.length) return 0;
  const paragraphMarks=[];
  for (const n of revisionNodes) {
    if (propertyChanges.has(n.localName)) continue;
    if (n.localName==='ins' || n.localName==='del') {
      const parent=n.parentNode;
      const mark=parent?.localName==='rPr' && parent.parentNode?.localName==='pPr';
      if (mark) { if (n.localName==='del') paragraphMarks.push(parent.parentNode.parentNode); continue; }
      if (parent?.localName==='trPr' || parent?.localName==='tcPr' || parent?.localName==='rPr' || elements(n,NS.w).some(e=>['p','tbl','tr','tc','sectPr'].includes(e.localName))) throw revisionError(part,n.localName);
      continue;
    }
    throw revisionError(part,n.localName);
  }
  // Paragraph marks can only join adjacent paragraphs within the same container.
  for (const p of paragraphMarks) {
    let next=p.nextSibling; while(next && next.nodeType!==1) next=next.nextSibling;
    if (!next || next.namespaceURI!==NS.w || next.localName!=='p' || elements(p,NS.w,'sectPr').length) throw revisionError(part,'删除段落标记');
  }
  // Remove old property snapshots first; they are not active document content.
  for (const n of revisionNodes.filter(n=>propertyChanges.has(n.localName))) n.parentNode?.removeChild(n);
  for (const n of revisionNodes.filter(n=>n.localName==='del')) {
    const parent=n.parentNode;if(!parent) continue;
    // Keep the existing comment itself and collapse a deleted anchor to this position.
    for(const anchor of elements(n,NS.w).filter(e=>['commentRangeStart','commentRangeEnd','commentReference'].includes(e.localName))) {
      if(anchor.localName==='commentReference') {
        const run=doc.createElementNS(NS.w,'w:r');
        const rPr=children(anchor.parentNode,NS.w,'rPr')[0];if(rPr) run.appendChild(rPr.cloneNode(true));
        run.appendChild(anchor);parent.insertBefore(run,n);
      } else parent.insertBefore(anchor,n);
    }
    parent.removeChild(n);
  }
  for (const n of revisionNodes.filter(n=>n.localName==='ins')) {
    const parent=n.parentNode; if (!parent) continue;
    while(n.firstChild) parent.insertBefore(n.firstChild,n);
    parent.removeChild(n);
  }
  for (const p of paragraphMarks) {
    if (!p.parentNode) continue;
    let next=p.nextSibling; while(next && next.nodeType!==1) next=next.nextSibling;
    const anchor=children(next).find(n=> !(n.namespaceURI===NS.w && n.localName==='pPr')) || null;
    for (const child of Array.from(p.childNodes)) if (!(child.namespaceURI===NS.w && child.localName==='pPr')) next.insertBefore(child,anchor);
    p.parentNode.removeChild(p);
  }
  return revisionNodes.length;
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
    for(const part of stories) {
      const doc=await xmlPart(zip,part);
      const count=acceptRevisions(doc,part);
      if(count) {result.preparation.acceptedRevisions+=count;result.preparation.changedParts.push(part);}
      result.parts.set(part,doc);collectPart(result,part,doc,null);
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
    return {name:doc.name,sha256:doc.sha256,blocks:doc.paragraphs.filter(p=>p.text.trim()).map(p=>({id:p.id,text:p.text,part:p.part,paragraphIndex:p.index,page:p.page})),warnings:doc.warnings};
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
