export const NS = Object.freeze({
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  xml: 'http://www.w3.org/XML/1998/namespace',
});

export function parseXml(text) {
  if (typeof text !== 'string' || /<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('XML 内容无效或含有不支持的外部声明。');
  let doc;
  try {
    const parser = new DOMParser({onError: (level, message) => { if (level !== 'warning') throw new Error(message); }});
    doc = parser.parseFromString(text, 'application/xml');
  } catch (error) { throw new Error(`XML 读取失败：${error.message}`); }
  if (!doc.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('XML 读取失败：文件结构不完整。');
  return doc;
}

export function serializeXml(doc) { return new XMLSerializer().serializeToString(doc); }

export function children(node, ns, local) {
  const out=[];
  for (let c=node?.firstChild;c;c=c.nextSibling) if (c.nodeType===1 && (ns==null || ns==='*' || c.namespaceURI===ns) && (local==null || local==='*' || c.localName===local)) out.push(c);
  return out;
}

export function elements(node, ns, local) {
  const out=[];
  function visit(parent) { for (let c=parent?.firstChild;c;c=c.nextSibling) if (c.nodeType===1) { if ((ns==null || ns==='*' || c.namespaceURI===ns) && (local==null || local==='*' || c.localName===local)) out.push(c); visit(c); } }
  visit(node); return out;
}

export function ancestor(node, ns, local) {
  for (let c=node?.parentNode;c;c=c.parentNode) if (c.namespaceURI===ns && c.localName===local) return c;
  return null;
}

// AlternateContent carries the same shape twice. Modern Office uses Choice.
export function isVisibleBranch(node) {
  for (let c=node;c;c=c.parentNode) {
    if (c.namespaceURI===NS.w && ['del','moveFrom'].includes(c.localName)) return false;
    if (c.namespaceURI===NS.mc && c.localName==='Fallback' && children(c.parentNode,NS.mc,'Choice').length) return false;
  }
  return true;
}

export function textOf(paragraph, kind='docx') {
  const ns=kind==='pptx'?NS.a:NS.w;
  let out='';
  function visit(node) {
    if (node!==paragraph && node.namespaceURI===ns && node.localName==='p') return;
    if (!isVisibleBranch(node)) return;
    if (node.namespaceURI===ns) {
      if (node.localName==='t') { out+=node.textContent; return; }
      if (node.localName==='tab') { out+='\t'; return; }
      if (['br','cr'].includes(node.localName)) { out+='\n'; return; }
      if (node.localName==='noBreakHyphen') { out+='\u2011'; return; }
      if (node.localName==='softHyphen') { out+='\u00ad'; return; }
      if (['instrText','delText'].includes(node.localName)) return;
    }
    for (let c=node.firstChild;c;c=c.nextSibling) if (c.nodeType===1) visit(c);
  }
  visit(paragraph); return out;
}

export async function sha256(bytes) {
  if (!globalThis.crypto?.subtle) throw new Error('请在 HTTPS 网页或本机服务中打开，以校验文件完整性。');
  const hash=await crypto.subtle.digest('SHA-256',bytes);
  return Array.from(new Uint8Array(hash),b=>b.toString(16).padStart(2,'0')).join('');
}
