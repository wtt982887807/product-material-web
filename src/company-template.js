import {NS,parseXml,serializeXml,children,elements,textOf} from './office-xml.js';

const labels={general:'通用版',ccb:'建行版'};
const headingVariants=new Map([['通用版本','general'],['通用版','general'],['建行版本','ccb'],['建行版','ccb']]);
const compact=text=>String(text).normalize('NFKC').replace(/\s/g,'');
const privatePart=path=>/^(?:docProps|customXml)(?:\/|$)/i.test(path.replace(/^\//,''));

function copiedBytes(input){
 if(input instanceof ArrayBuffer)return new Uint8Array(input.slice(0));
 if(ArrayBuffer.isView(input))return new Uint8Array(input.buffer,input.byteOffset,input.byteLength).slice();
 throw new Error('未取得公司介绍文件内容，请重新上传。');
}
function resolvedTarget(relsPart,target){
 const base=relsPart==='_rels/.rels'?'':relsPart.replace(/_rels\/[^/]+\.rels$/,'');
 const parts=(target.startsWith('/')?target.slice(1):base+target).split('/'),result=[];
 for(const part of parts){if(part==='..')result.pop();else if(part&&part!=='.')result.push(part);}
 return result.join('/');
}
async function removeMetadata(zip){
 for(const path of Object.keys(zip.files))if(privatePart(path))zip.remove(path);
 for(const path of Object.keys(zip.files).filter(path=>path.endsWith('.rels'))){
  const doc=parseXml(await zip.file(path).async('string'));let changed=false;
  for(const rel of elements(doc,NS.pr,'Relationship')){
   if(privatePart(resolvedTarget(path,rel.getAttribute('Target')))){
    rel.parentNode.removeChild(rel);changed=true;
   }
  }
  if(changed)zip.file(path,serializeXml(doc));
 }
 const entry=zip.file('[Content_Types].xml');
 if(entry){
  const doc=parseXml(await entry.async('string'));
  for(const override of elements(doc,NS.ct,'Override'))if(privatePart(override.getAttribute('PartName')))override.parentNode.removeChild(override);
  zip.file('[Content_Types].xml',serializeXml(doc));
 }
}

async function removeUnusedNotes(zip,document){
 const sources=[document];
 // A shared header/footer can also contain a note that is still displayed.
 for(const path of Object.keys(zip.files).filter(path=>/^word\/(?:header|footer)[^/]*\.xml$/.test(path)))sources.push(parseXml(await zip.file(path).async('string')));
 for(const [kind,part]of [['footnote','word/footnotes.xml'],['endnote','word/endnotes.xml'],['comment','word/comments.xml']]){
  const entry=zip.file(part);if(!entry)continue;
  const ids=new Set(sources.flatMap(source=>elements(source,NS.w,`${kind}Reference`).map(node=>node.getAttributeNS(NS.w,'id'))));
  const doc=parseXml(await entry.async('string'));
  for(const note of children(doc.documentElement,NS.w,kind)){
   const type=note.getAttributeNS(NS.w,'type');
   if(!ids.has(note.getAttributeNS(NS.w,'id'))&&!['separator','continuationSeparator','continuationNotice'].includes(type))note.parentNode.removeChild(note);
  }
  zip.file(part,serializeXml(doc));
 }
}

// Only explicit, standalone version headings select a variant. A single-version
// upload without them stays byte-for-byte intact for the later profile check.
export async function prepareCompanyTemplate(file,{variant='general'}={}){
 if(!Object.hasOwn(labels,variant))throw new Error('公司介绍版本无效，请选择通用版本或建行版本。');
 if(!/\.docx$/i.test(file?.name||''))throw new Error('公司介绍请使用 DOCX 文件。');
 const bytes=copiedBytes(file.bytes),zip=await JSZip.loadAsync(bytes),entry=zip.file('word/document.xml');
 if(!entry)throw new Error('公司介绍 DOCX 内容不完整，请重新选择文件。');
 const doc=parseXml(await entry.async('string')),body=elements(doc,NS.w,'body')[0];
 if(!body)throw new Error('公司介绍 DOCX 缺少正文，请重新选择文件。');
 const nodes=children(body),markers=nodes.flatMap((node,index)=>{
  const version=node.namespaceURI===NS.w&&node.localName==='p'&&headingVariants.get(compact(textOf(node)));
  return version?[{index,variant:version}]:[];
 });
 if(!markers.length)return{name:file.name,bytes};
 const ranges=markers.flatMap((marker,index)=>marker.variant===variant?[{start:marker.index,end:markers[index+1]?.index??nodes.length}]:[]);
 if(!ranges.length)throw new Error(`公司介绍文件中未找到${labels[variant]}本，请上传包含该版本的文件。`);
 for(const [index,node]of nodes.entries()){
  if(node.namespaceURI===NS.w&&node.localName==='sectPr')continue;
  if(!ranges.some(range=>index>=range.start&&index<range.end))body.removeChild(node);
 }
 zip.file('word/document.xml',serializeXml(doc));
 await removeUnusedNotes(zip,doc);
 await removeMetadata(zip);
 // Fixed package timestamps make the public downloads reproducible and avoid
 // carrying a local edit timestamp into a newly prepared template.
 for(const entry of Object.values(zip.files))entry.date=new Date('2000-01-01T00:00:00Z');
 const stem=file.name.replace(/\.docx$/i,'').replace(/[（(](?:通用|建行)版本?[）)]\s*$/,'').trim();
 return{name:`${stem}（${labels[variant]}）.docx`,bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'})};
}
