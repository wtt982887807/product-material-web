import {readTemplate,readContract} from './office-reader.js?v=20261027-holding-fees';
import {sha256,parseXml,elements} from './office-xml.js';

const zipSignature=bytes=>bytes[0]===0x50&&bytes[1]===0x4b;
const oleSignature=bytes=>[0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1].every((v,i)=>bytes[i]===v);
const rtfSignature=bytes=>/^\s*\{\\rtf\d/.test(new TextDecoder().decode(bytes.slice(0,80)));
const docxName=name=>String(name).replace(/\.[^.]+$/,'')+'.docx';
async function fullConversion(file,{onProgress,onWarning}){const {convertLegacyWord}=await import('./legacy-word-converter.js');return convertLegacyWord(file,{onWarning,onProgress:info=>onProgress('正在处理 Word 模板',info.message||String(info))});}
async function plainDocx(bytes){
  if(!zipSignature(bytes))return false;
  let zip;try{zip=await globalThis.JSZip.loadAsync(bytes);}catch{throw new Error('Word 文件包损坏，无法读取。');}
  if(!zip.file('word/document.xml')||!zip.file('[Content_Types].xml'))throw new Error('文件包不包含有效的 Word 正文。');
  const types=parseXml(await zip.file('[Content_Types].xml').async('string'));
  const main=elements(types,'http://schemas.openxmlformats.org/package/2006/content-types','Override').find(n=>n.getAttribute('PartName')==='/word/document.xml');
  if(main?.getAttribute('ContentType')!=='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')return false;
  if(Object.keys(zip.files).some(n=>/(?:^|\/)(?:vbaProject|vbaData)(?:[./]|$)/i.test(n)))return false;
  for(const name of Object.keys(zip.files).filter(n=>n.endsWith('.rels')))if(/\/(?:vbaProject|vbaData)["'\s<]/i.test(await zip.file(name).async('string')))return false;
  return true;
}

/** Prepare a copy of a template; ordinary DOCX/PPTX never start the large engine. */
export async function readWordTemplate(file,{convert=fullConversion,onProgress=()=>{}}={}) {
  const name=String(file?.name||'');
  const bytes=file.bytes instanceof ArrayBuffer?new Uint8Array(file.bytes):new Uint8Array(file.bytes.buffer,file.bytes.byteOffset,file.bytes.byteLength);
  if(/\.pptx$/i.test(name))return readTemplate({name,bytes});
  const sourceSha256=await sha256(bytes);
  const normalizedName=docxName(name);
  const conversionWarnings=[],conversionOptions={onProgress,onWarning:warning=>conversionWarnings.push(warning)};
  let prepared=bytes,converted=false,complex=false,recognized=false;
  if(zipSignature(bytes)) {
    recognized=!/\.docx$/i.test(name);
    if(!await plainDocx(bytes)){onProgress('正在规范 Word 格式','将模板转换为普通 DOCX，并接受旧修订。');prepared=await convert({name:normalizedName,bytes:bytes.slice()},conversionOptions);converted=true;}
  } else if(oleSignature(bytes)||rtfSignature(bytes)) {
    onProgress('正在转换 Word 模板','首次使用需下载约 50 MB 转换组件；文件仍在本机处理。');
    prepared=await convert({name,bytes:bytes.slice()},conversionOptions);converted=true;
  } else throw new Error('无法识别此 Word 文件，文件可能损坏；支持 DOCX、旧版 DOC 和 RTF。');
  if(!zipSignature(prepared)||converted&&!await plainDocx(prepared))throw new Error('Word 转换未生成有效 DOCX，请重试；原模板未改动。');
  let result;
  try{result=await readTemplate({name:normalizedName,bytes:prepared});}
  catch(error){
    if(converted||error.code!=='WORD_REVISION_COMPLEX')throw error;
    onProgress('正在接受复杂旧修订','使用完整转换组件接受旧修订后，再生成本轮修订。');
    prepared=await convert({name:normalizedName,bytes:bytes.slice()},conversionOptions);
    if(!zipSignature(prepared)||!await plainDocx(prepared))throw new Error('接受旧修订后未生成有效 DOCX，请重试。');
    result=await readTemplate({name:normalizedName,bytes:prepared});converted=true;complex=true;
  }
  result.sourceSha256=sourceSha256;result.originalName=name;
  if(converted)result.warnings.unshift(complex?'已自动接受复杂旧修订并转换为 DOCX，再生成本轮修订。':'已将 Word 模板自动转换为 DOCX 并接受旧修订；原文件不变。');
  else if(recognized)result.warnings.unshift('已按实际 DOCX 内容识别文件，输出统一使用 .docx 后缀。');
  result.warnings.push(...conversionWarnings);
  return result;
}

/** Contract text uses the quick DOC reader first; Office conversion is fallback. */
export async function readWordContract(file,{convert=fullConversion,onProgress=()=>{}}={}) {
  const name=String(file?.name||'');
  if(!/\.docx?$/i.test(name))throw new Error('合同请使用 DOCX 或旧 DOC 文件。');
  const bytes=file.bytes instanceof ArrayBuffer?new Uint8Array(file.bytes):ArrayBuffer.isView(file.bytes)?new Uint8Array(file.bytes.buffer,file.bytes.byteOffset,file.bytes.byteLength):new Uint8Array();
  if(!zipSignature(bytes)&&!oleSignature(bytes)&&!rtfSignature(bytes))throw new Error('无法识别合同正文，文件可能损坏；请使用有效 Word 合同。');
  let fastFallback=false;
  if(oleSignature(bytes)){
    try{
      const result=await readContract({name:name.replace(/\.docx?$/i,'.doc'),bytes});
      if(result.blocks.length)return {...result,name,sourceSha256:result.sha256};
    }catch{/* A recognized legacy container may need the complete Office reader. */}
    fastFallback=true;onProgress('正在兼容读取合同','快速读取未取得正文，正在使用完整 Word 引擎读取；原合同不变。');
  }
  const doc=await readWordTemplate({name,bytes},{convert,onProgress:(title,detail)=>onProgress(title.replace(/模板/g,'合同'),String(detail).replace(/模板/g,'合同'))});
  const blocks=doc.paragraphs.filter(p=>p.text.trim()).map(p=>({id:p.id,text:p.text,part:p.part,paragraphIndex:p.index,page:p.page}));
  if(!blocks.length)throw new Error('合同未包含可读取的正文；暂不支持空文档或图片扫描件。');
  const warnings=[...doc.warnings];
  if(fastFallback)warnings.unshift('旧 DOC 合同已通过完整 Word 引擎取得正文；原合同文件不变。');
  return {name,sha256:doc.sourceSha256,sourceSha256:doc.sourceSha256,preparedSha256:doc.sha256,blocks,warnings};
}
