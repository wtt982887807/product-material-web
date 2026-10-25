import {readTemplate} from './office-reader.js';
import {sha256,parseXml,elements} from './office-xml.js';

const zipSignature=bytes=>bytes[0]===0x50&&bytes[1]===0x4b;
const oleSignature=bytes=>[0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1].every((v,i)=>bytes[i]===v);
const rtfSignature=bytes=>/^\s*\{\\rtf\d/.test(new TextDecoder().decode(bytes.slice(0,80)));
const docxName=name=>String(name).replace(/\.[^.]+$/,'')+'.docx';
async function fullConversion(file,{onProgress}){const {convertLegacyWord}=await import('./legacy-word-converter.js');return convertLegacyWord(file,{onProgress:info=>onProgress('正在处理 Word 模板',info.message||String(info))});}
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
  let prepared=bytes,converted=false,complex=false,recognized=false;
  if(zipSignature(bytes)) {
    recognized=!/\.docx$/i.test(name);
    if(!await plainDocx(bytes)){onProgress('正在规范 Word 格式','将模板转换为普通 DOCX，并接受旧修订。');prepared=await convert({name:normalizedName,bytes:bytes.slice()},{onProgress});converted=true;}
  } else if(oleSignature(bytes)||rtfSignature(bytes)) {
    onProgress('正在转换 Word 模板','首次使用需下载约 50 MB 转换组件；文件仍在本机处理。');
    prepared=await convert({name,bytes:bytes.slice()},{onProgress});converted=true;
  } else throw new Error('无法识别此 Word 文件，文件可能损坏；支持 DOCX、旧版 DOC 和 RTF。');
  if(!zipSignature(prepared)||converted&&!await plainDocx(prepared))throw new Error('Word 转换未生成有效 DOCX，请重试；原模板未改动。');
  let result;
  try{result=await readTemplate({name:normalizedName,bytes:prepared});}
  catch(error){
    if(converted||error.code!=='WORD_REVISION_COMPLEX')throw error;
    onProgress('正在接受复杂旧修订','使用完整转换组件接受旧修订后，再生成本轮修订。');
    prepared=await convert({name:normalizedName,bytes:bytes.slice()},{onProgress});
    if(!zipSignature(prepared)||!await plainDocx(prepared))throw new Error('接受旧修订后未生成有效 DOCX，请重试。');
    result=await readTemplate({name:normalizedName,bytes:prepared});converted=true;complex=true;
  }
  result.sourceSha256=sourceSha256;result.originalName=name;
  if(converted)result.warnings.unshift(complex?'已自动接受复杂旧修订并转换为 DOCX，再生成本轮修订。':'已将 Word 模板自动转换为 DOCX 并接受旧修订；原文件不变。');
  else if(recognized)result.warnings.unshift('已按实际 DOCX 内容识别文件，输出统一使用 .docx 后缀。');
  return result;
}
