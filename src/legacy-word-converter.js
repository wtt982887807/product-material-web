import {normalizeLegacyDocShading} from './doc-shading.js';
import {NS,parseXml,serializeXml,elements} from './office-xml.js';
import {resolveTarget} from './office-package.js';

const MAX_BYTES = 60 * 1024 * 1024;
const TIMEOUT_MS = 240000;

function report(callback, stage, message) {
  try { callback?.({stage, message}); } catch { /* Progress must not alter a conversion. */ }
}

function inputCopy({name, bytes} = {}) {
  const extension = String(name || '').match(/\.(docx?|rtf)$/i)?.[1].toLowerCase();
  if (!extension) throw new Error('本地转换支持 DOC、RTF 和需要兼容处理的 DOCX；其他格式请先在 Word 中另存为 DOCX。');
  if (bytes instanceof ArrayBuffer) bytes = new Uint8Array(bytes);
  if (!(bytes instanceof Uint8Array) || !bytes.byteLength) throw new Error('待转换文件内容为空。');
  if (bytes.byteLength > MAX_BYTES) throw new Error('待转换文件请控制在 60 MB 以内。');
  const original=bytes.slice(),warnings=[];let prepared=original;
  try{prepared=normalizeLegacyDocShading(original);}
  catch(error){
    if(error.code!=='DOC_SHADING_UNSUPPORTED')throw error;
    warnings.push({code:'DOC_SHADING_FALLBACK',message:`旧版 Word 底纹兼容预处理未完成（${error.reason}），已将原始副本交由完整 Word 引擎转换；导出后请核对表格底色。`});
  }
  return {name: String(name), extension, bytes: prepared,warnings};
}

async function repairConvertedCommentReferences(zip,onWarning){
  const commentsParts=new Set(),relations=zip.file('word/_rels/document.xml.rels');
  if(relations){
    const doc=parseXml(await relations.async('string'));
    for(const rel of elements(doc,NS.pr,'Relationship'))if(rel.getAttribute('Type')===NS.r+'/comments'){
      if(rel.getAttribute('TargetMode')==='External')throw new Error('转换结果的批注指向外部文件。');
      const part=resolveTarget('word/document.xml',rel.getAttribute('Target'));
      if(!zip.file(part))throw new Error('转换结果缺少已声明的批注内容。');
      commentsParts.add(part);
    }
  }
  // LibreOffice normally uses this part; keep its contents even if a producer
  // omitted the relationship instead of mistaking those notes for empty ones.
  if(zip.file('word/comments.xml'))commentsParts.add('word/comments.xml');
  const key=node=>{const id=node.getAttributeNS(NS.w,'id');if(!/^[+-]?\d+$/.test(id||''))throw new Error('转换结果的旧批注编号无效。');return BigInt(id).toString();};
  const known=new Set();
  for(const part of commentsParts){
    const doc=parseXml(await zip.file(part).async('string'));
    if(doc.documentElement.namespaceURI!==NS.w||doc.documentElement.localName!=='comments')throw new Error('转换结果的批注结构无效。');
    for(const comment of elements(doc,NS.w,'comment'))known.add(key(comment));
  }
  const orphaned=new Set();let changed=false;
  for(const part of Object.keys(zip.files).filter(name=>/^word\/(?:document|header[^/]*|footer[^/]*|footnotes|endnotes)\.xml$/.test(name))){
    const doc=parseXml(await zip.file(part).async('string'));let touched=false;
    for(const local of ['commentRangeStart','commentRangeEnd','commentReference'])for(const node of elements(doc,NS.w,local)){
      const id=key(node);if(known.has(id))continue;
      orphaned.add(id);node.parentNode.removeChild(node);touched=true;
    }
    if(touched){zip.file(part,serializeXml(doc));changed=true;}
  }
  if(changed)onWarning({code:'WORD_ORPHAN_COMMENTS_REPAIRED',message:`完整转换后发现 ${orphaned.size} 个没有对应内容的旧批注引用，已清除这些空引用；现有正文、格式和有效旧批注保留。`});
  return changed;
}

async function validateDocx(bytes,onWarning) {
  if (!(bytes instanceof Uint8Array) || bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) throw new Error('转换结果不是有效 DOCX，已停止导出。');
  try {
    if (!globalThis.JSZip) throw new Error('DOCX 检查组件未加载');
    const zip = await globalThis.JSZip.loadAsync(bytes);
    const types = await zip.file('[Content_Types].xml')?.async('string');
    if (!zip.file('word/document.xml') || !types?.includes('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml') || zip.file('word/vbaProject.bin')) throw new Error('缺少 DOCX 正文或格式不符');
    if(await repairConvertedCommentReferences(zip,onWarning))return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  } catch (error) { throw new Error('转换结果未通过 DOCX 结构检查：' + error.message); }
  return bytes;
}

/** A single serial queue owns one isolated iframe and its LibreOffice workers. */
export function createLegacyWordConverter({
  createRuntime = createIframeRuntime,
  isIsolated = () => globalThis.crossOriginIsolated === true && typeof SharedArrayBuffer !== 'undefined',
  timeoutMs = TIMEOUT_MS,
} = {}) {
  let queue = Promise.resolve(), runtime = null, generation = 0;
  function reset() { generation++; runtime?.destroy(); runtime = null; }
  return function convert(file, {onProgress,onWarning} = {}) {
    let input;
    try { input = inputCopy(file); } catch (error) { return Promise.reject(error); }
    const result = queue.then(async () => {
      if (!isIsolated()) throw new Error('浏览器尚未启用本地转换所需的隔离环境。请刷新页面，并使用最新版桌面 Chrome 或 Edge 重试；原文件未改动。');
      const warn=warning=>{
        report(onProgress,'warning',warning.message);
        try{onWarning?.(warning);}catch{/* A warning consumer cannot stop conversion. */}
      };
      input.warnings.forEach(warn);
      const current = generation;
      let timer;
      const work = (async () => {
        if (!runtime) {
          report(onProgress, 'loading', '首次使用正在加载本地 Word 转换引擎，约 50 MB；文件仍留在本机。');
          const created = await createRuntime();
          if (current !== generation) { created.destroy(); throw new Error('转换已结束，请重新尝试。'); }
          runtime = created;
        }
        const bytes = await runtime.convert(input, {onProgress});
        return validateDocx(bytes,warn);
      })();
      try {
        return await Promise.race([work, new Promise((_, reject) => {
          timer = setTimeout(() => reject(new Error('本地 Word 转换超时，已释放本次转换。请使用桌面浏览器重试，或先用 Word 另存为 DOCX。')), timeoutMs);
        })]);
      } catch (error) { reset(); throw error; }
      finally { clearTimeout(timer); }
    });
    queue = result.catch(() => {});
    return result;
  };
}

function createIframeRuntime() {
  const frame = document.createElement('iframe');
  frame.hidden = true;
  frame.title = '本机 Word 格式转换';
  frame.setAttribute('aria-hidden', 'true');
  frame.setAttribute('allow', 'cross-origin-isolated');
  frame.src = new URL('../converter/frame.html', import.meta.url).href;
  const origin = new URL(frame.src).origin;
  const channel = new MessageChannel();
  let id = 0, pending = null, destroyed = false, readyResolve, readyReject;
  const ready = new Promise((resolve, reject) => { readyResolve = resolve; readyReject = reject; });
  // It can reject before convert starts awaiting startup.
  ready.catch(() => {});
  function fail(message) {
    const error = new Error(message || '本地转换引擎无法启动，请使用桌面浏览器重试。');
    readyReject(error); pending?.reject(error); pending = null;
  }
  channel.port1.onmessage = event => {
    const message = event.data || {};
    if (message.type === 'ready') readyResolve();
    else if (message.type === 'fatal') fail(message.message);
    else if (message.type === 'progress') report(pending?.onProgress, message.stage || 'loading', message.message);
    else if (pending && message.id === pending.id) {
      const request = pending; pending = null;
      if (message.type === 'result') request.resolve(new Uint8Array(message.bytes));
      else request.reject(new Error(message.message || 'Word 转换未成功，原文件未改动。'));
    }
  };
  frame.onload = () => {
    if (!destroyed) frame.contentWindow.postMessage({type: 'initialize-local-word-converter'}, origin, [channel.port2]);
  };
  frame.onerror = () => fail('本地转换组件加载失败，请刷新页面后重试。');
  document.body.append(frame);
  return {
    async convert(input, {onProgress} = {}) {
      if (destroyed) throw new Error('转换引擎已释放。');
      // Startup progress has no request id, but belongs to this sole queued call.
      report(onProgress, 'loading', '正在准备本地转换引擎；首次加载约 50 MB。');
      await ready;
      if (destroyed) throw new Error('转换引擎已释放。');
      return new Promise((resolve, reject) => {
        const next = ++id;
        pending = {id: next, resolve, reject, onProgress};
        report(onProgress, 'convert', '正在转换 Word 格式并接受已有修订。');
        const bytes = input.bytes.slice();
        channel.port1.postMessage({type: 'convert', id: next, extension: input.extension, bytes: bytes.buffer}, [bytes.buffer]);
      });
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      fail('转换引擎已释放。');
      channel.port1.close(); channel.port2.close(); frame.remove();
    },
  };
}

export const convertLegacyWord = createLegacyWordConverter();
