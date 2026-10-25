import {ZetaHelperMain} from './vendor/zetajs-1.2.0/zetaHelper.js';

// Official recommended endpoint is rolling; do not describe it as versioned.
const ENGINE_BASE = 'https://cdn.zetaoffice.net/zetaoffice_latest/';
let port, helper, active = null, ready = false;
function send(message, transfer) { port?.postMessage(message, transfer || []); }
function cleanup() {
  if (!active) return;
  for (const path of [active.from, active.to]) { try { helper.FS.unlink(path); } catch {} }
  active = null;
}
function fatal(message) { cleanup(); send({type: 'fatal', message: message || '本地转换引擎启动失败；请使用桌面浏览器重试。'}); }

window.addEventListener('message', event => {
  if (port || event.source !== parent || event.origin !== location.origin || event.data?.type !== 'initialize-local-word-converter' || !event.ports[0]) return;
  port = event.ports[0];
  if (!crossOriginIsolated || typeof SharedArrayBuffer === 'undefined') { fatal('本地转换隔离环境未就绪，请刷新网页后使用桌面浏览器重试。'); return; }
  port.onmessage = event => {
    const message = event.data || {};
    if (message.type !== 'convert') return;
    if (!ready || active || !Number.isSafeInteger(message.id) || message.id < 1 || !['doc', 'rtf', 'docx'].includes(message.extension)) {
      send({type: 'error', id: message.id, message: '转换请求无效或已有文件正在转换。'}); return;
    }
    active = {id: message.id, from: `/tmp/material-input-${message.id}.${message.extension}`, to: `/tmp/material-output-${message.id}.docx`};
    try {
      helper.FS.writeFile(active.from, new Uint8Array(message.bytes));
      helper.thrPort.postMessage({type: 'convert', ...active});
    } catch (error) { cleanup(); send({type: 'error', id: message.id, message: '无法准备本地转换文件：' + error.message}); }
  };
  try {
    helper = new ZetaHelperMain(new URL('./office-thread.js', import.meta.url).href, {
      threadJsType: 'module', wasmPkg: 'url:' + ENGINE_BASE, blockPageScroll: false,
    });
    helper.Module.onAbort = reason => fatal('本地 Word 引擎无法运行，可能是网络或浏览器内存不足。请使用桌面浏览器重试。');
    helper.start(() => {
      helper.thrPort.onmessage = event => {
        const message = event.data || {};
        if (message.type === 'worker-ready') { ready = true; send({type: 'ready'}); return; }
        if (!active || message.id !== active.id) return;
        const id = active.id;
        if (message.type === 'converted') {
          try {
            const bytes = helper.FS.readFile(active.to).slice();
            cleanup(); send({type: 'result', id, bytes: bytes.buffer}, [bytes.buffer]);
          } catch (error) { cleanup(); send({type: 'error', id, message: '无法读取转换结果：' + error.message}); }
        } else if (message.type === 'error') { cleanup(); send(message); }
      };
    });
  } catch (error) { fatal('本地 Word 引擎加载失败：' + error.message); }
});
window.addEventListener('error', event => fatal('本地转换组件加载失败，请检查网络后使用桌面浏览器重试。'));
window.addEventListener('unhandledrejection', event => fatal('本地 Word 引擎未能启动或转换失败，请使用桌面浏览器重试。'));
