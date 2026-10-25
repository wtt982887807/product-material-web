import {ZetaHelperThread} from './vendor/zetajs-1.2.0/zetaHelper.js';
import {convertUnoDocument} from './uno-convert.js';

const helper = new ZetaHelperThread();
const port = helper.thrPort;
port.onmessage = event => {
  const message = event.data || {};
  if (message.type !== 'convert') return;
  try {
    convertUnoDocument({inputUrl: 'file://' + message.from, outputUrl: 'file://' + message.to}, helper);
    port.postMessage({type: 'converted', id: message.id});
  } catch (error) {
    let detail = error?.message || '';
    if (!detail) {
      try { detail = helper.zetajs.catchUnoException(error).Message || ''; } catch {}
    }
    port.postMessage({type: 'error', id: message.id, message: '本地 Word 转换失败：' + (detail || '该文件格式或修订无法可靠处理。')});
  }
};
port.postMessage({type: 'worker-ready'});
