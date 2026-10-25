const DATABASE = 'product-material-company-profile';
const DATABASE_VERSION = 1;
const STORE = 'latest-profile';
const KEY = 'latest';
const MAX_BYTES = 60 * 1024 * 1024;
const MAX_PROFILE_CHARS = 5 * 1024 * 1024;

function failure(code, message, cause) {
  const error = new Error(message, cause ? {cause} : undefined);
  error.code = code;
  return error;
}

function storageFailure(error, action) {
  if (typeof error?.code === 'string' && error.code.startsWith('PROFILE_')) return error;
  if (error?.name === 'VersionError') {
    return failure('PROFILE_CACHE_INVALID', '公司介绍缓存版本不兼容，请使用更新后的页面读取。', error);
  }
  if (error?.name === 'QuotaExceededError') {
    return failure('PROFILE_STORAGE_FAILED', `浏览器存储空间不足，无法${action}公司介绍；原缓存未被替换。`, error);
  }
  return failure('PROFILE_STORAGE_FAILED', `无法${action}公司介绍缓存，请检查浏览器是否允许本地存储。`, error);
}

function copyRecord(record, cached = false) {
  const invalid = detail => {
    throw failure(cached ? 'PROFILE_CACHE_INVALID' : 'PROFILE_RECORD_INVALID',
      `公司介绍${cached ? '缓存' : '文件'}格式无效：${detail}${cached ? '，请清除缓存后重新上传。' : '。'}`);
  };
  if (!record || typeof record !== 'object' || Array.isArray(record)) invalid('缺少记录');
  if (record.schemaVersion !== 1) invalid('不支持此版本');
  if (typeof record.name !== 'string' || !record.name.trim() || record.name.length > 255 ||
      !/\.docx$/i.test(record.name) || /[\u0000-\u001f]/.test(record.name)) invalid('需要有效的 DOCX 文件名');
  if (typeof record.sha256 !== 'string' || !/^[a-f\d]{64}$/i.test(record.sha256)) invalid('缺少有效的文件摘要');
  if (typeof record.uploadedAt !== 'string' || record.uploadedAt.length > 128 || !Number.isFinite(Date.parse(record.uploadedAt))) invalid('上传时间无效');
  if (!(record.bytes instanceof Uint8Array) || record.bytes.byteLength === 0 || record.bytes.byteLength > MAX_BYTES) invalid('文件内容为空或超过 60 MB');
  if (!record.profile || typeof record.profile !== 'object' || Array.isArray(record.profile)) invalid('缺少已核对的公司资料');
  let profile;
  try {
    const serialized = JSON.stringify(record.profile, function(key, value) {
      // JSON calls toJSON before its replacer; inspect the original too so a
      // custom conversion cannot turn a valid-looking record into corrupt data.
      const original = this[key];
      if (original && typeof original === 'object' && typeof original.toJSON === 'function') throw new Error('Custom serialization is unsupported');
      const type = typeof value;
      if (['undefined', 'function', 'symbol', 'bigint'].includes(type) || type === 'number' && !Number.isFinite(value)) throw new Error('Unsupported value');
      if (value && type === 'object' && !Array.isArray(value) && ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new Error('Unsupported object');
      return value;
    });
    if (serialized.length > MAX_PROFILE_CHARS) invalid('公司资料内容过大');
    profile = JSON.parse(serialized);
  } catch (error) {
    if (typeof error?.code === 'string' && error.code.startsWith('PROFILE_')) throw error;
    invalid('公司资料无法完整保存');
  }
  return {
    schemaVersion: 1, name: record.name, sha256: record.sha256,
    uploadedAt: record.uploadedAt, bytes: record.bytes.slice(), profile,
  };
}

function openDatabase(indexedDB) {
  return new Promise((resolve, reject) => {
    if (!indexedDB || typeof indexedDB.open !== 'function') {
      reject(failure('PROFILE_STORAGE_UNAVAILABLE', '此浏览器无法使用本地存储，公司介绍不能在刷新后保留。'));
      return;
    }
    let settled = false;
    const fail = error => { settled = true; reject(storageFailure(error, '打开')); };
    let request;
    try { request = indexedDB.open(DATABASE, DATABASE_VERSION); }
    catch (error) { fail(error); return; }
    request.onupgradeneeded = event => {
      if (settled) { request.transaction.abort(); return; }
      try {
        if (event.oldVersion !== 0) throw failure('PROFILE_CACHE_INVALID', '公司介绍缓存版本不兼容，请更新页面后重试。');
        request.result.createObjectStore(STORE);
      } catch (error) {
        request.transaction.abort();
        fail(error);
      }
    };
    request.onblocked = () => fail(failure('PROFILE_STORAGE_BLOCKED', '公司介绍缓存被其他页面占用，请关闭其他物料页面后重试。'));
    request.onerror = () => fail(request.error);
    request.onsuccess = () => {
      const database = request.result;
      if (settled) { database.close(); return; }
      if (!database.objectStoreNames.contains(STORE)) {
        database.close();
        fail(failure('PROFILE_CACHE_INVALID', '公司介绍缓存结构不兼容，请使用更新后的页面读取。'));
        return;
      }
      database.onversionchange = () => database.close();
      settled = true;
      resolve(database);
    };
  });
}

async function transaction(indexedDB, mode, action, operation, onResult = value => value) {
  const database = await openDatabase(indexedDB);
  try {
    return await new Promise((resolve, reject) => {
      let tx;
      try {
        tx = database.transaction(STORE, mode);
        const store = tx.objectStore(STORE);
        if (store.keyPath !== null || store.autoIncrement) throw failure('PROFILE_CACHE_INVALID', '公司介绍缓存结构不兼容，请使用更新后的页面读取。');
        let result;
        let requestError;
        tx.oncomplete = () => resolve(result);
        tx.onabort = () => reject(storageFailure(tx.error || requestError, action));
        tx.onerror = () => { requestError = tx.error || requestError; };
        const request = operation(store);
        request.onsuccess = () => {
          try { result = onResult(request.result, store); }
          catch(error) { requestError=error;tx.abort(); }
        };
        request.onerror = () => { requestError = request.error; };
      } catch (error) {
        if (tx) { try { tx.abort(); } catch {} }
        reject(storageFailure(error, action));
      }
    });
  } finally {
    database.close();
  }
}

/** One validated company document, stored only in this browser's IndexedDB. */
export function createProfileStore({indexedDB = globalThis.indexedDB} = {}) {
  return {
    async load() {
      const record = await transaction(indexedDB, 'readonly', '读取', store => store.get(KEY));
      return record === undefined ? null : copyRecord(record, true);
    },
    async save(record) {
      // Snapshot before opening the database; callers may reuse their buffers.
      const copy = copyRecord(record);
      // A single put atomically replaces the old record. Do not clear first.
      await transaction(indexedDB, 'readwrite', '保存', store => store.put(copy, KEY));
    },
    async clear() {
      await transaction(indexedDB, 'readwrite', '清除', store => store.clear());
    },
    async clearIfSha256(expected) {
      // Read and conditional deletion share a single write transaction. A newer
      // upload in another tab can never be deleted based on an earlier read.
      return transaction(indexedDB,'readwrite','清除',store=>store.get(KEY),(record,store)=>{
        if(!record || record.sha256!==expected)return false;
        store.delete(KEY);return true;
      });
    },
  };
}
