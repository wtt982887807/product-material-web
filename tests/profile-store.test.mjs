import test from 'node:test';
import assert from 'node:assert/strict';

const module = await import('../src/profile-store.js').catch(error => {
  if (error.code === 'ERR_MODULE_NOT_FOUND') return {};
  throw error;
});
const createStore = options => {
  assert.equal(typeof module.createProfileStore, 'function', 'profile store entry point is implemented');
  return module.createProfileStore(options);
};
const record = (name = 'synthetic-company.docx', value = '示例公司') => ({
  schemaVersion: 1, name, sha256: 'a'.repeat(64), uploadedAt: '2026-10-25T02:03:04.000Z',
  bytes: new Uint8Array([80, 75, 3, 4]), profile: {company: value, paragraphs: ['合成测试资料'], version: 1},
});
const later = callback => setTimeout(callback, 0);

// Node has no IndexedDB. This boundary fake models async requests, structured
// cloning and commit/abort separately, including failures after put succeeds.
function fakeIndexedDB() {
  let exists = false;
  const stores = new Map();
  const control = {failCommit: false, failRequest: false, openError: null, blocked: false};
  const database = {
    objectStoreNames: {contains: name => stores.has(name)},
    createObjectStore(name) { stores.set(name, new Map()); },
    close() {},
    transaction(name, mode) {
      if (!stores.has(name)) throw new DOMException('Missing store', 'NotFoundError');
      const staged = structuredClone(stores.get(name));
      const tx = {error: null, aborted: false, abort() {
        if (this.aborted) return;
        this.aborted = true;
        later(() => this.onabort?.({target: this}));
      }};
      let pending = 0;
      let completionScheduled = false;
      const complete = () => {
        if (pending || completionScheduled) return;
        completionScheduled = true;
        later(() => {
          if (tx.aborted) return;
          if (control.failCommit && mode === 'readwrite') {
            control.failCommit = false;
            tx.error = new DOMException('Storage is full', 'QuotaExceededError');
            tx.abort();
          } else {
            if (mode === 'readwrite') stores.set(name, staged);
            tx.oncomplete?.({target: tx});
          }
        });
      };
      const request = action => {
        const req = {result: undefined, error: null};
        pending++;
        later(() => {
          if (tx.aborted) return;
          if (control.failRequest) {
            control.failRequest = false;
            req.error = new DOMException('Read or write failed', 'UnknownError');
            tx.error = req.error;
            req.onerror?.({target: req});
            tx.onerror?.({target: tx});
            tx.abort();
            return;
          }
          req.result = action();
          req.onsuccess?.({target: req});
          pending--;
          complete();
        });
        return req;
      };
      tx.objectStore = () => ({
        keyPath: null, autoIncrement: false,
        get: key => request(() => structuredClone(staged.get(key))),
        put: (value, key) => {
          assert.equal(mode, 'readwrite');
          const copy = structuredClone(value);
          return request(() => { staged.set(key, copy); return key; });
        },
        clear: () => {
          assert.equal(mode, 'readwrite');
          return request(() => staged.clear());
        },
        delete: key => { assert.equal(mode,'readwrite'); return request(()=>staged.delete(key)); },
      });
      later(complete);
      return tx;
    },
  };
  return {
    control,
    corrupt(transform) {
      for (const data of stores.values()) for (const [key, value] of data) data.set(key, transform(value));
    },
    open() {
      const req = {result: database, error: null, transaction: {abort() {}}};
      later(() => {
        if (control.openError) { req.error = control.openError; req.onerror?.({target: req}); return; }
        if (control.blocked) { req.onblocked?.({target: req}); return; }
        if (!exists) { exists = true; req.onupgradeneeded?.({oldVersion: 0, target: req}); }
        req.onsuccess?.({target: req});
      });
      return req;
    },
  };
}

test('an empty store is empty and a new store instance restores the saved document', async () => {
  const indexedDB = fakeIndexedDB();
  assert.equal(await createStore({indexedDB}).load(), null);
  await createStore({indexedDB}).save(record());
  const loaded = await createStore({indexedDB}).load();
  assert.deepEqual(loaded, record());
  assert.ok(loaded.bytes instanceof Uint8Array);
});

test('only the most recent saved document is returned and clear removes it persistently', async () => {
  const indexedDB = fakeIndexedDB();
  const store = createStore({indexedDB});
  await store.save(record());
  await store.save(record('next.docx', '新版公司'));
  assert.equal((await createStore({indexedDB}).load()).name, 'next.docx');
  await store.clear();
  assert.equal(await createStore({indexedDB}).load(), null);
});

test('general retains the legacy document while ccb starts independently and restores its own upload', async () => {
  const indexedDB = fakeIndexedDB(), original = record('legacy-general.docx');
  await createStore({indexedDB}).save(original);
  const general = createStore({indexedDB, variant: 'general'}), ccb = createStore({indexedDB, variant: 'ccb'});
  assert.deepEqual(await general.load(), original);
  assert.equal(await ccb.load(), null);
  const bank = {...record('ccb-company.docx', '建行口径公司'), sha256: 'b'.repeat(64)};
  await ccb.save(bank);
  assert.deepEqual(await createStore({indexedDB, variant: 'ccb'}).load(), bank);
  assert.deepEqual(await createStore({indexedDB}).load(), original);
});

test('clearing either version leaves the other saved company template intact', async () => {
  for (const selected of ['general', 'ccb']) {
    const indexedDB = fakeIndexedDB();
    const stores = Object.fromEntries(['general', 'ccb'].map(variant => [variant, createStore({indexedDB, variant})]));
    await stores.general.save(record('general.docx'));
    await stores.ccb.save(record('ccb.docx'));
    await stores[selected].clear();
    assert.equal(await createStore({indexedDB, variant: selected}).load(), null);
    const other = selected === 'general' ? 'ccb' : 'general';
    assert.deepEqual(await createStore({indexedDB, variant: other}).load(), record(other + '.docx'));
  }
});

test('conditional removal is scoped even when both versions have the same document hash', async () => {
  const indexedDB = fakeIndexedDB(), general = createStore({indexedDB}), ccb = createStore({indexedDB, variant: 'ccb'});
  await general.save(record('general.docx'));
  await ccb.save(record('ccb.docx'));
  assert.equal(await ccb.clearIfSha256('b'.repeat(64)), false);
  assert.equal(await ccb.clearIfSha256(record().sha256), true);
  assert.equal(await ccb.load(), null);
  assert.deepEqual(await general.load(), record('general.docx'));
});

test('a failed ccb replacement retains both committed versions', async () => {
  const indexedDB = fakeIndexedDB(), general = createStore({indexedDB}), ccb = createStore({indexedDB, variant: 'ccb'});
  await general.save(record('general.docx'));
  await ccb.save(record('ccb.docx'));
  indexedDB.control.failCommit = true;
  await assert.rejects(ccb.save(record('new-ccb.docx')), /保存|存储|空间/);
  assert.deepEqual(await general.load(), record('general.docx'));
  assert.deepEqual(await ccb.load(), record('ccb.docx'));
});

test('unknown version names are rejected before accessing storage', () => {
  const indexedDB = {open() { throw new Error('storage must not be opened'); }};
  for (const variant of ['', 'CCB', 'latest', 'unknown', 'toString', '__proto__', null, 1]) {
    assert.throws(() => createStore({indexedDB, variant}), /公司介绍.*版本/);
  }
});

test('conditional migration removes only the expected source and retains a newer saved document',async()=>{
 const store=createStore({indexedDB:fakeIndexedDB()});
 assert.equal(typeof store.clearIfSha256,'function');
 const old=record(),newer={...record('newer.docx'),sha256:'b'.repeat(64)};
 await store.save(old);await store.save(newer);
 assert.equal(await store.clearIfSha256(old.sha256),false);assert.deepEqual(await store.load(),newer);
 assert.equal(await store.clearIfSha256(newer.sha256),true);assert.equal(await store.load(),null);
 assert.equal(await store.clearIfSha256(newer.sha256),false);
});

test('conditional removal waits for commit and retains the prior document on failure',async()=>{
 const indexedDB=fakeIndexedDB(),store=createStore({indexedDB});await store.save(record());
 assert.equal(typeof store.clearIfSha256,'function');indexedDB.control.failCommit=true;
 await assert.rejects(store.clearIfSha256(record().sha256),/清除|存储|空间/);
 assert.deepEqual(await store.load(),record());
});

test('failed replacement after successful put rejects and preserves the previous committed document', async () => {
  const indexedDB = fakeIndexedDB();
  const store = createStore({indexedDB});
  await store.save(record());
  indexedDB.control.failCommit = true;
  await assert.rejects(store.save(record('next.docx')), /保存|存储|空间/);
  assert.deepEqual(await createStore({indexedDB}).load(), record());
});

test('a failed clear is reported and does not remove the saved document', async () => {
  const indexedDB = fakeIndexedDB();
  const store = createStore({indexedDB});
  await store.save(record());
  indexedDB.control.failCommit = true;
  await assert.rejects(store.clear(), /清除|存储|空间/);
  assert.deepEqual(await store.load(), record());
});

test('save snapshots bytes and nested data before asynchronous work', async () => {
  const indexedDB = fakeIndexedDB();
  const source = record();
  const promise = createStore({indexedDB}).save(source);
  source.bytes[0] = 0;
  source.profile.paragraphs[0] = 'changed after save';
  await promise;
  const loaded = await createStore({indexedDB}).load();
  assert.deepEqual(loaded, record());
  loaded.profile.paragraphs[0] = 'changed after load';
  loaded.bytes[0] = 1;
  assert.deepEqual(await createStore({indexedDB}).load(), record());
});

test('invalid replacement records are rejected before replacing a usable cache', async () => {
  const indexedDB = fakeIndexedDB();
  const store = createStore({indexedDB});
  await store.save(record());
  const invalid = [
    null, {...record(), schemaVersion: 2}, {...record(), name: ''}, {...record(), name: 'x'.repeat(256) + '.docx'},
    {...record(), name: 'file.pptx'}, {...record(), sha256: ''}, {...record(), sha256: 'x'.repeat(64)},
    {...record(), uploadedAt: 'not a date'}, {...record(), bytes: []}, {...record(), bytes: new Uint8Array()},
    {...record(), bytes: new Uint8Array(60 * 1024 * 1024 + 1)}, {...record(), profile: null},
    {...record(), profile: []}, {...record(), profile: {bad: () => {}}}, {...record(), profile: {bad: undefined}},
    {...record(), profile: {toJSON() { return 'lost object'; }}}, {...record(), profile: new Date()},
  ];
  const cycle = {}; cycle.self = cycle;
  invalid.push({...record(), profile: cycle});
  for (const candidate of invalid) await assert.rejects(store.save(candidate), /公司介绍|格式|文件|版本/);
  assert.deepEqual(await store.load(), record());
});

test('stale or corrupt stored records are surfaced instead of being treated as a fresh empty cache', async () => {
  for (const corrupt of [value => ({...value, schemaVersion: 9}), value => ({...value, bytes: null}), () => null]) {
    const indexedDB = fakeIndexedDB();
    const store = createStore({indexedDB});
    await store.save(record());
    indexedDB.corrupt(corrupt);
    await assert.rejects(store.load(), /缓存|版本|格式/);
    await store.clear();
    assert.equal(await store.load(), null);
  }
});

test('missing, denied, blocked and incompatible browser storage produces readable failures', async () => {
  for (const indexedDB of [null, {}, {open() {throw new DOMException('Denied', 'SecurityError');}}]) {
    const store = createStore({indexedDB});
    for (const operation of [() => store.load(), () => store.save(record()), () => store.clear()]) {
      await assert.rejects(operation(), /浏览器|存储|缓存/);
    }
  }
  for (const name of ['SecurityError', 'VersionError']) {
    const indexedDB = fakeIndexedDB();
    indexedDB.control.openError = new DOMException('Unavailable', name);
    await assert.rejects(createStore({indexedDB}).load(), /浏览器|存储|缓存|版本/);
  }
  const indexedDB = fakeIndexedDB();
  indexedDB.control.blocked = true;
  await assert.rejects(createStore({indexedDB}).load(), /浏览器|页面|占用/);
});

test('a request failure is surfaced and the same store can retry successfully', async () => {
  const indexedDB = fakeIndexedDB();
  const store = createStore({indexedDB});
  await store.save(record());
  indexedDB.control.failRequest = true;
  await assert.rejects(store.load(), /读取|存储|缓存/);
  assert.deepEqual(await store.load(), record());
});
