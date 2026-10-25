import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import JSZip from 'jszip';
import * as naming from '../src/naming.js';
import {setSupplement} from '../src/review-engine.js';

const appSource = readFileSync(new URL('../src/app.js', import.meta.url), 'utf8');
const page = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const ids = [...page.matchAll(/\bid="([^"]+)"/g)].map(match => match[1]);

class Element {
  constructor(tagName = 'div') {
    this.tagName = tagName;
    this.children = [];
    this.listeners = new Map();
    this.value = '';
    this.textContent = '';
    this.hidden = false;
    this.disabled = false;
    this.dataset = {};
    this.classList = {add() {}, remove() {}, toggle() {}};
  }
  append(...children) { this.children.push(...children); }
  replaceChildren(...children) { this.children = children; }
  setAttribute(name, value) { this[name] = value; }
  addEventListener(type, callback) {
    if (!this.listeners.has(type)) this.listeners.set(type, []);
    this.listeners.get(type).push(callback);
  }
  async dispatch(type, detail = {}) {
    const event = {target: this, preventDefault() {}, ...detail};
    await Promise.all((this.listeners.get(type) || []).map(callback => callback(event)));
  }
  click() { return this.onclick?.({target: this, preventDefault() {}}); }
}

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return {promise, resolve};
}
function descendants(element) {
  return element.children.flatMap(child => child instanceof Element ? [child, ...descendants(child)] : []);
}
const file = (name, byte = 1) => ({name, size: 1, lastModified: 1, arrayBuffer: async () => new Uint8Array([byte]).buffer});
const profile = (name, digit = 'a') => ({schemaVersion: 1, name, sha256: digit.repeat(64), sections: []});
const record = value => ({...value, bytes: new Uint8Array([1])});

// Execute the real app event handlers with small DOM/import boundary fakes.
// No browser, storage, Office files or network is used; naming and supplements
// retain their production implementations.
function createApp({store, parseCompanyProfile, onAnalyze, onWrite} = {}) {
  const elements = new Map(ids.map(id => [id, new Element()]));
  const document = {
    getElementById(id) { assert.ok(elements.has(id), `Missing page element: ${id}`); return elements.get(id); },
    createElement: tag => new Element(tag),
    querySelectorAll: () => [],
  };
  const saved = store || {load: async () => null, save: async () => {}, clear: async () => {}};
  const modules = {
    './default-company-profile.js': {loadDefaultCompanyProfile: () => profile('builtin-company.docx', 'c')},
    './profile-store.js': {createProfileStore: () => saved},
    './company-profile.js': {
      parseCompanyProfile: parseCompanyProfile || (async ({name}) => profile(name)),
      applyCompanyProfile(task, current) { task.companyProfile = {name: current.name}; return task; },
    },
    './office-reader.js': {
      readContract: async ({name}) => ({name, sha256: 'contract'}),
      readTemplate: async ({name}) => ({id: 'template', name, sha256: 'template', kind: 'docx', paragraphs: []}),
    },
    './word-input.js': {
      readWordTemplate: async ({name}) => ({id: 'template', name:name.replace(/\.(doc|rtf)$/i,'.docx'), sha256: 'template', kind: 'docx', paragraphs: []}),
    },
    './review-engine.js': {
      analyzeMaterials(contract, templates, options) {
        onAnalyze?.(options);
        return {
          target: {name: '示例稳健1号集合资产管理计划', shortName: '示例稳健1号'},
          contract, templates, month: options.month, warnings: [], _records: [],
          facts: [{id: 'risk', key: 'risk_rating', label: '风险等级', status: 'pending', evidence: [], occurrences: []}],
          materials: templates.map(template => ({template, kind: 'word', fields: [], operations: [], warnings: []})),
        };
      },
      setSupplement,
    },
    './office-writer.js': {
      async writeTask(task) {
        onWrite?.(task);
        return {files: [{name: `示例_${task.month}.docx`, bytes: new Uint8Array([1]), mime: 'application/octet-stream'}]};
      },
    },
  };
  let urlNumber = 0;
  const context = vm.createContext({
    ...naming, document, Date, Intl, Uint8Array, Blob, JSZip,
    URL: {createObjectURL: () => `blob:test-${++urlNumber}`, revokeObjectURL() {}},
    requestAnimationFrame: callback => queueMicrotask(callback),
    window: {scrollTo() {}},
    importModule: async path => { assert.ok(modules[path], `Unexpected app import: ${path}`); return modules[path]; },
  });
  const source = appSource
    .replace(/^import[^\n]+\n/, '')
    .replace(/\bimport\(/g, 'importModule(')
    .replace(/restoreCompany\(\);(?=\s*$)/, 'globalThis.appReady=restoreCompany();');
  vm.runInContext(source, context, {filename: 'src/app.js'});
  assert.ok(context.appReady, 'App startup still exposes its initial restore boundary');
  return {elements, state: vm.runInContext('state', context), ready: context.appReady};
}

async function chooseMaterials(app) {
  await app.elements.get('contract-input').dispatch('change', {target: {files: [file('示例合同.docx')]}});
  await app.elements.get('templates-input').dispatch('change', {target: {files: [file('示例模板.docx', 2)]}});
}

test('manual month survives analysis, user supplement, and export; later edits update both download names', async () => {
  const analyzed = [], written = [];
  const app = createApp({onAnalyze: options => analyzed.push(options.month), onWrite: task => written.push(task.month)});
  await app.ready;
  const month = app.elements.get('month-value');
  month.value = '2032-04';
  await month.dispatch('change');
  await chooseMaterials(app);
  await app.elements.get('analyze-button').click();
  assert.deepEqual(analyzed, ['203204']);
  assert.equal(app.state.task.month, '203204');
  const supplement = descendants(app.elements.get('facts-list')).find(element => element.tagName === 'form');
  assert.ok(supplement, 'Pending facts expose the actual supplement handler');
  supplement.children.find(element => element.tagName === 'input').value = 'R2';
  await supplement.onsubmit({preventDefault() {}});
  assert.equal(app.state.task.month, '203204');
  assert.equal(app.state.task.facts[0].value, 'R2');
  await app.elements.get('generate-button').click();
  assert.deepEqual(written, ['203204']);
  assert.deepEqual(app.elements.get('download-list').children.filter(node => node.tagName === 'a').map(node => node.download), ['示例稳健1号_203204.zip', '示例_203204.docx']);
  month.value = '2032-05';
  await month.dispatch('change');
  assert.equal(app.elements.get('download-list').children.length, 0);
  await app.elements.get('generate-button').click();
  assert.deepEqual(written, ['203204', '203205']);
  assert.ok(app.elements.get('download-list').children.filter(node => node.tagName === 'a').every(node => node.download.includes('203205')));
});

test('dropping a company file while restore is pending cannot race the restored profile', async () => {
  const old = profile('old-company.docx'), latest = profile('latest-company.docx', 'b');
  const entered = deferred(), restore = deferred();
  const parsed = [], saved = [];
  const app = createApp({
    store: {load: async () => record(old), save: async value => saved.push(value.name), clear: async () => {}},
    parseCompanyProfile: async ({name}) => {
      parsed.push(name);
      if (name === old.name) { entered.resolve(); return restore.promise; }
      return latest;
    },
  });
  await entered.promise;
  try {
    assert.equal(app.elements.get('company-input').disabled, true);
    await app.elements.get('company-drop').dispatch('drop', {dataTransfer: {files: [file(latest.name)]}});
    // Let a mistakenly accepted upload reach its parse/save boundary too.
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(parsed, [old.name]);
    assert.deepEqual(saved, []);
    assert.match(app.elements.get('error-message').textContent, /稍后|正在读取/);
  } finally { restore.resolve(old); await app.ready; }
  assert.equal(app.state.company.name, old.name);
  await app.elements.get('company-input').dispatch('change', {target: {files: [file(latest.name)]}});
  assert.equal(app.state.company.name, latest.name);
  assert.equal(app.state.companyPersisted, true);
  assert.deepEqual(saved, [latest.name]);
});

test('an invalid cached record exposes a working clear action', async () => {
  let clears = 0;
  const error = Object.assign(new Error('缓存格式无效，请清除缓存后重新上传。'), {code: 'PROFILE_CACHE_INVALID'});
  const app = createApp({store: {load: async () => { throw error; }, save: async () => {}, clear: async () => { clears++; }}});
  await app.ready;
  const clear = app.elements.get('company-clear');
  assert.equal(clear.hidden, false);
  assert.equal(clear.disabled, false);
  assert.match(app.elements.get('company-status').textContent, /未能恢复/);
  await clear.click();
  assert.equal(clears, 1);
  assert.equal(app.state.company.name, 'builtin-company.docx');
  assert.equal(app.state.companyStored, false);
  assert.equal(clear.hidden, true);
  assert.doesNotMatch(app.elements.get('company-status').textContent, /未能恢复/);
});

test('a fresh browser starts with the built-in company reference without uploading', async () => {
  const app = createApp();
  await app.ready;
  assert.equal(app.state.company?.name, 'builtin-company.docx');
  assert.match(app.elements.get('company-status').textContent, /网页内置/);
  assert.equal(app.elements.get('company-clear').hidden, true);
  await chooseMaterials(app);
  await app.elements.get('analyze-button').click();
  assert.equal(app.state.task.companyProfile.name, 'builtin-company.docx');
});

test('an uploaded company reference overrides the built-in version and clearing returns to it', async () => {
  const saved = profile('uploaded-company.docx');
  let cleared = false;
  const app = createApp({store: {load: async () => record(saved), save: async () => {}, clear: async () => { cleared = true; }}});
  await app.ready;
  assert.equal(app.state.company.name, saved.name);
  assert.match(app.elements.get('company-status').textContent, /已保存在此浏览器/);
  await app.elements.get('company-clear').click();
  assert.equal(cleared, true);
  assert.equal(app.state.company.name, 'builtin-company.docx');
  assert.match(app.elements.get('company-status').textContent, /网页内置/);
});

test('unavailable browser storage still leaves the built-in reference usable', async () => {
  const app = createApp({store: {load: async () => { throw new Error('无法使用浏览器保存空间'); }, save: async () => {}, clear: async () => {}}});
  await app.ready;
  assert.equal(app.state.company?.name, 'builtin-company.docx');
  assert.match(app.elements.get('company-status').textContent, /网页内置/);
  await chooseMaterials(app);
  assert.equal(app.elements.get('analyze-button').disabled, false);
});

test('the upload flow accepts a legacy Word template and analyzes the prepared DOCX',async()=>{
  const app=createApp();await app.ready;
  await app.elements.get('contract-input').dispatch('change',{target:{files:[file('示例合同.docx')]}});
  await app.elements.get('templates-input').dispatch('change',{target:{files:[file('旧模板.doc',2)]}});
  assert.equal(app.state.templates.length,1);
  await app.elements.get('analyze-button').click();
  assert.equal(app.state.task.templates[0].name,'旧模板.docx');
});

test('a failed save uses the new company only for this session without claiming persistence', async () => {
  const old = profile('old-company.docx'), latest = profile('latest-company.docx', 'b');
  let attempts = 0;
  const app = createApp({
    store: {load: async () => record(old), save: async () => { attempts++; throw new Error('存储空间不足'); }, clear: async () => {}},
    parseCompanyProfile: async ({name}) => name === old.name ? old : latest,
  });
  await app.ready;
  await app.elements.get('company-input').dispatch('change', {target: {files: [file(latest.name)]}});
  assert.equal(attempts, 1);
  assert.equal(app.state.company.name, latest.name);
  assert.equal(app.state.companyPersisted, false);
  const status = app.elements.get('company-status').textContent;
  assert.match(status, /latest-company\.docx/);
  assert.match(status, /仅本次使用.*尚未保存/);
  assert.match(status, /未能保存/);
  assert.doesNotMatch(status, /已保存在此浏览器|下次自动沿用/);
  await chooseMaterials(app);
  await app.elements.get('analyze-button').click();
  assert.equal(app.state.task.companyProfile.name, latest.name);
});
