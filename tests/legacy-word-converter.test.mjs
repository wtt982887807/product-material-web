import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {createLegacyWordConverter} from '../src/legacy-word-converter.js';
import {convertUnoDocument} from '../converter/uno-convert.js';

globalThis.JSZip = JSZip;
async function docx() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types><Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('word/document.xml', '<w:document/>');
  return zip.generateAsync({type: 'uint8array'});
}

test('conversion is lazy, serial, reuses one runtime, and snapshots caller bytes', async () => {
  const output = await docx();
  let factories = 0, active = 0, maximum = 0;
  const seen = [];
  const convert = createLegacyWordConverter({isIsolated: () => true, createRuntime: async () => {
    factories++;
    return {destroy() {}, async convert(input, {onProgress}) {
      active++; maximum = Math.max(maximum, active); seen.push(input);
      onProgress?.({stage: 'convert', message: '转换中'});
      await new Promise(resolve => setImmediate(resolve)); active--; return output;
    }};
  }});
  assert.equal(factories, 0);
  const bytes = new Uint8Array([1, 2, 3]);
  const messages = [];
  const first = convert({name: '../../a.doc', bytes}, {onProgress: item => messages.push(item.message)});
  bytes[0] = 9;
  const second = convert({name: 'b.rtf', bytes: new Uint8Array([4])});
  const third = convert({name: 'complex.docx', bytes: new Uint8Array([5])});
  await Promise.all([first, second, third]);
  assert.equal(factories, 1); assert.equal(maximum, 1);
  assert.equal(seen[0].bytes[0], 1);
  assert.deepEqual(seen.map(item => item.extension), ['doc', 'rtf', 'docx']);
  assert.ok(messages.includes('转换中'));
});

test('unavailable isolation and unsupported inputs fail before loading the engine', async () => {
  let calls = 0;
  const convert = createLegacyWordConverter({isIsolated: () => false, createRuntime: async () => { calls++; }});
  await assert.rejects(convert({name: 'a.doc', bytes: new Uint8Array([1])}), /隔离|桌面浏览器/);
  const isolated = createLegacyWordConverter({isIsolated: () => true, createRuntime: async () => { calls++; }});
  await assert.rejects(isolated({name: 'a.xml', bytes: new Uint8Array([1])}), /DOC|RTF/);
  await assert.rejects(isolated({name: 'a.doc', bytes: new Uint8Array()}), /为空/);
  assert.equal(calls, 0);
});

test('a timed out runtime is destroyed and the next request starts a fresh runtime', async () => {
  let starts = 0, destroys = 0;
  const output = await docx();
  const convert = createLegacyWordConverter({timeoutMs: 15, isIsolated: () => true, createRuntime: async () => {
    starts++; const current = starts;
    return {destroy() { destroys++; }, convert: async () => current === 1 ? new Promise(() => {}) : output};
  }});
  await assert.rejects(convert({name: 'a.doc', bytes: new Uint8Array([1])}), /超时/);
  await convert({name: 'b.doc', bytes: new Uint8Array([2])});
  assert.equal(starts, 2); assert.equal(destroys, 1);
});

test('an invalid engine result is rejected instead of being renamed DOCX', async () => {
  let destroyed = false;
  const convert = createLegacyWordConverter({isIsolated: () => true, createRuntime: async () => ({
    destroy() { destroyed = true; }, convert: async () => new Uint8Array([1, 2, 3]),
  })});
  await assert.rejects(convert({name: 'a.doc', bytes: new Uint8Array([1])}), /DOCX|转换结果/);
  assert.equal(destroyed, true);
});

function uno({accepts = true} = {}) {
  const calls = []; let revisions = 2;
  const model = {
    supportsService: name => name === 'com.sun.star.text.TextDocument',
    getRedlines: () => ({getCount: () => revisions}),
    getCurrentController: () => ({getFrame: () => ({})}),
    storeToURL(url, values) { calls.push(['store', url, values]); },
    close() { calls.push(['close']); },
  };
  const css = {
    beans: {PropertyValue: class {constructor(value) {Object.assign(this, value);}}},
    task: {XInteractionHandler: {}, XInteractionAbort: {}},
    frame: {DispatchHelper: {create: () => ({executeDispatch(_frame, command) {calls.push(['dispatch', command]); if (accepts) revisions = 0;}})}},
  };
  const zetajs = {type: {short: 'short', interface: value => value}, Any: class {constructor(type, value) {this.type = type; this.value = value;}}, unoObject: (_types, value) => value};
  const desktop = {loadComponentFromURL(url, target, flags, values) {calls.push(['load', url, values]); return model;}};
  return {calls, css, zetajs, desktop, context: {}};
}

test('UNO accepts revisions before DOCX export and disables macros and link updates', () => {
  const env = uno();
  convertUnoDocument({inputUrl: 'file:///tmp/input.doc', outputUrl: 'file:///tmp/output.docx'}, env);
  const properties = Object.fromEntries(env.calls.find(call => call[0] === 'load')[2].map(item => [item.Name, item.Value]));
  assert.equal(properties.Hidden, true); assert.equal(properties.ReadOnly, false);
  assert.equal(properties.MacroExecutionMode.type, 'short'); assert.equal(properties.MacroExecutionMode.value, 0);
  assert.equal(properties.UpdateDocMode.type, 'short'); assert.equal(properties.UpdateDocMode.value, 0);
  assert.ok(properties.InteractionHandler);
  assert.deepEqual(env.calls.map(call => call[0]), ['load', 'dispatch', 'store', 'close']);
  assert.equal(env.calls[1][1], '.uno:AcceptAllTrackedChanges');
  assert.equal(env.calls[2][2].find(item => item.Name === 'FilterName').Value, 'Office Open XML Text');
});

test('UNO refuses to export when the accept command leaves tracked changes', () => {
  const env = uno({accepts: false});
  assert.throws(() => convertUnoDocument({inputUrl: 'file:///tmp/input.doc', outputUrl: 'file:///tmp/output.docx'}, env), /修订/);
  assert.equal(env.calls.some(call => call[0] === 'store'), false);
  assert.equal(env.calls.at(-1)[0], 'close');
});
