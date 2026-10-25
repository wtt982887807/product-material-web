import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createLegacyWordConverter} from '../src/legacy-word-converter.js';
import {convertUnoDocument} from '../converter/uno-convert.js';

globalThis.JSZip = JSZip;
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PR='http://schemas.openxmlformats.org/package/2006/relationships';
async function docx() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types><Override ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>合同正文</w:t></w:r></w:p></w:body></w:document>`);
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

test('unsupported optional DOC shading uses untouched input in the full engine and reports a warning',async()=>{
 const bytes=new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1,0,1]),original=bytes.slice(),output=await docx(),warnings=[],progress=[];
 let calls=0;
 const convert=createLegacyWordConverter({isIsolated:()=>true,createRuntime:async()=>({destroy(){},async convert(input){calls++;assert.deepEqual(input.bytes,original);return output;}})});
 assert.deepEqual(await convert({name:'旧底纹模板.doc',bytes},{onWarning:item=>warnings.push(item),onProgress:item=>progress.push(item)}),output);
 assert.equal(calls,1);assert.deepEqual(bytes,original);
 assert.equal(warnings.length,1);assert.equal(warnings[0].code,'DOC_SHADING_FALLBACK');
 assert.match(warnings[0].message,/底纹|底色/);assert.match(warnings[0].message,/完整.*引擎/);
 assert.ok(progress.some(item=>item.stage==='warning'));
});

test('shading fallback cannot bypass a failed or invalid full conversion',async()=>{
 const bytes=new Uint8Array([0xd0,0xcf,0x11,0xe0,0xa1,0xb1,0x1a,0xe1,0,1]);let called=0;
 const convert=createLegacyWordConverter({isIsolated:()=>true,createRuntime:async()=>({destroy(){},async convert(){called++;return bytes;}})});
 await assert.rejects(convert({name:'旧底纹模板.doc',bytes}),/转换结果不是有效 DOCX/);assert.equal(called,1);
});

test('converted orphan comment anchors are repaired without deleting existing comments, text or package parts',async()=>{
 const zip=await JSZip.loadAsync(await docx());
 const paragraph=`<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>应当保留的正文</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeStart w:id="9"/><w:r><w:rPr><w:b/></w:rPr><w:t>以及原有格式</w:t></w:r><w:commentRangeEnd w:id="9"/><w:r><w:commentReference w:id="9"/></w:r></w:p>`;
 zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${paragraph}</w:body></w:document>`);
 const comments=`<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="旧作者"><w:p><w:r><w:t>必须完整保留的旧批注</w:t></w:r></w:p></w:comment></w:comments>`;
 zip.file('word/annotations.xml',comments);
 zip.file('word/_rels/document.xml.rels',`<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="${R}/comments" Target="annotations.xml"/></Relationships>`);
 zip.file('word/footer1.xml',`<w:ftr xmlns:w="${W}"><w:p><w:r><w:t>页脚原文</w:t><w:commentReference w:id="9"/></w:r></w:p></w:ftr>`);
 zip.file('word/media/retained.bin',new Uint8Array([1,2,3]));const output=await zip.generateAsync({type:'uint8array'}),original=output.slice(),warnings=[];
 const convert=createLegacyWordConverter({isIsolated:()=>true,createRuntime:async()=>({destroy(){},async convert(){return output;}})});
 const result=await JSZip.loadAsync(await convert({name:'example.rtf',bytes:new TextEncoder().encode('{\\rtf1 example}')},{onWarning:item=>warnings.push(item)}));
 assert.deepEqual(output,original);
 const parser=new DOMParser();const document=parser.parseFromString(await result.file('word/document.xml').async('string'),'application/xml');
 assert.equal(Array.from(document.getElementsByTagNameNS(W,'t')).map(n=>n.textContent).join(''),'应当保留的正文以及原有格式');assert.equal(document.getElementsByTagNameNS(W,'b').length,1);
 for(const local of ['commentRangeStart','commentRangeEnd','commentReference'])assert.deepEqual(Array.from(document.getElementsByTagNameNS(W,local)).map(n=>n.getAttributeNS(W,'id')),['1']);
 assert.equal(await result.file('word/annotations.xml').async('string'),comments);
 assert.equal(parser.parseFromString(await result.file('word/footer1.xml').async('string'),'application/xml').getElementsByTagNameNS(W,'commentReference').length,0);
 assert.deepEqual(await result.file('word/media/retained.bin').async('uint8array'),new Uint8Array([1,2,3]));
 assert.deepEqual(Object.keys(result.files),Object.keys(zip.files));
 assert.equal(warnings.length,1);assert.equal(warnings[0].code,'WORD_ORPHAN_COMMENTS_REPAIRED');assert.match(warnings[0].message,/旧批注|空引用/);
});

test('converted malformed XML cannot be hidden by comment repair',async()=>{
 const zip=await JSZip.loadAsync(await docx());zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body><w:p></w:body></w:document>`);
 const output=await zip.generateAsync({type:'uint8array'}),convert=createLegacyWordConverter({isIsolated:()=>true,createRuntime:async()=>({destroy(){},async convert(){return output;}})});
 await assert.rejects(convert({name:'example.rtf',bytes:new Uint8Array([1])}),/XML|结构检查/);
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
