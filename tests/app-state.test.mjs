import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import JSZip from 'jszip';
import * as naming from '../src/naming.js';
import * as reviewEngine from '../src/review-engine.js';
import * as companyDefaults from '../src/default-company-profile.js';
import {calendarSupplementText,calendarOpeningRange} from '../src/calendar-import.js';

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
    const classes = new Set();
    this.classList = {add: name => classes.add(name), remove: name => classes.delete(name), contains: name => classes.has(name), toggle(name, active) { if(active ?? !classes.has(name)) classes.add(name); else classes.delete(name); }};
  }
  append(...children) { this.children.push(...children); }
  contains(element) { return element===this || descendants(this).includes(element); }
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
  scrollIntoView() { this.scrolledIntoView = true; }
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
function createApp({store, parseCompanyProfile, onAnalyze, onWrite, onDownload, readWordTemplate, readWordContract, taskFactory, calendar} = {}) {
  const elements = new Map(ids.map(id => [id, new Element()]));
  const document = {
    getElementById(id) { assert.ok(elements.has(id), `Missing page element: ${id}`); return elements.get(id); },
    createElement: tag => { const element=new Element(tag); if(tag==='a'&&onDownload)element.click=()=>onDownload(element); return element; },
    querySelectorAll: () => [],
  };
  const saved = store || {load: async () => null, save: async () => {}, clear: async () => {}};
  const modules = {
    './default-company-profile.js': {loadDefaultCompanyProfile: () => profile('builtin-company.docx', 'c'), isSupersededCompanyProfile: companyDefaults.isSupersededCompanyProfile},
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
      readWordContract: readWordContract || (async ({name})=>({name,sha256:'contract'})),
      readWordTemplate: readWordTemplate || (async ({name}) => ({id: 'template', name:name.replace(/\.(doc|rtf)$/i,'.docx'), sha256: 'template', kind: 'docx', paragraphs: []})),
    },
    './review-engine.js': {
      ...reviewEngine,
      analyzeMaterials(contract, templates, options) {
        onAnalyze?.(options);
        if(taskFactory)return taskFactory(contract,templates,options);
        return {
          target: {name: '示例稳健1号集合资产管理计划', shortName: '示例稳健1号'},
          contract, templates, month: options.month, warnings: [], _records: [],
          facts: [{id: 'risk', key: 'risk_rating', label: '风险等级', status: 'pending', evidence: [], occurrences: []}],
          materials: templates.map(template => ({template, kind: 'word', fields: [], operations: [], warnings: []})),
        };
      },
    },
    './calendar-import.js':{calendarSupplementText,calendarOpeningRange,...(calendar||{readProductCalendar:async()=>({name:'产品日历.xlsx',rows:[],warnings:[]}),matchCalendarRows:()=>[]})},
    './office-writer.js': {
      async writeTask(task) {
        const result=await onWrite?.(task);
        return result?.files ? result : {files: [{name: `示例_${task.month}.docx`, bytes: new Uint8Array([1]), mime: 'application/octet-stream'}]};
      },
    },
  };
  let urlNumber = 0;const blobs=new Map();
  const context = vm.createContext({
    ...naming, document, Date, Intl, Uint8Array, Blob, JSZip,
    URL: {createObjectURL: blob => {const url=`blob:test-${++urlNumber}`;blobs.set(url,blob);return url;}, revokeObjectURL() {}},
    requestAnimationFrame: callback => queueMicrotask(callback),
    window: {scrollTo() {}},
    importModule: async path => { const modulePath=path.split('?')[0]; assert.ok(modules[modulePath], `Unexpected app import: ${path}`); return modules[modulePath]; },
  });
  const source = appSource
    .replace(/^import[^\n]+\n/, '')
    .replace(/\bimport\(/g, 'importModule(')
    .replace(/restoreCompany\(\);(?=\s*$)/, 'globalThis.appReady=restoreCompany();');
  vm.runInContext(source, context, {filename: 'src/app.js'});
  assert.ok(context.appReady, 'App startup still exposes its initial restore boundary');
  return {elements, state: vm.runInContext('state', context), ready: context.appReady,blobs,renderReview:()=>vm.runInContext('renderReview()',context)};
}

async function chooseMaterials(app) {
  await app.elements.get('contract-input').dispatch('change', {target: {files: [file('示例合同.docx')]}});
  await app.elements.get('templates-input').dispatch('change', {target: {files: [file('示例模板.docx', 2)]}});
}

function appendSupplementFacts(app){
 const scope={product:app.state.task.target.name,period:'announcement',condition:'opening|date|2026/8/12'};
 for(const [id,key,label,condition]of [['date-a','opening_announced_date','本次开放日','opening|start|8月12日'],['date-b','opening_announced_date','本次开放日','opening|date|2026/8/12-2026/8/14'],['date-c','opening_announced_date','本次开放日','opening|end|8月14日'],['code','product_code','产品代码']])app.state.task.facts.push({id,key,label,status:'pending',value:null,scope:condition?{...scope,condition}:{product:scope.product},evidence:[],occurrences:[]});
 app.renderReview();
}
const buttonIn=(element,text)=>descendants(element).find(n=>n.tagName==='button'&&n.textContent===text);

test('four independent rows exclude contract opening rules and apply only the parsed row after confirmation',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);
 assert.equal(descendants(app.elements.get('facts-list')).filter(n=>n.tagName==='strong'&&n.textContent==='本次开放日').length,1);
 assert.equal(app.elements.has('supplement-opening'),false);
 assert.equal(app.elements.has('supplement-parse-opening'),false);
 for(const kind of ['date','benchmark','code'])assert.equal(app.elements.get('supplement-'+kind).value,'');
 for(const key of ['threshold1','ratio1','threshold2','ratio2'])assert.equal(app.elements.get('performance-'+key).value,'');
 app.elements.get('supplement-date').value='2026年11月3日至2026年11月5日';
 app.elements.get('supplement-code').value='000123';
 await app.elements.get('supplement-parse-code').click();assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
 assert.ok(buttonIn(app.elements.get('supplement-preview'),'确认并应用补充'));
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,'000123');
 assert.ok(app.state.task.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='pending'));
 await app.elements.get('supplement-parse-date').click();
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 assert.ok(app.state.task.facts.filter(f=>f.key==='opening_announced_date').every(f=>f.status==='verified'));
});

test('confirm-original and undo actions invalidate existing downloads and never invent a replacement value',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();await app.elements.get('generate-button').click();
 assert.ok(app.elements.get('download-list').children.length);
 await buttonIn(app.elements.get('facts-list'),'确认无误').click();
 assert.equal(app.state.task.facts[0].userConfirmed,true);assert.equal(app.state.task.facts[0].value,undefined);
 assert.equal(app.elements.get('download-list').children.length,0);
 await buttonIn(app.elements.get('facts-list'),'撤销确认').click();assert.equal(app.state.task.facts[0].status,'pending');assert.equal(app.state.task.facts[0].userConfirmed,undefined);
});

test('Excel matching row previews with provenance; another product is not silently selected',async()=>{
 const row={id:'row-2',productName:'示例稳健1号',productCode:'000123',source:{file:'日历.xlsx',sheet:'Sheet1',row:2}};
 const app=createApp({calendar:{readProductCalendar:async()=>({name:'日历.xlsx',rows:[row],warnings:[]}),matchCalendarRows:()=>[row]}});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);
 await app.elements.get('calendar-input').dispatch('change',{target:{files:[file('日历.xlsx')]}});
 assert.equal(app.elements.get('calendar-row').value,'row-2');assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 const code=app.state.task.facts.find(f=>f.id==='code');assert.equal(code.value,'000123');assert.equal(code.source,'user_calendar');assert.equal(code.evidence[0].source.row,2);
 const other=createApp({calendar:{readProductCalendar:async()=>({name:'日历.xlsx',rows:[{...row,productName:'示例稳健9号'}],warnings:[]}),matchCalendarRows:()=>[]}});await other.ready;await chooseMaterials(other);await other.elements.get('analyze-button').click();appendSupplementFacts(other);
 await other.elements.get('calendar-input').dispatch('change',{target:{files:[file('日历.xlsx')]}});assert.equal(other.elements.get('calendar-row').value,'');assert.equal(other.state.supplementPreview,null);assert.match(other.elements.get('calendar-status').textContent,/未找到合同产品/);
});

test('Excel preview includes both schedule bounds and the complete window rule before applying',async()=>{
 const row={id:'row-2',productName:'示例稳健1号',openingDate:'2026/9/8',openingStart:'2026/9/8',openingEnd:'2026/9/14',openingDateNote:'前2个开放日开放参与和退出；后3个开放日仅开放参与',source:{file:'日历.xlsx',sheet:'Sheet1',row:2}};
 const app=await appForCalendar({readProductCalendar:async()=>({name:'日历.xlsx',rows:[row],warnings:[]}),matchCalendarRows:()=>[row]});
 await app.elements.get('calendar-input').dispatch('change',{target:{files:[file('日历.xlsx')]}});
 assert.match(app.state.supplementPreview.text,/本次开放日：2026\/9\/8至2026\/9\/14/);
 assert.match(app.state.supplementPreview.text,/分段规则：前2.*后3/);
 assert.ok(!descendants(app.elements.get('calendar-row-detail')).some(n=>n.textContent.includes('不作为本次开放日填入')));
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 assert.equal(app.state.task.facts.find(f=>f.id==='date-c').status,'verified');
 assert.equal(app.state.task.facts.find(f=>f.id==='date-c').value,'2026/9/8至2026/9/14');
});

test('editing a row after preview cannot apply the stale parsed values',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);
 app.elements.get('supplement-code').value='产品代码：000123';await app.elements.get('supplement-parse-code').click();const oldButton=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');
 app.elements.get('supplement-code').value='产品代码：000999';await app.elements.get('supplement-code').dispatch('input');await oldButton.click();assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
});

test('an empty or failed row parse cannot leave an earlier preview applicable',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);
 for(const value of ['', 'not-a-date']){
  app.elements.get('supplement-code').value='000123';await app.elements.get('supplement-parse-code').click();const oldButton=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');
  app.elements.get('supplement-date').value=value;await app.elements.get('supplement-parse-date').click();await oldButton.click();
  assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);assert.equal(app.state.supplementPreview,null);
 }
});

function appendPerformanceFacts(app,tiered=false){
 for(const key of ['performance_R1','performance_ratio_1',...(tiered?['performance_R2','performance_ratio_2']:[])])app.state.task.facts.push({id:key,key,label:key,value:null,status:'pending',scope:{product:app.state.task.target.name},evidence:[],occurrences:[]});app.renderReview();
}
test('single-tier numeric fee input previews two numbers and leaves the other rows alone',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);appendPerformanceFacts(app);
 assert.equal(app.elements.get('performance-mode').value,'single');assert.equal(app.elements.get('performance-threshold2-wrap').hidden,true);
 app.elements.get('performance-threshold1').value='2.50';app.elements.get('performance-ratio1').value='60';app.elements.get('supplement-code').value='009999';
 await app.elements.get('supplement-parse-performance').click();assert.equal(app.state.task.facts.find(f=>f.id==='performance_R1').value,null);
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 assert.equal(app.state.task.facts.find(f=>f.id==='performance_R1').value,'2.5%');assert.equal(app.state.task.facts.find(f=>f.id==='performance_ratio_1').value,'60%');assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
});
test('tiered numeric input uses all four fields and mode or value edits invalidate preview',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendPerformanceFacts(app,true);
 app.elements.get('performance-mode').value='tiered';await app.elements.get('performance-mode').dispatch('change');assert.equal(app.elements.get('performance-threshold2-wrap').hidden,false);
 for(const [key,value]of Object.entries({threshold1:'2',threshold2:'3',ratio1:'20',ratio2:'40'}))app.elements.get('performance-'+key).value=value;
 await app.elements.get('supplement-parse-performance').click();const previous=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');
 app.elements.get('performance-ratio2').value='50';await app.elements.get('performance-ratio2').dispatch('input');await previous.click();assert.ok(app.state.task.facts.filter(f=>f.key.startsWith('performance_')).every(f=>f.value===null));
 await app.elements.get('supplement-parse-performance').click();await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();assert.equal(app.state.task.facts.find(f=>f.key==='performance_ratio_2').value,'50%');
 await app.elements.get('supplement-parse-performance').click();const old=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');app.elements.get('performance-mode').value='single';await app.elements.get('performance-mode').dispatch('change');await old.click();assert.equal(app.state.supplementPreview,null);
});
test('empty and incomplete numeric fee inputs cannot apply another row preview',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);appendPerformanceFacts(app);
 for(const first of ['', '2.5']){
  app.elements.get('supplement-code').value='000123';await app.elements.get('supplement-parse-code').click();const old=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');app.elements.get('performance-threshold1').value=first;
  await app.elements.get('supplement-parse-performance').click();await old.click();assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);assert.equal(app.state.supplementPreview,null);
 }
});

test('editing an individual field invalidates its previous preview',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);
 const input=descendants(app.elements.get('facts-list')).find(n=>n['aria-label']==='补充产品代码');
 const form=descendants(app.elements.get('facts-list')).find(n=>n.tagName==='form'&&n.children.includes(input));
 input.value='000123';await form.onsubmit({preventDefault(){}});const oldButton=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');
 input.value='000999';await input.dispatch('input');await oldButton.click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);assert.equal(app.elements.get('supplement-preview').hidden,true);
});

const calendarFixture=()=>{const row={id:'row-2',productName:'示例稳健1号',productCode:'000123',source:{file:'日历.xlsx',sheet:'Sheet1',row:2}};return {readProductCalendar:async({name})=>({name,rows:[row],warnings:[]}),matchCalendarRows:()=>[row]};};
async function appForCalendar(calendar=calendarFixture()){const app=createApp({calendar});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);return app;}

test('dropping a calendar shares click-import preview and requires confirmation before applying',async()=>{
 const app=await appForCalendar(),zone=app.elements.get('calendar-drop');let prevented=false;
 await zone.dispatch('dragover',{dataTransfer:{types:['Files']},preventDefault(){prevented=true;}});
 assert.equal(prevented,true);assert.equal(zone.classList.contains('dragging'),true);
 await zone.dispatch('drop',{dataTransfer:{files:[file('日历.xlsx')]}});
 assert.equal(zone.classList.contains('dragging'),false);assert.equal(app.elements.get('calendar-row').value,'row-2');
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,'000123');assert.equal(app.state.task.facts.find(f=>f.id==='code').source,'user_calendar');
});

test('failed replacement clears old unapplied calendar rows and preview, while keeping applied facts',async()=>{
 const calendar=calendarFixture(),read=calendar.readProductCalendar;calendar.readProductCalendar=async input=>{if(input.name==='损坏.xlsx')throw new Error('无法读取产品日历');return read(input);};
 const app=await appForCalendar(calendar);await app.elements.get('calendar-input').dispatch('change',{target:{files:[file('日历.xlsx')]}});
 await buttonIn(app.elements.get('supplement-preview'),'确认并应用补充').click();await app.elements.get('calendar-row').dispatch('change');
 const oldButton=buttonIn(app.elements.get('supplement-preview'),'确认并应用补充');
 await app.elements.get('calendar-drop').dispatch('drop',{dataTransfer:{files:[file('损坏.xlsx')]}});await oldButton.click();
 assert.equal(app.state.calendar,null);assert.equal(app.elements.get('calendar-row').hidden,true);assert.equal(app.elements.get('calendar-row-detail').children.length,0);
 assert.equal(app.state.supplementPreview,null);assert.equal(app.state.task.facts.find(f=>f.id==='code').value,'000123');assert.equal(app.state.busy,false);assert.match(app.elements.get('calendar-status').textContent,/未导入/);
});

test('calendar drop validates one XLSX and size; empty drops and busy drops cannot replace the preview',async()=>{
 const app=await appForCalendar(),zone=app.elements.get('calendar-drop');
 for(const files of [[file('a.xlsx'),file('b.xlsx')],[file('日历.docx')],[{...file('大文件.xlsx'),size:31*1024*1024}]]){
  await zone.dispatch('drop',{dataTransfer:{files}});assert.equal(app.state.calendar,null);assert.equal(app.state.busy,false);assert.equal(app.elements.get('error-panel').hidden,false);
 }
 await zone.dispatch('drop',{dataTransfer:{files:[file('日历.xlsx')]}});const preview=app.state.supplementPreview;
 await zone.dispatch('drop',{dataTransfer:{files:[]}});assert.equal(app.state.supplementPreview,preview);
 app.state.busy=true;await zone.dispatch('drop',{dataTransfer:{files:[file('新日历.xlsx')]}});assert.equal(app.state.supplementPreview,preview);assert.equal(app.state.calendar.name,'日历.xlsx');
});

test('calendar drop cannot race an in-progress import and supports importing the same file again',async()=>{
 const gate=deferred(),calendar=calendarFixture(),read=calendar.readProductCalendar;let reads=0;
 calendar.readProductCalendar=async input=>{reads++;if(reads===1)await gate.promise;return read(input);};
 const app=await appForCalendar(calendar),zone=app.elements.get('calendar-drop');const first=zone.dispatch('drop',{dataTransfer:{files:[file('日历.xlsx')]}});
 await new Promise(resolve=>setImmediate(resolve));assert.equal(app.state.busy,true);
 await zone.dispatch('drop',{dataTransfer:{files:[file('别的日历.xlsx')]}});gate.resolve();await first;assert.equal(reads,1);assert.equal(app.state.calendar.name,'日历.xlsx');
 await app.elements.get('calendar-input').dispatch('change',{target:{files:[file('日历.xlsx')]}});assert.equal(reads,2);assert.equal(app.elements.get('calendar-input').value,'');assert.equal(app.elements.get('error-panel').hidden,true);
});

test('one unreadable template does not prevent reviewing and downloading the other template',async()=>{
 const app=createApp({readWordTemplate:async({name})=>{if(name.includes('损坏'))throw new Error('文件包损坏');return {id:name,name,sha256:name,kind:'docx',paragraphs:[]};}});
 await app.ready;await chooseMaterials(app);
 await app.elements.get('templates-input').dispatch('change',{target:{files:[file('损坏模板.docx',3)]}});
 await app.elements.get('analyze-button').click();
 assert.equal(app.state.task.materials.length,1);
 assert.equal(app.state.task.inputIssues[0].sourceName,'损坏模板.docx');
 assert.match(descendants(app.elements.get('warnings')).map(e=>e.textContent).join('\n'),/损坏模板.*未处理.*文件包损坏/);
 await app.elements.get('generate-button').click();
 assert.match(app.elements.get('generation-status').textContent,/1.*未生成/);
 assert.ok(app.elements.get('download-list').children.some(n=>n.tagName==='a'));
});

test('duplicate templates and the accidentally selected contract are skipped without blocking valid material',async()=>{
 const app=createApp({readWordTemplate:async({name})=>({id:name,name,sha256:name.includes('合同')?'contract':'same-template',kind:'docx',paragraphs:[]})});
 await app.ready;await chooseMaterials(app);
 await app.elements.get('templates-input').dispatch('change',{target:{files:[file('副本模板.docx',3),file('误选合同.docx',4)]}});
 await app.elements.get('analyze-button').click();
 assert.equal(app.state.task.materials.length,1);
 assert.deepEqual(Array.from(app.state.task.inputIssues,i=>i.reason),['duplicate','contract']);
});

test('when no valid templates remain the app gives file-specific reasons instead of an empty success',async()=>{
 const app=createApp({readWordTemplate:async()=>{throw new Error('缺少正文');}});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 assert.equal(app.state.task,null);assert.match(app.elements.get('error-message').textContent,/没有可更新.*示例模板.*缺少正文/);
});

test('partial generation exposes every missing file while preserving successful downloads and review notes',async()=>{
 const app=createApp({onWrite:()=>({files:[{name:'正常.docx',bytes:new Uint8Array([1]),mime:'application/octet-stream',report:{repairs:['已移除失效跳转']}}],failures:[{sourceName:'未完成.pptx',message:'图片缺失'}]}),onDownload:()=>{}});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 await app.elements.get('generate-button').click();
 assert.match(app.elements.get('generation-status').textContent,/1.*未生成/);
 const text=descendants(app.elements.get('download-list')).map(e=>e.textContent).join('\n');
 assert.match(text,/未完成\.pptx.*图片缺失/);assert.match(text,/已移除失效跳转/);
 assert.ok(app.elements.get('download-list').children.some(n=>n.download==='正常.docx'));
 const bundle=app.elements.get('download-list').children.find(n=>n.download?.endsWith('.zip'));
 const zip=await JSZip.loadAsync(await app.blobs.get(bundle.href).arrayBuffer());
 assert.match(await zip.file('审阅清单.txt').async('string'),/未生成 未完成\.pptx.*图片缺失/);
 assert.match(await zip.file('审阅清单.txt').async('string'),/已移除失效跳转/);
});

test('the contract entry uses the same signature-aware Word compatibility reader',async()=>{
 let reads=0;const app=createApp({readWordContract:async({name})=>{reads++;return {name,sha256:'contract',warnings:['已自动兼容合同格式']};}});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 assert.equal(reads,1);assert.equal(app.state.task.contract.warnings[0],'已自动兼容合同格式');
});

test('delivery previews use contract fallback and distinguish colliding template names',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);
 await app.elements.get('analyze-button').click();
 const material=app.state.task.materials[0];
 material.template.name='【一页通202608】旧产品2号X期.docx';
 app.state.task.materials.push({...material,template:{...material.template,name:'【一页通202608】未知产品7号.docx'}});
 app.elements.get('month-value').value='2026-10';
 await app.elements.get('month-value').dispatch('change');
 const names=descendants(app.elements.get('delivery-list')).filter(n=>n.tagName==='strong').map(n=>n.textContent);
 assert.deepEqual(names,['【一页通202610】示例稳健1号.docx','【一页通202610】示例稳健1号（2）.docx']);
});

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
  await descendants(app.elements.get('supplement-preview')).find(element=>element.tagName==='button'&&element.textContent==='确认并应用补充').click();
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

test('generation progress and failures remain visible beside the download button and permit retry',async()=>{
  const entered=deferred(),finish=deferred();let attempts=0;
  const app=createApp({onWrite:async()=>{if(++attempts===1){entered.resolve();await finish.promise;throw new Error('文件名中的原产品名称无法对应所选模板');}}});
  await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
  const status=app.elements.get('generation-status');
  assert.ok(status,'The output panel has a visible generation status');
  const pending=app.elements.get('generate-button').click();await entered.promise;
  assert.equal(status.hidden,false);assert.match(status.textContent,/正在生成/);
  assert.equal(app.elements.get('generate-button').disabled,true);
  finish.resolve();await pending;
  assert.match(status.textContent,/生成失败.*文件名/);assert.equal(status.hidden,false);
  assert.equal(app.elements.get('generate-button').disabled,false);
  await app.elements.get('generate-button').click();
  assert.match(status.textContent,/已生成/);assert.doesNotMatch(status.textContent,/生成失败/);
});

test('automatic download denial leaves every direct download link usable and never reports generation failure',async()=>{
  const app=createApp({onDownload:()=>{throw new Error('automatic download blocked');}});
  await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
  await app.elements.get('generate-button').click();
  const links=app.elements.get('download-list').children.filter(n=>n.tagName==='a');
  assert.equal(links.length,2,'Bundle and original-format file links both survive');
  assert.equal(app.elements.get('error-panel').hidden,true);
  const status=app.elements.get('generation-status');assert.ok(status);
  assert.match(status.textContent,/已生成.*点击/);
  assert.equal(app.elements.get('download-list').scrolledIntoView,true);
  app.elements.get('month-value').value='2030-02';await app.elements.get('month-value').dispatch('change');
  assert.equal(status.hidden,true,'Stale download confirmation disappears when output naming changes');
});

test('only the known superseded company source is migrated to the corrected built-in profile',async()=>{
  const old={...profile('公司介绍2026.10版本.docx'),sha256:'e627c53304a4e01b72435054d64f373589ca724884177144986070f4f8356f6f'};
  let clears=0;
  const app=createApp({store:{load:async()=>record(old),save:async()=>{},clearIfSha256:async hash=>{assert.equal(hash,old.sha256);clears++;return true;}},parseCompanyProfile:async()=>old});
  await app.ready;assert.equal(app.state.companyIsDefault,true);assert.equal(clears,1);
  assert.equal(app.state.companyStored,false);assert.match(app.elements.get('company-status').textContent,/更正/);
  const distinct=profile(old.name,'d');let distinctClears=0;
  const newer=createApp({store:{load:async()=>record(distinct),save:async()=>{},clear:async()=>{distinctClears++;}},parseCompanyProfile:async()=>distinct});
  await newer.ready;assert.equal(newer.state.company.sha256,distinct.sha256);assert.equal(distinctClears,0);
});

test('a failed removal of the known old company cache cannot override the corrected built-in source',async()=>{
  const old={...profile('公司介绍2026.10版本.docx'),sha256:'e627c53304a4e01b72435054d64f373589ca724884177144986070f4f8356f6f'};
  const app=createApp({store:{load:async()=>record(old),save:async()=>{},clearIfSha256:async()=>{throw new Error('存储不可写');}},parseCompanyProfile:async()=>old});
  await app.ready;assert.equal(app.state.companyIsDefault,true);assert.equal(app.state.companyStored,true);
  assert.match(app.elements.get('company-status').textContent,/网页内置.*更正/);
  assert.equal(app.elements.get('company-clear').hidden,false);
});

test('migration cannot delete a different company version saved by another page during parsing',async()=>{
 const old={...profile('old.docx'),sha256:'e627c53304a4e01b72435054d64f373589ca724884177144986070f4f8356f6f'};
 const newer=profile('newer.docx','d'),entered=deferred(),finish=deferred();let stored=record(old);
 const store={load:async()=>stored,save:async value=>{stored=value;},clear:async()=>{stored=null;},clearIfSha256:async hash=>{if(stored?.sha256!==hash)return false;stored=null;return true;}};
 const app=createApp({store,parseCompanyProfile:async()=>{entered.resolve();await finish.promise;return old;}});
 await entered.promise;stored=record(newer);finish.resolve();await app.ready;
 assert.equal(stored?.sha256,newer.sha256);assert.equal(app.state.companyStored,true);
 assert.equal(app.state.companyIsDefault,true);
});
