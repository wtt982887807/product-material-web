import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
import JSZip from 'jszip';
import * as naming from '../src/naming.js';
import * as productionDates from '../src/production-date.js';
import * as introductionTasks from '../src/introduction-task.js';
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
function createApp({store, storeFactory, parseCompanyProfile, prepareCompanyTemplate, managerStore, parseManagerProfile, onApplyManager, clock=Date, onAnalyze, onWrite, onDownload, readWordTemplate, readWordContract, taskFactory, calendar} = {}) {
  const elements = new Map(ids.map(id => [id, new Element()]));
  const document = {
    getElementById(id) { assert.ok(elements.has(id), `Missing page element: ${id}`); return elements.get(id); },
    createElement: tag => { const element=new Element(tag); if(tag==='a'&&onDownload)element.click=()=>onDownload(element); return element; },
    querySelectorAll: () => [],
  };
  const saved = store || {load: async () => null, save: async () => {}, clear: async () => {}};
  const modules = {
    './introduction-task.js':introductionTasks,
    './ppt-company-profile.js':{applyPptCompanyProfile:task=>task},
    './ppt-manager-profile.js':{applyPptManagerProfile:task=>task},
    './default-manager-profile.js': {DEFAULT_MANAGER_PROFILE:{schemaVersion:1,name:'内置经理.xlsx',sha256:'d'.repeat(64),records:[]},DEFAULT_MANAGER_TEMPLATE:{name:'内置经理.xlsx',url:'assets/manager-introductions.xlsx'}},
    './manager-store.js': {createManagerStore:()=>managerStore||{load:async()=>null,save:async()=>{},clear:async()=>{}}},
    './manager-data.js': {parseManagerProfile:parseManagerProfile|| (async({name})=>({schemaVersion:1,name,sha256:'e'.repeat(64),records:[]}))},
    './manager-profile.js': {applyManagerProfile(task,current,options){onApplyManager?.(task,current,options);task.managerProfile={name:current.name,productionDate:options.productionDate};return task;}},
    './default-company-profile.js': {loadDefaultCompanyProfile: variant => profile(variant==='ccb'?'builtin-ccb.docx':'builtin-company.docx', 'c'), getDefaultCompanyTemplate:variant=>({name:variant==='ccb'?'builtin-ccb.docx':'builtin-company.docx',url:'https://example.test/'+(variant||'general')+'.docx'}),isSupersededCompanyProfile: companyDefaults.isSupersededCompanyProfile},
    './profile-store.js': {createProfileStore: options => storeFactory?storeFactory(options):saved},
    './company-template.js': {prepareCompanyTemplate:prepareCompanyTemplate|| (async file=>file)},
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
    ...naming, ...productionDates, document, Date:clock, Intl, Uint8Array, Blob, JSZip,
    URL: {createObjectURL: blob => {const url=`blob:test-${++urlNumber}`;blobs.set(url,blob);return url;}, revokeObjectURL() {}},
    requestAnimationFrame: callback => queueMicrotask(callback),
    window: {scrollTo() {}},
    importModule: async path => { const modulePath=path.split('?')[0]; assert.ok(modules[modulePath], `Unexpected app import: ${path}`); return modules[modulePath]; },
  });
  const source = appSource
    .replace(/^import[^\n]+\n/gm, '')
    .replace(/\bimport\(/g, 'importModule(')
    .replace(/const appReady=Promise.all\(\[restoreCompany\(\),restoreManager\(\)\]\);/, 'globalThis.appReady=Promise.all([restoreCompany(),restoreManager()]);');
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

test('introduction-only is off by default and full mode still requires a contract',async()=>{
 const app=createApp();await app.ready;
 assert.equal(app.elements.get('introductions-only').checked,false);
 await app.elements.get('templates-input').dispatch('change',{target:{files:[file('甲产品.docx',2)]}});
 assert.equal(app.elements.get('analyze-button').disabled,true);
 const toggle=app.elements.get('introductions-only');toggle.checked=true;await toggle.dispatch('change');
 assert.equal(app.elements.get('analyze-button').disabled,false);assert.equal(app.elements.get('contract-drop').hidden,true);
 toggle.checked=false;await toggle.dispatch('change');assert.equal(app.elements.get('analyze-button').disabled,true);assert.equal(app.elements.get('contract-drop').hidden,false);
});

test('introduction-only bypasses even a previously selected contract and prepares direct downloads for multiple managers',async()=>{
 let readContracts=0,analyses=0,writes=0,downloads=0;
 const app=createApp({readWordContract:async()=>{readContracts++;throw new Error('must not read');},onAnalyze:()=>analyses++,onWrite:task=>{writes++;assert.equal(task.updateMode,'introductions');assert.equal(task.facts.length,2);},onDownload:()=>downloads++,onApplyManager(task){
  if(task.updateMode!=='introductions')return;
  task.facts=[];for(const m of task.materials){m.fields=[];for(const name of ['甲经理','乙经理']){const id=m.materialId+name;task.facts.push({id,key:'manager:'+name,label:'投资经理 · '+name,source:'manager',status:'verified',value:name+'完整简介',evidence:[],scope:{manager:name},occurrences:[{materialId:m.materialId,paragraphId:id,fieldId:id}]});m.fields.push({id,label:name,source:'manager',status:'verified',factIds:[id],paragraphIds:[id],before:'旧简介',after:name+'完整简介'});}}
 }});await app.ready;await chooseMaterials(app);
 const toggle=app.elements.get('introductions-only');toggle.checked=true;await toggle.dispatch('change');
 await app.elements.get('analyze-button').click();
 assert.equal(readContracts,0);assert.equal(analyses,0);assert.equal(writes,1);assert.equal(downloads,0,'prepared links should not silently start a download');
 assert.ok(app.elements.get('download-list').children.some(n=>n.tagName==='a'));
 assert.equal(app.elements.get('supplement-panel').hidden,true);assert.equal(app.state.filter,'all');
 assert.equal(app.state.task.target.name,'公司和经理介绍更新');
 app.elements.get('month-value').value='2027-01-30';await app.elements.get('month-value').dispatch('change');
 assert.equal(writes,2,'a date change refreshes introduction files');assert.equal(downloads,0);
 assert.equal(app.state.task.productionDate,'2027-01-30');assert.equal(app.state.task.month,'202701');
 assert.equal(readContracts,0);assert.equal(analyses,0);assert.equal(app.state.task.facts.length,2);
 assert.match(app.elements.get('generation-status').textContent,/直接点击/);
 toggle.checked=false;await toggle.dispatch('change');assert.equal(app.state.task,null);assert.equal(app.elements.get('download-list').children.length,0);
});

test('same manager source is shown once across Word and PPT while different biographies remain separate',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const make=(id,scope,value='完整简介')=>({id,key:'manager:张三',label:'投资经理 · 张三',source:'manager',status:'verified',scope:{manager:'张三',recordId:'row-2',...scope},value,evidence:[{blockId:'manager-profile-hash-row-2'}],occurrences:[]});
 app.state.task.facts.push(make('manager-word',{}),make('manager-ppt',{page:2,pages:[2]}),make('manager-different',{page:3,pages:[3]},'另一份完整简介'));
 app.state.filter='all';app.renderReview();
 const cards=app.elements.get('facts-list').children.filter(n=>n.className.includes('manager-review'));
 assert.equal(cards.length,2);assert.ok(descendants(cards[0]).some(n=>n.textContent==='完整简介'));assert.ok(descendants(cards[1]).some(n=>n.textContent==='另一份完整简介'));
});

test('manager Excel can be uploaded, downloaded and restored independently of company templates',async()=>{
 const saved=[];const app=createApp({managerStore:{load:async()=>null,save:async r=>saved.push(r),clear:async()=>{}}});await app.ready;
 const companyBefore=app.state.company;
 assert.ok(app.elements.get('manager-input'),'Manager upload is available');
 await app.elements.get('manager-input').dispatch('change',{target:{files:[file('经理最新版.xlsx',7)]}});
 assert.equal(app.state.manager.name,'经理最新版.xlsx');assert.equal(saved.length,1);
 assert.deepEqual([...saved[0].bytes],[7]);assert.equal(app.state.company,companyBefore);
 const link=app.elements.get('manager-download');assert.equal(link.download,'经理最新版.xlsx');
 assert.deepEqual([...new Uint8Array(await app.blobs.get(link.href).arrayBuffer())],[7]);
 const reopened=createApp({managerStore:{load:async()=>saved[0]}});await reopened.ready;
 assert.equal(reopened.state.manager.name,'经理最新版.xlsx');assert.equal(reopened.elements.get('manager-download').download,'经理最新版.xlsx');
 await app.elements.get('manager-clear').click();assert.equal(app.state.manager.name,'内置经理.xlsx');
 assert.equal(app.elements.get('manager-download').href,'assets/manager-introductions.xlsx');
});

test('production date recalculates manager profiles while preserving confirmed product values',async()=>{
 const dates=[];const app=createApp({onApplyManager:(task,profile,options)=>dates.push(options.productionDate)});await app.ready;await chooseMaterials(app);
 const input=app.elements.get('month-value');input.value='2026-10-29';await input.dispatch('change');await app.elements.get('analyze-button').click();
 const task=app.state.task;task.facts[0].value='R2';task.facts[0].userConfirmed=true;
 await app.elements.get('generate-button').click();assert.ok(app.elements.get('download-list').children.length);
 input.value='2026-11-07';await input.dispatch('change');
 assert.equal(app.state.task,task);assert.equal(task.facts[0].value,'R2');assert.equal(task.facts[0].userConfirmed,true);
 assert.equal(task.productionDate,'2026-11-07');assert.equal(task.month,'202611');
 assert.deepEqual(dates,['2026-10-29','2026-11-07']);assert.equal(app.elements.get('download-list').children.length,0);
});

test('a fresh page and reset use today, but an open task does not silently change at midnight',async()=>{
 let now='2026-10-29T15:59:00Z';class Clock extends Date{constructor(...args){super(...(args.length?args:[now]));}}
 const app=createApp({clock:Clock});await app.ready;assert.equal(app.state.productionDate,'2026-10-29');
 now='2026-10-29T16:01:00Z';await chooseMaterials(app);await app.elements.get('analyze-button').click();
 assert.equal(app.state.task.productionDate,'2026-10-29');
 const reopened=createApp({clock:Clock});await reopened.ready;assert.equal(reopened.state.productionDate,'2026-10-30');
 await app.elements.get('month-reset').click();assert.equal(app.state.productionDate,'2026-10-30');
 assert.equal(app.state.task.productionDate,'2026-10-30');
});

test('manager drag upload and failed replacement retain the usable current template without a blocking error',async()=>{
 const app=createApp({parseManagerProfile:async({name})=>{if(name==='bad.xlsx')throw new Error('bad zip');return{schemaVersion:1,name,sha256:'e'.repeat(64),records:[]};}});await app.ready;
 await app.elements.get('manager-drop').dispatch('drop',{dataTransfer:{files:[file('拖入经理.xlsx')]}});
 assert.equal(app.state.manager.name,'拖入经理.xlsx');const current=app.state.manager;
 await app.elements.get('manager-input').dispatch('change',{target:{files:[file('bad.xlsx')]}});
 assert.equal(app.state.manager,current);assert.equal(app.elements.get('manager-download').download,'拖入经理.xlsx');
 assert.equal(app.elements.get('error-panel').hidden,true);
});

test('an invalid manager cache falls back to builtin and leaves a working clear action',async()=>{
 let cleared=0;const app=createApp({managerStore:{load:async()=>({name:'损坏缓存.xlsx',sha256:'0'.repeat(64),bytes:new Uint8Array([1])}),clear:async()=>{cleared++;}}});await app.ready;
 assert.equal(app.state.manager.name,'内置经理.xlsx');assert.equal(app.elements.get('manager-clear').hidden,false);
 await app.elements.get('manager-clear').click();assert.equal(cleared,1);assert.equal(app.elements.get('manager-clear').hidden,true);
});

test('clearing the production date cannot export under an unintended date',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();const original=app.state.task.productionDate;
 app.elements.get('month-value').value='';await app.elements.get('month-value').dispatch('change');assert.equal(app.state.task.productionDate,original);assert.equal(app.elements.get('generate-button').disabled,true);
 await app.elements.get('month-reset').click();assert.equal(app.elements.get('generate-button').disabled,false);
});

function narrativeTask(contract,templates,options){
 const rows=[['产品名称','晨星稳健3号集合资产管理计划'],['投资范围','本计划投资债券、货币市场工具。'],['投资范围','本计划可投资公募基金。'],['投资比例','固定收益资产占80%以上。'],['投资比例','权益类资产不超过20%。'],['A类份额投资范围','本份额投资债券。'],['B类份额投资范围','本份额投资基金。']];
 const t={...templates[0],id:'narrative',rows:[],paragraphs:[],slides:[],warnings:[]};
 for(const [i,[label,value]]of rows.entries()){
  const ids=['l'+i,'v'+i];[label,value].forEach((text,cellIndex)=>t.paragraphs.push({id:ids[cellIndex],text,part:'word/document.xml',page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));
  t.rows.push({id:'r'+i,label,value,part:'word/document.xml',page:1,tableId:'t',rowIndex:i,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]});
 }
 return reviewEngine.analyzeMaterials({...contract,warnings:[],blocks:[{id:'name',text:'产品名称：晨星稳健9号集合资产管理计划'}]},[t],options);
}

const allocationClauses=[
 '（1）固定收益类资产：占集合计划资产总值的20%～100%（不含）；',
 '（2）权益类资产：占集合计划资产总值的0%～40%（不含）；',
 '（3）期货和衍生品类账户权益：占集合计划资产总值的0%～20%（不含）；',
 '（4）公募基金：占集合计划资产总值的0%～80%；',
 '（5）境外投资：占集合计划资产总值的0%～50%（不含）。',
];
function incompleteAllocationTask(contract,templates,options){
 const base=narrativeTask(contract,templates,options),t=base.templates[0];
 const fields=[['产品名称','晨星稳健3号集合资产管理计划'],...allocationClauses.slice(0,3).map(text=>['投资比例',text])];
 t.rows=[];t.paragraphs=[];
 for(const [i,[label,value]]of fields.entries()){
  const ids=['label'+i,'value'+i];[label,value].forEach((text,cellIndex)=>t.paragraphs.push({id:ids[cellIndex],text,part:'word/document.xml',page:1,index:i*2+cellIndex,tableId:'t',rowIndex:i,cellIndex}));
  t.rows.push({id:'row'+i,label,value,tableId:'t',part:'word/document.xml',page:1,paragraphIds:ids,cells:[{text:label,paragraphIds:[ids[0]]},{text:value,paragraphIds:[ids[1]]}]});
 }
 const texts=['产品名称：晨星稳健9号集合资产管理计划','3、投资比例',...allocationClauses,'上述投资比例应当按照穿透原则合并计算。','4、产品风险等级','本计划风险等级为R3。'];
 return reviewEngine.analyzeMaterials({...contract,blocks:texts.map((text,i)=>({id:'contract-'+i,text})),warnings:[]},[t],options);
}
function sectionCaseTask(contract,templates,options,{label,value,clauses}){
 const t={...templates[0],id:'section-case',rows:[],paragraphs:[],slides:[],warnings:[]};
 for(const [i,[key,text]]of [['产品名称','晨星稳健9号集合资产管理计划'],[label,value]].entries()){
  const ids=['section-label'+i,'section-value'+i];[key,text].forEach((text,j)=>t.paragraphs.push({id:ids[j],text,index:i*2+j,part:'word/document.xml',page:1,tableId:'t',rowIndex:i,cellIndex:j}));
  t.rows.push({id:'row'+i,label:key,value:text,tableId:'t',part:'word/document.xml',page:1,paragraphIds:ids,cells:[{text:key,paragraphIds:[ids[0]]},{text,paragraphIds:[ids[1]]}]});
 }
 return reviewEngine.analyzeMaterials({...contract,blocks:['产品名称：晨星稳健9号集合资产管理计划',...clauses].map((text,i)=>({id:'sc'+i,text})),warnings:[]},[t],options);
}
test('a missing strategy title has a working on-card preview, apply and undo',async()=>{
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'投资策略',value:'1、资产配置策略；2、基金策略',clauses:['三、投资策略','1、资产配置策略','合理分配。','2、基金策略','优选基金。','3、境外结构化票据投资策略','谨慎管理票据风险。','四、投资限制']})});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const card=()=>app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资策略'));
 assert.ok(!card().children.some(n=>n.tagName==='form'));assert.ok(descendants(card()).some(n=>n.tagName==='summary'&&n.textContent==='分别补充原文'));await buttonIn(card(),'预览补齐策略标题').click();
 assert.ok(descendants(card()).some(n=>n.textContent==='1、资产配置策略；2、基金策略；3、境外结构化票据投资策略'));
 await buttonIn(card(),'确认补齐策略标题').click();assert.ok(app.state.task.materials[0].fields.find(f=>f.label==='投资策略').after.endsWith('3、境外结构化票据投资策略'));
 assert.ok(descendants(card()).some(n=>n.textContent.includes('已采用合同策略标题')));await buttonIn(card(),'撤销采用合同内容').click();assert.equal(app.state.task.materials[0].fields.find(f=>f.label==='投资策略').after,'1、资产配置策略；2、基金策略');
});
test('large restriction differences give a concise complete-evidence cue instead of broken paragraph supplements',async()=>{
 const clauses=['三、投资限制',...Array.from({length:15},(_,i)=>`（${i+1}）不得投资类别${i+1}的特定资产。`),'上述投资比例按照穿透原则合并计算。','四、风险揭示'];
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'投资限制',value:'不得进行违规投资。',clauses})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const card=app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资限制')),nodes=descendants(card);
 assert.ok(nodes.some(n=>/建议使用完整合同依据.*15项.*1段/.test(n.textContent)));assert.ok(!card.children.some(n=>n.tagName==='form'));assert.ok(!nodes.some(n=>n.textContent==='逐项查看合同与当前物料'));
 const evidence=nodes.filter(n=>n.tagName==='blockquote').map(n=>n.textContent).join('\n');assert.ok(evidence.includes('（15）'));assert.ok(evidence.includes('穿透原则'));
});

test('complete section coverage stays visible even when every existing template clause is verified',async()=>{
 const app=createApp({taskFactory:incompleteAllocationTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const task=app.state.task,allocations=task.facts.filter(f=>f.key==='investment_ratios');assert.ok(allocations.every(f=>f.status==='verified'));
 const card=app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资比例'));
 assert.ok(card,'A missing contract clause must appear in the default focus filter');
 const nodes=descendants(card);assert.ok(nodes.some(n=>n.className==='section-review-summary'&&/公募基金/.test(n.textContent)&&/境外投资/.test(n.textContent)));
 const evidence=nodes.filter(n=>n.tagName==='blockquote').map(n=>n.textContent).join('\n');
 for(const text of [...allocationClauses,'上述投资比例应当按照穿透原则合并计算。'])assert.ok(evidence.includes(text),text);
 assert.ok(!evidence.includes('本计划风险等级为R3'));
 assert.ok(Number(app.elements.get('count-pending').textContent)>0);
 const before=JSON.stringify(task.materials.map(m=>m.operations));
 await buttonIn(card,'已核对差异，保留当前表述').click();
 assert.equal(JSON.stringify(task.materials.map(m=>m.operations)),before,'Coverage acknowledgement cannot rewrite any Office content');
 assert.equal(Number(app.elements.get('count-pending').textContent),0);
 app.state.filter='pending';app.renderReview();assert.ok(!app.elements.get('facts-list').children.some(n=>n.tagName==='article'));
 app.state.filter='focus';app.renderReview();
 const confirmedCard=app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.textContent==='已核对差异，保留当前表述'));
 assert.ok(confirmedCard,'Acknowledged coverage remains inspectable in focus');
 await buttonIn(confirmedCard,'撤销差异确认').click();assert.ok(Number(app.elements.get('count-pending').textContent)>0);
});

test('changing the acknowledged wording or tail explanation invalidates coverage acknowledgement',async()=>{
 const app=createApp({taskFactory:incompleteAllocationTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const findCard=()=>app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资比例'));
 await buttonIn(findCard(),'已核对差异，保留当前表述').click();assert.equal(Number(app.elements.get('count-pending').textContent),0);
 const field=app.state.task.materials[0].fields.find(f=>f.label==='投资比例');
 field.after+='\n上述比例按资产净值计算。';app.renderReview();
 assert.ok(Number(app.elements.get('count-pending').textContent)>0);
 assert.ok(buttonIn(findCard(),'已核对差异，保留当前表述'));
 assert.ok(!buttonIn(findCard(),'撤销差异确认'));
});

test('ratio adoption previews complete content, applies only on confirmation and can be undone',async()=>{
 const app=createApp({taskFactory:incompleteAllocationTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const card=()=>app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资比例'));
 const before=JSON.stringify(app.state.task.materials.map(m=>m.fields.map(f=>f.after)));
 await buttonIn(card(),'预览采用合同投资比例').click();assert.equal(JSON.stringify(app.state.task.materials.map(m=>m.fields.map(f=>f.after))),before);
 assert.ok(descendants(card()).some(n=>n.textContent.includes('境外投资')));
 await buttonIn(card(),'确认采用合同投资比例').click();assert.equal(Number(app.elements.get('count-pending').textContent),0);
 assert.ok(app.state.task.materials[0].fields.map(f=>f.after).join('\n').includes('境外投资'));
 await buttonIn(card(),'撤销采用合同内容').click();assert.equal(JSON.stringify(app.state.task.materials.map(m=>m.fields.map(f=>f.after))),before);
});

test('fixed monthly opening review shows full body evidence and supports preview, adoption and undo',async()=>{
 const current='本集合计划自成立次日起每月1日、15日（如遇节假日顺延至下一个工作日）开放办理参与和退出业务。合同另有约定的，从其约定。除因根据合同变更或者展期等需要设立临时开放期供投资者退出外，本集合计划原则上不设立特别开放期。';
 const old='本集合计划每月10号和25号开放办理参与和退出业务，非交易日顺延至下一个交易日；每笔份额持有期为180个自然日。';
 const definition='指本集合计划办理参与和退出业务的期间。';
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'开放期',value:old,clauses:['一、释义','开放期',definition,'四、封闭期、开放期','2、开放期：'+current]})});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const card=()=>app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='开放期'));
 const field=()=>app.state.task.materials[0].fields.find(f=>f.label==='开放期');
 assert.equal(field().status,'pending');assert.equal(field().after,old);
 const nodes=descendants(card()),evidence=nodes.filter(n=>n.tagName==='blockquote').map(n=>n.textContent).join('\n');
 assert.ok(evidence.includes(current),'The full body rule, including the special-opening tail, remains inspectable');
 assert.ok(!evidence.includes(definition),'The glossary definition is not offered as the actionable rule');
 assert.ok(nodes.some(n=>n.className==='section-review-summary'&&/工作日/.test(n.textContent)&&/持有/.test(n.textContent)));
 assert.ok(!nodes.some(n=>n.tagName==='form'));
 const preview=buttonIn(card(),'预览采用合同开放期');assert.ok(preview);await preview.click();
 assert.equal(field().after,old,'Preview does not write the selected body rule');
 assert.ok(descendants(card()).some(n=>n.textContent===current));
 await buttonIn(card(),'确认采用合同开放期').click();
 assert.equal(field().after,current);assert.equal(field().status,'verified');assert.equal(Number(app.elements.get('count-pending').textContent),0);
 assert.doesNotMatch(field().after,/180|10号|25号|下一个交易日/);
 assert.ok(descendants(card()).some(n=>n.textContent.includes('已采用合同开放期规则')));
 await buttonIn(card(),'撤销采用合同内容').click();
 assert.equal(field().after,old);assert.equal(field().status,'pending');assert.ok(Number(app.elements.get('count-pending').textContent)>0);
 assert.ok(buttonIn(card(),'预览采用合同开放期'));assert.ok(!buttonIn(card(),'撤销采用合同内容'));
});

test('no-fee review never displays old charging parameters or an irrelevant supplement form',async()=>{
 const app=createApp({taskFactory:(contract,templates,options)=>{
  const task=narrativeTask(contract,templates,options),m=task.materials[0];
  const f={id:'policy',key:'performance_policy',label:'业绩报酬',status:'verified',value:'本集合计划不提取业绩报酬。',policyMode:'none',source:'contract',scope:{product:task.target.name},evidence:[{blockId:'c',quote:'本集合计划不提取业绩报酬。'}],occurrences:[{materialId:m.materialId,paragraphId:'oldfee'}]};task.facts=[f];
  m.fields=[{id:'fee',label:'业绩报酬',before:'超过5%的部分提取20%作为业绩报酬。',after:f.value,factIds:[f.id],paragraphIds:['oldfee']}];m.operations=[{kind:'replace',factIds:[f.id]}];return task;
 }});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const card=app.elements.get('facts-list').children.find(n=>n.tagName==='article');const nodes=descendants(card);
 assert.ok(nodes.some(n=>n.textContent==='本集合计划不提取业绩报酬。'));assert.ok(!nodes.some(n=>/5%|20%/.test(n.textContent)));assert.ok(!nodes.some(n=>n.tagName==='form'));assert.ok(!buttonIn(card,'确认无误'));
});

test('download review notes retain full section evidence and distinguish unresolved coverage from confirmed wording',async()=>{
 const app=createApp({taskFactory:incompleteAllocationTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 await app.elements.get('generate-button').click();
 const link=app.elements.get('download-list').children.find(n=>n.tagName==='a'&&n.download.endsWith('.zip'));
 const zip=await JSZip.loadAsync(await app.blobs.get(link.href).arrayBuffer());
 const notes=await zip.file('审阅清单.txt').async('string');
 assert.match(notes,/章节覆盖待核/);assert.match(notes,/公募基金/);assert.match(notes,/境外投资/);assert.match(notes,/穿透原则合并计算/);
 assert.equal(app.elements.get('error-panel').hidden,true,'Coverage warning must not block download');
});
test('narrative review cards combine display and confirmation while keeping each original supplement separate',async()=>{
 const app=createApp({taskFactory:narrativeTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const cards=()=>app.elements.get('facts-list').children.filter(n=>n.tagName==='article');
 const card=label=>cards().find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent===label));
 for(const label of ['投资范围','投资比例'])assert.equal(cards().filter(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent===label)).length,1);
 assert.ok(card('A类份额投资范围 · A类份额'));assert.ok(card('B类份额投资范围 · B类份额'));
 const task=app.state.task,range=task.facts.filter(f=>f.key==='investment_scope'&&!f.scope.share);assert.equal(range.length,2);
 const rangeCard=card('投资范围'),forms=descendants(rangeCard).filter(n=>n.tagName==='form');assert.equal(forms.length,2);
 for(const original of ['本计划投资债券、货币市场工具。','本计划可投资公募基金。'])assert.ok(descendants(rangeCard).some(n=>n.textContent===original));
 const firstForm=forms[0],input=firstForm.children.find(n=>['input','textarea'].includes(n.tagName));input.value='本计划主要投资债券。';
 firstForm.onsubmit({preventDefault(){}});await buttonIn(firstForm,'确认并应用补充').click();
 assert.equal(range[0].value,'本计划主要投资债券。');assert.equal(range[1].value,null);assert.equal(range[1].status,'pending');
 const material=task.materials[0];assert.ok(material.operations.some(o=>o.kind==='replace'&&o.paragraphId==='v1'));assert.ok(!material.operations.some(o=>o.kind==='replace'&&o.paragraphId==='v2'));
 await buttonIn(card('投资范围'),'确认无误').click();assert.equal(range[1].userConfirmed,true);assert.ok(!range[0].userConfirmed);
 await buttonIn(card('投资范围'),'撤销确认').click();assert.equal(range[1].status,'pending');assert.equal(range[0].value,'本计划主要投资债券。');
 assert.ok(task.facts.filter(f=>f.key==='investment_ratios').every(f=>f.status==='pending'));
});

test('review display keeps same-label channel share and semantic conditions separate',async()=>{
 const app=createApp({taskFactory:narrativeTask});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const task=app.state.task,fs=task.facts.filter(f=>f.key==='investment_scope');
 const count=()=>app.elements.get('facts-list').children.filter(n=>n.tagName==='article'&&descendants(n).some(c=>c.tagName==='strong'&&c.textContent.startsWith('投资范围'))).length;
 for(const f of fs)f.label='投资范围';app.renderReview();assert.equal(count(),3,'common group plus A and B');
 const common=fs.filter(f=>!f.scope.share);common[0].scope.channel='银行甲';common[1].scope.channel='银行乙';app.renderReview();assert.equal(count(),4);
 for(const f of common)f.scope.channel='';common[0].scope.condition='aggregate_layers';app.renderReview();assert.equal(count(),4,'semantic conditions must not disappear');
});

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

test('Excel with only an end date labels the cutoff and sends it for range preview without applying',async()=>{
 const row={id:'row-2',productName:'示例稳健1号',openingEnd:'2026/9/14',source:{file:'日历.xlsx',sheet:'Sheet1',row:2}};
 const app=await appForCalendar({readProductCalendar:async()=>({name:'日历.xlsx',rows:[row],warnings:[]}),matchCalendarRows:()=>[row]});
 await app.elements.get('calendar-drop').dispatch('drop',{dataTransfer:{files:[file('日历.xlsx')]}});
 const details=descendants(app.elements.get('calendar-row-detail')).map(n=>n.textContent).join('\n');
 assert.match(details,/开放截止日：2026\/9\/14/);assert.doesNotMatch(details,/不自动套用/);
 assert.match(app.state.supplementPreview.text,/本次开放截止日：2026\/9\/14/);
 assert.equal(app.state.task.facts.find(f=>f.id==='date-a').value,null);
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
 input.value='000123';await form.onsubmit({preventDefault(){}});const oldButton=buttonIn(form,'确认并应用补充');
 input.value='000999';await input.dispatch('input');await oldButton.click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);assert.equal(app.elements.get('supplement-preview').hidden,true);
});

function supplementEntry(app,label){
 const input=descendants(app.elements.get('facts-list')).find(n=>n['aria-label']==='补充'+label);
 assert.ok(input,`Missing supplement input: ${label}`);
 const form=descendants(app.elements.get('facts-list')).find(n=>n.tagName==='form'&&n.children.includes(input));
 return {input,form,feedback:()=>descendants(form).find(n=>n.className==='supplement-preview card-supplement-feedback')};
}
async function pendingApp(){const app=createApp();await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();appendSupplementFacts(app);return app;}
test('card supplements show blank and invalid input feedback beside the input, preserving text',async()=>{
 const app=await pendingApp(),entry=supplementEntry(app,'风险等级');
 await entry.form.onsubmit({preventDefault(){}});
 assert.ok(entry.feedback()&&!entry.feedback().hidden,'Blank input needs visible feedback on its card');
 assert.ok(descendants(entry.feedback()).some(n=>/请填写/.test(n.textContent)));
 entry.input.value='错误风险';await entry.form.onsubmit({preventDefault(){}});
 assert.equal(entry.input.value,'错误风险');assert.ok(descendants(entry.feedback()).some(n=>/R1至R5/.test(n.textContent)));
 assert.equal(app.state.supplementPreview,null);assert.equal(app.elements.get('error-panel').hidden,true);
 assert.ok(!buttonIn(entry.feedback(),'确认并应用补充'));
});
test('card preview, cancel and apply stay on the selected card',async()=>{
 const app=await pendingApp(),entry=supplementEntry(app,'产品代码');entry.input.value='000123';
 await entry.form.onsubmit({preventDefault(){}});assert.ok(entry.feedback()&&!entry.feedback().hidden);
 assert.ok(descendants(entry.feedback()).some(n=>n.textContent.includes('000123')));
 assert.equal(app.elements.get('supplement-preview').hidden,true);
 const old=buttonIn(entry.feedback(),'确认并应用补充');await buttonIn(entry.feedback(),'取消').click();await old.click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null);
 await entry.form.onsubmit({preventDefault(){}});await buttonIn(entry.feedback(),'确认并应用补充').click();
 assert.equal(app.state.task.facts.find(f=>f.id==='code').value,'000123');
});
test('a successfully supplemented text card remains editable for a second correction',async()=>{
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'风险等级',value:'R3（中风险）',clauses:[]})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 let entry=supplementEntry(app,'风险等级');entry.input.value='R2';await entry.form.onsubmit({preventDefault(){}});await buttonIn(entry.form,'确认并应用补充').click();
 entry=supplementEntry(app,'风险等级');entry.input.value='R1';await entry.form.onsubmit({preventDefault(){}});await buttonIn(entry.form,'确认并应用补充').click();
 assert.equal(app.state.task.materials[0].fields.find(f=>f.label==='风险等级').after,'R1（低风险）');
});
test('contract holding-day fee tiers share one card with an optional full-text supplement',async()=>{
 const original='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'退出费率',value:original.replace('1%','2%').replace('0%','1%'),clauses:[original]})});
 await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const cards=app.elements.get('facts-list').children.filter(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='退出费率'));
 assert.equal(cards.length,1);assert.ok(app.state.task.facts.filter(f=>f.key==='redemption_fee').every(f=>f.status==='verified'));
 const entry=supplementEntry(app,'退出费率');entry.input.value=original;await entry.form.onsubmit({preventDefault(){}});
 assert.equal(app.state.supplementPreview.updates.length,2);assert.ok(buttonIn(entry.form,'确认并应用补充'));
 await buttonIn(entry.form,'确认并应用补充').click();assert.equal(app.state.task.materials[0].fields.find(f=>f.label==='退出费率').after,original);
});
test('fee card accepts complete wording over a single-rate template and offers undo after confirmation',async()=>{
 const full='持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%。';
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'退出费率',value:'退出费率0.5%。',clauses:[]})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 let entry=supplementEntry(app,'退出费率');entry.input.value=full;await entry.form.onsubmit({preventDefault(){}});
 assert.ok(app.state.supplementPreview?.feeText);assert.ok(descendants(entry.feedback()).some(n=>n.textContent.includes('完整文字')));
 await buttonIn(entry.form,'确认并应用补充').click();assert.equal(app.state.task.materials[0].fields.find(f=>f.label==='退出费率').after,full);
 entry=supplementEntry(app,'退出费率');assert.ok(entry.input);await buttonIn(app.elements.get('facts-list'),'撤销整段修改').click();assert.equal(app.state.task.materials[0].fields.find(f=>f.label==='退出费率').after,'退出费率0.5%。');
});
test('investment range, ratio, strategy and restriction cards offer an explicit complete custom replacement with preview and undo',async()=>{
 for(const [label,old,replacement]of [
  ['投资比例','（1）债券：0%—80%；\n（2）公募基金：0%—80%；\n（3）衍生品：0%—20%。','（1）债券：20%—100%；\n（2）货币市场工具：0%—80%。'],
  ['投资范围','（1）债券。\n（2）公募基金。\n（3）金融衍生品。','（1）债券。\n（2）货币市场工具。'],
  ['投资策略','（1）资产配置策略\n（2）债券投资策略\n（3）基金投资策略','（1）资产配置策略\n根据市场变化调整配置。\n（2）债券投资策略'],
  ['投资限制','（1）不得从事违规交易。\n（2）单一资产不超过10%。\n（3）不得投资未上市股权。','（1）不得从事违规交易。\n（2）单一资产不超过15%。'],
 ]){
  const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label,value:old,clauses:[]})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
  const card=()=>app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent===label));
  assert.ok(descendants(card()).some(n=>n.tagName==='summary'&&n.textContent==='分别补充原文'));
  assert.ok(descendants(card()).some(n=>n.tagName==='summary'&&n.textContent==='自定义替换'));
  let input=descendants(card()).find(n=>n['aria-label']==='自定义替换'+label);assert.ok(input);
  input.value=replacement;
  await buttonIn(card(),'预览整项替换').click();assert.equal(app.state.task.materials[0].fields[1].after,old);
  const confirmed=buttonIn(card(),'确认整项替换');assert.ok(confirmed);await confirmed.click();
  assert.equal(app.state.task.materials[0].fields[1].after,input.value);
  assert.ok(descendants(card()).some(n=>n.textContent==='已自定义替换'));
  assert.ok(!buttonIn(card(),'预览采用合同投资比例'));
  await buttonIn(card(),'撤销自定义替换').click();assert.equal(app.state.task.materials[0].fields[1].after,old);
 }
});

test('custom replacement ignores empty input and invalidates preview after text changes',async()=>{
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'投资比例',value:'原投资比例条目。',clauses:[]})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const box=app.elements.get('facts-list'),input=descendants(box).find(n=>n['aria-label']==='自定义替换投资比例');assert.ok(input);
 await buttonIn(box,'预览整项替换').click();assert.ok(!buttonIn(box,'确认整项替换'));assert.equal(app.state.task.materials[0].fields[1].after,'原投资比例条目。');
 input.value='本计划投资债券80%以上。';await buttonIn(box,'预览整项替换').click();const apply=buttonIn(box,'确认整项替换');assert.ok(apply);
 input.value='本计划投资债券90%以上。';await input.dispatch('input');await apply.click();assert.equal(app.state.task.materials[0].fields[1].after,'原投资比例条目。');
});

test('investment evidence shows one full version and folds repeated other positions with differences',async()=>{
 const app=createApp({taskFactory:(c,t,o)=>sectionCaseTask(c,t,o,{label:'投资范围',value:'旧范围：债券与基金。',clauses:[]})});await app.ready;await chooseMaterials(app);await app.elements.get('analyze-button').click();
 const primary='投资范围\n债券。\n货币市场工具。\n特别揭示：债券回购风险。',alternate='投资范围\n债券。\n货币市场工具。';
 app.state.reviewApi={...app.state.reviewApi,getSectionReview:(t,ids)=>t.facts.find(f=>f.id===ids[0])?.key==='investment_scope'?{id:'scope-view',key:'investment_scope',label:'投资范围',factIds:ids,materials:[],summary:'请核对完整范围。',needsReview:true,evidence:[{blockId:'body',quote:primary,itemCount:2,noteCount:1},{blockId:'annex',quote:alternate,itemCount:2,noteCount:0}]}:null};
 app.renderReview();const card=app.elements.get('facts-list').children.find(n=>descendants(n).some(c=>c.tagName==='strong'&&c.textContent==='投资范围'));
 const details=descendants(card).find(n=>n.tagName==='details'&&n.children.some(c=>c.tagName==='summary'&&/查看.*依据.*位置/.test(c.textContent)));
 assert.equal(details.children.filter(n=>n.tagName==='blockquote'&&n.textContent===primary).length,1);
 assert.equal(details.children.filter(n=>n.tagName==='blockquote'&&n.textContent===alternate).length,0,'alternate full quote must be folded');
 const other=details.children.find(n=>n.tagName==='details'&&descendants(n).some(c=>c.textContent===alternate));assert.ok(other);
 assert.ok(descendants(other).some(n=>n.tagName==='summary'&&/2段相同/.test(n.textContent)));
 assert.ok(descendants(other).some(n=>/本处未含/.test(n.textContent)&&/特别揭示/.test(n.textContent)));
 assert.ok(details.children.some(n=>n.tagName==='h4'&&n.textContent==='物料原文与当前结果'));
});

test('no-match and apply errors remain visible beside the triggering card',async()=>{
 const app=await pendingApp(),entry=supplementEntry(app,'产品代码');entry.input.value='000123';
 const api=app.state.reviewApi;app.state.reviewApi={...api,previewSupplement:()=>({updates:[],warnings:['未定位到对应位置']})};
 await entry.form.onsubmit({preventDefault(){}});assert.ok(entry.feedback()&&!entry.feedback().hidden);
 assert.ok(descendants(entry.feedback()).some(n=>/没有可直接应用/.test(n.textContent)));assert.ok(entry.feedback().scrolledIntoView);
 app.state.reviewApi={...api,applySupplement:()=>{throw new Error('本次补充已失效，请重新预览。');}};
 await entry.form.onsubmit({preventDefault(){}});await buttonIn(entry.feedback(),'确认并应用补充').click();
 assert.ok(descendants(entry.feedback()).some(n=>/重新预览/.test(n.textContent)));assert.equal(app.state.supplementPreview,null);
 assert.ok(!buttonIn(entry.feedback(),'确认并应用补充'));
});
test('empty resubmit, another card and upper inputs invalidate previous card apply buttons',async()=>{
 for(const action of ['empty','invalid','other-card','upper-input','switch-tab']){
  const app=await pendingApp(),entry=supplementEntry(app,'产品代码');entry.input.value='000123';await entry.form.onsubmit({preventDefault(){}});
  const old=buttonIn(entry.feedback(),'确认并应用补充');assert.ok(old);
  if(action==='empty'||action==='invalid'){entry.input.value=action==='empty'?'':'?';await entry.form.onsubmit({preventDefault(){}});}
  if(action==='other-card'){const other=supplementEntry(app,'风险等级');other.input.value='R2';await other.form.onsubmit({preventDefault(){}});}
  if(action==='upper-input')await app.elements.get('supplement-date').dispatch('input');
  if(action==='switch-tab')await app.elements.get('supplement-excel-tab').click();
  await old.click();assert.equal(app.state.task.facts.find(f=>f.id==='code').value,null,action);
  assert.ok(!buttonIn(entry.feedback(),'确认并应用补充'),action+' clears the obsolete button');
  if(!['empty','invalid'].includes(action))assert.equal(entry.feedback().hidden,true,action);
 }
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
 app.elements.get('month-value').value='2026-10-29';
 await app.elements.get('month-value').dispatch('change');
 const names=descendants(app.elements.get('delivery-list')).filter(n=>n.tagName==='strong').map(n=>n.textContent);
 assert.deepEqual(names,['【一页通202610】示例稳健1号.docx','【一页通202610】示例稳健1号（2）.docx']);
});

test('manual month survives analysis, user supplement, and export; later edits update both download names', async () => {
  const analyzed = [], written = [];
  const app = createApp({onAnalyze: options => analyzed.push(options.month), onWrite: task => written.push(task.month)});
  await app.ready;
  const month = app.elements.get('month-value');
  month.value = '2032-04-15';
  await month.dispatch('change');
  await chooseMaterials(app);
  await app.elements.get('analyze-button').click();
  assert.deepEqual(analyzed, ['203204']);
  assert.equal(app.state.task.month, '203204');
  const supplement = descendants(app.elements.get('facts-list')).find(element => element.tagName === 'form');
  assert.ok(supplement, 'Pending facts expose the actual supplement handler');
  supplement.children.find(element => element.tagName === 'input').value = 'R2';
  await supplement.onsubmit({preventDefault() {}});
  await buttonIn(supplement,'确认并应用补充').click();
  assert.equal(app.state.task.month, '203204');
  assert.equal(app.state.task.facts[0].value, 'R2');
  await app.elements.get('generate-button').click();
  assert.deepEqual(written, ['203204']);
  assert.deepEqual(app.elements.get('download-list').children.filter(node => node.tagName === 'a').map(node => node.download), ['示例稳健1号_203204.zip', '示例_203204.docx']);
  month.value = '2032-05-15';
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

test('company version starts general, switches both evidence and download, and a fresh page defaults to general',async()=>{
 const app=createApp();await app.ready;
 assert.equal(app.elements.get('company-variant')?.value,'general');
 assert.equal(app.elements.get('company-download').href,'https://example.test/general.docx');
 app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 assert.equal(app.state.company.name,'builtin-ccb.docx');assert.equal(app.elements.get('company-download').download,'builtin-ccb.docx');
 assert.equal(app.elements.get('company-download').href,'https://example.test/ccb.docx');assert.match(app.elements.get('company-status').textContent,/建行版/);
 await chooseMaterials(app);await app.elements.get('analyze-button').click();assert.equal(app.state.task.companyProfile.name,'builtin-ccb.docx');
 const fresh=createApp();await fresh.ready;assert.equal(fresh.state.companyVariant,'general');
});

test('company uploads, download bytes and clear stay isolated between variants',async()=>{
 const records=new Map(),calls=[],storeFactory=({variant})=>({load:async()=>records.get(variant)||null,save:async r=>{calls.push(variant);records.set(variant,r);},clear:async()=>records.delete(variant)});
 const app=createApp({storeFactory});await app.ready;
 await app.elements.get('company-input').dispatch('change',{target:{files:[file('generic-upload.docx',7)]}});
 assert.equal(app.elements.get('company-download').download,'generic-upload.docx');
 assert.deepEqual(new Uint8Array(await app.blobs.get(app.elements.get('company-download').href).arrayBuffer()),new Uint8Array([7]));
 app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 await app.elements.get('company-input').dispatch('change',{target:{files:[file('ccb-upload.docx',8)]}});
 assert.deepEqual(calls,['general','ccb']);assert.equal(records.size,2);
 await app.elements.get('company-clear').click();assert.equal(records.has('general'),true);assert.equal(records.has('ccb'),false);assert.equal(app.state.company.name,'builtin-ccb.docx');
 app.elements.get('company-variant').value='general';await app.elements.get('company-variant').dispatch('change');
 assert.equal(app.state.company.name,'generic-upload.docx');assert.equal(app.elements.get('company-download').download,'generic-upload.docx');
});

test('unsaved company uploads survive switching within the page and busy switches cannot change evidence',async()=>{
 const app=createApp({store:{load:async()=>null,save:async()=>{throw Error('空间不足');},clear:async()=>{}}});await app.ready;
 await app.elements.get('company-input').dispatch('change',{target:{files:[file('session-only.docx',9)]}});
 app.state.busy=true;app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 assert.equal(app.state.companyVariant,'general');assert.equal(app.elements.get('company-variant').value,'general');app.state.busy=false;
 app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 app.elements.get('company-variant').value='general';await app.elements.get('company-variant').dispatch('change');
 assert.equal(app.state.company.name,'session-only.docx');assert.equal(app.state.companyPersisted,false);assert.match(app.elements.get('company-status').textContent,/未能保存/);
 assert.deepEqual(new Uint8Array(await app.blobs.get(app.elements.get('company-download').href).arrayBuffer()),new Uint8Array([9]));
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

test('restoring version caches passes the selected variant and rejects an invalid replacement without losing its download',async()=>{
 const calls=[],saved={general:record(profile('saved-general.docx')),ccb:record(profile('saved-ccb.docx','b'))};
 const app=createApp({storeFactory:({variant})=>({load:async()=>saved[variant],save:async()=>{},clear:async()=>{}}),parseCompanyProfile:async(file,options)=>{calls.push(options.variant);if(file.name==='invalid.docx')throw Error('无法识别公司介绍');return saved[options.variant];}});
 await app.ready;assert.equal(app.elements.get('company-download').download,'saved-general.docx');
 app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 const oldLink=app.elements.get('company-download').href;
 await app.elements.get('company-input').dispatch('change',{target:{files:[file('invalid.docx')]}});
 assert.deepEqual(calls,['general','ccb','ccb']);assert.equal(app.state.company.name,'saved-ccb.docx');assert.equal(app.elements.get('company-download').href,oldLink);assert.match(app.elements.get('error-message').textContent,/仍保留此前版本/);
});

test('switching company version invalidates the prior generated task while retaining inputs and production month',async()=>{
 const app=createApp();await app.ready;await chooseMaterials(app);
 app.elements.get('month-value').value='2027-03-15';await app.elements.get('month-value').dispatch('change');
 await app.elements.get('analyze-button').click();await app.elements.get('generate-button').click();assert.ok(app.state.urls.length);
 const contract=app.state.contract,templates=app.state.templates;
 app.elements.get('company-variant').value='ccb';await app.elements.get('company-variant').dispatch('change');
 assert.equal(app.state.task,null);assert.equal(app.state.urls.length,0);assert.equal(app.state.contract,contract);assert.equal(app.state.templates,templates);assert.equal(app.state.month,'202703');
 await app.elements.get('analyze-button').click();assert.equal(app.state.task.companyProfile.name,'builtin-ccb.docx');assert.match(app.elements.get('review-subtitle').textContent,/公司介绍（建行版）/);
});

test('a legacy mixed-version cache is verified before splitting, and download uses the selected bytes',async()=>{
 const old=record(profile('mixed.docx')),calls=[];
 const app=createApp({store:{load:async()=>old,save:async()=>{},clear:async()=>{}},parseCompanyProfile:async file=>{calls.push(file.name);return profile(file.name,file.name==='mixed.docx'?'a':'b');},prepareCompanyTemplate:async(file,{variant})=>{assert.equal(variant,'general');return{name:'selected-general.docx',bytes:new Uint8Array([7])};}});
 await app.ready;assert.deepEqual(calls,['mixed.docx','selected-general.docx']);assert.equal(app.state.company.name,'selected-general.docx');
 assert.deepEqual(new Uint8Array(await app.blobs.get(app.elements.get('company-download').href).arrayBuffer()),new Uint8Array([7]));
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
  app.elements.get('month-value').value='2030-02-15';await app.elements.get('month-value').dispatch('change');
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
