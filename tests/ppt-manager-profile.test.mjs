import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {NS} from '../src/office-xml.js';
import {readTemplate} from '../src/office-reader.js';
import {applyPptManagerProfile} from '../src/ppt-manager-profile.js';
import {profile} from './manager-fixture.mjs';
Object.assign(globalThis,{JSZip,DOMParser,XMLSerializer});
const esc=s=>s.replaceAll('&','&amp;').replaceAll('<','&lt;');
const shape=(id,texts,[x,y,w,h]=[100,200,350,200])=>`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="shape-${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${texts.map(t=>`<a:p><a:r><a:rPr sz="1800"/><a:t>${esc(t)}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;
async function fixture(...pages){
 const z=new JSZip();z.file('ppt/presentation.xml',`<p:presentation xmlns:p="${NS.p}" xmlns:r="${NS.r}"><p:sldIdLst>${pages.map((_,i)=>`<p:sldId id="${256+i}" r:id="r${i+1}"/>`).join('')}</p:sldIdLst></p:presentation>`);
 z.file('ppt/_rels/presentation.xml.rels',`<Relationships xmlns="${NS.pr}">${pages.map((_,i)=>`<Relationship Id="r${i+1}" Type="${NS.r}/slide" Target="slides/slide${i+1}.xml"/>`).join('')}</Relationships>`);
 pages.forEach((s,i)=>z.file(`ppt/slides/slide${i+1}.xml`,`<p:sld xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:cSld><p:spTree>${s}</p:spTree></p:cSld></p:sld>`));
 return readTemplate({name:'【PPT202609】示例产品.pptx',bytes:await z.generateAsync({type:'uint8array'})});
}
const task=template=>({facts:[],materials:[{materialId:template.id,kind:'ppt',template,operations:[],fields:[],warnings:[]}],stats:{},warnings:[]});
const source=profile([{name:'张三',biography:'10年证券从业经验。金融博士。\n现任甲公司固定收益投资部总经理。\n擅长自下而上研究，专注稳定回报。'},{name:'李四',biography:'8年证券从业经验。经济学硕士。\n现任乙公司多资产投资部投资经理。\n擅长全球大类资产配置。'}]);
const body=(id,x,text='金融硕士。')=>shape(id,[text,'现任旧公司投资经理。','旧的投资风格。'],[x,180,360,300]);
const name=(id,x,text)=>shape(id,[text],[x,90,300,45]);
const paragraphText=(m,id)=>{let text=m.template.paragraphs.find(p=>p.id===id).text;for(const o of m.operations.filter(o=>o.paragraphId===id).sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;};
const resultText=(m,ids)=>ids.map(id=>paragraphText(m,id)).filter(Boolean).join('\n');
const assertSourceComplete=(m,field)=>{const output=resultText(m,field.biographyParagraphIds);for(const segment of field.sourceSegments)assert.ok(output.includes(segment.text),segment.text);};

test('semantic matching keeps separate education, experience, employment and style paragraph slots',async()=>{
 const education='甲大学数学系本科，乙大学金融工程硕士。',experience='14年证券从业经验，其中10年投资经验。',style='投资风格稳健，重视长期稳定的绝对收益。';
 const data=profile([{biography:experience+education+'\n现任甲公司固定收益投资部（私募）总经理、投资经理。历任乙银行衍生品顾问。\n'+style}]);
 const t=await fixture(shape('caption',['固定收益投资部（上海）总经理','张三'],[20,500,270,70])+shape('body',[education,experience,'现任甲公司固定收益投资部（上海）总经理兼任固定收益投资部（私募）总经理、投资经理。',style],[350,180,600,300]));
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0],ps=t.paragraphs.filter(p=>p.shapeId==='body');
 assert.equal(paragraphText(m,ps[0].id),education);assert.equal(paragraphText(m,ps[1].id),experience);assert.equal(paragraphText(m,ps[3].id),style);assert.ok(paragraphText(m,ps[2].id).includes('历任乙银行衍生品顾问。'));
 assert.ok(!m.operations.some(o=>[ps[0].id,ps[1].id,ps[3].id].includes(o.paragraphId)));assert.ok(m.operations.every(o=>!o.profileParagraphs&&!o.profileDeleteParagraph));
 assert.equal(paragraphText(m,t.paragraphs.find(p=>p.text==='固定收益投资部（上海）总经理').id),'固定收益投资部（私募）总经理');
 assert.ok(!m.operations.some(o=>o.before.includes('固定收益投资部')&&o.before===t.paragraphs.find(p=>p.id===o.paragraphId).text));assertSourceComplete(m,m.fields[0]);
});

test('education remains before tenure in one original paragraph and unchanged career text receives no edits',async()=>{
 const education='甲大学计量经济学博士，特许金融分析师',career='现任甲公司综合策略投资部投资经理，从事全球资产研究工作。加入甲公司之前，曾就职于乙银行，历任金融工程师、研究员。',style='拥有严谨的数据框架，成熟的量化体系；投研交易经验覆盖：国债期货、可转债套利和全球配置策略。';
 const data=profile([{biography:'7年证券从业经验，其中1年投资经验。'+education+'。\n'+career+'\n'+style}]);
 const t=await fixture(name('n',100,'张三')+shape('b',[''+education+'；10年金融证券从业经验，超过5年大类资产投资管理经验。',''+career,''+style,''],[100,180,600,300]));
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0],ps=t.paragraphs.filter(p=>p.shapeId==='b');
 assert.equal(paragraphText(m,ps[0].id),''+education+'；7年证券从业经验，其中1年投资经验。');assert.equal(paragraphText(m,ps[1].id),''+career);assert.equal(paragraphText(m,ps[2].id),''+style);assert.equal(paragraphText(m,ps[3].id),'');
 assert.ok(m.operations.every(o=>o.paragraphId===ps[0].id));assert.ok(m.operations.every(o=>!o.before.includes(education)&&!o.after.includes(education)));assertSourceComplete(m,m.fields[0]);
});

test('a photo caption containing only a name never gains a job title from the full biography',async()=>{
 const t=await fixture(shape('photo-caption',['张三'],[20,500,270,70])+body('body',350));
 const x=applyPptManagerProfile(task(t),source),m=x.materials[0],caption=t.paragraphs.find(p=>p.shapeId==='photo-caption');
 assert.equal(paragraphText(m,caption.id),'张三');assert.ok(!m.operations.some(o=>o.paragraphId===caption.id));assert.ok(!m.fields[0].paragraphIds.includes(caption.id));assertSourceComplete(m,m.fields[0]);
});

test('an unchanged existing primary-post caption stays untouched although the source lists additional posts',async()=>{
 const data=profile([{biography:'金融博士。\n现任甲公司固定收益投资部（私募）总经理、投资经理。\n擅长量化投资。'}]),t=await fixture(shape('photo-caption',['固定收益投资部（私募）总经理','张三'],[20,500,270,70])+body('body',350));
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0];assert.ok(!m.operations.some(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='photo-caption'));
 assert.ok(m.fields[0].after.includes('总经理、投资经理'));assertSourceComplete(m,m.fields[0]);
});

test('an obsolete existing primary-post caption changes only the stale words and retains its abbreviated scope',async()=>{
 const data=profile([{biography:'金融博士。\n现任甲公司固定收益投资部（私募）总经理、投资经理。\n擅长量化投资。'}]),t=await fixture(shape('photo-caption',['固定收益投资部（上海）总经理','张三'],[20,500,270,70])+body('body',350));
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0],caption=t.paragraphs.find(p=>p.text==='固定收益投资部（上海）总经理'),ops=m.operations.filter(o=>o.paragraphId===caption.id);
 assert.equal(paragraphText(m,caption.id),'固定收益投资部（私募）总经理');assert.deepEqual(ops.map(o=>({before:o.before,after:o.after})),[{before:'上海',after:'私募'}]);assert.ok(m.fields[0].after.includes('总经理、投资经理'));assertSourceComplete(m,m.fields[0]);
});

test('an existing secondary-post caption is changed to the first current post while the body keeps every post',async()=>{
 const data=profile([{biography:'金融博士。\n现任甲公司固定收益投资部（私募）总经理、投资经理。\n擅长量化投资。'}]),t=await fixture(shape('photo-caption',['固定收益投资部（私募）投资经理','张三'],[20,500,270,70])+body('body',350));
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0],caption=t.paragraphs.find(p=>p.text==='固定收益投资部（私募）投资经理');assert.equal(paragraphText(m,caption.id),'固定收益投资部（私募）总经理');assert.ok(m.fields[0].after.includes('总经理、投资经理'));assertSourceComplete(m,m.fields[0]);assert.ok(m.operations.filter(o=>o.paragraphId===caption.id).every(o=>!o.profileParagraphs&&!o.profileDeleteParagraph));
});

test('multiple existing caption posts collapse to the first complete source post across common separators',async()=>{
 const cases=[
  ['固定收益投资部总经理、投资经理','固定收益投资部总经理'],
  ['固定收益投资部总经理，投资经理','固定收益投资部总经理'],
  ['固定收益投资部总经理,投资经理','固定收益投资部总经理'],
  ['固定收益投资部总经理兼任投资经理','固定收益投资部总经理'],
  ['固定收益投资部总经理兼投资经理','固定收益投资部总经理'],
  ['固定收益部、权益投资部总经理、投资经理','固定收益部、权益投资部总经理'],
  ['固定收益部/权益部总经理、投资经理','固定收益部/权益部总经理'],
  ['固定收益部，权益投资部总经理、投资经理','固定收益部，权益投资部总经理'],
  ['固定收益部（公募、私募）总经理、投资经理','固定收益部（公募、私募）总经理'],
  ['固定收益部总经理助理、投资经理','固定收益部总经理助理'],
  ['固定收益部总经理（主持工作）、投资经理','固定收益部总经理（主持工作）'],
  ['固定收益部总经理，负责债券投资研究','固定收益部总经理'],
 ];
 for(const [employment,expected]of cases){
  const biography='金融博士。\n现任甲公司'+employment+'。\n擅长量化投资。',t=await fixture(shape('photo-caption',['旧固定收益投资部总经理、投资经理','张三'],[20,500,270,70])+body('body',350)),x=applyPptManagerProfile(task(t),profile([{biography}])),m=x.materials[0],caption=t.paragraphs.find(p=>p.text==='旧固定收益投资部总经理、投资经理');
  assert.equal(paragraphText(m,caption.id),expected,employment);assert.equal(x.facts[0].value,biography);assertSourceComplete(m,m.fields[0]);assert.ok(m.operations.every(o=>!o.profileParagraphs&&!o.profileDeleteParagraph));
 }
});

test('comma-separated original caption titles are still updated in place to one first post',async()=>{
 const data=profile([{biography:'金融博士。\n现任甲公司固定收益投资部总经理、投资经理。'}]),t=await fixture(shape('photo-caption',['固定收益投资部总经理，投资经理','张三'],[20,500,270,70])+body('body',350)),x=applyPptManagerProfile(task(t),data),m=x.materials[0],caption=t.paragraphs.find(p=>p.text==='固定收益投资部总经理，投资经理');
 assert.equal(paragraphText(m,caption.id),'固定收益投资部总经理');assertSourceComplete(m,m.fields[0]);
});

test('a generic job line beside the photo name counts as an existing job while a slide name header stays intact',async()=>{
 const data=profile([{biography:'金融博士。\n现任甲公司固定收益投资部总经理、投资经理。'}]),t=await fixture(shape('photo-caption',['投资经理','张三'],[20,500,270,70])+body('body',350),name('header',100,'投资经理：张三')+body('body',100)),x=applyPptManagerProfile(task(t),data),m=x.materials[0];
 assert.equal(paragraphText(m,t.paragraphs.find(p=>p.text==='投资经理').id),'固定收益投资部总经理');assert.equal(paragraphText(m,t.paragraphs.find(p=>p.text==='投资经理：张三').id),'投资经理：张三');
});

test('an inline name followed by education and current employment without a full stop remains a biography',async()=>{
 const t=await fixture(shape('inline',['张三：金融硕士，现任旧公司投资经理'])),x=applyPptManagerProfile(task(t),source),m=x.materials[0];
 assert.deepEqual(m.profilePages.manager,[1]);assert.equal(x.facts.length,1);assert.equal(x.facts[0].value,source.records[0].biography);assertSourceComplete(m,m.fields[0]);assert.ok(paragraphText(m,t.paragraphs[0].id).startsWith('张三：'));assert.ok(m.operations.every(o=>o.start>=3&&!o.profileParagraphs&&!o.profileDeleteParagraph));
});

test('two managers on one slide use shape geometry, despite interleaved XML order',async()=>{
 const t=await fixture(shape('heading',['王牌团队个人简介'],[20,10,900,45])+body('right',580)+name('left-name',80,'张三')+body('left',80)+name('right-name',580,'李四'));
 const x=applyPptManagerProfile(task(t),source,{productionDate:'2026-10-29'}),m=x.materials[0];
 assert.deepEqual(m.profilePages.manager,[1]);assert.deepEqual(x.facts.map(f=>f.scope.manager).sort(),['张三','李四']);
 for(const field of m.fields){assertSourceComplete(m,field);assert.equal(resultText(m,field.biographyParagraphIds),field.after);assert.equal(x.facts.find(f=>field.factIds.includes(f.id)).value,source.records.find(r=>r.name===field.managerName).biography);}
 assert.ok(m.operations.every(o=>o.profilePpt===true&&o.before===t.paragraphs.find(p=>p.id===o.paragraphId).text.slice(o.start,o.start+o.before.length)&&!o.profileParagraphs&&!o.profileDeleteParagraph));
 assert.ok(m.operations.every(o=>!['left-name','right-name','heading'].includes(t.paragraphs.find(p=>p.id===o.paragraphId).shapeId)));
});

test('each manager and repeated manager on multiple pages retains a complete source biography',async()=>{
 const t=await fixture(name('n',100,'投资经理：张三')+body('b',100),name('n',100,'李四')+body('b',100),name('n',100,'张三（续）')+body('b',100));
 const x=applyPptManagerProfile(task(t),source),m=x.materials[0];assert.deepEqual(m.profilePages.manager,[1,2,3]);assert.equal(x.facts.length,3);
 assert.deepEqual(x.facts.map(f=>f.scope.manager),['张三','李四','张三']);for(const f of m.fields)assert.equal(resultText(m,f.biographyParagraphIds),f.after);
});

test('an explicit adjacent unnamed continuation shares one full biography across both pages',async()=>{
 const t=await fixture(name('n',100,'张三')+shape('b',['金融硕士。','现任旧公司经理。'],[100,180,600,300]),shape('h',['投资经理简介（续）'],[30,10,900,45])+shape('b',['擅长历史量化投资研究，重视资产风险控制。','坚持长期稳定的投资理念。'],[100,180,600,300]));
 const x=applyPptManagerProfile(task(t),source),m=x.materials[0];assert.deepEqual(m.profilePages.manager,[1,2]);assert.equal(x.facts.length,1);assert.deepEqual(x.facts[0].scope.pages,[1,2]);assertSourceComplete(m,m.fields[0]);assert.equal(x.facts[0].value,source.records[0].biography);
 for(const page of [1,2])assert.ok(m.fields[0].biographyParagraphIds.some(id=>t.paragraphs.find(p=>p.id===id).page===page&&paragraphText(m,id)));
 assert.equal(x.facts[0].occurrences.length,2);
});

test('ambiguous and ordinary team follow-on pages never infer a continuation manager',async()=>{
 const continuation=shape('h',['投资经理简介（续）'],[30,10,900,45])+body('b',100);
 for(const previous of [name('a',80,'张三')+body('a-body',80)+name('b',580,'李四')+body('b-body',580),name('a',80,'张三')+body('a-body',80)+name('b',580,'王五')+body('b-body',580)]){
  const x=applyPptManagerProfile(task(await fixture(previous,continuation)),source);assert.deepEqual(x.materials[0].profilePages.manager,[1]);assert.ok(x.materials[0].operations.every(o=>o.part==='ppt/slides/slide1.xml'));
 }
 const x=applyPptManagerProfile(task(await fixture(name('a',100,'张三')+body('b',100),shape('h',['投资团队介绍'],[30,10,900,45])+body('b',100))),source);assert.deepEqual(x.materials[0].profilePages.manager,[1]);
});

test('full text is never shortened when a long source has more paragraphs than the original body',async()=>{
 const biography='金融博士。\n'+Array.from({length:12},(_,i)=>`第${i+1}项独立投资研究经历，覆盖不同投资市场与策略。`).join('\n');
 const x=applyPptManagerProfile(task(await fixture(name('n',100,'张三')+shape('b',['金融硕士。'],[100,180,600,300]))),profile([{biography}])),m=x.materials[0];
 assert.equal(resultText(m,m.fields[0].biographyParagraphIds),biography.replace(/\n/g,''));assert.equal(m.fields[0].after,biography.replace(/\n/g,''));assert.equal(x.facts[0].value,biography);assertSourceComplete(m,m.fields[0]);assert.ok(m.operations.every(o=>!o.profileParagraphs&&!o.profileDeleteParagraph));
});

test('date calculation uses full anniversaries and preserves Excel years without a start date',async()=>{
 const data=profile([{biography:'9年证券从业经验，5年投资管理经验。金融硕士。',workStart:'2016-10-30'}]);
 const x=task(await fixture(name('n',100,'张三')+body('b',100)));applyPptManagerProfile(x,data,{productionDate:'2026-10-29'});assert.match(x.facts[0].value,/9年证券从业经验，5年投资管理经验/);
 applyPptManagerProfile(x,data,{productionDate:'2026-10-30'});assert.match(x.facts[0].value,/10年证券从业经验，5年投资管理经验/);assert.equal(x.facts.length,1);assert.equal(x.facts[0].calculations[0].years,10);
});

test('title captions are synchronized from current employment while names and promotional headings remain intact',async()=>{
 const t=await fixture(shape('caption',['固定收益投资部总经理助理','张三'],[20,500,270,70])+shape('slogan',['专注信用增强策略，投资风格稳健'],[350,120,550,40])+body('body',350));
 const x=applyPptManagerProfile(task(t),source),m=x.materials[0];
 const caption=t.paragraphs.find(p=>p.text==='固定收益投资部总经理助理');assert.equal(paragraphText(m,caption.id),'固定收益投资部总经理');assert.deepEqual(m.operations.filter(o=>o.paragraphId===caption.id).map(o=>({before:o.before,after:o.after})),[{before:'助理',after:''}]);
 assert.ok(!m.operations.some(o=>o.before==='张三'||o.before.includes('专注信用增强策略')));assertSourceComplete(m,m.fields[0]);
});

test('a combined job title/name caption preserves the name and changes only evidenced title',async()=>{
 const x=applyPptManagerProfile(task(await fixture(name('n',100,'固定收益投资部总经理助理 张三')+body('b',100))),source);
 const m=x.materials[0],p=m.template.paragraphs.find(p=>p.text.includes(' 张三'));assert.equal(paragraphText(m,p.id),'固定收益投资部总经理 张三');assert.ok(m.operations.filter(o=>o.paragraphId===p.id).every(o=>!o.before.includes('张三')&&!o.after.includes('张三')));
});

test('identical biographies still record their relevant page with no replacement operations',async()=>{
 const t=await fixture(name('n',100,'张三')+shape('b',source.records[0].biography.split('\n'),[100,180,600,300]));
 const x=applyPptManagerProfile(task(t),source);assert.deepEqual(x.materials[0].profilePages.manager,[1]);assert.equal(x.facts.length,1);assert.equal(x.materials[0].operations.length,0);
});

test('team organization, awards, product rows, and commentary names are not biography regions',async()=>{
 const t=await fixture(shape('t',['投资团队组织架构','张三','李四','团队共18人，覆盖股票债券研究。']),shape('a',['荣誉奖项','张三','年度优秀投资经理奖。']),shape('p',['产品要素','投资经理：张三','风险等级：R3']),shape('c',['市场观点','张三表示，当前市场估值合理。','建议长期投资。']));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,0);assert.deepEqual(x.materials[0].profilePages.manager,[]);assert.equal(x.materials[0].operations.length,0);
});

test('unknown manager cards are isolated and never overwritten with the known manager biography',async()=>{
 const t=await fixture(name('known',80,'张三')+name('unknown',580,'王五')+body('unknown-body',580,'海外金融学博士。')+body('known-body',80));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='known-body'));
});

test('an unknown manager in a combined job caption remains an independent card boundary',async()=>{
 const t=await fixture(name('known',80,'张三')+name('unknown',580,'固定收益投资部总经理 王五')+body('unknown-body',580,'海外金融学博士。')+body('known-body',80));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='known-body'));
});

test('unknown three-character names with honorifics partition separate biography cards',async()=>{
 for(const caption of ['王小五先生','王小五女士']){
  const t=await fixture(name('known',80,'张三')+name('unknown',580,caption)+body('unknown-body',580,'海外金融学博士。')+body('known-body',80));
  const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='known-body'));
 }
});

test('organization-chart education and headcount metrics are not a biography for the only named manager',async()=>{
 const t=await fixture(shape('heading',['投资团队组织架构'],[20,10,900,45])+name('name',100,'张三')+shape('role',['投资总监'],[100,140,300,30])+shape('metrics',['团队共18人，全部具有硕士及以上学历，覆盖股票与债券研究。'],[100,200,800,300]));
 const x=applyPptManagerProfile(task(t),source);assert.deepEqual(x.materials[0].profilePages.manager,[]);assert.equal(x.facts.length,0);assert.equal(x.materials[0].operations.length,0);
});

test('mixed-page company and award sections retain their separate heading and body shapes',async()=>{
 const t=await fixture(name('name',100,'张三')+body('manager-body',100)+shape('company-heading',['公司介绍'],[600,90,300,45])+shape('company-body',['公司共有30名投研人员，均具有硕士及以上学历。'],[600,180,360,300])+shape('award-heading',['荣誉奖项'],[1100,90,300,45])+shape('award-body',['金融博士论坛年度优秀投资荣誉。'],[1100,180,360,300]));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='manager-body'));assertSourceComplete(x.materials[0],x.materials[0].fields[0]);
});

test('collective education statistics without a section heading never become a personal biography',async()=>{
 const t=await fixture(name('name',100,'张三')+body('manager-body',100)+shape('metrics',['公司管理团队均具备硕士学历。','18名投研人员中博士占比50%。'],[600,180,360,300]));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='manager-body'));
});

test('a separate company or awards box on a mixed biography slide is not absorbed as manager text',async()=>{
 const t=await fixture(name('known',80,'张三')+body('known-body',80)+shape('company',['公司介绍','公司管理团队均具备硕士学历。'],[550,180,400,200])+shape('awards',['荣誉奖项','张三荣获金融投资博士论坛年度大奖。'],[550,450,400,180]));
 const x=applyPptManagerProfile(task(t),source);assert.equal(x.facts.length,1);assert.ok(x.materials[0].operations.every(o=>t.paragraphs.find(p=>p.id===o.paragraphId).shapeId==='known-body'));
});

test('duplicate source names choose the closest biography and grouped coordinates preserve card ownership',async()=>{
 const group=(id,x,content)=>`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="500" cy="600"/><a:chOff x="0" y="0"/><a:chExt cx="500" cy="600"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;
 const t=await fixture(group('left',0,name('left-name',50,'张三')+body('left-body',50,'复旦大学金融学硕士。'))+group('right',550,name('right-name',50,'李四')+body('right-body',50,'北京大学医学博士。'))),data=profile([{name:'张三',biography:'北京大学医学博士。\n现任甲公司投资经理。'},{name:'张三',biography:'复旦大学金融学硕士。\n现任乙公司投资经理。'},{name:'李四',biography:'北京大学医学博士。\n现任丙公司投资经理。'}]);
 const x=applyPptManagerProfile(task(t),data),m=x.materials[0];assert.equal(x.facts.find(f=>f.scope.manager==='张三').scope.row,3);assert.equal(m.fields.length,2);
 for(const f of m.fields)assert.ok(f.biographyParagraphIds.every(id=>t.paragraphs.find(p=>p.id===id).shapeId===(f.managerName==='张三'?'left-body':'right-body')));
});

test('PPT reapplication only clears its own facts and operations, preserving Word and other source objects',async()=>{
 const t=await fixture(name('n',100,'张三')+body('b',100)),x=task(t),wordFact={id:'word-fact',source:'manager',status:'verified',scope:{manager:'李四'},occurrences:[{materialId:'word'}]},wordOp={source:'manager',id:'word-op'},wordField={source:'manager',id:'word-field'},otherFact={id:'user-fact',source:'user',status:'verified'},otherOp={id:'company-op',source:'company'},otherField={id:'company-field',source:'company'};
 const word={kind:'word',materialId:'word',operations:[wordOp],fields:[wordField]};x.materials.push(word);x.facts.push(wordFact,otherFact);x.materials[0].operations.push(otherOp);x.materials[0].fields.push(otherField);
 applyPptManagerProfile(x,source);const first=JSON.stringify(x.materials[0].operations);applyPptManagerProfile(x,source);
 assert.equal(JSON.stringify(x.materials[0].operations),first);assert.ok(x.facts.includes(wordFact)&&x.facts.includes(otherFact));assert.deepEqual(word.operations,[wordOp]);assert.deepEqual(word.fields,[wordField]);assert.ok(x.materials[0].operations.includes(otherOp));assert.ok(x.materials[0].fields.includes(otherField));
 applyPptManagerProfile(x,profile([{name:'无匹配',biography:'金融博士。'}]));assert.deepEqual(x.materials[0].profilePages.manager,[]);assert.deepEqual(x.materials[0].operations,[otherOp]);assert.equal(x.facts.length,2);
});

test('two managers in one text box and inline name/biography preserve each name',async()=>{
 const x=applyPptManagerProfile(task(await fixture(shape('b',['张三：金融硕士。','现任旧公司经理。','李四','经济学博士。','现任旧机构经理。']))),source),m=x.materials[0];
 assert.equal(x.facts.length,2);const inline=m.template.paragraphs.find(p=>p.text.startsWith('张三：'));assert.ok(paragraphText(m,inline.id).startsWith('张三：金融博士。'));assert.ok(m.operations.filter(o=>o.paragraphId===inline.id).every(o=>o.start>=3));assert.ok(!m.operations.some(o=>o.before==='李四'));for(const field of m.fields)assertSourceComplete(m,field);
 const zs=m.fields.find(f=>f.managerName==='张三'),ls=m.fields.find(f=>f.managerName==='李四');assert.ok(!zs.biographyParagraphIds.some(id=>ls.biographyParagraphIds.includes(id)));
});

test('obsolete biography text is cleared in place while native paragraphs and unrelated notes remain',async()=>{
 const t=await fixture(name('n',100,'张三')+shape('b',['金融硕士。','现任旧公司经理。','擅长研究。','注：历史业绩不代表未来表现。'],[100,180,600,300]));
 const x=applyPptManagerProfile(task(t),profile([{biography:'金融博士。'}])),m=x.materials[0],ops=m.operations;assert.equal(ops.filter(o=>o.after===''&&!o.profileDeleteParagraph).length,2);assert.ok(!ops.some(o=>o.before.startsWith('注：')));assert.ok(ops.every(o=>!o.profileDeleteParagraph&&!o.profileParagraphs));
});

test('an existing trailing decorative bullet is preserved exactly with the original template',async()=>{
 const x=applyPptManagerProfile(task(await fixture(name('n',100,'张三')+shape('b',['金融硕士。',''],[100,180,600,300]))),profile([{biography:'金融博士。'}])),m=x.materials[0],bullet=m.template.paragraphs.find(p=>p.text==='');assert.equal(paragraphText(m,bullet.id),'');assert.ok(!m.operations.some(o=>o.paragraphId===bullet.id));
});
