import {readTemplate} from './office-reader.js';
import {NS,elements,ancestor} from './office-xml.js';

// These rules identify layouts and clause types; company names and numbers always
// come from the uploaded authority document, never from a built-in answer table.
const compact=s=>String(s||'').normalize('NFKC').replace(/\s/g,'');
const canonical=s=>compact(s).replace(/^(?:股东|投资顾问|管理人|受托人)[：:]?/,'').replace(/(?:股份有限公司|有限责任公司|有限公司)$/,'').replace(/证券资产管理$/,'资产管理').replace(/资管$/,'资产管理');
const sameSubject=(a,b)=>{const x=canonical(a),y=canonical(b);return !!x&&!!y&&x===y;};
const unique=a=>[...new Set(a)];
const noteRE=/数据来源|金牛奖颁奖机构|统计时间|近两年管理人类获奖情况/;
const kpiKey=s=>({资管规模:'aum',金牛奖:'awards',荣获金牛奖:'awards',资管经验:'experience'}[compact(s)]||null);
const normalName=s=>compact(s).replace(/^(?:股东|投资顾问|管理人|受托人)[：:]?/,'');
const explicitName=s=>{const n=normalName(s);return n.length>=4&&n.length<=60&&!/[，。；、：]/.test(n)&&/(?:资产管理|资管|证券|信托|银行|有限公司)$/.test(n)?n:null;};
const sourceNames=s=>[...String(s).matchAll(/数据来源[：:]\s*([^。；，,\n]+)/g)].flatMap(m=>m[1].split(/[、和]/)).map(s=>s.trim()).filter(s=>explicitName(s));

function atom(p,key,start,text){return{key,text,start,paragraphId:p.id,part:p.part,index:p.index};}
function noteAtoms(p){
 const out=[];
 const rules=[['aum_date',/(?:资管规模数据时间|规模数据|数据时间)?截至\d{4}[.年/-]\d{1,2}[.月/-]\d{1,2}日?/g],['award_issuer',/金牛奖颁奖机构为[^，,。；;\n]+/g],['award_period',/统计时间为[^，,。；;\n]+/g],['award_recent',/近两年管理人类获奖情况为[^。；;\n]+/g]];
 for(const [key,re]of rules)for(const m of p.text.matchAll(re))out.push(atom(p,key,m.index,m[0]));
 const dates=out.filter(a=>a.key==='aum_date');
 const labeled=dates.filter(a=>/^(?:资管规模数据时间|规模数据)/.test(a.text));
 return labeled.length?out.filter(a=>a.key!=='aum_date'||labeled.includes(a)):out;
}
function companyAtoms(p,role='company'){
 if(noteRE.test(p.text)||/更名为|变更为|发行与管理|兑付责任/.test(p.text))return[];
 if(role==='shareholder'){
  const text=p.text.trim(),start=p.text.indexOf(text);
  const key=/业务隔离|股东不直接/.test(text)?'shareholder_isolation':/分支机构|服务范围覆盖|资本实力位居/.test(text)?'shareholder_network':/(?:有限公司|金融服务商)/.test(text)?'shareholder_description':null;
  return key&&text?[atom(p,key,start,text)]:[];
 }
 const out=[],aum=p.text.match(/截至[^。；;\n]*?(?:资产管理|资管)规模[^。；;\n]*/);
 if(aum)out.push(atom(p,'aum',aum.index,aum[0]));
 for(const m of p.text.matchAll(/[^，,。；;\n]+/g)){
  if(out.some(a=>m.index>=a.start&&m.index<a.start+a.text.length))continue;
  const text=m[0].replace(/^[\s•·●◆]+/,'').trim(),start=m.index+m[0].indexOf(text);
  const key=/注册资本/.test(text)?'capital':/首批.*资产管理子公司/.test(text)?'qualification':/全能型资产管理机构/.test(text)?'institution':/\d+多年?资产管理经验|\d+年资产管理经验/.test(text)?'experience':/投研团队/.test(text)?'team':/产品体系/.test(text)?'product_system':/荣获.*金牛奖/.test(text)?'awards':null;
  if(key&&text)out.push(atom(p,key,start,text));
 }
 return out.sort((a,b)=>a.start-b.start);
}
function tables(t){
 const map=new Map();for(const p of t.paragraphs)if(p.tableId&&p.part==='word/document.xml'){if(!map.has(p.tableId))map.set(p.tableId,[]);map.get(p.tableId).push(p);}
 return[...map.entries()].map(([id,paragraphs])=>({id,paragraphs}));
}
function followingNotes(t,ps){
 const last=Math.max(...ps.map(p=>p.index)),result=[];
 for(const p of t.paragraphs.filter(p=>p.part==='word/document.xml'&&p.index>last).slice(0,5)){
  if(p.tableId)break;if(!p.text.trim())continue;if(!noteRE.test(p.text))break;result.push(p);
 }
 return result;
}
function noteIdentity(p){for(const kind of ['footnote','endnote']){const owner=ancestor(p.node,NS.w,kind);if(owner)return`${kind}:${owner.getAttributeNS(NS.w,'id')}`;}return null;}
function references(p){return['footnote','endnote'].flatMap(kind=>elements(p.node,NS.w,`${kind}Reference`).map(n=>`${kind}:${n.getAttributeNS(NS.w,'id')}`));}
function referencedNotes(t,ps){const ids=new Set(ps.flatMap(references));return t.paragraphs.filter(p=>ids.has(noteIdentity(p))&&noteRE.test(p.text));}
function getKpi(t,tbl){
 const labels=tbl.paragraphs.filter(p=>kpiKey(p.text)),keys=unique(labels.map(p=>kpiKey(p.text)));
 if(keys.length!==3)return null;
 const entries=[];
 for(const label of labels){const values=tbl.paragraphs.filter(p=>p.cellIndex===label.cellIndex&&p.rowIndex<label.rowIndex&&p.text.trim());if(values.length!==1)return null;const p=values[0];entries.push(atom(p,kpiKey(label.text),p.text.indexOf(p.text.trim()),p.text.trim()));}
 if(entries.length!==3)return null;
 for(const label of labels)entries.push(atom(label,`${kpiKey(label.text)}_label`,label.text.indexOf(label.text.trim()),label.text.trim()));
 const notes=[...followingNotes(t,tbl.paragraphs),...referencedNotes(t,tbl.paragraphs)],names=unique(notes.flatMap(p=>sourceNames(p.text)));
 const previous=t.paragraphs.filter(p=>p.part==='word/document.xml'&&p.index<tbl.paragraphs[0].index&&!p.tableId&&p.text.trim()).at(-1);
 return{layout:'kpi',tableId:tbl.id,heading:previous?.text||'',subject:names.length===1?names[0]:'',aliases:names,columns:[{role:'company',subject:names.length===1?names[0]:'',entries}],notes,paragraphs:tbl.paragraphs};
}
function getNarrative(t,tbl){
 const ps=tbl.paragraphs,notes=[...ps.filter(p=>noteRE.test(p.text)),...referencedNotes(t,ps)],allCore=ps.flatMap(p=>companyAtoms(p));
 if(unique(allCore.map(a=>a.key)).length<3)return null;
 const firstCore=Math.min(...allCore.map(a=>ps.find(p=>p.id===a.paragraphId).rowIndex));
 const headers=ps.filter(p=>p.rowIndex<firstCore),names=unique(notes.flatMap(p=>sourceNames(p.text))),columns=[];
 for(const cell of unique(ps.map(p=>p.cellIndex))){
  const head=headers.filter(p=>p.cellIndex===cell),headerText=head.map(p=>p.text).join(''),role=/股东/.test(headerText)?'shareholder':'company';
  const explicit=head.map(p=>explicitName(p.text)).filter(Boolean);
  let subject=explicit.length===1?explicit[0]:'';
  const body=ps.filter(p=>p.cellIndex===cell&&p.rowIndex>=firstCore&&!notes.includes(p));
  if(!subject&&role==='shareholder'){const first=body.find(p=>/^\s*[^，。；]{3,55}(?:股份有限公司|有限责任公司|有限公司)/.test(p.text));subject=first?.text.match(/^\s*([^，。；]{3,55}?(?:股份有限公司|有限责任公司|有限公司))/)?.[1]||'';}
  if(!subject&&role==='company'&&/管理人|投资顾问|资管/.test(headerText)){const matching=names.filter(n=>/资产管理|资管/.test(n));if(matching.length===1)subject=matching[0];}
  const entries=body.flatMap(p=>companyAtoms(p,role));
  if(entries.length&&subject)columns.push({role,subject,cellIndex:cell,entries,paragraphIds:body.map(p=>p.id)});
 }
 if(!columns.length)return{layout:'narrative',tableId:tbl.id,heading:'',subject:'',aliases:names,columns:[],notes,paragraphs:ps};
 const company=columns.find(c=>c.role==='company'&&c.entries.some(a=>a.key==='qualification'))||columns.find(c=>c.role==='company');
 return{layout:'narrative',tableId:tbl.id,heading:headers.find(p=>/公司实力|实力支持/.test(p.text))?.text||'',subject:company?.subject||'',aliases:names,columns,notes,paragraphs:ps};
}
function sections(t){return tables(t).map(tbl=>getKpi(t,tbl)||getNarrative(t,tbl)).filter(Boolean);}
function companyNote(t,section,p){
 if(section.layout==='kpi')return true;
 const noteId=noteIdentity(p);
 if(noteId){const referring=section.paragraphs.filter(q=>references(q).includes(noteId));return referring.length>0&&referring.every(q=>section.columns.some(c=>c.role==='company'&&sameSubject(c.subject,section.subject)&&c.entries.some(e=>e.paragraphId===q.id)));}
 const row=t.rows.find(r=>r.tableId===p.tableId&&r.rowIndex===p.rowIndex);
 if((row?.cells.length||0)<=1)return true;
 const column=section.columns.find(c=>c.cellIndex===p.cellIndex);
 return !!column&&column.role==='company'&&sameSubject(column.subject,section.subject);
}
function plainSection(s,i){
 return{id:`section-${i+1}`,layout:s.layout,heading:s.heading,subject:s.subject,aliases:s.aliases,columns:s.columns.map(c=>({role:c.role,subject:c.subject,entries:c.entries})),notes:s.notes.map(p=>({paragraphId:p.id,part:p.part,index:p.index,text:p.text,entries:noteAtoms(p)})),sourceTexts:s.paragraphs.concat(s.notes.filter(p=>!s.paragraphs.includes(p))).map(p=>({paragraphId:p.id,text:p.text}))};
}

function selectedVersion(t,variant){
 if(variant===undefined)return t;
 const labels={general:'通用版本',ccb:'建行版本'};
 if(!['general','ccb'].includes(variant))throw new Error('公司介绍版本无效，请选择通用版本或建行版本。');
 const versionNames=new Map([['通用版本','general'],['通用版','general'],['建行版本','ccb'],['建行版','ccb']]),versionOf=text=>versionNames.get(compact(text));
 // A version marker is a standalone body heading, not a label inside a
 // company table. Keep original offsets so all evidence still points to source.
 const markers=t.paragraphs.filter(p=>p.part==='word/document.xml'&&!p.tableId&&versionOf(p.text));
 if(!markers.length)return t;
 const ranges=markers.flatMap((p,i)=>versionOf(p.text)===variant?[{start:p.index,end:markers[i+1]?.index??Infinity}]:[]);
 if(!ranges.length)throw new Error('公司介绍文件中未找到'+labels[variant]+'，请上传包含该版本的文件。');
 const body=t.paragraphs.filter(p=>p.part==='word/document.xml'&&ranges.some(r=>p.index>=r.start&&p.index<r.end)),ids=new Set(body.flatMap(references)),bodyIds=new Set(body.map(p=>p.id));
 const paragraphs=t.paragraphs.filter(p=>bodyIds.has(p.id)||(p.part!=='word/document.xml'&&ids.has(noteIdentity(p))));
 return{...t,paragraphs};
}

export async function parseCompanyProfile(file,{variant}={}){
 if(!/\.docx$/i.test(file?.name||''))throw new Error('公司介绍请使用 DOCX 文件。');
 const t=selectedVersion(await readTemplate(file),variant),found=sections(t).filter(s=>s.columns.length&&s.subject);
 if(!found.length)throw new Error('未识别到公司介绍双栏或三项指标版式，请检查公司介绍文件。');
 const date=String(file.name).match(/(20\d{2})[.年_-]?(0?[1-9]|1[0-2])(?=[^\d]|$)/);
 return{schemaVersion:1,name:t.name,version:date?`${date[1]}.${date[2].padStart(2,'0')}`:'未标注月份',sha256:t.sha256,sections:found.map((s,i)=>plainSection({...s,notes:s.notes.filter(p=>companyNote(t,s,p))},i)),warnings:[...t.warnings],sourceTexts:t.paragraphs.filter(p=>p.text.trim()).map(p=>({paragraphId:p.id,part:p.part,index:p.index,text:p.text}))};
}

// LCS produces disjoint edits against original paragraph offsets. A pure insertion
// borrows one adjacent unchanged character because the Word writer requires an
// existing, nonempty anchor; overlapping anchors are then merged.
function differences(before,after){
 if(before===after)return[];
 const n=before.length,m=after.length;
 if(n*m>250000)return singleDifference(before,after);
 const grid=Array.from({length:n+1},()=>new Uint16Array(m+1));
 for(let i=n-1;i>=0;i--)for(let j=m-1;j>=0;j--)grid[i][j]=before[i]===after[j]?grid[i+1][j+1]+1:Math.max(grid[i+1][j],grid[i][j+1]);
 const raw=[];let i=0,j=0,pending=null;
 const flush=()=>{if(pending){raw.push(pending);pending=null;}};
 while(i<n||j<m){if(i<n&&j<m&&before[i]===after[j]){flush();i++;j++;continue;}if(!pending)pending={start:i,end:i,after:''};if(j<m&&(i===n||grid[i][j+1]>=grid[i+1][j]))pending.after+=after[j++];else{pending.end=++i;}}
 flush();
 for(const d of raw)if(d.start===d.end){if(d.start>0){d.start--;d.after=before[d.start]+d.after;}else if(n){d.end++;d.after+=before[0];}}
 const merged=[];for(const d of raw){const last=merged.at(-1);if(last&&d.start<last.end){const overlap=last.end-d.start;last.after+=d.after.slice(overlap);last.end=Math.max(last.end,d.end);}else merged.push({...d});}
 return merged.filter(d=>d.end>d.start).map(d=>({start:d.start,before:before.slice(d.start,d.end),after:d.after}));
}
function singleDifference(a,b){let start=0;while(start<a.length&&start<b.length&&a[start]===b[start])start++;let end=a.length,tail=b.length;while(end>start&&tail>start&&a[end-1]===b[tail-1]){end--;tail--;}if(start===end){if(start){start--;}else if(end<a.length){end++;tail++;}}return end>start?[{start,before:a.slice(start,end),after:b.slice(start,tail)}]:[];}
const labelFor=key=>({capital:'注册资本',qualification:'公司资质',institution:'公司定位',experience:'资管经验',team:'投研团队',product_system:'产品体系',awards:'公司奖项',aum:'资管规模',aum_label:'规模指标名称',awards_label:'奖项指标名称',experience_label:'经验指标名称',aum_date:'规模数据日期',award_issuer:'奖项来源',award_period:'奖项统计时间',award_recent:'近期获奖情况',shareholder_description:'股东介绍',shareholder_network:'股东实力',shareholder_isolation:'股东隔离声明'}[key]||key);
function warning(task,m,text){const value=`公司介绍：${text}`;if(!m.warnings.includes(value))m.warnings.push(value);if(!task.warnings.includes(value))task.warnings.push(value);}
function overlap(a,b){return a.paragraphId===b.paragraphId&&a.start<b.start+b.before.length&&b.start<a.start+a.before.length;}
function addMatch(task,m,profile,section,column,target,source,after=source.text){
 const p=m.template.paragraphs.find(p=>p.id===target.paragraphId);if(!p||!target.text)return;
 const edits=differences(target.text,after).map(d=>({...d,start:target.start+d.start,paragraphId:p.id}));
 if(edits.some(e=>m.operations.some(o=>overlap(e,o)))){warning(task,m,`“${labelFor(target.key)}”与已有核对位置重叠，保留待核。`);return;}
 const index=task.facts.filter(f=>f.source==='company').length+1,id=`company-fact-${index}`;
 const evidence=[{blockId:`company-profile-${profile.sha256.slice(0,16)}-${section.id}-${source.paragraphId}`,quote:source.text}];
 const reason=`依据公司介绍 ${profile.version} 的${section.layout==='kpi'?'指标':'双栏'}版式，按公司主体与条目更新。`;
 const fieldId=`${m.materialId}-company-field-${index}`;
 task.facts.push({id,key:`company:${target.key}`,label:labelFor(target.key),value:after,status:'verified',source:'company',evidence,scope:{company:column.subject,role:column.role,layout:section.layout},occurrences:[{materialId:m.materialId,paragraphId:p.id,fieldId}],reason,reasonType:'company_profile'});
 m.fields.push({id:fieldId,label:`公司介绍 · ${labelFor(target.key)}`,before:p.text,after:'',status:'verified',source:'company',factIds:[id],paragraphIds:[p.id],reason});
 for(const edit of edits)m.operations.push({id:`company-op-${m.materialId}-${index}-${m.operations.length}`,kind:'replace',source:'company',paragraphId:p.id,part:p.part,index:p.index,start:edit.start,before:edit.before,after:edit.after,factIds:[id],reason,evidence});
}
function expandedSource(source,sourceColumn,targetColumn){
 const siblings=sourceColumn.entries.filter(a=>a.paragraphId===source.paragraphId);
 if(siblings.length<2)return source.text;
 // Attach a missing neighbouring clause to the matching text in the same source
 // paragraph. This keeps the target's paragraph/run layout intact.
 const same=targetColumn.entries.filter(a=>siblings.some(s=>s.key===a.key));
 if(same.length!==1)return source.text;
 if(siblings.some(a=>targetColumn.entries.some(t=>t.key===a.key&&t!==same[0])))return source.text;
 return siblings.map(a=>a.text).join('，');
}
function finish(task){
 for(const m of task.materials)for(const field of m.fields.filter(f=>f.source==='company'))field.after=field.paragraphIds.map(id=>{let text=m.template.paragraphs.find(p=>p.id===id)?.text||'';for(const o of m.operations.filter(o=>o.kind==='replace'&&o.paragraphId===id).sort((a,b)=>b.start-a.start))text=text.slice(0,o.start)+o.after+text.slice(o.start+o.before.length);return text;}).join('\n');
 task.stats={...task.stats,facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};return task;
}

export function applyCompanyProfile(task,profile){
 if(!profile||profile.schemaVersion!==1||!Array.isArray(profile.sections)||!/^[0-9a-f]{64}$/.test(profile.sha256||''))throw new Error('公司介绍资料无效，请重新导入。');
 const isCompanyWarning=w=>(typeof w==='string'?w:w?.message||'').startsWith('公司介绍：');
 task.facts=task.facts.filter(f=>f.source!=='company');task.warnings=(task.warnings||[]).filter(w=>!isCompanyWarning(w));
 for(const m of task.materials){
  m.operations=m.operations.filter(o=>o.source!=='company');m.fields=m.fields.filter(f=>f.source!=='company');m.warnings=(m.warnings||[]).filter(w=>!isCompanyWarning(w));
  if(m.kind!=='word'||m.template.kind!=='docx')continue;
  const found=sections(m.template);if(!found.length){warning(task,m,'未找到可确认主体与版式的公司介绍区域，已保留原文。');continue;}
  let matched=false;
  for(const target of found){
   const candidates=profile.sections.filter(s=>s.layout===target.layout&&s.columns.some(c=>target.columns.some(t=>t.role===c.role&&sameSubject(t.subject,c.subject))));
   if(candidates.length!==1){warning(task,m,`${target.layout==='kpi'?'指标':'双栏'}区域无法唯一匹配公司主体，已保留原文。`);continue;}
   const source=candidates[0];let companyMatched=false;
   for(const col of target.columns){
    const sources=source.columns.filter(c=>c.role===col.role&&sameSubject(c.subject,col.subject));if(sources.length!==1)continue;
    const sc=sources[0];matched=true;if(col.role==='company')companyMatched=true;
    for(const entry of col.entries){const candidates=sc.entries.filter(e=>e.key===entry.key);if(candidates.length!==1){warning(task,m,`“${labelFor(entry.key)}”没有唯一对应的新版条目，已保留原文。`);continue;}const s=candidates[0];addMatch(task,m,profile,source,sc,entry,s,expandedSource(s,sc,col));}
    const absent=sc.entries.filter(s=>!col.entries.some(e=>e.key===s.key)&&!sc.entries.some(sibling=>sibling.paragraphId===s.paragraphId&&col.entries.some(e=>e.key===sibling.key)));
    if(absent.length)warning(task,m,`未找到新版${unique(absent.map(a=>labelFor(a.key))).join('、')}的可对应位置，请人工补齐；原附加内容保留。`);
    const known=new Set(col.entries.map(e=>e.paragraphId));const unknown=target.paragraphs.filter(p=>col.paragraphIds?.includes(p.id)&&p.text.trim()&&!known.has(p.id)&&!noteRE.test(p.text)&&!/更名为|变更为|发行与管理|兑付责任/.test(p.text));
    if(unknown.length)warning(task,m,'公司栏存在新版中未匹配的附加介绍，已保留原文，请核对。');
   }
   if(companyMatched){
    const sourceNotes=source.notes.flatMap(n=>n.entries),company=source.columns.find(c=>c.role==='company');
    for(const p of target.notes){
     if(!companyNote(m.template,target,p))continue;
     const names=sourceNames(p.text);if(names.length&&!names.some(n=>sameSubject(n,company.subject)||source.aliases.some(a=>sameSubject(n,a))))continue;
     if([...p.text.matchAll(/截至\d{4}[.年/-]\d{1,2}[.月/-]\d{1,2}日?/g)].length>1)warning(task,m,'来源说明含多个日期，仅核对可明确对应的规模日期；其他日期保留待核。');
     for(const entry of noteAtoms(p)){const options=sourceNotes.filter(e=>e.key===entry.key);if(options.length===1)addMatch(task,m,profile,source,company,entry,options[0]);else if(options.length>1)warning(task,m,`“${labelFor(entry.key)}”有多个来源，已保留原文。`);}
    }
   }
  }
  if(!matched)warning(task,m,'没有找到与已保存公司介绍一致的主体，已保留全部公司原文。');
 }
 task.companyProfile={name:profile.name,version:profile.version,sha256:profile.sha256};return finish(task);
}
