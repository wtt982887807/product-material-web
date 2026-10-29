import {NS,elements,ancestor,children} from './office-xml.js';
import {managerTextDifferences} from './manager-profile.js';

// Only the uploaded profile supplies factual text. These patterns identify the
// introduction's role and layout; a heading alone is never evidence of a match.
const compact=s=>String(s||'').normalize('NFKC').replace(/\s/g,'');
const canonical=s=>compact(s).replace(/^(?:管理人|投资顾问|受托人|股东)(?:介绍)?[：:]?/,'').replace(/(?:股份有限公司|有限责任公司|有限公司)$/,'').replace(/证券资产管理$/,'资产管理').replace(/资管$/,'资产管理');
const unique=a=>[...new Set(a)];
const noteRE=/数据来源|金牛奖颁奖机构|统计时间|近两年管理人类获奖情况/;
const isolationRE=/业务隔离|股东不直接/;
const companyRules=[['capital',/注册资本/],['qualification',/首批.*资产管理/],['institution',/全能型资产管理机构/],['experience',/\d+多年?资产管理经验|\d+年资产管理经验/],['aum',/(?:资产管理|资管)规模/],['awards',/荣获.*金牛奖/],['team',/专业的?投研团队/],['product_system',/产品体系/]];
const companyKeys=text=>companyRules.flatMap(([key,re])=>re.test(text)?[key]:[]);
const semanticKeys=(text,role)=>role==='company'?companyKeys(text):[['shareholder_description',/(?:金融服务商|证券股份有限公司.*(?:提供|是|由))/],['shareholder_network',/分支机构|服务范围覆盖|资本实力位居/],['shareholder_isolation',isolationRE]].flatMap(([key,re])=>re.test(text)?[key]:[]);
const namesInNote=text=>[...text.matchAll(/数据来源[：:]([^。；;\n]+)/g)].flatMap(m=>m[1].split(/[、，,和]/)).map(s=>s.trim()).filter(Boolean);
const sourceBody=(section,column)=>unique(column.entries.map(e=>e.paragraphId)).map(id=>({paragraphId:id,text:section.sourceTexts?.find(p=>p.paragraphId===id)?.text||unique(column.entries.filter(e=>e.paragraphId===id).map(e=>e.text)).join('，'),keys:column.entries.filter(e=>e.paragraphId===id).map(e=>e.key)})).filter(p=>p.text.trim());
const kpiKey=text=>({资管规模:'aum',金牛奖:'awards',荣获金牛奖:'awards',资管经验:'experience'}[compact(text)]||null);
const kpiValue=text=>/^(?:荣获)?\d+(?:\.\d+)?(?:亿元|万亿元|座|年)$/.test(compact(text));

function sourceRecords(profile){
 const records=profile.sections.filter(s=>s.layout==='narrative').flatMap(section=>section.columns.filter(c=>['company','shareholder'].includes(c.role)).map(column=>{
  const subject=canonical(column.subject),aliases=new Set([subject]),renames=new Map();
  if(column.role==='company')for(const name of section.aliases||[])if(canonical(name)===subject)aliases.add(canonical(name));
  const texts=[...(profile.sourceTexts||[]),...(section.sourceTexts||[])].map(p=>p.text||'');
  for(const text of texts){
   const rename=text.match(/^\s*([^，,。；;]*?(?:股份有限公司|有限责任公司|有限公司)).*?更名为[“「"]([^”」"]+(?:股份有限公司|有限责任公司|有限公司))[”」"]/);
   if(rename&&canonical(rename[2]).endsWith(subject)){
    const old=canonical(rename[1]),current=canonical(rename[2]);aliases.add(old);renames.set(old,{text:rename[2],evidence:text});
    const prefix=current.slice(0,current.length-subject.length);
    if(prefix&&old.startsWith(prefix)){const short=old.slice(prefix.length);aliases.add(short);renames.set(short,{text:column.subject,evidence:text});}
   }
   // Full legal names are aliases only when the authority itself supplies them.
   // A longer brand found solely in a template must never inherit this match.
   for(const part of text.split(/[，,。；;“”「」\n]/)){
    const name=compact(part).match(/^([^：:（）()]*?(?:股份有限公司|有限责任公司|有限公司))/)?.[1];
    if(name&&canonical(name).endsWith(subject))aliases.add(canonical(name));
   }
  }
  return{section,column,aliases,renames,body:sourceBody(section,column)};
 }));
 // Repeated identical authorities are the same evidence; genuinely different
 // versions for one subject remain separate and therefore fail closed.
 const seen=new Set();return records.filter(record=>{const key=JSON.stringify([record.column.role,canonical(record.column.subject),record.body.map(p=>p.text),(record.section.notes||[]).map(p=>p.text)]);if(seen.has(key))return false;seen.add(key);return true;});
}
function matchesName(text,record){return record.aliases.has(canonical(text));}
function identity(text,record){
 if(matchesName(text,record))return true;
 const cleaned=compact(text).replace(/^[•·●◆]+/,'').replace(/^(?:管理人|投资顾问|受托人|股东)(?:介绍)?[：:]/,'');
 const legal=cleaned.match(/^(.+?(?:股份有限公司|有限责任公司|有限公司))(?=成立|由|是|提供|作为|于|拥有|$)/)?.[1];
 return !!legal&&matchesName(legal,record);
}
function rect(p){
 const owner=ancestor(p.node,NS.p,'sp')||ancestor(p.node,NS.p,'graphicFrame');
 const properties=children(owner,NS.p,'spPr')[0];
 const transform=elements(properties,NS.a,'xfrm')[0]||children(owner,NS.p,'xfrm')[0];
 const off=children(transform,NS.a,'off')[0],ext=children(transform,NS.a,'ext')[0];
 if(!off||!ext)return null;
 return{x:Number(off.getAttribute('x')),y:Number(off.getAttribute('y')),w:Number(ext.getAttribute('cx')),h:Number(ext.getAttribute('cy'))};
}
function groups(paragraphs){const map=new Map();for(const p of paragraphs){const key=p.tableId?`${p.tableId}:${p.cellIndex}`:p.shapeId||p.id;if(!map.has(key))map.set(key,{key,paragraphs:[],rect:rect(p)});map.get(key).paragraphs.push(p);}return[...map.values()].map(g=>({...g,text:g.paragraphs.map(p=>p.text).join('\n')}));}
function distance(body,anchor){
 if(body===anchor)return 0;
 if(body.key.includes('#table-')&&anchor.key===body.key)return 0;
 const a=body.rect,b=anchor.rect;if(!a||!b)return Infinity;
 const overlap=Math.min(a.x+a.w,b.x+b.w)-Math.max(a.x,b.x);
 if(overlap<=0||b.y>a.y+a.h)return Infinity;
 return Math.abs(a.y-b.y)+Math.abs(a.x+a.w/2-b.x-b.w/2);
}
function explicitLegalName(text){return compact(text).match(/^([^：:，,。；;]*?(?:股份有限公司|有限责任公司|有限公司))(?=成立|由|是|提供|作为|于|拥有|$)/)?.[1];}
function subjectCaption(text){
 const value=compact(text).replace(/^(?:管理人|投资顾问|受托人|股东)(?:介绍)?[：:]?/,'');
 return value.length>=4&&value.length<=60&&!/[，,。；;：:（）()]/.test(value)&&/(?:资产管理|资管|证券|信托|银行|有限公司)$/.test(value)?value:null;
}
function chooseRecord(body,records,allGroups){
 const explicitSubjects=body.paragraphs.map(p=>subjectCaption(p.text)||explicitLegalName(p.text)).filter(Boolean);
 // A known opening sentence cannot claim later paragraphs about an unknown
 // entity in the same native text box. Retain that mixed box for review.
 if(explicitSubjects.some(name=>!records.some(record=>matchesName(name,record))))return null;
 const direct=records.filter(r=>body.paragraphs.some(p=>identity(p.text,r)));
 if(direct.length===1)return direct[0];
 if(direct.length>1)return null;
 if(body.paragraphs.some(p=>explicitLegalName(p.text)||subjectCaption(p.text)))return null;
 // Known and unknown entity captions are both ownership boundaries. Determine
 // the subject/role before considering phrases such as registered capital.
 const anchors=allGroups.filter(g=>g.paragraphs.some(p=>subjectCaption(p.text)||explicitLegalName(p.text)));
 const scored=anchors.map(group=>({group,records:records.filter(r=>group.paragraphs.some(p=>identity(p.text,r))),score:distance(body,group)})).filter(x=>Number.isFinite(x.score)).sort((a,b)=>a.score-b.score);
 if(scored.length){
  const best=scored.filter(x=>x.score===scored[0].score);if(best.some(x=>x.records.length!==1))return null;
  const candidates=unique(best.flatMap(x=>x.records));return candidates.length===1?candidates[0]:null;
 }
 if(anchors.length)return null;
 const noteMatches=records.filter(r=>allGroups.some(g=>namesInNote(g.text).some(n=>matchesName(n,r))));
 return noteMatches.length===1?noteMatches[0]:null;
}
function warning(task,m,message){const text=`PPT公司介绍：${message}`;if(!m.warnings.includes(text))m.warnings.push(text);if(!task.warnings.includes(text))task.warnings.push(text);}
function addBlock(task,m,profile,record,paragraphs,source,label,evidenceSource=source){
 const targets=paragraphs.filter(p=>p.text.trim());if(!targets.length||!source.length)return;
 // Callers supply one revised value per original paragraph. Never redistribute
 // the authority's short Word bullets across a PPT's narrative paragraphs.
 if(targets.length!==source.length)throw new Error('PPT公司介绍必须保留原段落边界。');
 const before=targets.map(p=>p.text).join('\n'),after=source.map(p=>p.text).join('\n');
 const edits=targets.flatMap((p,i)=>managerTextDifferences(p.text,source[i].text).map(edit=>({...edit,paragraphId:p.id,part:p.part,index:p.index,profilePpt:true})));
 if(edits.some(e=>m.operations.some(o=>o.paragraphId===e.paragraphId))){warning(task,m,'介绍正文与已有核对位置重叠，已保留原文待核。');return;}
 const id='ppt-company-fact-'+m.materialId+'-'+(m.fields.filter(f=>f.profilePpt&&f.source==='company').length+1),fieldId=id+'-field';
 const evidence=evidenceSource.map(p=>({blockId:'company-profile-'+profile.sha256.slice(0,16)+'-'+record.section.id+'-'+p.paragraphId,quote:p.text}));
 const reason='依据已保存公司介绍 '+profile.version+'，按公司主体和对应条目作最小差异更新，保留原段落及原话术。';
 task.facts.push({id,key:'company:'+record.column.role,label,value:after,status:'verified',source:'company',profilePpt:true,evidence,scope:{company:record.column.subject,role:record.column.role,layout:record.section.layout},occurrences:targets.map(p=>({materialId:m.materialId,paragraphId:p.id,fieldId})),reason,reasonType:'company_profile'});
 m.fields.push({id:fieldId,label:'公司介绍 · '+label,before,after,status:'verified',source:'company',profilePpt:true,factIds:[id],paragraphIds:targets.map(p=>p.id),reason});
 for(const [i,edit]of edits.entries())m.operations.push({...edit,id:id+'-op-'+i,kind:'replace',source:'company',factIds:[id],reason,evidence});
}
function bodyParagraphs(group,record){return group.paragraphs.filter(p=>p.text.trim()&&!noteRE.test(p.text)&&!(record.column.role==='company'&&isolationRE.test(p.text))&&!matchesName(p.text,record)&&!/^(?:管理人|投资顾问|受托人|股东|公司)(?:介绍|实力)[：:]?$/.test(compact(p.text)));}
const atomRules=[
 ['capital',/注册资本(?:金)?\s*(\d+(?:\.\d+)?(?:万)?亿元)/g,1],
 ['qualification',/首批[^，,。；;\n]{0,15}资产管理(?:子)?公司/g],
 ['institution',/全能型资产管理机构/g],
 ['experience',/\d+(?:\.\d+)?多?年(?:多)?(?:的)?资产管理经验/g],
 ['team',/(?:专业|优秀|强大|领先)[^，,。；;\n]{0,6}投研团队/g],
 ['product_system',/(?:全面|丰富|覆盖面广|完善)[^，,。；;\n]{0,6}产品体系/g],
 ['awards',/荣获[^，,。；;\n]*?金牛奖[”」"]?/g],
 ['aum',/(?:截至[^。；;\n]*?[，,]?\s*)?(?:公司)?(?:资产管理|资管)规模(?:超|超过|达|达到|约|为)?\s*\d+(?:\.\d+)?(?:万)?亿元/g],
];
const noteRules=[['aum_date',/(?:资管规模数据时间|规模数据|数据时间)?截至\d{4}[.年/-]\d{1,2}[.月/-]\d{1,2}日?/g],['award_issuer',/金牛奖颁奖机构为[^，,。；;\n]+/g],['award_period',/统计时间为[^，,。；;\n]+/g],['award_recent',/近两年管理人类获奖情况为[^。；;\n]+/g]];
function atoms(text,rules){return rules.flatMap(([key,re,capture])=>[...text.matchAll(re)].map(match=>({key,start:match.index+(capture?match[0].indexOf(match[capture]):0),text:capture?match[capture]:match[0]})));}
function sourceAtoms(record,rules){
 const out=[];for(const source of record.column.entries)for(const atom of atoms(source.text,rules).filter(a=>a.key===source.key))if(!out.some(a=>a.key===atom.key&&a.text===atom.text))out.push({...atom,paragraphId:source.paragraphId,evidence:source});return out;
}
function renameText(text,record){
 let after=text;for(const [old,rename]of record.renames){
  const full=rename.evidence.match(/^\s*([^，,。；;]*?(?:股份有限公司|有限责任公司|有限公司))/)?.[1];
  if(full&&canonical(full)===old)after=after.replaceAll(full,rename.text);
 }
 return after;
}
function reconcileParagraph(task,m,profile,record,p,targets,sources,label){
 let after=p.text,proof=[],used=[];
 for(const target of targets.sort((a,b)=>b.start-a.start)){
  const candidates=sources.filter(s=>s.key===target.key);if(candidates.length!==1)continue;
  const source=candidates[0];if(used.some(t=>target.start<t.start+t.text.length&&t.start<target.start+target.text.length))continue;
  after=after.slice(0,target.start)+source.text+after.slice(target.start+target.text.length);used.push(target);proof.push(source.evidence||source);
 }
 const renamed=renameText(after,record);if(renamed!==after){after=renamed;for(const rename of record.renames.values())proof.push({paragraphId:'rename',text:rename.evidence});}
 if(!proof.length)return;
 addBlock(task,m,profile,record,[p],[{paragraphId:p.id,text:after}],label,unique(proof));
}
function reconcileBody(task,m,profile,record,paragraphs){
 if(record.column.role==='company'){
  const sources=sourceAtoms(record,atomRules);
  for(const p of paragraphs)reconcileParagraph(task,m,profile,record,p,atoms(p.text,atomRules),sources,'管理人介绍');
  return;
 }
 for(const p of paragraphs){
  if(isolationRE.test(p.text)){
   // The existing declaration often uses legal names and a different sentence
   // order. Equivalent isolation/reference statements must remain verbatim.
   const source=record.column.entries.find(e=>e.key==='shareholder_isolation');if(!source)continue;
   let after=renameText(p.text,record);const missing=[/[^，,。；;（）()]*业务隔离[^，,。；;（）()]*/,/股东不直接[^，,。；;（）()]*/,/股东情况仅供参考/].flatMap(re=>{const clause=source.text.match(re)?.[0];return clause&&!re.test(after)?[clause]:[];});
   if(missing.length)after=after.replace(/[。]?$/,'；'+missing.join('；')+'。');
   addBlock(task,m,profile,record,[p],[{paragraphId:p.id,text:after}],'股东隔离声明',[source]);continue;
  }
  const targets=[];
  const network=p.text.match(/(?:公司)?分支机构[^。；;\n]*[。]?|服务范围覆盖[^。；;\n]*[。]?|资本实力位居[^。；;\n]*[。]?/);
  if(network){let end=network.index+network[0].length;const tail=p.text.slice(end).match(/^(?:服务范围覆盖|资本实力位居)[^。；;\n]*[。]?/);if(tail)end+=tail[0].length;targets.push({key:'shareholder_network',start:network.index,text:p.text.slice(network.index,end)});}
  else{const description=p.text.match(/[^。；;\n]*?(?:有限公司)[^。；;\n]*[。]?|[^。；;\n]*金融服务商[^。；;\n]*[。]?/);if(description)targets.push({key:'shareholder_description',start:description.index,text:description[0]});}
  reconcileParagraph(task,m,profile,record,p,targets,record.column.entries.map(e=>({...e,evidence:e})),'股东介绍');
 }
}
function sourceNotes(task,m,profile,record,allGroups,used){
 const notes=record.section.notes||[],sources=[];
 for(const note of notes)for(const atom of atoms(note.text,noteRules))if(!sources.some(a=>a.key===atom.key&&a.text===atom.text))sources.push({...atom,paragraphId:note.paragraphId,evidence:note});
 if(!sources.length)return;
 for(const group of allGroups){
  const names=namesInNote(group.text);if(!names.some(name=>matchesName(name,record)))continue;
  for(const p of group.paragraphs){if(used.has(p.id)||!noteRE.test(p.text))continue;
   const own=namesInNote(p.text);if(own.length&&!own.some(name=>matchesName(name,record)))continue;
   let targets=atoms(p.text,noteRules);const dates=targets.filter(a=>a.key==='aum_date'),labeled=dates.filter(a=>/^(?:资管规模数据时间|规模数据)/.test(a.text));if(labeled.length)targets=targets.filter(a=>a.key!=='aum_date'||labeled.includes(a));else if(dates.length>1)targets=targets.filter(a=>a.key!=='aum_date');
   reconcileParagraph(task,m,profile,record,p,targets,sources,'来源说明');used.add(p.id);
  }
 }
}
function detectSlide(m,records,slide,previous){
 const ps=m.template.paragraphs.filter(p=>p.page===slide.page),allGroups=groups(ps),matches=[],used=new Set();
 const titleGroups=allGroups.filter(g=>g.paragraphs.some(p=>p.text.trim())&&g.rect).sort((a,b)=>a.rect.y-b.rect.y);
 const title=titleGroups[0]?.text||'';
 const continuation=ps.some(p=>/^(?:公司|管理人|股东|投资顾问)介绍[（(]?续[）)]?$/.test(compact(p.text)))&&previous?.matches.length>0;
 const result={page:slide.page,allGroups,matches,used,continuation};
 if(/团队|荣誉|奖项|获奖|股权结构/.test(title)&&!/(?:公司|管理人|股东)介绍/.test(title))return result;
 for(const group of allGroups){
  if(group.paragraphs.filter(p=>p.text.trim()).every(p=>kpiKey(p.text)||kpiValue(p.text)))continue;
  const core=group.paragraphs.filter(p=>!noteRE.test(p.text)&&!isolationRE.test(p.text)).map(p=>p.text).join('\n');
  const role=companyKeys(core).length?'company':semanticKeys(core,'shareholder').length?'shareholder':null;
  if(!role)continue;
  let record=chooseRecord(group,records,allGroups);
  if(!record&&continuation&&!ps.some(p=>subjectCaption(p.text)||explicitLegalName(p.text))){
   const candidates=unique(previous.matches.filter(m=>m.record.column.role===role).map(m=>m.record));if(candidates.length===1)record=candidates[0];
  }
  if(!record)continue;
  const targets=bodyParagraphs(group,record);if(!targets.length)continue;
  matches.push({group,record,targets});
 }
 // An isolated award/team phrase is insufficient even when a source note names
 // the authority. Several introduction characteristics must agree on a subject.
 result.matches=matches.filter(({record})=>record.column.role==='shareholder'||continuation||unique(matches.filter(m=>m.record===record).flatMap(m=>companyKeys(m.group.text)).filter(k=>!['team','product_system'].includes(k))).length>=2);
 return result;
}
function processBatch(task,m,profile,records,batch){
 for(const page of batch){const {allGroups,matches,used}=page;
  for(const {record,targets}of matches){reconcileBody(task,m,profile,record,targets);targets.forEach(p=>used.add(p.id));}
  for(const record of unique(matches.map(match=>match.record))){
   if(record.column.role==='shareholder')for(const group of allGroups)for(const p of group.paragraphs){if(!used.has(p.id)&&isolationRE.test(p.text)&&chooseRecord(group,records,allGroups)===record){reconcileBody(task,m,profile,record,[p]);used.add(p.id);}}
   for(const caption of allGroups.flatMap(g=>g.paragraphs).filter(p=>!used.has(p.id)&&matchesName(p.text,record))){
    const rename=record.renames.get(canonical(caption.text));if(!rename)continue;
    const prefix=caption.text.match(/^\s*(?:管理人|投资顾问|受托人|股东)(?:介绍)?[：:]?/)?.[0]||'';
    addBlock(task,m,profile,record,[caption],[{paragraphId:'rename',text:prefix+rename.text}],'公司名称',[{paragraphId:'rename',text:rename.evidence}]);used.add(caption.id);
   }
   if(record.column.role==='company')sourceNotes(task,m,profile,record,allGroups,used);
  }
 }
}

function processKpi(task,m,profile,slide){
 const paragraphs=m.template.paragraphs.filter(p=>p.page===slide.page),labels=paragraphs.filter(p=>kpiKey(p.text));
 if(labels.length!==3||unique(labels.map(p=>kpiKey(p.text))).length!==3)return false;
 const names=paragraphs.flatMap(p=>namesInNote(p.text));
 const sections=profile.sections.filter(s=>s.layout==='kpi'&&names.some(name=>canonical(name)===canonical(s.subject)));
 if(sections.length!==1)return false;
 const section=sections[0],column=section.columns.find(c=>c.role==='company');if(!column)return false;
 const record={section,column,aliases:new Set([canonical(section.subject),...(section.aliases||[]).map(canonical)]),renames:new Map()},mapping=[];
 for(const label of labels){
  const key=kpiKey(label.text),labelRect=rect(label),same=paragraphs.filter(p=>p.shapeId===label.shapeId&&p.tableId===label.tableId&&p.cellIndex===label.cellIndex&&p.index<label.index&&kpiValue(p.text));
  let target=same.length===1?same[0]:null;
  if(!target&&labelRect){
   const choices=paragraphs.filter(p=>kpiValue(p.text)).map(p=>({p,r:rect(p)})).filter(({r})=>r&&r.y<labelRect.y&&Math.min(r.x+r.w,labelRect.x+labelRect.w)>Math.max(r.x,labelRect.x)).map(({p,r})=>({p,d:labelRect.y-r.y+Math.abs(labelRect.x+labelRect.w/2-r.x-r.w/2)})).sort((a,b)=>a.d-b.d);
   if(choices.length&&(!choices[1]||choices[0].d<choices[1].d))target=choices[0].p;
  }
  const source=column.entries.filter(e=>e.key===key),sourceLabel=column.entries.filter(e=>e.key===`${key}_label`);
  if(!target||source.length!==1||sourceLabel.length!==1)return false;
  mapping.push({target,label,source:source[0],sourceLabel:sourceLabel[0]});
 }
 if(unique(mapping.map(x=>x.target.id)).length!==3)return false;
 for(const {target,label,source,sourceLabel}of mapping){addBlock(task,m,profile,record,[target],[source],'公司指标');addBlock(task,m,profile,record,[label],[sourceLabel],'指标名称');}
 sourceNotes(task,m,profile,record,groups(paragraphs),new Set());return true;
}

export function applyPptCompanyProfile(task,profile){
 if(!profile||profile.schemaVersion!==1||!Array.isArray(profile.sections)||!/^[0-9a-f]{64}$/.test(profile.sha256||''))throw new Error('公司介绍资料无效，请重新导入。');
 const materials=task.materials.filter(m=>m.kind==='ppt'&&m.template?.kind==='pptx'),ids=new Set(materials.map(m=>m.materialId));
 const isOurs=f=>f.source==='company'&&(f.profilePpt||f.occurrences?.some(o=>ids.has(o.materialId)));
 task.facts=task.facts.filter(f=>!isOurs(f));task.warnings=(task.warnings||[]).filter(w=>!(typeof w==='string'&&w.startsWith('PPT公司介绍：')));
 const records=sourceRecords(profile);
 for(const m of materials){
  m.operations=m.operations.filter(o=>o.source!=='company');m.fields=m.fields.filter(f=>f.source!=='company');m.warnings=(m.warnings||[]).filter(w=>!(typeof w==='string'&&w.startsWith('PPT公司介绍：')));
  m.profilePages={...m.profilePages,company:[]};
  const pages=[];for(const slide of m.template.slides)pages.push(detectSlide(m,records,slide,pages.at(-1)));
  let batch=[];const flush=()=>{if(batch.length)processBatch(task,m,profile,records,batch);batch=[];};
  for(const page of pages){if(!page.continuation)flush();if(page.matches.length){batch.push(page);m.profilePages.company.push(page.page);}else flush();}flush();
  for(const slide of m.template.slides)if(!m.profilePages.company.includes(slide.page)&&processKpi(task,m,profile,slide))m.profilePages.company.push(slide.page);
  m.profilePages.company.sort((a,b)=>a-b);
 }
 task.companyProfile={name:profile.name,version:profile.version,sha256:profile.sha256};
 task.stats={...task.stats,facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};
 return task;
}
