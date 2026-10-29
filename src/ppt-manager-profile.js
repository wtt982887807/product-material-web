import {NS,ancestor,children,elements} from './office-xml.js';
import {prepareManagerBiography} from './manager-data.js';
import {currentProductionDate} from './production-date.js';
import {managerTextDifferences} from './manager-profile.js';

const trim=value=>String(value??'').replace(/\r\n?/g,'\n').trim();
const compact=value=>trim(value).normalize('NFKC').replace(/[\s•·●◆\uf0d8]/g,'');
const biographyIdentity=/学士|硕士|博士|本科|大学|毕业|学历|CFA|FRM|\d+(?:\.\d+)?年[^。\n]{0,30}(?:从业|投研|投资|金融市场)[^。\n]{0,15}经验|(?:^|[。\n])(?:现任|历任|曾任|曾就职|曾任职)/i;
const biographyWords=/擅长|擅于|投资风格|投资理念|研究|投研|负责|深耕|精通|具备|具有|拥有|重视|注重|坚持|专注|从事|致力|追求|组合|风控|策略|风险|资产/;
const boundary=/^(?:[一二三四五六七八九十\d]+[、.．\s]*)?(?:产品(?:要素|介绍|特色|亮点|信息)|风险(?:提示|揭示|等级)|重要提示|免责声明|公司(?:实力|介绍|简介)|投资团队|团队(?:介绍|组织|架构)|组织架构|荣誉奖项|获奖情况|历史业绩|业绩表现|数据来源|注[：:]|资料来源|备注[：:])/;
const jobTitle=text=>{const s=trim(text);return s.length>0&&s.length<100&&/经理|总监|主管|负责人|分析师|研究员|首席/.test(s)&&!/[。；;，,\n]|^(?:现任|历任|曾任|曾就职|曾任职)/.test(s)&&!/(?:奖|产品|观点|表示|认为)/.test(s);};
const genericTitle=/^(?:(?:拟任|拟聘)?(?:基金|投资)经理(?:简介)?[：:]?|(?:基金|投资)?经理简介)$/;
const withoutBullets=text=>trim(text).replace(/^[\s•·●◆\uf0d8]+/,'');
// Education/experience words alone do not identify a person. Collective
// subjects and headcount ratios remain organizational facts even beside a
// single known name or inside a larger text box containing a biography.
const collectiveSubject=/^(?:(?:本|该)?公司(?:管理团队)?|(?:投资|投研|研究|管理)?团队|(?:投资|研究|业务)?部门|(?:投研|研究|投资)?人员|员工|全体人员|全体员工)(?:共|现有|均|全部|拥有|具有|具备|人数|平均|在|获|荣获|被|为|是|已|累计|旗下|整体|中)|^(?:共有|拥有|现有)?\d+(?:余|多)?[名人](?:投研|研究|投资|专业)?(?:人员|员工|团队成员)|(?:硕士|博士)(?:及以上)?(?:学历)?(?:人员)?(?:占比|比例)|学历(?:人员)?(?:占比|比例)/;
const isBody=text=>{const s=withoutBullets(text);return !!s&&!boundary.test(s)&&!jobTitle(s)&&!collectiveSubject.test(s);};
const emptyBullet=text=>/^[\s•·●◆\uf0d8]*[•·●◆\uf0d8][\s•·●◆\uf0d8]*$/.test(text);
const continuationHeading=/^(?:(?:投资|基金)?经理|个人)(?:介绍|简介)[（(]续(?:页|完)?[）)]$/;
const owner=p=>ancestor(p.node,NS.a,'tc')||ancestor(p.node,NS.p,'sp')||ancestor(p.node,NS.p,'graphicFrame')||p.node?.parentNode;

// Read the native shape transform, including ancestor group transforms. XML
// order is drawing order and cannot establish which person's card owns text.
function rectangle(node){
 const shape=node?.namespaceURI===NS.p&&['sp','graphicFrame'].includes(node.localName)?node:ancestor(node,NS.p,'sp')||ancestor(node,NS.p,'graphicFrame');
 if(!shape)return null;
 const props=children(shape,NS.p,'spPr')[0],xfrm=children(props,NS.a,'xfrm')[0]||children(shape,NS.p,'xfrm')[0];
 const values=x=>{const off=children(x,NS.a,'off')[0],ext=children(x,NS.a,'ext')[0];return off&&ext?{x:Number(off.getAttribute('x')),y:Number(off.getAttribute('y')),w:Number(ext.getAttribute('cx')),h:Number(ext.getAttribute('cy'))}:null;};
 let r=values(xfrm);if(!r||!Object.values(r).every(Number.isFinite)||r.w<=0||r.h<=0)return null;
 for(let group=shape.parentNode;group;group=group.parentNode){
  if(group.namespaceURI!==NS.p||group.localName!=='grpSp')continue;
  const transform=children(children(group,NS.p,'grpSpPr')[0],NS.a,'xfrm')[0],outer=values(transform),off=children(transform,NS.a,'chOff')[0],ext=children(transform,NS.a,'chExt')[0];
  if(!outer||!off||!ext)continue;const w=Number(ext.getAttribute('cx')),h=Number(ext.getAttribute('cy'));if(w<=0||h<=0)continue;
  r={x:outer.x+(r.x-Number(off.getAttribute('x')))*outer.w/w,y:outer.y+(r.y-Number(off.getAttribute('y')))*outer.h/h,w:r.w*outer.w/w,h:r.h*outer.h/h};
 }
 return r;
}

function findAnchor(p,names){
 const raw=p.text;
 for(const name of names){
  const at=raw.indexOf(name);if(at<0)continue;
  const before=withoutBullets(raw.slice(0,at)).replace(/[：:]$/,'').trim(),tail=raw.slice(at+name.length);
  if(before&&!genericTitle.test(before)&&!jobTitle(before))continue;
  const suffix=tail.match(/^[\s：:,，]*(?:(?:先生|女士)|[（(](?:配照片|附照片|照片|图片|续|续页|续完)[）)])?[\s：:,，]*/)?.[0]||'';
  const rest=tail.slice(suffix.length),start=at+name.length+suffix.length;
  if(/^(?:表示|认为|指出|荣获|获奖|获评|荣膺|负责本|管理本)/.test(rest))continue;
  if(rest&&!biographyIdentity.test(rest)&&!jobTitle(rest))continue;
  if(rest&&jobTitle(rest))return {p,name,known:true,start:raw.length,prefix:raw,rect:rectangle(owner(p))};
  return {p,name,known:true,start:rest?start:raw.length,prefix:rest?raw.slice(0,start):raw,rect:rectangle(owner(p))};
 }
 // An unknown person's caption is still a partition boundary. Otherwise a
 // nearby known manager could overwrite that person's education and career.
 const s=withoutBullets(raw);
 const namedTitle=s.match(/^(.{2,80}?(?:经理|总监|主管|负责人|研究员|分析师))[\s：:]+([\p{Script=Han}]{2,4})(?:先生|女士)?$/u)||s.match(/^([\p{Script=Han}]{2,4})[\s：:]+(.{2,80}?(?:经理|总监|主管|负责人|研究员|分析师))$/u);
 if(namedTitle){const name=jobTitle(namedTitle[1])?namedTitle[2]:namedTitle[1];return {p,name,known:false,start:raw.length,prefix:raw,rect:rectangle(owner(p))};}
 const shortName=s.replace(/(?:先生|女士)$/,'').trim();
 if(/^[\p{Script=Han}]{2,4}$/u.test(shortName)&&!biographyIdentity.test(shortName)&&!biographyWords.test(shortName)&&!boundary.test(shortName)&&!/(?:简介|介绍|姓名|经理|团队|荣誉|学历|专业|经历|照片|观点|备注|投资|证券|金融|公司|收益)/.test(shortName))return {p,name:shortName,known:false,start:raw.length,prefix:raw,rect:rectangle(owner(p))};
 return null;
}

function similarity(a,b){
 a=compact(a);b=compact(b);if(a===b)return 1;if(!a||!b)return 0;
 const grams=s=>{const out=new Map();for(let i=0;i<s.length-1;i++)out.set(s.slice(i,i+2),(out.get(s.slice(i,i+2))||0)+1);return out;};
 const x=grams(a),y=grams(b);let count=0;for(const [key,n]of x)count+=Math.min(n,y.get(key)||0);return 2*count/Math.max(1,a.length+b.length-2);
}
function distance(a,b){
 const dx=Math.max(0,a.x-b.x-b.w,b.x-a.x-a.w),dy=Math.max(0,a.y-b.y-b.h,b.y-a.y-a.h);
 return Math.hypot(dx,dy)+Math.hypot((a.x+a.w/2)-(b.x+b.w/2),(a.y+a.h/2)-(b.y+b.h/2))*.05;
}
function nearest(rect,anchors){
 if(anchors.length===1)return anchors[0];
 if(!rect||anchors.some(a=>!a.rect))return null;
 const scores=anchors.map(a=>({a,score:distance(rect,a.rect)})).sort((a,b)=>a.score-b.score);
 if(scores[1]&&Math.abs(scores[1].score-scores[0].score)<Math.max(1,scores[0].score*.01))return null;
 return scores[0]?.a;
}
function blockedBySection(rect,target,headings){
 if(!headings.length)return false;
 if(!rect||!target.rect)return true;
 const personalDistance=distance(rect,target.rect);
 return headings.some(heading=>{
  const r=heading.rect;if(!r)return false;
  const overlap=Math.min(rect.x+rect.w,r.x+r.w)-Math.max(rect.x,r.x);
  // A separately drawn heading establishes a section for the body below it.
  // Compare it with the person's caption instead of assuming the only name
  // owns every textbox on the page.
  return r.y<=rect.y&&overlap>=Math.min(rect.w,r.w)*.5&&distance(rect,r)<personalDistance;
 });
}
function entry(p,start=0,prefix=''){return {p,start,prefix,text:p.text.slice(start)};}
function pageRegions(paragraphs,names,inherited=null){
 const shapes=new Map();for(const p of paragraphs){const key=owner(p);if(!shapes.has(key))shapes.set(key,[]);shapes.get(key).push(p);}
 const anchors=paragraphs.map(p=>findAnchor(p,names)).filter(Boolean);
 if(!anchors.length&&inherited){const p=paragraphs.find(p=>continuationHeading.test(compact(p.text)));if(p)anchors.push({p,name:inherited,known:true,start:p.text.length,prefix:p.text,rect:rectangle(owner(p))});}
 const byParagraph=new Map(anchors.map(a=>[a.p.id,a]));
 const sectionHeads=paragraphs.filter(p=>{const s=withoutBullets(p.text);return s.length<=45&&boundary.test(s)&&!biographyIdentity.test(s)&&!/^(?:数据来源|注[：:]|资料来源|备注)/.test(s);}).map(p=>({p,rect:rectangle(owner(p))}));
 const enough=entries=>entries.some(e=>biographyIdentity.test(withoutBullets(e.text))||inherited&&e.text.length>=15&&biographyWords.test(e.text));
 const regions=anchors.map(a=>({...a,entries:[],captions:[]})),byAnchor=new Map(regions.map(r=>[r.p.id,r]));
 for(const [shape,ps]of shapes){
  const local=ps.map(p=>byParagraph.get(p.id)).filter(Boolean);
  if(local.length){
   for(const anchor of local){
    const r=byAnchor.get(anchor.p.id),at=ps.indexOf(anchor.p),next=ps.findIndex((p,i)=>i>at&&byParagraph.has(p.id)),last=next<0?ps.length:next;
    if(anchor.start<anchor.p.text.length)r.entries.push(entry(anchor.p,anchor.start,anchor.prefix));
    for(const p of ps.slice(at+1,last)){if(boundary.test(withoutBullets(p.text)))break;if(isBody(p.text)||emptyBullet(p.text))r.entries.push(entry(p));}
    if(!enough(r.entries))r.entries=[];
   }
   continue;
  }
  const entries=[];for(const p of ps){if(boundary.test(withoutBullets(p.text)))break;if(isBody(p.text)||emptyBullet(p.text))entries.push(entry(p));}
  if(!enough(entries))continue;
  const rect=rectangle(shape),target=nearest(rect,regions);if(target&&!blockedBySection(rect,target,sectionHeads.filter(h=>owner(h.p)!==shape)))target.entries.push(...entries);
 }
 const accepted=regions.filter(r=>r.known&&r.entries.length);
 for(const r of accepted){
  // Native paragraph order is meaningful inside a shape; geometry orders
  // independent body boxes. Captions never become biography destinations.
  r.entries.sort((a,b)=>{const ra=rectangle(owner(a.p)),rb=rectangle(owner(b.p));return owner(a.p)===owner(b.p)?a.p.index-b.p.index:ra&&rb?(ra.y-rb.y||ra.x-rb.x||a.p.index-b.p.index):a.p.index-b.p.index;});
  for(const p of paragraphs){
   if(r.entries.some(e=>e.p===p))continue;
   const other=byParagraph.get(p.id);if(other&&other.name!==r.name)continue;
   const stripped=other?p.text.replace(other.name,'').replace(/[（(](?:续|照片)[）)]/g,'').trim():trim(p.text);
   if(!jobTitle(stripped)||genericTitle.test(stripped))continue;
   const rect=rectangle(owner(p)),target=nearest(rect,regions);if(target!==r||blockedBySection(rect,r,sectionHeads))continue;
   const same=owner(p)===owner(r.p),close=rect&&r.rect&&distance(rect,r.rect)<=Math.max(rect.h,r.rect.h)*3;
   if(same||close)r.captions.push(p);
  }
 }
 return accepted;
}

function blockKind(text){
 const s=trim(text);
 if(/^现任/.test(s))return'current';
 if(/^(?:历任|曾任|曾就职|曾在|曾于|先后|加入.{0,40}之前|\d{4}年.{0,20}加入)/.test(s))return'history';
 if(/^(?:其中)?(?:超过|超|近|约)?\d+(?:\.\d+)?年[^。\n]{0,45}(?:经验|从业)/.test(s))return /从业|金融市场|投研经验/.test(s)?'work':'investment';
 if(/学士|硕士|博士|本科|大学|毕业|学历|CFA|FRM/i.test(s))return'education';
 if(/经历/.test(s))return'history';
 if(biographyWords.test(s))return'style';
 return'other';
}
function biographyBlocks(text){
 // Excel line breaks do not define the presentation's paragraph structure.
 // Split only semantic sentences/clauses, retaining their source text and the
 // original delimiters so the template can keep its own ordering and spacing.
 const boundary=/(?:[。；;\n]+|[，,](?=(?:(?:其中)?(?:超过|超|近|约)?\d+(?:\.\d+)?年|历任|曾任(?:职)?|曾就职|曾在|曾于|现任|擅长|专注|投资风格|拥有|具有|加入[^，,。；;\n]{0,35}之前)))/g;
 const out=[];let start=0;
 for(const match of text.matchAll(boundary)){const value=trim(text.slice(start,match.index));if(value)out.push({text:value,separator:match[0],kind:blockKind(value)});start=match.index+match[0].length;}
 const value=trim(text.slice(start));if(value)out.push({text:value,separator:'',kind:blockKind(value)});return out;
}
function semanticAllocation(canonical,entries){
 const slots=[],parts=new Map();
 for(const e of entries){
  const leading=e.text.match(/^[\s•·●◆\uf0d8]*/)?.[0]||'',trailing=e.text.slice(leading.length).match(/\s*$/)?.[0]||'';
  const body=e.text.slice(leading.length,trailing?e.text.length-trailing.length:undefined),blocks=biographyBlocks(body);
  parts.set(e,{leading,trailing,blocks});for(const block of blocks)slots.push({entry:e,block,assigned:[],index:slots.length});
 }
 const sources=biographyBlocks(canonical).map((source,index)=>({...source,index})),sourceSegments=[];
 for(const source of sources){
  const affinity=target=>source.kind===target.kind?5:source.kind==='history'&&target.kind==='current'?3:source.kind==='current'&&target.kind==='history'?2.8:['work','investment','education'].includes(source.kind)&&['work','investment','education'].includes(target.kind)?2.4:source.kind==='style'&&target.kind==='current'?1:.1;
  const ranked=slots.map(slot=>({slot,score:affinity(slot.block)+similarity(source.text,slot.block.text)*4})).sort((a,b)=>b.score-a.score||a.slot.index-b.slot.index);
  const target=ranked[0]?.slot;if(!target)continue;
  target.assigned.push(source);sourceSegments.push({text:source.text,kind:source.kind,paragraphId:target.entry.p.id,sourceIndex:source.index});
 }
 const allocation=new Map();
 for(const e of entries){
  const {leading,trailing}=parts.get(e),local=slots.filter(s=>s.entry===e),retained=local.filter(s=>s.assigned.length);let text='';
  for(const [slotIndex,slot]of retained.entries()){
   // Existing semantic blocks stay in their template order. New related
   // details follow the matched block; source order only resolves additions.
   const assigned=slot.assigned.sort((a,b)=>(a.kind===slot.block.kind?0:1)-(b.kind===slot.block.kind?0:1)||a.index-b.index);
   for(const [i,source]of assigned.entries()){
    text+=source.text;
    if(i<assigned.length-1)text+=source.separator.replace(/\n/g,'')||'。';
    else if(slotIndex<retained.length-1)text+=slot.block.separator.replace(/\n/g,'')||source.separator.replace(/\n/g,'')||'。';
    else text+=slot.block.separator.replace(/\n/g,'')||source.separator.replace(/\n/g,'');
   }
  }
  // Empty spacer/bullet paragraphs are presentation structure, not obsolete
  // source content. Keep them exactly as supplied by the template.
  allocation.set(e,local.length?e.prefix+leading+text+trailing:e.p.text);
 }
 return{allocation,sourceSegments};
}

function currentTitle(biography){
 const clause=biography.match(/现任\s*([^。；;\n，,]+)/)?.[1]?.trim();if(!clause)return null;
 // A caption omits the employer; the full current-employment sentence still
 // appears verbatim in the authoritative biography beside it.
 const title=clause.replace(/^.*?(?:有限责任公司|股份有限公司|有限公司|资管|公司)/,'').trim();
 return jobTitle(title)?title:null;
}
function captionAfter(p,region,title){
 const at=p.text.indexOf(region.name),leading=p.text.match(/^\s*/)?.[0]||'',trailing=p.text.match(/\s*$/)?.[0]||'';
 const oldTitle=(at<0?p.text:p.text.replace(region.name,'')).trim(),count=oldTitle.split(/、|兼任/).length;
 const supported=new Set();let department='';
 for(const role of title.split(/、|兼任/).map(trim)){
  supported.add(compact(role));
  const match=role.match(/^(.*?)(副?总经理(?:助理)?|(?:高级|首席)?投资经理|基金经理|投资总监|研究员|负责人|主管|分析师|董事长|总裁)$/);
  if(match){if(match[1])department=match[1];supported.add(compact(department+match[2]));supported.add(compact(match[2]));}
 }
 // A template may deliberately name a secondary post. If that abbreviated
 // post still appears in the source, do not promote it to the first role.
 if(oldTitle.split(/、|兼任/).every(role=>supported.has(compact(role))))return p.text;
 title=title.split(/、|兼任/).slice(0,count).join('、');
 if(at<0)return leading+title+trailing;
 const before=p.text.slice(0,at),after=p.text.slice(at+region.name.length);
 if(jobTitle(before.trim()))return leading+title+(before.match(/\s*$/)?.[0]||' ')+region.name+after;
 if(jobTitle(after.trim()))return before+region.name+(after.match(/^\s*/)?.[0]||' ')+title+trailing;
 return p.text;
}
function applyRegion(task,m,profile,region,records,productionDate){
 const before=region.entries.map(e=>withoutBullets(e.text)).join('\n'),ranked=records.filter(r=>r.name===region.name).map((record,index)=>({record,index,score:similarity(before,record.biography)})).sort((a,b)=>b.score-a.score||a.index-b.index);
 const record=ranked[0]?.record;if(!record)return false;
 const prepared=prepareManagerBiography(record,productionDate),canonical=trim(prepared.text);if(!canonical)return false;
 const number=m.fields.filter(f=>f.source==='manager'&&f.profilePpt).length+1,id=`manager-ppt-fact-${m.materialId}-${number}`,fieldId=`${m.materialId}-manager-ppt-field-${number}`;
 const evidence=[{blockId:`manager-profile-${profile.sha256.slice(0,16)}-${record.id}`,source:{file:profile.name,sheet:record.sheet,row:record.row,name:record.name},fileName:profile.name,sheet:record.sheet,row:record.row,quote:`${profile.name} · ${record.sheet} · 第${record.row}行 · ${record.name}\n${record.biography}`,calculations:prepared.calculations}];
 const reason=`依据经理资料“${record.sheet}”第${record.row}行，按${region.name}及页面位置匹配并更新完整简介。`+(prepared.calculations.length?`经验年限按制作日期${productionDate}计算。`:'');
 const operations=[],add=(p,after)=>{for(const difference of managerTextDifferences(p.text,after))operations.push({id:`manager-ppt-op-${m.materialId}-${number}-${operations.length+1}`,kind:'replace',source:'manager',profilePpt:true,paragraphId:p.id,part:p.part,index:p.index,...difference,factIds:[id],reason,evidence});};
 const entries=region.entries,{allocation,sourceSegments}=semanticAllocation(canonical,entries);
 for(const e of entries)add(e.p,allocation.get(e));
 const title=currentTitle(canonical);if(title)for(const p of region.captions)add(p,captionAfter(p,region,title));
 if(operations.some(o=>m.operations.some(other=>other.paragraphId===o.paragraphId&&other.kind==='replace'))){
  m.warnings||=[];const warning=`PPT经理简介：第${region.p.page}页${region.name}的文字已有其他更新，请核对后重新生成。`;if(!m.warnings.includes(warning))m.warnings.push(warning);return false;
 }
 const paragraphIds=[...new Set(entries.map(e=>e.p.id).concat(region.captions.map(p=>p.id)))];
 const anchors=region.anchors||[region.p],pages=[...new Set(anchors.map(p=>p.page))];
 task.facts.push({id,key:`manager:${record.name}`,label:`投资经理 · ${record.name}`,value:canonical,status:'verified',source:'manager',profilePpt:true,evidence,scope:{manager:record.name,recordId:record.id,sheet:record.sheet,row:record.row,page:region.p.page,pages},calculations:prepared.calculations,occurrences:anchors.map(p=>({materialId:m.materialId,paragraphId:p.id,fieldId,page:p.page})),reason,reasonType:'manager_profile'});
 m.fields.push({id:fieldId,label:`投资经理 · ${record.name}`,before,after:entries.map(e=>allocation.get(e)).filter(Boolean).join('\n'),status:'verified',source:'manager',profilePpt:true,managerName:record.name,page:region.p.page,pages,factIds:[id],paragraphIds,biographyParagraphIds:entries.map(e=>e.p.id),sourceSegments,reason});
 m.operations.push(...operations);return true;
}

/** Add PPT-only manager state without clearing the existing Word adapter's state. */
export function applyPptManagerProfile(task,profile,{productionDate=currentProductionDate()}={}){
 if(!profile||profile.schemaVersion!==1||!Array.isArray(profile.records)||!/^[0-9a-f]{64}$/.test(profile.sha256||''))throw new Error('经理资料无效，请重新导入 Excel。');
 const records=profile.records.filter(r=>trim(r.name)&&trim(r.biography)),names=[...new Set(records.map(r=>r.name))].sort((a,b)=>b.length-a.length);
 const materials=task.materials.filter(m=>m.kind==='ppt'&&m.template?.kind==='pptx'),ids=new Set(materials.map(m=>m.materialId));
 task.facts=task.facts.filter(f=>!(f.source==='manager'&&f.profilePpt&&f.occurrences?.every(o=>ids.has(o.materialId))));
 for(const m of materials){
  m.operations=m.operations.filter(o=>!(o.source==='manager'&&o.profilePpt));m.fields=m.fields.filter(f=>!(f.source==='manager'&&f.profilePpt));m.profilePages={...m.profilePages,manager:[]};
  if(m.warnings)m.warnings=m.warnings.filter(w=>typeof w!=='string'||!w.startsWith('PPT经理简介：'));
  const groups=[];let previous=null;
  for(const slide of m.template.slides){
   const paragraphs=m.template.paragraphs.filter(p=>p.part===slide.part),anchors=paragraphs.map(p=>findAnchor(p,names)).filter(Boolean),continuation=previous&&previous.page===slide.page-1&&!anchors.length&&paragraphs.some(p=>continuationHeading.test(compact(p.text)));
   const regions=pageRegions(paragraphs,names,continuation?previous.region.name:null);
   if(continuation&&regions.length===1){
    const region=previous.region;region.entries.push(...regions[0].entries);region.anchors.push(regions[0].p);region.captions.push(...regions[0].captions);previous={region,page:slide.page};continue;
   }
   for(const region of regions){region.anchors=[region.p];groups.push(region);}
   previous=regions.length===1&&new Set(anchors.map(a=>a.name)).size===1?{region:regions[0],page:slide.page}:null;
  }
  for(const region of groups)if(applyRegion(task,m,profile,region,records,productionDate)){
   for(const p of region.anchors)if(!m.profilePages.manager.includes(p.page))m.profilePages.manager.push(p.page);
  }
  m.profilePages.manager.sort((a,b)=>a-b);
 }
 const managers=task.facts.filter(f=>f.source==='manager');
 task.managerProfile={...task.managerProfile,name:profile.name,sha256:profile.sha256,productionDate,recordCount:records.length,matchedCount:managers.length,matchedNames:[...new Set(managers.map(f=>f.scope?.manager).filter(Boolean))]};
 task.stats={...task.stats,facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};
 return task;
}
