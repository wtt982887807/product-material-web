import {NS,ancestor} from './office-xml.js';
import {prepareManagerBiography} from './manager-data.js';
import {currentProductionDate} from './production-date.js';

const trim=s=>String(s||'').replace(/\r\n?/g,'\n').trim();
const compact=s=>trim(s).normalize('NFKC').replace(/[\s•·●◆\uf0d8]/g,'');
const bullet=/^[\s•·●◆\uf0d8]*/;
const escape=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const heading=/^(?:[一二三四五六七八九十\d]+[、.．\s]*)?(?:产品(?:要素|介绍|特色|亮点|信息)|风险(?:提示|揭示)|重要提示|免责声明|公司(?:实力|介绍|简介)|投资(?:团队|顾问)|基金投资团队|团队介绍|认购|申购|赎回|费用|业绩表现|历史业绩|数据来源|注[：:])/;
const biographyWords=/经验|学历|学士|硕士|博士|大学|现任|历任|曾任|任职|毕业|擅长|擅于|投资风格|投资理念|研究|投研|负责|深耕|精通|具备|具有|拥有|重视|注重|坚持|专注|从事|致力|追求|管理|组合|风控|策略|风险|资产|CFA|FRM/i;
const biographyIdentity=/学士|硕士|博士|学历|毕业|CFA|FRM|^\d{1,2}年[^。\n]{0,25}(?:经验|从业)|^(?:现任|历任|曾任|曾在|曾于|先后|任职|从事)/i;
const internalHeading=/^(?:投资|管理|研究|个人)(?:特点|风格|理念|方法|方法论|经历|专长|荣誉)[：:]?$/;
function isBoundary(text){const s=trim(text).replace(/^[\s\x60_—–=~～|｜#*\-]+|[\s\x60_—–=~～|｜#*\-]+$/g,'');if(internalHeading.test(s))return false;return heading.test(s)||/^.{2,25}资管\s*FOF[：:\s]/i.test(s)||s.length<35&&/(?:团队|舰队|公司实力|产品要素)/.test(s)&&!/[。；，,]/.test(s)&&!/^(?:现任|历任|曾任|曾在|曾于)/.test(s)||s.length<35&&!biographyWords.test(s)&&!/[。；，,]/.test(s);}
function anchor(p,names){
 for(const name of names){
  const match=p.text.match(new RegExp('^[\\s•·●◆\\uf0d8]*(?:(?:拟任|拟聘)?(?:投资|基金)经理(?:简介)?[：:\\s]*)?'+escape(name)+'(?=$|[\\s：:,，(（])'));
  if(!match)continue;
  let end=match[0].length;
  const suffix=p.text.slice(end).match(/^\s*(?:[（(](?:配|附)?(?:照片|图片)[）)])?(?:先生|女士)?[\s：:,，]*/)?.[0]||'';
  end+=suffix.length;
  const remainder=p.text.slice(end).trim();
  if(remainder&&!biographyIdentity.test(remainder))continue;
  // A name inside a prose/attribution sentence is not a biography header.
  if(!suffix&&!p.text.slice(end).trim()&&match.index===0||suffix||end===p.text.length)return{p,name,prefix:p.text.slice(0,end),start:end};
 }
 return null;
}
function owner(p){return ancestor(p.node,NS.w,'txbxContent')||ancestor(p.node,NS.w,'tc')||p.node?.parentNode;}
function inRegion(a,p){
 if(a.p.part!==p.part)return false;
 if(owner(a.p)===owner(p))return true;
 // A common biography card puts the name/photo in its left cell and the text
 // in the right cell. Never cross its row into another person's card.
 return !!a.p.tableId&&a.p.tableId===p.tableId&&a.p.rowIndex===p.rowIndex&&!ancestor(a.p.node,NS.w,'txbxContent')&&!ancestor(p.node,NS.w,'txbxContent');
}
function regions(template,names){
 const ps=template.paragraphs.filter(p=>p.part==='word/document.xml'),anchors=ps.map(p=>anchor(p,names)).filter(Boolean),byId=new Map(anchors.map(a=>[a.p.id,a]));
 return anchors.map(a=>{
  const entries=[],initial=a.p.text.slice(a.start);
  if(initial.trim())entries.push({p:a.p,start:a.start,text:initial,prefix:''});
  for(const p of ps.filter(p=>p.index>a.p.index)){
   if(byId.has(p.id))break;
   if(!inRegion(a,p)){
    // Ignore an enclosing textbox paragraph, but never continue through a new
    // table or text container after leaving the manager's card.
    break;
   }
   if(!p.text.trim())continue;
   if(isBoundary(p.text))break;
   const prefix=p.text.match(bullet)?.[0]||'';
   entries.push({p,start:prefix.length,text:p.text.slice(prefix.length),prefix});
  }
  return{...a,entries};
 }).filter(a=>a.entries.length||!a.p.tableId||/[（(](?:配|附)?(?:照片|图片)[）)]/.test(a.prefix));
}
function similarity(a,b){
 a=compact(a);b=compact(b);if(a===b)return 1;if(!a||!b)return 0;
 const grams=s=>{const out=new Map();for(let i=0;i<s.length-1;i++)out.set(s.slice(i,i+2),(out.get(s.slice(i,i+2))||0)+1);return out;};
 const x=grams(a),y=grams(b);let count=0;for(const [key,n]of x)count+=Math.min(n,y.get(key)||0);
 return 2*count/Math.max(1,a.length+b.length-2);
}
// Exact character edits retain all common text, including punctuation and
// formatting runs. Empty anchors are intentional insertions, never a delete/
// insert of a neighbouring unchanged character.
export function managerTextDifferences(before,after){
 if(before===after)return[];
 const n=before.length,m=after.length,rows=Array.from({length:n+1},()=>new Uint32Array(m+1));
 for(let i=n-1;i>=0;i--)for(let j=m-1;j>=0;j--)rows[i][j]=before[i]===after[j]?1+rows[i+1][j+1]:Math.max(rows[i+1][j],rows[i][j+1]);
 const edits=[];let i=0,j=0,change;
 const flush=()=>{if(change){edits.push({...change,before:before.slice(change.start,i)});change=null;}};
 while(i<n||j<m){if(i<n&&j<m&&before[i]===after[j]){flush();i++;j++;continue;}change||={start:i,after:''};if(j<m&&(i===n||rows[i][j+1]>=rows[i+1][j]))change.after+=after[j++];else i++;}
 flush();return edits;
}
function align(before,after){
 const n=before.length,m=after.length,dp=Array.from({length:n+1},()=>new Float64Array(m+1)),path=Array.from({length:n+1},()=>new Uint8Array(m+1));
 for(let i=1;i<=n;i++){dp[i][0]=i*.7;path[i][0]=1;}for(let j=1;j<=m;j++){dp[0][j]=j*.7;path[0][j]=2;}
 for(let i=1;i<=n;i++)for(let j=1;j<=m;j++){
  const costs=[dp[i-1][j-1]+1.3*(1-similarity(before[i-1].text,after[j-1])),dp[i-1][j]+.7,dp[i][j-1]+.7];
  dp[i][j]=Math.min(...costs);path[i][j]=costs.indexOf(dp[i][j]);
 }
 const out=[];let i=n,j=m;while(i||j){const step=path[i][j];if(step===0)out.push({target:before[--i],text:after[--j]});else if(step===1)out.push({target:before[--i],text:null});else out.push({target:null,text:after[--j]});}return out.reverse();
}
function canonicalLines(text,entries){
 let lines=trim(text).split('\n').map(trim).filter(Boolean);
 if(entries.length===1&&entries[0].text.includes('\n'))return[lines.join('\n')];
 if(lines.length===1&&entries.length>1)lines=lines[0].match(/[^。；;]+[。；;]?/g)||lines;
 return lines;
}
function evidence(profile,record,prepared){return[{blockId:`manager-profile-${profile.sha256.slice(0,16)}-${record.id}`,source:{file:profile.name,sheet:record.sheet,row:record.row,name:record.name},fileName:profile.name,sheet:record.sheet,row:record.row,quote:`${profile.name} · ${record.sheet} · 第${record.row}行 · ${record.name}\n${record.biography}`,calculations:prepared.calculations}];}
function addBiography(task,m,profile,region,record,productionDate){
 const prepared=prepareManagerBiography(record,productionDate),canonical=trim(prepared.text);if(!canonical)return;
 const number=task.facts.filter(f=>f.source==='manager').length+1,id=`manager-fact-${number}`,fieldId=`${m.materialId}-manager-field-${number}`,proof=evidence(profile,record,prepared);
 const reason=`依据经理资料“${record.sheet}”第${record.row}行，按${region.name}匹配并更新完整简介。`+(prepared.calculations.length?`经验年限按制作日期${productionDate}计算。`:'');
 const operations=[],add=(p,change,extra={})=>operations.push({id:`manager-op-${m.materialId}-${number}-${operations.length+1}`,kind:'replace',source:'manager',paragraphId:p.id,part:p.part,index:p.index,...change,...extra,factIds:[id],reason,evidence:proof});
 const sourceLines=canonicalLines(canonical,region.entries),pairs=align(region.entries,sourceLines);
 for(const pair of pairs)if(pair.target){
  const entry=pair.target,after=pair.text??'';
  if(pair.text===null&&entry.p!==region.p){add(entry.p,{start:0,before:entry.p.text,after:''},{managerDeleteParagraph:true});continue;}
  for(const diff of managerTextDifferences(entry.text,after))add(entry.p,{...diff,start:entry.start+diff.start});
  if(pair.text===null&&entry.start===0&&entry.text===entry.p.text){
   const op=operations.at(-1);if(op?.paragraphId===entry.p.id)op.managerDeleteParagraph=true;
  }
 }
 for(let i=0;i<pairs.length;i++)if(!pairs[i].target){
  const lines=[];let j=i;while(j<pairs.length&&!pairs[j].target)lines.push(pairs[j++].text);
  const next=pairs.slice(j).find(p=>p.target)?.target,previous=pairs.slice(0,i).reverse().find(p=>p.target)?.target;
  const anchor=next?.p||previous?.p||region.p,position=next?'before':'after';
  if(position==='before'&&anchor===region.p){add(anchor,{start:region.start,before:'',after:lines.join('\n')+'\n'});i=j-1;continue;}
  add(anchor,{start:position==='before'?0:anchor.text.length,before:'',after:position==='before'?lines.join('\n')+'\n':'\n'+lines.join('\n')},{managerParagraphs:{position,texts:lines}});i=j-1;
 }
 // A biography region is independent from contract facts. A true conflicting
 // edit is an internal ownership error; do not discard the user's supplement.
 if(operations.some(o=>m.operations.some(other=>o.paragraphId===other.paragraphId&&other.kind==='replace'&&o.start<other.start+other.before.length&&other.start<o.start+o.before.length)))return;
 const paragraphIds=[...new Set(region.entries.map(e=>e.p.id).concat(operations.map(o=>o.paragraphId)))];
 task.facts.push({id,key:`manager:${record.name}`,label:`投资经理 · ${record.name}`,value:canonical,status:'verified',source:'manager',evidence:proof,scope:{manager:record.name,recordId:record.id,sheet:record.sheet,row:record.row},calculations:prepared.calculations,occurrences:[{materialId:m.materialId,paragraphId:region.p.id,fieldId}],reason,reasonType:'manager_profile'});
 m.fields.push({id:fieldId,label:`投资经理 · ${record.name}`,before:region.entries.map(e=>e.text).join('\n'),after:canonical,status:'verified',source:'manager',managerName:record.name,factIds:[id],paragraphIds,reason});
 m.operations.push(...operations);
}
export function applyManagerProfile(task,profile,{productionDate=currentProductionDate()}={}){
 if(!profile||profile.schemaVersion!==1||!Array.isArray(profile.records)||!/^[0-9a-f]{64}$/.test(profile.sha256||''))throw new Error('经理资料无效，请重新导入 Excel。');
 const records=profile.records.filter(r=>trim(r.name)&&trim(r.biography)),names=[...new Set(records.map(r=>r.name))].sort((a,b)=>b.length-a.length);
 task.facts=task.facts.filter(f=>f.source!=='manager');
 for(const m of task.materials){
  m.operations=m.operations.filter(o=>o.source!=='manager');m.fields=m.fields.filter(f=>f.source!=='manager');
  if(m.kind!=='word'||m.template.kind!=='docx')continue;
  for(const region of regions(m.template,names)){
   const before=region.entries.map(e=>e.text).join('\n');
   const candidates=records.filter(r=>r.name===region.name).map((record,index)=>({record,index,score:similarity(before,record.biography)})).sort((a,b)=>b.score-a.score||a.index-b.index);
   if(candidates.length)addBiography(task,m,profile,region,candidates[0].record,productionDate);
  }
 }
 const managers=task.facts.filter(f=>f.source==='manager');
 task.managerProfile={name:profile.name,sha256:profile.sha256,productionDate,recordCount:records.length,matchedCount:managers.length,matchedNames:[...new Set(managers.map(f=>f.scope.manager))]};
 task.stats={...task.stats,facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};
 return task;
}
