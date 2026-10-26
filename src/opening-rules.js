// Added branch for fixed multiple dates within each month. Other opening rules
// intentionally remain in the existing atom resolver.
const KEY='opening_fixed_monthdays';
const unique=a=>[...new Set(a)];
const clean=s=>String(s??'').normalize('NFKC').replace(/\s/g,'').replace(/[，,。；;：:、【】\[\]“”"']/g,'');
const scopeKeys=['share','channel','issuer','period'];
const sameScope=(a={},b={})=>scopeKeys.every(k=>(a[k]||'')===(b[k]||''));
const MONTH_LIST=/每(?:个)?(?:自然)?月(?:的)?\s*(\d{1,2}\s*[日号]?\s*(?:[、，,和及与]\s*\d{1,2}\s*[日号]?)+)/g;
const excluded=/底层|标的(?:基金|产品|计划)|所投资(?:的)?(?:基金|产品|计划)|拟投资(?:的)?|拟投向(?:的)?|举例|例如|示例|假设/;
const majorHeading=/^(?:第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十]+[、.．])/;

export function fixedMonthlyDatePatterns(value){
 const text=String(value??''),out=[];
 for(const match of text.matchAll(MONTH_LIST)){
  const list=match[1],tokens=[...list.matchAll(/\d{1,2}/g)].map(m=>({start:match.index+match[0].indexOf(list)+m.index,before:m[0],value:Number(m[0])}));
  if(!/[日号]\s*$/.test(list)||tokens.length<2)continue;
  const tail=text.slice(match.index+match[0].length);
  const warning=tokens.some(t=>t.value<1||t.value>31)||new Set(tokens.map(t=>t.value)).size!==tokens.length?'月内固定日期有无效或重复日期，需核对完整规则。':/^\s*(?:[、和及与]\s*(?:月末|最后|月底|\d)|(?:至|到|[-~～])\s*\d)/.test(tail)?'月内固定日期枚举尚未完整识别，需核对完整规则。':'';
  out.push({start:match.index,before:match[0],tokens,dates:tokens.map(t=>t.value),warning});
 }
 return out;
}

function stripLabel(text){return String(text).trim().replace(/^(?:(?:[（(]?\d+[）)、.．]|[一二三四五六七八九十]+[、.．])\s*)?开放(?:期|日|安排)\s*[：:]\s*/,'');}
function canonical(text,mask=false){
 let value=stripLabel(text);for(const p of fixedMonthlyDatePatterns(value).reverse())value=value.slice(0,p.start)+'月内日期['+(mask?'#':p.dates.join('/'))+']'+value.slice(p.start+p.before.length);
 return clean(value).replace(/本集合计划|本资产管理计划|本产品/g,'本计划');
}
function activeSentence(text,pattern){
 const left=text.slice(0,pattern.start),start=Math.max(...['。','；',';','\n'].map(x=>left.lastIndexOf(x)))+1,right=text.slice(pattern.start),end=right.search(/[。；;\n]/);
 return text.slice(start,pattern.start)+(end<0?right:right.slice(0,end));
}
function ruleContext(text,pattern,heading){
 const sentence=activeSentence(text,pattern),p=fixedMonthlyDatePatterns(sentence)[0];if(!p)return false;
 const right=sentence.slice(p.start+p.before.length),action=right.search(/开放|办理(?:参与|退出|申购|赎回)/),disclosure=right.search(/披露|公布|公告|估值|净值/);
 if(disclosure>=0&&(action<0||disclosure<action))return false;
 return action>=0||/^(?:开放期|开放日|开放安排)\s*[:：]?$/.test(heading)&&!right.trim();
}
function continuingRule(text){
 return /^(?:合同另有约定|除因|因根据合同变更|本(?:集合)?计划原则上不设立特别开放期|(?:每笔份额|本(?:集合)?计划)?(?:最短持有期|持有期|锁定期|封闭期))/.test(text)||/^[^。；;\n]{0,24}持有(?:满|至少|不少于|\s*\d)/.test(text);
}
function scopedCandidates(contract,scope={}){
 const blocks=(contract?.blocks||[]).filter(b=>b.text?.trim()),out=[];let glossary=false,share='',channel='',excludedHeading=false,previous='';
 for(let i=0;i<blocks.length;i++){
  const b=blocks[i],raw=b.text.trim(),short=raw.length<80&&!/[。；;%％]/.test(raw),major=short&&majorHeading.test(raw);
  if(short&&/^(?:(?:第[一二三四五六七八九十\d]+章|[一二三四五六七八九十]+[、.．])\s*)?释义\s*$/.test(raw)){glossary=true;previous=raw;continue;}
  if(major&&!/释义/.test(raw)){glossary=false;share='';channel='';excludedHeading=false;}
  const inlineShare=raw.match(/([A-Z])类(?:份额)?/)?.[1]||'',inlineChannel=raw.match(/渠道\s*[:：]\s*([^，。；;\s]{1,20})/)?.[1]||'';
  if(short&&!fixedMonthlyDatePatterns(raw).length){if(inlineShare)share=inlineShare;if(inlineChannel)channel=inlineChannel;if(excluded.test(raw))excludedHeading=true;}
  const effective={share:inlineShare||share,channel:inlineChannel||channel},mismatch=['share','channel'].some(k=>effective[k]&&scope[k]&&effective[k]!==scope[k]),scopeUnknown=['share','channel'].some(k=>effective[k]&&!scope[k]);
  const text=stripLabel(raw),patterns=fixedMonthlyDatePatterns(text);
  if(!patterns.length){previous=raw;continue;}
  if(glossary||excludedHeading||excluded.test(raw)||mismatch||/^(?:指|是指|系指)/.test(text)||/[:：]\s*(?:指|是指|系指)/.test(raw)||patterns.some(p=>!ruleContext(text,p,previous))){previous=raw;continue;}
  // Preserve the complete paragraph and immediately continuing rule notes.
  const lines=[text],ids=[b.id];
  for(let j=i+1;j<blocks.length;j++){
   const next=blocks[j].text.trim();
   const nextShare=next.match(/([A-Z])类(?:份额)?/)?.[1];
   if(!continuingRule(next)||excluded.test(next)||nextShare&&nextShare!==effective.share||fixedMonthlyDatePatterns(next).length)break;
   if(!lines.some(line=>clean(line).includes(clean(next))))lines.push(next);ids.push(blocks[j].id);
  }
  const full=unique(lines).join('\n'),core=patterns.map(p=>canonical(activeSentence(text,p))).join('|'),explicit=/^(?:(?:[（(]?\d+[）)、.．]|[一二三四五六七八九十]+[、.．])\s*)?开放(?:期|日|安排)\s*[：:]/.test(raw);
  out.push({text:full,lines:unique(lines),patterns:fixedMonthlyDatePatterns(full),core,priority:explicit?2:/开放期|开放安排|办理时间/.test(previous)&&previous.length<60?1:0,scopeUnknown,warning:patterns.find(p=>p.warning)?.warning,evidence:{blockId:b.id,blockIds:ids,quote:full,itemCount:1,noteCount:0}});
  previous=raw;
 }
 return out;
}

function bodyRule(contract,scope){
 const candidates=scopedCandidates(contract,scope),evidence=[...new Map(candidates.map(c=>[canonical(c.text),c.evidence])).values()];
 if(!candidates.length)return{status:'pending',evidence,reason:'合同正文尚未找到同一产品及份额适用的月内固定多日期开放规则；未采用释义、披露日期或底层产品日期。'};
 if(candidates.some(c=>c.scopeUnknown||c.warning))return{status:'pending',evidence,reason:'合同月内多日期规则的份额、渠道或完整日期枚举尚不明确，请核对适用范围。'};
 if(new Set(candidates.map(c=>c.core)).size>1)return{status:'conflict',evidence,reason:'合同正文存在同一范围的不同月内固定日期或开放条件，不能自动选择。'};
 // Shorter cross-references may omit a trailing note. A conflicting condition
 // in any body occurrence still matters, even when its heading is less direct.
 if(candidates.some((a,i)=>candidates.slice(i+1).some(b=>!canonical(a.text).startsWith(canonical(b.text))&&!canonical(b.text).startsWith(canonical(a.text)))))return{status:'conflict',evidence,reason:'合同正文日期相同，但附加持有期限或开放条件存在不同版本，请核对全部正文依据。'};
 const max=Math.max(...candidates.map(c=>c.priority)),best=candidates.filter(c=>c.priority===max),sorted=best.sort((a,b)=>b.text.length-a.text.length);let selected=sorted[0];
 if(best.some(c=>!canonical(selected.text).startsWith(canonical(c.text))))return{status:'conflict',evidence,reason:'合同正文的月内日期一致，但附加开放规则有不同表述，请核对完整依据。'};
 const fullest=[...candidates].sort((a,b)=>canonical(b.text).length-canonical(a.text).length)[0],extra=canonical(fullest.text).slice(canonical(selected.text).length);
 // A direct heading locates the rule, but cannot erase additional operative
 // conditions repeated elsewhere. Only this pure procedural cross-reference
 // may remain in evidence without expanding a matching product-elements row.
 const proceduralReference=/^(?:因根据合同变更或者展期等需要而?设立临时开放期供投资者退出的)?程序及披露比照.{1,60}执行$/;
 if(extra&&!proceduralReference.test(extra))selected=fullest;
 return{status:'verified',selected,evidence,reason:'合同正文提供完整月内固定多日期开放规则。'};
}
function ruleDifferences(before,after){
 const a=fixedMonthlyDatePatterns(before),b=fixedMonthlyDatePatterns(after),differences=[];
 if(JSON.stringify(a.map(p=>p.dates))!==JSON.stringify(b.map(p=>p.dates)))differences.push({kind:'dates',label:'月内固定日期',before:a.map(p=>p.dates.join('、')+'日').join('；'),after:b.map(p=>p.dates.join('、')+'日').join('；')});
 const rollover=s=>(s.match(/(?:非交易日|非工作日|节假日)?[^。；;\n]{0,18}顺延[^。；;\n]{0,25}/g)||[]).join('；');
 if(clean(rollover(before))!==clean(rollover(after)))differences.push({kind:'rollover',label:'节假日顺延及工作日/交易日条件',before:rollover(before),after:rollover(after)});
 const holding=s=>(s.match(/[^。；;\n]{0,24}(?:持有期|持有时间|(?:需|须)?持有(?:满|至少|不少于)|锁定期|封闭期)[^。；;\n]*/g)||[]).join('；');
 if(clean(holding(before))!==clean(holding(after)))differences.push({kind:'holding',label:'持有或封闭期限附加限制',before:holding(before),after:holding(after)});
 if(canonical(before,true)!==canonical(after,true))differences.push({kind:'conditions',label:'开放条件或附加说明',before,after});
 return differences;
}

export function resolveMonthlyDateRule(contract,templateText,scope={}){
 const before=String(templateText??''),patterns=fixedMonthlyDatePatterns(before);if(!patterns.length)return null;
 const body=bodyRule(contract,scope),base={matched:true,status:body.status,value:null,evidence:body.evidence,reason:body.reason,reasonType:'needs_review',changes:[],differences:[]};
 if(!body.selected)return base;
 const selected=body.selected,other=selected.patterns,differences=ruleDifferences(before,selected.text),compatible=!patterns.some(p=>p.warning)&&patterns.length===other.length&&patterns.every((p,i)=>p.tokens.length===other[i].tokens.length)&&canonical(before,true)===canonical(selected.text,true);
 const result={...base,status:compatible?'verified':'pending',differences,replacement:{lines:selected.lines},reason:compatible?'完整开放规则与合同正文对应，仅按原模板话术更新各月内日期。':'月内固定日期与完整开放规则需一并核对：'+unique(differences.map(d=>d.label)).join('、')+'。保留原文，可预览采用完整合同规则。'};
 if(compatible){
  for(let i=0;i<patterns.length;i++)for(let j=0;j<patterns[i].tokens.length;j++){const token=patterns[i].tokens[j],value=String(other[i].tokens[j].value),after=/^0\d/.test(token.before)?value.padStart(token.before.length,'0'):value;if(token.before!==after)result.changes.push({start:token.start,before:token.before,after});}
  let after=before;for(const c of [...result.changes].reverse())after=after.slice(0,c.start)+c.after+after.slice(c.start+c.before.length);result.value=after;result.reasonType='confirmed';
 }
 return result;
}

export function getOpeningRuleReview(task,factIds){
 const ids=new Set(factIds||[]),facts=(task?.facts||[]).filter(f=>ids.has(f.id)&&f.key===KEY);if(!facts.length)return null;
 const scope=facts[0].scope||{},scopeConflict=facts.some(f=>!sameScope(f.scope,scope)),body=bodyRule(task.contract,scope),materials=[];
 for(const m of task.materials||[]){
  const fields=(m.fields||[]).filter(field=>field.factIds.some(id=>task.facts.some(f=>f.id===id&&f.key===KEY&&sameScope(f.scope,scope))));if(!fields.length)continue;
  const items=fields.map(field=>{
   const text=field.after??field.before,result=resolveMonthlyDateRule(task.contract,text,scope),status=scopeConflict||!body.selected?'uncertain':result?.status==='verified'?'same':'structure';
   return{title:'开放期',contractText:body.selected?.text||'',templateText:text,status,reason:result?.reason||body.reason,differences:result?.differences||[]};
  });
  materials.push({materialId:m.materialId,name:m.template?.name||m.name||m.materialId,fieldIds:fields.map(f=>f.id),matched:items.filter(i=>i.status==='same').length,total:items.length,items,extra:[]});
 }
 const uncertain=scopeConflict||!body.selected,needsReview=uncertain||materials.some(m=>m.items.some(i=>i.status!=='same')),kind=uncertain?'uncertain':needsReview?'structure':'aligned',labels=unique(materials.flatMap(m=>m.items.flatMap(i=>i.differences.map(d=>d.label))));
 return{id:'section-'+KEY+'-'+JSON.stringify(Object.fromEntries(scopeKeys.map(k=>[k,scope[k]||'']))),key:KEY,label:'开放期',evidence:body.evidence,materials,kind,needsReview,summary:uncertain?body.reason:needsReview?'月内固定日期需连同'+(labels.join('、')||'完整开放规则')+'核对；不会仅验证第一日期或持有天数。':'月内固定多个日期及完整开放规则已对应；保留模板话术与格式。',...(!uncertain?{replacement:{lines:body.selected.lines}}:{})};
}
