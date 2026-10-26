// Read-only chapter evidence and coverage. This module never builds replacement operations.
const LABELS={investment_scope:'投资范围',investment_ratios:'投资比例',investment_restrictions:'投资限制',investment_strategy:'投资策略',investment_objective:'投资目标'};
const CACHE=new WeakMap();
const raw=s=>String(s??'').trim();
const clean=s=>String(s??'').normalize('NFKC').replace(/\u00a0/g,' ').trim();
const compact=s=>clean(s).replace(/\s/g,'');
const unique=a=>[...new Set(a)];
const prefix=/^\s*(?:(第[一二三四五六七八九十百\d]+[章节])\s*|([一二三四五六七八九十百]+)[、.．]\s*|\(([一二三四五六七八九十百]+)\)\s*|(\d+(?:\.\d+)*)(?:[、.．](?!\d)|\s+)\s*|\((\d+)\)\s*)/;
function numbering(text){
 const simple=clean(text).match(/^\d+\)\s*/);if(simple)return {text:raw(text).slice(simple[0].length).trim(),rank:6,style:'paren-decimal'};
 const m=clean(text).match(prefix);if(!m)return {text:raw(text),rank:0,style:''};
 const rank=m[1]?1:m[2]?2:m[3]?3:m[4]?4+m[4].split('.').length-1:6;
 return {text:raw(text).slice(m[0].length).trim(),rank,style:m[1]?'chapter':m[2]?'chinese':m[3]?'paren-chinese':m[4]?'decimal-'+m[4].split('.').length:'paren-decimal'};
}
function textScope(text){
 return {share:clean(text).match(/\b([A-Z])类(?:份额)?/)?.[1]||'',channel:clean(text).match(/渠道\s*[:：]\s*([^,，。;；\n]{1,20})/)?.[1]?.trim()||''};
}
function heading(text){
 const n=numbering(text),s=n.text;
 for(const [key,label] of Object.entries(LABELS)){
  const pos=s.indexOf(label);if(pos<0)continue;
  const before=s.slice(0,pos),rest=s.slice(pos+label.length);
  if(before&&!/^(?:[A-Z]类(?:份额)?\s*|(?:适用于)?渠道[:：][^\n]{1,20}\s*)$/.test(before))continue;
  if(rest&&!/^\s*[:：]/.test(rest))continue;
  return {key,label,rank:n.rank||3,scope:textScope(before),body:rest.replace(/^\s*[:：]\s*/,''),numbered:!!n.rank};
 }
 return null;
}
const referenceHeading=text=>/^(?:底层基金|底层产品|标的基金|标的产品|投资标的|产品举例|示例产品|历史产品|同策略产品|其他产品)$/.test(text)||/^(?:底层|标的基金|标的产品|投资标的|产品举例|示例产品|历史产品|同策略产品|其他产品)[^。;；:：,，%％]{0,70}(?:概况|介绍|基本信息|产品要素|投资要素|投资策略|投资情况|情况|说明)$/.test(text);
function boundary(text,current){
 if(heading(text))return true;
 const n=numbering(text);
 if(n.text.length<=140&&!/[。;；%％]/.test(n.text)&&(/^(?:适用于)?[A-Z]类(?:份额)?$/.test(compact(n.text))||/^(?:适用于)?渠道\s*[:：]/.test(n.text)||referenceHeading(n.text)))return true;
 if(/^(?:附件|附录|第[一二三四五六七八九十百\d]+[章节])/.test(n.text))return true;
 if(n.rank&&n.rank<=current.rank&&n.text.length<70&&!/[。;；:%％]/.test(n.text))return true;
 return /^(?:风险揭示|收益分配|资产估值|费用与税收|合同的变更|合同终止|特别约定|投资政策|投资经理|估值原则|关联交易|建仓期)(?:[:：])?$/.test(n.text);
}
const noteLine=text=>/^(?:本(?:集合|信托)?计划|本产品|上述|以上|前述|注[:：]?|说明|托管人|管理人)/.test(clean(text))&&/(?:穿透|合并|计算|监控|监管|上述比例|前述比例|运作规定|投资比例)/.test(text);
function sectionItems(lines){
 const content=lines.filter(l=>l.text.trim());
 const numbered=content.map(l=>numbering(l.text)).filter(n=>n.rank);
 const topRank=numbered.length?Math.min(...numbered.map(n=>n.rank)):0;
 const items=[],notes=[];let active=null;
 for(const line of content){
  const n=numbering(line.text);
  if(noteLine(n.text)&&!n.rank){notes.push(line);active=null;continue;}
  if(topRank&&n.rank===topRank){active={text:n.text,blockIds:[line.id]};items.push(active);}
  else if(active){active.text+='\n'+line.text;active.blockIds.push(line.id);}
  else if(!topRank){items.push({text:n.text,blockIds:[line.id]});}
  else notes.push(line);
 }
 return {items,notes};
}
function chapters(contract){
 const signature=JSON.stringify((contract.blocks||[]).map(b=>[b.id,b.text]));
 if(CACHE.get(contract)?.signature===signature)return CACHE.get(contract).sections;
 const lines=(contract.blocks||[]).flatMap(b=>raw(b.text).split(/\r?\n/).map(text=>({...b,text:text.trim()}))).filter(b=>b.text);
 const found=[];let inAttachment=false,contexts=[];
 for(let i=0;i<lines.length;i++){
  if(/^(?:附件|附表)(?:[一二三四五六七八九十\d\s：:]|$)/.test(lines[i].text))inAttachment=true;
  const n=numbering(lines[i].text);let h=heading(lines[i].text);
  const titleLike=n.text.length<=140&&!/[。;；%％]/.test(n.text);
  // Unnumbered element headings inherit an evidenced parent scope. A numbered
  // peer ends the previous parent; asset-class words inside body prose do not.
  if(n.rank&&titleLike)contexts=contexts.filter(c=>c.rank?c.rank<n.rank:!(c.excluded&&!h));
  if(titleLike&&/^(?:本集合计划|本信托计划|本计划|本产品).*(?:概况|介绍|基本信息|投资安排|产品要素)$/.test(n.text))contexts=contexts.filter(c=>!c.excluded);
  if(!h&&titleLike){
   const share=compact(n.text).match(/^(?:适用于)?([A-Z])类(?:份额)?$/)?.[1],channel=clean(n.text).match(/^(?:适用于)?渠道\s*[:：]\s*([^,，。;；]{1,30})$/)?.[1]?.trim();
   if(share){contexts=contexts.filter(c=>!c.scope?.share);contexts.push({rank:n.rank,scope:{share}});}
   if(channel){contexts=contexts.filter(c=>!c.scope?.channel);contexts.push({rank:n.rank,scope:{channel}});}
   if(referenceHeading(n.text))contexts.push({rank:n.rank,excluded:true});
  }
  if(!h||contexts.some(c=>c.excluded))continue;
  h={...h,scope:{share:'',channel:'',...Object.assign({},...contexts.map(c=>c.scope||{})),...Object.fromEntries(Object.entries(h.scope).filter(([,v])=>v))}};
  let end=i+1;while(end<lines.length&&!boundary(lines[end].text,h)){
   if(/^[一二三四五六七八九十百]+$/.test(lines[end].text)&&boundary(lines[end+1]?.text||'',h))break;
   end++;
  }
  const content=[...(h.body?[{...lines[i],text:h.body}]:[]),...lines.slice(i+1,end)];
  if(!content.length)continue;
  const parsed=sectionItems(content);if(!parsed.items.length)continue;
  const quote=[lines[i].text,...lines.slice(i+1,end).map(l=>l.text)].join('\n');
  const signature=h.key+'|'+JSON.stringify(h.scope)+'|'+compact(content.map(l=>l.text).join('\n'));
  const same=found.find(s=>s.signature===signature);
  const ids=unique([lines[i].id,...content.map(l=>l.id)]);
  if(same){same.evidence.blockIds=unique([...same.evidence.blockIds,...ids]);same.evidence.duplicateCount++;}
  else found.push({...h,signature,content,isAttachment:inAttachment,...parsed,evidence:{blockId:lines[i].id,blockIds:ids,quote,itemCount:parsed.items.length,noteCount:parsed.notes.length,duplicateCount:0}});
  i=end-1;
 }
 CACHE.set(contract,{signature,sections:found});return found;
}
const categories=[
 ['固定收益类资产',/固定收益(?:类)?(?:资产)?|债权类资产/g],
 ['权益类资产',/权益类(?:资产)?|股票类资产/g],
 ['期货和衍生品类账户权益',/期货和衍生品类账户权益|(?:商品及)?(?:期货|金融衍生品|衍生品)(?:类资产|类账户权益)?/g],
 ['公募基金',/公募(?:证券投资)?基金|公开募集证券投资基金/g],
 ['境外投资',/境外投资|境外资产|海外投资/g],
 ['现金及货币市场工具',/现金(?:及货币市场工具)?|货币市场工具/g],
 ['正回购融资',/正回购(?:融入资金余额)?|回购融入资金余额/g],
];
function semantic(text){
 const s=numbering(text).text;
 const colon=s.search(/[:：]/);const head=colon>=0?s.slice(0,colon):s;
 const hits=categories.flatMap(([title,re])=>[...head.matchAll(new RegExp(re.source,'g'))].map(m=>({title,index:m.index,length:m[0].length}))).sort((a,b)=>a.index-b.index||b.length-a.length);
 if(hits.length)return {title:hits[0].title,key:hits[0].title};
 const first=head.split(/[。;；\n]/)[0].replace(/(?:本集合计划|本计划|本产品)/g,'').trim();
 const title=first.replace(/\d+(?:\.\d+)?\s*%/g,'').replace(/[0-9\s~～至<>≤≥=()%％,，-]/g,'').slice(0,70);
 return {title:title||'本节说明',key:compact(title)};
}
function materialChunks(text){
 const parts=raw(text).split(/\r?\n|[;；]/).map(s=>s.trim()).filter(s=>s&&!noteLine(s));
 const chunks=[];
 for(const line of parts){
  // An abbreviated row may combine several classes separated by commas.
  const hits=categories.flatMap(([,re])=>[...line.matchAll(new RegExp(re.source,'g'))].map(m=>({index:m.index,length:m[0].length}))).sort((a,b)=>a.index-b.index||b.length-a.length);
  const starts=unique(hits.filter((h,i)=>!i||h.index>=hits[i-1].index+hits[i-1].length).map(h=>h.index));
  if(starts.length>1&&starts.slice(1).every(p=>/[,，、。]\s*$/.test(line.slice(0,p)))){
   starts.forEach((p,i)=>chunks.push(line.slice(i? p:0,starts[i+1]??line.length).replace(/[,，、。]\s*$/,'')));
  }else chunks.push(line);
 }
 return chunks.map(text=>({text,...semantic(text)}));
}
function canonical(text){
 return clean(numbering(text).text).replace(/本集合计划|本信托计划|集合计划|本计划|本产品/g,'计划').replace(/资产总值|资产总额/g,'资产总值').replace(/资产净值|资产净额/g,'资产净值').replace(/百分之/g,'%').replace(/(\d+)\.0+(?=\s*%)/g,'$1').replace(/～|~|—|–|至/g,'-').replace(/[\s,，。;；:：、]/g,'');
}
const withoutNumbers=text=>canonical(text).replace(/\d+(?:\.\d+)?/g,'#');
function compare(contractText,templateText){
 if(canonical(contractText)===canonical(templateText))return {status:'same',reason:'对应条目内容一致，保持模板表述。'};
 const base=s=>/资产总(?:值|额)|总资产/.test(s)?'total':/资产净(?:值|额)|净资产/.test(s)?'net':'';
 const a=base(contractText),b=base(templateText);
 if(a&&b&&a!==b)return {status:'structure',reason:'计量口径不同（资产总值／资产净值），需逐项核对。'};
 const prohibited=s=>/(?:不得|不能|不可|禁止)投资/.test(s),permitted=s=>/(?:可以|允许|可)投资/.test(s);
 if(prohibited(contractText)&&permitted(templateText)||permitted(contractText)&&prohibited(templateText))return {status:'structure',reason:'是否允许投资的约束不同，需逐项核对。'};
 const edges=s=>[...clean(s).matchAll(/不含|不包括|不超过|不得超过|不低于|不得低于|含|包括|小于|大于|[<>≤≥]/g)].map(m=>m[0]);
 const ca=edges(contractText),cb=edges(templateText);
 if(ca.length&&cb.length&&ca.join('|')!==cb.join('|'))return {status:'structure',reason:'上下限或包含／不包含条件不同，需核对适用口径。'};
 if(withoutNumbers(contractText)===withoutNumbers(templateText))return {status:'local',reason:'对应类别和口径一致，仅有数值差异；继续沿用原局部更新规则。'};
 // Allow only controlled omissions. A shared asset label does not prove that
 // additional restrictions, exceptions or investment conditions are equivalent.
 const abbreviated=s=>canonical(s).replace(/占(?:计划)?(?:资产总值|资产净值|总资产|净资产)的?/g,'');
 const ac=abbreviated(contractText),at=abbreviated(templateText);
 const titleOnly=canonical(templateText)===canonical(semantic(contractText).title);
 if(titleOnly||ac===at)return {status:'summary',reason:'模板以摘要表述对应本项；未要求逐字照搬合同，仍可核对完整数值及条件。'};
 if(ac.replace(/\d+(?:\.\d+)?/g,'#')===at.replace(/\d+(?:\.\d+)?/g,'#'))return {status:'local',reason:'受控摘要省略外的类别和条件一致，仅有数值差异；继续沿用原局部更新规则。'};
 return {status:'uncertain',reason:'已找到相近条目，但表述差异无法仅按数值判断，请结合完整合同核对。'};
}
function sameScope(a={},b={}){return ['share','channel','issuer','period','condition'].every(k=>(a[k]||'')===(b[k]||''));}
function factScope(task,fact){
 const scope={...(fact.scope||{})};if(!scope.condition)return scope;
 const provenanceNorm=s=>clean(s).replace(/[\s·•,，。;；:：、【】\[\]"“”'‘’]/g,'');
 const sourceTexts=(fact.occurrences||[]).flatMap(o=>{
  const material=task.materials?.find(m=>m.materialId===o.materialId);
  const paragraph=material?.template?.paragraphs?.find(p=>p.id===o.paragraphId);
  return paragraph?[paragraph.text]:[];
 });
 // The existing engine stores an atom's original text in condition to preserve
 // provenance. Remove only that exact text, never a semantic scope condition.
 if(sourceTexts.some(t=>provenanceNorm(t)===provenanceNorm(scope.condition)))scope.condition='';
 return scope;
}
function fieldMatches(task,field,key,scope){
 return (field.factIds||[]).some(id=>task.facts.some(f=>f.id===id&&f.key===key&&sameScope(factScope(task,f),scope)));
}
function materialReview(task,m,key,scope,items){
 const fields=(m.fields||[]).filter(f=>fieldMatches(task,f,key,scope));if(!fields.length)return null;
 const source=unique(fields.map(f=>String(f.after??f.before??''))).join('\n');
 const chunks=[...new Map(materialChunks(source).map(c=>[canonical(c.text),c])).values()],used=new Set();
 const rows=items.map(item=>{
  const s=semantic(item.text);let candidates=chunks.map((c,i)=>({...c,index:i})).filter(c=>c.key&&c.key===s.key);
  if(!candidates.length){const exact=chunks.map((c,i)=>({...c,index:i})).filter(c=>canonical(c.text).includes(canonical(item.text))||canonical(item.text).includes(canonical(c.text))&&canonical(c.text).length>=12);candidates=exact;}
  if(!candidates.length)return {title:s.title,contractText:item.text,templateText:'',status:'missing',reason:'本份物料未呈现该项，可能是摘要；请核对是否需要补充，不据此直接判错或改写。'};
  const selected=candidates.find(c=>canonical(c.text)===canonical(item.text))||candidates[0];
  used.add(selected.index);
  return {title:s.title,contractText:item.text,templateText:selected.text,...compare(item.text,selected.text)};
 });
 const extra=chunks.filter((c,i)=>!used.has(i)&&!noteLine(c.text)&&!heading(c.text)).map(c=>({title:c.title,contractText:'',templateText:c.text,status:'uncertain',reason:'模板存在本节尚未对应的条目，可能来自补充说明或其他章节，请核对。'}));
 return {materialId:m.materialId,name:m.template?.name||m.name||m.materialId,fieldIds:fields.map(f=>f.id),matched:rows.filter(r=>r.status!=='missing').length,total:items.length,items:rows,extra};
}

const strategyName=text=>clean(numbering(text).text).replace(/[\s。;；:：]+$/g,'').replace(/\s/g,'').replace(/投资策略$/,'策略');
const conditionalStrategy=text=>/(?:发生情形|发生情况|情形和条件|适用条件|特殊情形|超过.*(?:%|％)|突破.*(?:%|％))/.test(text);
function shortStrategyTitles(text){
 const parts=raw(text).replace(/策略[、]/g,'策略\n').split(/[\r\n;；,，]+/).map(raw).filter(Boolean);
 if(!parts.length)return null;
 const titles=parts.map(p=>({text:p,title:numbering(p).text.replace(/[。;；:：]+$/g,'').trim(),key:strategyName(p)}));
 if(titles.some(t=>t.title.length>55||!t.title.endsWith('策略')||/[。;；:：%％]/.test(t.title)||conditionalStrategy(t.title)))return null;
 return titles;
}
function strategyHierarchy(section,templateTitles){
 const lines=section.content||[],root={title:'',rank:0,children:[]},nodes=[root],stack=[root];
 for(let i=0;i<lines.length;i++){
  const n=numbering(lines[i].text),inline=n.text.match(/^([^：:]{2,55}策略)[:：]/),title=inline?inline[1]:n.text;
  const standalone=!n.rank&&title.length<=55&&title.endsWith('策略');
  if(!(n.rank||standalone)||title.length>140||/[。;；:：]/.test(title))continue;
  const rank=n.rank||section.rank+1;
  while(stack.length>1&&stack.at(-1).rank>=rank)stack.pop().end=i;
  const node={title,rank,start:i,end:lines.length,children:[],parent:stack.at(-1)};
  node.parent.children.push(node);nodes.push(node);stack.push(node);
 }
 const wanted=new Set(templateTitles.map(t=>t.key));
 const governance=/^(?:决策依据|投资决策程序|投资管理的方法和标准|投资交易程序|投资风险的监控与绩效评估)$/;
 const groups=nodes.map(parent=>{
  const items=parent.children.filter(n=>!conditionalStrategy(n.title)&&!governance.test(n.title));
  if(!items.some(n=>n.title.endsWith('策略')))return null;
  const score=items.filter(n=>wanted.has(strategyName(n.title))).length;
  return {items,notes:parent.children.filter(n=>conditionalStrategy(n.title)),score};
 }).filter(Boolean);
 if(!groups.length)return null;
 const score=Math.max(...groups.map(g=>g.score)),best=groups.filter(g=>g.score===score);
 if(best.length!==1)return {ambiguous:true,items:[],notes:[]};
 const describe=n=>({title:n.title,text:lines.slice(n.start,n.end).map(l=>l.text).join('\n'),blockIds:unique(lines.slice(n.start,n.end).map(l=>l.id))});
 return {ambiguous:false,items:best[0].items.map(describe),notes:best[0].notes.map(describe)};
}
function strategyContext(contract,titles,scope={}){
 let sections=chapters(contract).filter(s=>s.key==='investment_strategy');if(!sections.length)return null;
 sections=sections.filter(s=>!['share','channel'].some(k=>s.scope[k]&&scope[k]&&s.scope[k]!==scope[k]));if(!sections.length)return null;
 const versions=sections.map(section=>({section,tree:strategyHierarchy(section,titles)}));
 if(versions.every(v=>!v.tree))return null;
 const conditionNorm=s=>clean(s).replace(/[\s·•,，。;；:：、【】\[\]"“”'‘’]/g,'');
 const semanticCondition=scope.condition&&!titles.some(t=>conditionNorm(t.text)===conditionNorm(scope.condition))&&conditionNorm(titles.map(t=>t.text).join('；'))!==conditionNorm(scope.condition);
 const ambiguous=!!semanticCondition||versions.some(v=>!v.tree||v.tree.ambiguous)||sections.some(s=>['share','channel'].some(k=>s.scope[k]&&!scope[k]));
 const evidence=versions.map(({section,tree})=>({...section.evidence,blockIds:[...section.evidence.blockIds],...(tree&&!tree.ambiguous?{itemCount:tree.items.length,noteCount:tree.notes.length}: {})}));
 const items=[...new Map(versions.flatMap(v=>v.tree?.items||[]).map(i=>[strategyName(i.title)+'|'+canonical(i.text),i])).values()];
 const notes=[...new Map(versions.flatMap(v=>v.tree?.notes||[]).map(i=>[canonical(i.text),{title:i.title,contractText:i.text,blockIds:[...i.blockIds]}])).values()];
 return {ambiguous,conflict:sections.length>1,evidence,items,notes};
}
function strategyRows(context,titles){
 const keys=new Set(titles.map(t=>t.key)),known=new Set(context.items.map(i=>strategyName(i.title)));
 const items=context.items.map(item=>{
  const title=numbering(item.title).text,found=titles.find(t=>t.key===strategyName(title));
  return {title,contractText:title,templateText:found?.text||'',status:found?(context.ambiguous||context.conflict?'uncertain':'same'):'missing',reason:found?'策略标题与合同对应，保持标题摘要及模板原话术；完整正文在合同依据中。':'模板未呈现该策略标题，请核对是否补充；不会自动展开策略正文。'};
 });
 const extra=titles.filter(t=>!known.has(t.key)).map(t=>({title:t.title,contractText:'',templateText:t.text,status:'uncertain',reason:'合同对应策略层级未找到该标题，请核对，不按相同资产类别替代另一策略。'}));
 return {items,extra,matched:context.items.filter(i=>keys.has(strategyName(i.title))).length,total:context.items.length};
}

function strategyTitleReplacement(context,source){
 if(context.ambiguous||context.conflict||!context.items.length)return undefined;
 const titles=shortStrategyTitles(source);if(!titles)return undefined;
 const known=new Set(context.items.map(i=>strategyName(i.title))),removedTitles=unique(titles.filter(t=>!known.has(t.key)).map(t=>t.title));
 const original=String(source),tokens=[];let start=0,separator='';
 for(const m of original.matchAll(/[\r\n;；,，]+\s*|(?<=策略)、\s*/g)){
  const text=original.slice(start,m.index);if(text.trim())tokens.push({text,key:strategyName(text),separator});
  separator=m[0];start=m.index+m[0].length;
 }
 const last=original.slice(start);if(last.trim())tokens.push({text:last,key:strategyName(last),separator});
 const trailing=last.trim()?'':separator;
 const nodes=tokens.filter(t=>known.has(t.key)),kept=new Set(nodes.map(n=>n.key)),addedTitles=[];
 const usualSeparator=tokens.find(t=>t.separator)?.separator||'；';
 const prefixPattern=/^(\s*[（(]?)([\d一二三四五六七八九十百]+)([）)、.．]\s*)/;
 const sample=tokens.map(t=>t.text.match(prefixPattern)).find(Boolean);
 const numeral=(n,chinese)=>!chinese?String(n):n<10?'一二三四五六七八九'[n-1]:n===10?'十':n<20?'十'+'一二三四五六七八九'[n-11]:n<100?'一二三四五六七八九'[Math.floor(n/10)-1]+'十'+(n%10?'一二三四五六七八九'[n%10-1]:''):String(n);
 for(const [i,item]of context.items.entries()){
  const key=strategyName(item.title);if(kept.has(key))continue;
  let pos=-1;for(let p=i-1;p>=0;p--){const found=nodes.findIndex(n=>n.key===strategyName(context.items[p].title));if(found>=0){pos=found+1;break;}}
  if(pos<0){pos=nodes.findIndex(n=>context.items.slice(i+1).some(next=>n.key===strategyName(next.title)));if(pos<0)pos=nodes.length;}
  const style=sample?sample[1]+numeral(pos+1,!/^\d+$/.test(sample[2]))+sample[3]:'';
  nodes.splice(pos,0,{text:style+item.title,key,separator:pos?nodes[pos]?.separator||nodes[pos-1]?.separator||usualSeparator:''});
  kept.add(key);addedTitles.push(item.title);
 }
 if(addedTitles.length||removedTitles.length)nodes.forEach((node,i)=>{node.text=node.text.replace(prefixPattern,(_,a,b,c)=>a+numeral(i+1,!/^\d+$/.test(b))+c);});
 const text=addedTitles.length||removedTitles.length?nodes.map((n,i)=>(i?n.separator||usualSeparator:'')+n.text).join('')+trailing:original;
 const warnings=removedTitles.length?['拟采用标题摘要将移除合同对应层级不存在的旧标题：'+removedTitles.join('、')+'。请在预览中确认。']:[];
 return {text,lines:text.split(/\r?\n/),addedTitles,removedTitles,warnings};
}

export function resolveStrategySummary(contract,text,scope={}){
 const titles=shortStrategyTitles(text);if(!titles||!contract)return null;
 const context=strategyContext(contract,titles,scope);if(!context)return null;
 const rows=strategyRows(context,titles),missing=rows.items.filter(i=>i.status==='missing').map(i=>i.title);
 const status=context.conflict?'conflict':context.ambiguous||missing.length||rows.extra.length?'pending':'verified';
 const reason=status==='verified'?`合同对应${rows.total}个策略标题，模板采用标题摘要，原文字保持不变。`:
 context.conflict||context.ambiguous?'投资策略存在不同章节版本、标题层级或适用范围歧义，请核对完整依据。':
 [missing.length?'模板未呈现：'+missing.join('、'):'',rows.extra.length?'合同未对应：'+rows.extra.map(i=>i.title).join('、'):''].filter(Boolean).join('；')+'。请核对标题增减，不自动扩写正文。';
 return {status,value:status==='verified'?text:null,evidence:context.evidence,reason,reasonType:status==='verified'?'confirmed':'needs_review'};
}
export function resolveStrategyReplacement(contract,text,scope={}){
 if(!contract)return undefined;
 const titles=shortStrategyTitles(text);if(!titles)return undefined;
 const context=strategyContext(contract,titles,scope);if(!context)return undefined;
 return strategyTitleReplacement(context,text);
}
function strategySectionReview(task,facts,scope){
 const targets=(task.materials||[]).map(m=>({m,fields:(m.fields||[]).filter(f=>fieldMatches(task,f,'investment_strategy',scope))})).filter(t=>t.fields.length);
 for(const target of targets){target.source=unique(target.fields.map(f=>String(f.after??f.before??''))).join('\n');target.titles=shortStrategyTitles(target.source);}
 const titles=targets.flatMap(t=>t.titles||[]);if(!titles.length)return null;
 const context=strategyContext(task.contract,titles,scope);if(!context)return null;
 if(facts.some(f=>!sameScope(factScope(task,f),scope)))context.ambiguous=true;
 const materials=targets.map(({m,titles,source,fields})=>{
  if(!titles)return materialReview(task,m,'investment_strategy',scope,context.items);
  const ownContext=strategyContext(task.contract,titles,scope),replacement=!context.ambiguous&&!context.conflict&&ownContext?strategyTitleReplacement(ownContext,source):undefined;
  return {materialId:m.materialId,name:m.template?.name||m.name||m.materialId,fieldIds:fields.map(f=>f.id),...strategyRows(context,titles),...(replacement?{replacement}:{})};
 }).filter(Boolean);
 if(!materials.length)return null;
 const statuses=materials.flatMap(m=>[...m.items,...m.extra]).map(i=>i.status),missing=unique(materials.flatMap(m=>m.items.filter(i=>i.status==='missing').map(i=>i.title)));
 const kind=context.conflict||context.ambiguous||statuses.includes('uncertain')?'uncertain':statuses.includes('missing')?'structure':'aligned';
 const summary=context.conflict||context.ambiguous?'投资策略存在不同章节版本、标题层级或适用范围歧义，请核对完整依据。':missing.length?'策略标题摘要未呈现：'+missing.join('、')+'。请核对是否补充；不会自动扩写正文。':statuses.includes('uncertain')?'模板包含合同对应层级未能确认的策略标题，请核对。':`合同对应${context.items.length}个策略标题，当前模板标题摘要已对应；保留原话术及格式。`;
 return {id:'section-investment_strategy-'+JSON.stringify({share:scope.share||'',channel:scope.channel||'',issuer:scope.issuer||'',period:scope.period||'',condition:scope.condition||''}),key:'investment_strategy',label:'投资策略',evidence:context.evidence,materials,notes:context.notes,kind,needsReview:kind!=='aligned',summary:summary+(context.notes.length?`另有${context.notes.length}条条件性说明，已单列并保留完整依据。`:'')};
}

export function getSectionReview(task,factIds){
 const ids=new Set(factIds||[]),facts=(task?.facts||[]).filter(f=>ids.has(f.id)&&LABELS[f.key]);
 if(!facts.length||!task.contract||new Set(facts.map(f=>f.key)).size!==1)return null;
 const key=facts[0].key,label=LABELS[key],scope=factScope(task,facts[0]);
 if(key==='investment_strategy'){const strategy=strategySectionReview(task,facts,scope);if(strategy)return strategy;}
 let sections=chapters(task.contract).filter(s=>s.key===key);if(!sections.length)return null;
 const scopeUncertain=!!scope.condition||facts.some(f=>!sameScope(factScope(task,f),scope));
 const explicitMismatch=s=>['share','channel'].some(k=>s.scope[k]&&scope[k]&&s.scope[k]!==scope[k]);
 sections=sections.filter(s=>!explicitMismatch(s));if(!sections.length)return null;
 const ambiguousScope=scopeUncertain||sections.some(s=>['share','channel'].some(k=>s.scope[k]&&!scope[k]));
 const ambiguous=ambiguousScope||sections.length>1;
 const primary=sections.find(s=>!s.isAttachment)||sections[0];
 const items=key==='investment_restrictions'?primary.items:ambiguous?[...new Map(sections.flatMap(s=>s.items).map(item=>[canonical(item.text),item])).values()]:sections[0].items;
 const materials=(task.materials||[]).map(m=>materialReview(task,m,key,scope,items)).filter(Boolean);
 if(!materials.length)return null;
 if(ambiguous)for(const material of materials)for(const item of material.items){if(item.status!=='missing'){item.status='uncertain';item.reason='存在不同合同章节或适用范围，列出各版本供核对，未选择其中一个作为更新依据。';}}
 const statuses=materials.flatMap(m=>[...m.items,...m.extra]).map(i=>i.status);
 const kind=ambiguous||statuses.includes('uncertain')?'uncertain':statuses.some(s=>s==='structure'||s==='missing')?'structure':statuses.includes('local')?'local':'aligned';
 const missing=unique(materials.flatMap(m=>m.items.filter(i=>i.status==='missing').map(i=>i.title)));
 const needsReview=kind==='structure'||kind==='uncertain';
 const preferCompleteEvidence=key==='investment_restrictions'&&materials.some(m=>m.total>=3&&m.items.filter(i=>['missing','structure','uncertain'].includes(i.status)).length>=Math.ceil(m.total/2));
 const summary=preferCompleteEvidence?`差异较大，建议使用完整合同依据 · 共${primary.items.length}项`+(primary.notes.length?`＋${primary.notes.length}段补充说明`:'')+(sections.length>1?'；另有其他版本依据，已分别保留供展开核对。':'。'):ambiguous?`${label}存在多个不同章节或适用范围尚不明确，请对照完整依据确认，未自动选择或合并。`:missing.length?`${label}完整依据共${items.length}项；部分物料未呈现：${missing.join('、')}。可能是摘要，请逐份核对是否补充。`:kind==='structure'?`${label}发现计量口径或条件差异，请查看逐项对照。`:kind==='uncertain'?`${label}存在尚不能确认对应关系的条目，请查看完整依据。`:statuses.includes('summary')?`${label}完整依据共${items.length}项，模板包含摘要表述；类别对应不代表已核实全部数值及条件，可展开完整依据核对。`:`${label}完整依据共${items.length}项，当前物料均已对应；保留原有局部更新规则。`;
 const replacement=key==='investment_ratios'&&!ambiguous?{lines:sections[0].content.map(line=>line.text)}:undefined;
 return {id:'section-'+key+'-'+JSON.stringify({share:scope.share||'',channel:scope.channel||'',issuer:scope.issuer||'',period:scope.period||'',condition:scope.condition||''}),key,label,evidence:sections.map(s=>({...s.evidence,blockIds:[...s.evidence.blockIds],...(key==='investment_restrictions'?{variantRole:s===primary?'primary':'other_version'}:{})})),materials,kind,needsReview,summary,...(replacement?{replacement}:{}),...(preferCompleteEvidence?{preferCompleteEvidence:true}:{})};
}
