// Holding-period redemption rates are a separate evidence path. A bare rate,
// waiver or allocation percentage never becomes a tier. Month/year intervals
// keep their own units; no conversion to days (or between months and years).
const norm=s=>String(s??'').normalize('NFKC').replace(/\s+/g,'').trim();
const percent=/\d+(?:\.\d+)?\s*[%％]/g;
const feeLabel=/(?:退出|赎回)费(?:费率|率)?/;
const restricted=/临时|优惠|促销|特定投资者|部分投资者|特殊投资者|减免|以.{0,12}公告.{0,8}为准/;
const reference=/底层(?:产品|基金|资管|资产)|标的(?:产品|基金)|目标(?:产品|基金)|举例|例如|比如|假设|计算示例|历史产品|同策略产品/;
const extraApplicability=/违约|提前退出|强制退出|仅适用|仅限|特殊投资者|特定投资者|通过.{0,16}(?:银行|渠道)|额外条件|同时满足|且(?:需|须|满足)/;
const unitPattern='(自然日|交易日|工作日|天|日|个自然月|自然月|个月|月|年)?';
const numberPattern='(\\d+(?:\\.\\d+)?)';
const comparison='(<=|>=|<|>|≤|≥|=)';
const dimensions=['share','channel','issuer','period'];
const pending=(reason,evidence=[])=>({status:'pending',value:null,evidence,reason,reasonType:'needs_review'});

function scopeAt(text,position=String(text).length){
 const before=String(text).slice(0,position).normalize('NFKC'),scope={};
 const share=[...before.matchAll(/([A-Z])类(?:份额)?/g)].at(-1);if(share)scope.share=share[1];
 const actor=[...before.matchAll(/(代销机构|销售机构|销售银行|管理人|受托人)(?=[^，,。；;\n]{0,14}(?:收取|费用|费率|退出费|赎回费))/g)].at(-1);if(actor)scope.issuer=/代销|销售/.test(actor[1])?'distributor':'manager';
 const channels=[...before.matchAll(/(?:渠道[:：]\s*([^，,。；;:\n]{1,24})|(?:^|[，,。；;\n])\s*([\u4e00-\u9fffA-Za-z0-9]{2,20})渠道)/g)].at(-1);if(channels)scope.channel=(channels[1]||channels[2]).trim();
 return scope;
}
function heading(text){
 const raw=String(text).trim(),n=raw.normalize('NFKC'),m=n.match(/^(第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十百]+、|\([一二三四五六七八九十百]+\)|\d+[、.])\s*(.*)$/);
 const title=m?.[2]||n,rank=m?/^第/.test(m[1])?1:/^[一二三四五六七八九十百]+、/.test(m[1])?2:/^\(/.test(m[1])?3:4:0;
 if(title.length>90||/[%％。；;]/.test(title))return null;
 const scope=scopeAt(title),isScope=/^[A-Z]类(?:份额)?$/.test(title)||/^渠道[:：]/.test(title)||/渠道$/.test(title)||/^(?:代销机构|销售机构|销售银行|管理人|受托人)(?:的)?(?:费用|费率|退出费率|赎回费率)$/.test(title)||!!m&&Object.keys(scope).length>0;
 const excluded=reference.test(title)||restricted.test(title);
 const own=/^(?:本集合计划|本计划|本信托计划|本产品)(?:的)?(?:费用|费率|概况|产品要素|赎回费用|退出费用|基本信息)$/.test(title);
 return rank||isScope||excluded||own?{rank,scope:isScope?scope:{},excluded,own}:null;
}
function blockContexts(blocks){
 const contexts=[],result=new Map();
 for(const block of blocks){const h=!block.tableId?heading(block.text):null;
  if(h){
   if(h.rank)for(let i=contexts.length-1;i>=0;i--)if(contexts[i].rank>=h.rank&&contexts[i].rank)contexts.splice(i,1);
   if(h.own)for(let i=contexts.length-1;i>=0;i--)if(contexts[i].excluded)contexts.splice(i,1);
   for(const dim of Object.keys(h.scope))for(let i=contexts.length-1;i>=0;i--)if(contexts[i].scope[dim]){const next={...contexts[i].scope};delete next[dim];contexts[i]={...contexts[i],scope:next};}
   if(h.excluded||Object.keys(h.scope).length)contexts.push(h);
  }
  result.set(block,{scope:Object.assign({},...contexts.map(c=>c.scope)),excluded:contexts.some(c=>c.excluded)});
 }
 return result;
}
function clauseAt(text,start,length){
 const left=text.slice(0,start),a=Math.max(...['。','；',';','\n'].map(s=>left.lastIndexOf(s)))+1,tail=text.slice(start+length),next=tail.search(/[。；;\n]/),end=next<0?text.length:start+length+next;
 return{start:a,end,text:text.slice(a,end)};
}
function allocationAtom(text,start,length){return /^(?:全部|全额|均|的部分)?(?:归入|归属|计入|返还|分配给)/.test(norm(text.slice(start+length)));}
function exactRateAtom(text,start,length,clause){
 const left=text.slice(clause.start,start),labels=[...left.matchAll(new RegExp(feeLabel.source,'g'))],label=labels.at(-1);if(!label)return false;
 const bridge=norm(left.slice(label.index+label[0].length)),right=norm(text.slice(start+length,clause.end));
 if(/不超过|不高于|不低于|至少|至多|最高|最低|上限|下限|标准|原费率|基准|折|减免|优惠|约|超过|[%％×*]/.test(bridge))return false;
 if(/^(?:至|到|[-~～]|[×*]|的\d|起|及以上|以上|及以下|以下|左右)/.test(right))return false;
 return true;
}
function unitOf(units,source,defaultUnit){
 const values=units.filter(Boolean),kinds=new Set(values.map(u=>u==='交易日'?'trading_day':u==='工作日'?'working_day':/月/.test(u)?'calendar_month':u==='年'?'calendar_year':'calendar_day'));
 if(kinds.size>1)return{warning:'持有期限混合不同的日期单位，请核对适用口径。'};
 if(kinds.size)return{unit:[...kinds][0]};
 const header=norm(source).match(/持有(?:期限|时间|期)\((?:单位[:：])?(个自然月|自然月|个月|月|年)\)/);if(header)return{unit:header[1]==='年'?'calendar_year':'calendar_month'};
 if(/持有月数/.test(source))return{unit:'calendar_month'};if(/持有年数/.test(source))return{unit:'calendar_year'};
 if(/交易日/.test(source))return{unit:'trading_day'};
 if(/工作日/.test(source))return{unit:'working_day'};
 if(/自然日|持有天数/.test(source))return{unit:'calendar_day'};
 return defaultUnit?{unit:defaultUnit}:{warning:'持有期限未明确单位，请核对自然日、交易日、月或年。'};
}
function bound(range,operator,value,reverse=false){
 const op=reverse?({'<':'>','>':'<','≤':'≥','≥':'≤','<=':'>=','>=':'<=','=':'='})[operator]:operator,n=Number(value);
 if(['>','≥','>=','='].includes(op)){range.lower=n;range.lowerInclusive=op!=='>';}
 if(['<','≤','<=','='].includes(op)){range.upper=n;range.upperInclusive=op!=='<';}
}
function intervalKey(r){return`${r.unit}:${r.lowerInclusive?'[':'('}${r.lower??'-inf'},${r.upper??'+inf'}${r.upperInclusive?']':')'}`;}
function intervalLabel(r){
 const unit=({calendar_day:'自然日',trading_day:'交易日',working_day:'工作日',calendar_month:'个月',calendar_year:'年'})[r.unit],parts=[];
 if(r.lower!==null)parts.push(`${r.lower}${unit}${r.lowerInclusive?'≤':'＜'}`);
 parts.push(/month|year/.test(r.unit)?'持有期限':'持有天数');if(r.upper!==null)parts.push(`${r.upperInclusive?'≤':'＜'}${r.upper}${unit}`);
 return parts.join('');
}
function parseCondition(source,defaultUnit){
 const original=String(source),raw=original.normalize('NFKC').replace(/\s/g,''),dayHint=/持有|(?:^|[^A-Za-z])[DNT](?:[<>=≤≥]|$)|[<>=≤≥][DNT]|\d+(?:自然日|交易日|工作日|天|日|个月|月|年)/.test(raw);
 if(!dayHint&&!defaultUnit)return null;
 let text=raw.replace(/持有(?:天数|月数|年数|期限|时间|期)?(?:\(?[DNT]\)?)?/g,'D').replace(/不少于|不低于|不小于|至少/g,'≥').replace(/不超过|不高于|不大于/g,'≤').replace(/不足|不满|未满|少于|小于/g,'<').replace(/超过|高于|大于/g,'>').replace(/D满/g,'D≥').replace(/(^|[:：,，])满(?=\d)/g,'$1D≥');
 const marked=new RegExp('(?:[DNT](?:在)?)?'+numberPattern+unitPattern+'\\((不含|含)\\)(?:至|到|[-~～])'+numberPattern+unitPattern+'\\((不含|含)\\)','g');
 text=text.replace(marked,(_,a,au,ai,b,bu,bi)=>a+(au||'')+(ai==='含'?'≤':'<')+'D'+(bi==='含'?'≤':'<')+b+(bu||''));
 // A range contains the variable only once; match it before a one-sided bound.
 const chain=new RegExp(numberPattern+unitPattern+comparison+'[DNT]'+comparison+numberPattern+unitPattern),direct=new RegExp('[DNT]'+comparison+numberPattern+unitPattern),reverse=new RegExp(numberPattern+unitPattern+comparison+'[DNT]'),sequence=new RegExp('[DNT]'+comparison+numberPattern+unitPattern+'(?:但|且|并且|并|,|，|、)?(?:[DNT])?'+comparison+numberPattern+unitPattern);
 const candidates=[],occupied=[];
 for(const m of text.matchAll(new RegExp(sequence.source,'g'))){
  const lower=op=>['>','≥','>='].includes(op),upper=op=>['<','≤','<='].includes(op);if(!((lower(m[1])&&upper(m[4]))||(upper(m[1])&&lower(m[4]))))return{condition:'uncertain:'+norm(original),label:original.trim(),warning:'持有期限的多个界限无法明确组合为一个区间。'};
  const range={lower:null,lowerInclusive:false,upper:null,upperInclusive:false};bound(range,m[1],m[2]);bound(range,m[4],m[5]);candidates.push({range,units:[m[3],m[6]],match:m[0],index:m.index});occupied.push([m.index,m.index+m[0].length]);
 }
 for(const m of text.matchAll(new RegExp(chain.source,'g'))){const range={lower:null,lowerInclusive:false,upper:null,upperInclusive:false};bound(range,m[3],m[1],true);bound(range,m[4],m[5]);candidates.push({range,units:[m[2],m[6]],match:m[0],index:m.index});occupied.push([m.index,m.index+m[0].length]);}
 for(const [re,reversed]of [[direct,false],[reverse,true]])for(const m of text.matchAll(new RegExp(re.source,'g'))){if(occupied.some(([a,b])=>m.index>=a&&m.index<b))continue;const range={lower:null,lowerInclusive:false,upper:null,upperInclusive:false};bound(range,reversed?m[3]:m[1],reversed?m[1]:m[2],reversed);candidates.push({range,units:[reversed?m[2]:m[3]],match:m[0],index:m.index});}
 if(!candidates.length){
  const suffix=new RegExp('(?:[DNT])?'+numberPattern+unitPattern+'(及以上|以上|及以下|以下|以内)');
  for(const m of text.matchAll(new RegExp(suffix.source,'g'))){const range={lower:null,lowerInclusive:false,upper:null,upperInclusive:false};bound(range,/以上/.test(m[3])?'≥':'≤',m[1]);candidates.push({range,units:[m[2]],match:m[0],index:m.index});}
 }
 if(!candidates.length&&!/[DNT][<>=≤≥]|[<>=≤≥][DNT]|(?:持有|[DNT]).{0,20}(?:不足|不满|未满|少于|小于|不少于|不超过|至少|满|以上|以下|以内)|\d+(?:自然日|交易日|工作日|天|日|个月|月|年)?(?:至|到|[-~～])\d/.test(raw))return null;
 if(candidates.length!==1)return{condition:'uncertain:'+norm(original),label:original.trim(),warning:'持有天数条件无法唯一识别，请核对区间边界和适用条件。'};
 const remainder=text.slice(0,candidates[0].index)+text.slice(candidates[0].index+candidates[0].match.length);if(/[<>=≤≥]\d+(?:\.\d+)?(?:自然日|交易日|工作日|天|日|个月|月|年)|(?:或|或者).{0,12}[<>=≤≥]|(?:且|并且)D/.test(remainder))return{condition:'uncertain:'+norm(original),label:original.trim(),warning:'持有期限还有无法对应的区间界限或其他持有条件，不能仅采用其中一端。'};
 const item=candidates[0],units=unitOf(item.units,original,defaultUnit);
 if(units.warning)return{condition:'uncertain:'+norm(original),label:original.trim(),warning:units.warning};
 const range={...item.range,unit:units.unit};
 if(range.lower!==null&&range.upper!==null&&(range.lower>range.upper||range.lower===range.upper&&!(range.lowerInclusive&&range.upperInclusive)))return{condition:'uncertain:'+norm(original),label:original.trim(),warning:'持有天数区间为空或边界顺序有误。'};
 return{condition:intervalKey(range),label:intervalLabel(range),range};
}

export function matchTemplateRedemptionTier(text,atom){
 const raw=String(text??''),start=atom?.start,before=String(atom?.before??'');if(!Number.isInteger(start)||start<0||raw.slice(start,start+before.length)!==before)return null;
 if(allocationAtom(raw,start,before.length))return null;
 const clause=clauseAt(raw,start,before.length);
 const condition=parseCondition(clause.text);if(!condition)return null;
 const scope=scopeAt(raw,start),rateWarning=feeLabel.test(clause.text)&&!exactRateAtom(raw,start,before.length,clause)?'退出费率是上限、折扣或其他非确定数值，不能作为固定分档费率。':null;
 return{...condition,scope,...(extraApplicability.test(clause.text)?{warning:'退出费率含额外适用条件，不能仅按持有天数对应。'}:rateWarning?{warning:rateWarning}:{})};
}
function evidence(blocks,quote){return{blockId:blocks[0].id,blockIds:blocks.map(b=>b.id),quote};}
function mergeEvidence(evidence){const seen=new Set();return evidence.filter(e=>{const key=JSON.stringify([e.blockIds||[e.blockId],e.quote]);if(seen.has(key))return false;seen.add(key);return true;});}
function addEntry(index,tier,scope,value,evidence,excluded=false){
 if(!tier||tier.warning){if(tier)index.issues.push({tier,scope,evidence,excluded});return;}
 if(!/^\d+(?:\.\d+)?\s*[%％]$/.test(String(value).trim()))return;
 const rate=Number(String(value).replace(/[^\d.]/g,''));if(!Number.isFinite(rate)||rate<0||rate>100)return;
 const normalized=rate+'%',key=JSON.stringify([tier.condition,scope,normalized,excluded]);let entry=index.entries.find(e=>e.identity===key);
 if(entry)entry.evidence=mergeEvidence([...entry.evidence,...evidence]);else index.entries.push({identity:key,...tier,scope,value:normalized,evidence,excluded});
}
function addUniformPolicy(index,block,context){
 const subject='(?:本集合计划|本信托计划|本计划|本产品)(?:的)?',none=new RegExp('^'+subject+'(?:统一|无条件)?(?:不收取|不收|免收)(?:退出|赎回)费(?:用)?$'),uniform=new RegExp('^'+subject+'(?:退出|赎回)费率(?:统一|均)(?:为|是|[:：])?(\\d+(?:\\.\\d+)?%)$');
 for(const part of block.text.split(/[。；;\n]/)){
  const text=norm(part).replace(/^(?:\d+[、.]|[一二三四五六七八九十]+、)/,''),match=text.match(uniform),value=none.test(text)?'0%':match&&Number(match[1].replace('%',''))<=100?Number(match[1].replace('%',''))+'%':null;
  const scope={...context.scope,...scopeAt(part)},excluded=context.excluded||restricted.test(block.text)||reference.test(block.text)||extraApplicability.test(part);
  if(value)index.uniform.push({value,scope,excluded,evidence:[evidence([block],block.text)]});
  else if(new RegExp('^'+subject+'(?:退出|赎回)费率(?:统一|均)').test(text))index.uniformIssues.push({scope,excluded,evidence:[evidence([block],block.text)]});
 }
}
export function buildRedemptionTierIndex(contract){
 const blocks=(contract?.blocks||[]).filter(b=>b.text?.trim()),index={entries:[],issues:[],uniform:[],uniformIssues:[]},contexts=blockContexts(blocks),tables=new Map();
 for(const block of blocks){
  const context=contexts.get(block);
  if(block.tableId){const key=(block.part||'')+'|'+block.tableId;if(!tables.has(key))tables.set(key,[]);tables.get(key).push(block);continue;}
  addUniformPolicy(index,block,context);
  for(const m of block.text.matchAll(percent)){
   const clause=clauseAt(block.text,m.index,m[0].length);if(!feeLabel.test(clause.text)||allocationAtom(block.text,m.index,m[0].length))continue;
   let tier=matchTemplateRedemptionTier(block.text,{start:m.index,before:m[0]});if(!tier)continue;
   if(!exactRateAtom(block.text,m.index,m[0].length,clause))tier={...tier,warning:'退出费率是上限、折扣或其他非确定数值，不能作为固定分档费率。'};
   const scope={...context.scope,...tier.scope},prefix=block.text.slice(0,clause.end),excluded=context.excluded||restricted.test(prefix)||reference.test(prefix);
   addEntry(index,tier,scope,m[0],[evidence([block],block.text)],excluded);
  }
  for(const m of block.text.matchAll(/(?:不收取|不收|免收)(?:退出|赎回)费(?:用)?/g)){
   const clause=clauseAt(block.text,m.index,m[0].length),tier=parseCondition(clause.text);if(!tier)continue;
   const prefix=block.text.slice(0,clause.end),scope={...context.scope,...scopeAt(block.text,m.index)},qualified=extraApplicability.test(clause.text)?{...tier,warning:'免收退出费用条款含额外适用条件，需核对完整口径。'}:tier;
   addEntry(index,qualified,scope,'0%',[evidence([block],block.text)],context.excluded||restricted.test(prefix)||reference.test(prefix));
  }
 }
 for(const table of tables.values()){
  const rows=new Map();for(const block of table){if(!Number.isInteger(block.rowIndex)||!Number.isInteger(block.cellIndex))continue;const cells=rows.get(block.rowIndex)||new Map(),bs=cells.get(block.cellIndex)||[];bs.push(block);cells.set(block.cellIndex,bs);rows.set(block.rowIndex,cells);}
  let holdingColumn=null,feeColumns=[],headerUnit,tableScope={},tableExcluded=false;
  for(const [,cells]of [...rows].sort((a,b)=>a[0]-b[0])){
   const entries=[...cells].map(([column,bs])=>({column,blocks:bs,text:bs.map(b=>b.text).join('\n')}));
   if(!entries.some(e=>/[%％]/.test(e.text)))for(const cell of entries){const h=heading(cell.text);if(h){if(h.own)tableExcluded=false;if(h.excluded)tableExcluded=true;tableScope={...tableScope,...h.scope};}}
   const rowScopes=entries.map(e=>scopeAt(e.text)),ambiguousScope=dimensions.some(k=>new Set(rowScopes.map(s=>s[k]).filter(Boolean)).size>1);
   if(!ambiguousScope)for(const cell of entries)for(const block of cell.blocks)addUniformPolicy(index,block,{scope:{...contexts.get(block).scope,...tableScope,...Object.assign({},...rowScopes)},excluded:contexts.get(block).excluded||tableExcluded});
   const holding=entries.find(e=>/持有(?:天数|月数|年数|期限|时间|期)/.test(e.text)&&!/[<>=≤≥]|不足|不满|少于|不少于/.test(e.text));
   const fees=entries.filter(e=>feeLabel.test(e.text)&&!/%|％/.test(e.text));
   if(holding&&fees.length){holdingColumn=holding.column;feeColumns=fees;headerUnit=unitOf([],holding.text).unit;continue;}
   if(holdingColumn===null||!feeColumns.length)continue;
   const condition=entries.find(e=>e.column===holdingColumn);if(!condition)continue;const tier=parseCondition(condition.text,headerUnit);if(!tier)continue;
   for(const column of feeColumns){const rate=entries.find(e=>e.column===column.column);if(!rate)continue;const value=/^(?:无|不收取|不收|免收)(?:(?:退出|赎回)费(?:用)?)?[。.]?$/.test(norm(rate.text))?'0%':rate.text,context=contexts.get(condition.blocks[0]),scope={...context.scope,...tableScope,...scopeAt(column.text),...scopeAt(condition.text)},quote=column.text+'\n'+condition.text+'\n'+rate.text,qualified=extraApplicability.test(condition.text+' '+column.text)?{...tier,warning:'退出费率表含额外适用条件，需核对完整口径。'}:!/^\d+(?:\.\d+)?\s*[%％]$/.test(value.trim())?{...tier,warning:'此档退出费率表未给出可确定的单个费率。'}:tier;
    addEntry(index,qualified,scope,value,[evidence([...condition.blocks,...rate.blocks],quote)],context.excluded||tableExcluded||restricted.test(condition.text+' '+column.text)||reference.test(condition.text+' '+column.text));
   }
  }
 }
 return index;
}
function overlaps(a,b){
 if(a.unit!==b.unit)return false;
 for(const [left,right]of [[a,b],[b,a]])if(left.upper!==null&&right.lower!==null&&(left.upper<right.lower||left.upper===right.lower&&!(left.upperInclusive&&right.lowerInclusive)))return false;
 return true;
}
function scopeCompatible(entry,scope){return dimensions.every(k=>!entry[k]||!scope[k]||norm(entry[k])===norm(scope[k]));}
export function resolveRedemptionTier(index,tier,scope={}){
 if(!tier)return pending('模板未明确持有天数分档。');
 if(tier.warning||!tier.range)return pending(tier.warning||'持有天数分档条件尚不明确。');
 if(dimensions.some(k=>scope[k]&&tier.scope?.[k]&&norm(scope[k])!==norm(tier.scope[k])))return pending('模板中的份额或渠道与当前核对口径不一致。');
 const effective={...scope,...tier.scope},entries=(index?.entries||[]).filter(e=>!e.excluded&&scopeCompatible(e.scope,effective)),related=entries.filter(e=>overlaps(e.range,tier.range));
 if(entries.some(e=>e.range.unit!==tier.range.unit))return pending('合同与模板或不同合同条款使用了不同的持有期限单位，不能自动换算或混用。',mergeEvidence(entries.flatMap(e=>e.evidence)));
 const issues=(index?.issues||[]).filter(i=>!i.excluded&&scopeCompatible(i.scope,effective)&&(!i.tier.range||overlaps(i.tier.range,tier.range)));if(issues.length)return pending('合同还有单位或适用条件不明确的退出费率条款，请核对完整口径。',mergeEvidence(issues.flatMap(i=>i.evidence)));
 if(related.some(e=>dimensions.some(k=>e.scope[k]&&!effective[k])))return pending('合同退出费率存在不同份额或渠道口径，模板未明确适用对象。',mergeEvidence(related.flatMap(e=>e.evidence)));
 const exact=entries.filter(e=>e.condition===tier.condition),evidence=mergeEvidence((exact.length?[...exact,...related]:related).flatMap(e=>e.evidence));
 if(!entries.length){
  const policies=(index?.uniform||[]).filter(p=>!p.excluded&&scopeCompatible(p.scope,effective)),uncertain=(index?.uniformIssues||[]).filter(p=>!p.excluded&&scopeCompatible(p.scope,effective)),policyEvidence=mergeEvidence([...policies,...uncertain].flatMap(p=>p.evidence));
  if(uncertain.length)return pending('合同统一退出费政策未明确可直接采用的固定费率。',policyEvidence);
  if(policies.some(p=>dimensions.some(k=>p.scope[k]&&!effective[k])))return pending('合同统一退出费政策限定份额或收费主体，模板未明确适用口径。',policyEvidence);
  const values=new Set(policies.map(p=>p.value));if(values.size>1)return{status:'conflict',value:null,evidence:policyEvidence,reason:'同一适用口径存在不同的统一退出费政策。',reasonType:'needs_review'};
  if(values.size===1)return{status:'verified',value:[...values][0],evidence:policyEvidence,reason:'合同明确本产品统一退出费政策，且未发现适用的持有天数分档条款；统一费率适用于模板各持有区间。',reasonType:'confirmed'};
 }
 if(!exact.length)return pending('合同未找到与“'+tier.label+'”完全对应的退出费率；保留模板天数和原值待核。',evidence);
 const values=new Set(exact.map(e=>e.value));
 if(values.size!==1||related.some(e=>!values.has(e.value)))return{status:'conflict',value:null,evidence,reason:'同一持有天数区间或与之重叠的区间存在不同退出费率，请核对适用条款。',reasonType:'needs_review'};
 return{status:'verified',value:[...values][0],evidence:mergeEvidence(exact.flatMap(e=>e.evidence)),reason:'合同按相同持有天数区间明确了退出费率；仅核对费率，不改变模板天数及表述。',reasonType:'confirmed'};
}
