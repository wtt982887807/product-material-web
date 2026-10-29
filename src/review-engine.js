// Local evidence rules. No network calls and no product-specific answer tables.
export {getSectionReview} from './review-sections.js?v=20261026-opening-strategy';
import {resolveStrategySummary} from './contract-sections.js?v=20261026-opening-strategy';
import {previewSectionAdoption,adoptSection,refreshSectionOperations,releaseSection} from './section-adoption.js?v=20261026-opening-strategy';
import {NS,elements} from './office-xml.js';
import {parseDateInput,formatDateLike,parsePerformanceInput,parseOpeningInput} from './supplement-parser.js?v=20261026-rows';
import {monthlyWeekdayPatterns,parseMonthlyWeekdayInput,monthlyWeekdayChanges} from './monthly-weekday.js?v=20261026-monthly';
import {resolveMonthlyDateRule} from './opening-rules.js?v=20261026-opening-strategy';
export {getOpeningRuleReview} from './opening-rules.js?v=20261026-opening-strategy';
import {getTradingDays} from './trading-calendar.js?v=20261026-openings';
import {buildRedemptionTierIndex,matchTemplateRedemptionTier,resolveRedemptionTier} from './redemption-tiers.js?v=20261027-custom-investment';
import {previewFeeText,applyFeeText,refreshFeeTextOperations,releaseFeeText,assertFeeActionScope} from './fee-text-supplement.js?v=20261027-custom-investment';
import {simplifyEqualFeeTiers} from './equal-fee-tiers.js?v=20261027-custom-investment';
import {previewCustomSection as previewCustomSectionText,applyCustomSectionOverride,refreshCustomSectionOperations,refreshCustomSectionFields,releaseCustomSection} from './custom-section.js?v=20261027-custom-investment';
const PRODUCT_END='(?:集合资金信托计划|集合资产信托计划|集合资产管理计划|单一资产管理计划|资产管理计划|证券投资基金|信托计划)';
const NAME_LABEL=/^(?:产品名称|产品全称|信托名称|信托计划名称|基金名称|计划名称|资产管理计划名称|产品简称)$/;
const MARKETING=/历史业绩|产品表现|累计增长|最新净值|成立以来|年化收益|产品举例|同策略|代表产品|获奖|金牛奖/;
const ELEMENT=/产品要素|基本要素|要素表|主要条款|产品条款|产品概要/;
const UNDERLYING=/底层|投资标的|标的产品/;
const PERCENT=/\d+(?:\.\d+)?\s*[】\]]?\s*[%％]/g;
const norm=s=>String(s??'').normalize('NFKC').replace(/[\s·•,，。;；:：、【】\[\]"“”'‘’]/g,'');
const clauseNorm=s=>norm(String(s).replace(/^\s*[（(]?\d+[）).、]\s*/,'').replace(/[.。；;]+\s*$/,''));
function equivalentClause(s){return clauseNorm(s).replace(/集合资产管理计划|本集合计划|该计划|本计划/g,'本计划').replace(/本计划可参与债券正回购业务/g,'').replace(/债券正回购融入资金余额|债券正回购余额/g,'债券正回购融入资金余额').replace(/计划资产净值/g,'本计划资产净值').replace(/本本计划/g,'本计划');}
function nearbyEvidence(original,rs){const grams=s=>new Set([...s].slice(1).map((x,i)=>s[i]+x));const a=grams(clauseNorm(original));return rs.filter(b=>!b.joined).map(b=>{const z=grams(clauseNorm(b.text));const common=[...a].filter(x=>z.has(x)).length;return{b,score:common/Math.max(a.size,z.size,1)};}).filter(x=>x.score>.32).sort((x,y)=>y.score-x.score).slice(0,3).map(x=>quote(x.b.base||x.b,x.b.text));}
const labelNorm=s=>norm(s).replace(/[()（）].*?[()（）]/g,'').replace(/[:：*＊]/g,'');
const unique=a=>[...new Set(a)];
const esc=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const number=s=>Number(String(s).replace(/[^\d.\-]/g,''));
const equivalent=(a,b)=>norm(a)===norm(b)||(/^\d+(?:\.\d+)?[%％]?$/.test(norm(a))&&/^\d+(?:\.\d+)?[%％]?$/.test(norm(b))&&number(a)===number(b));
const cnum=s=>({一:1,二:2,两:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10}[s]??Number(s));
const quote=(b,text=b.text)=>({blockId:b.id,quote:text.trim().slice(0,1600)});
function keyFor(label){
 const l=labelNorm(label);
 if(/^(?:业绩比较基准|业绩参考基准|业绩基准|参考业绩基准|比较基准)$/.test(l))return 'benchmark';
 if(l.length>32||/实力|简介|数据来源|奖项|股东/.test(l))return `custom:${l}`;
 if(/(?:开售|发售|募集)(?:日期|日|期)/.test(l))return 'opening';
 const mappings=[['product_name',NAME_LABEL],['risk_rating',/风险(?:等级|评级)|产品评级/],['custodian',/托管人|保管人|托管行|保管行|集合托管人/],['manager',/管理人|受托人/],['advisor',/投资顾问/],['distributor',/销售机构|代销机构/],['product_code',/产品代码|产品编码/],['product_type',/产品类型|产品类别/],['performance',/业绩报酬|浮动.*费|浮动.*报酬/],['fees',/产品费率|管理、托管|费用/],['management_fee',/管理费|固定信托报酬/],['custody_fee',/托管费|保管费/],['subscription_fee',/认.*购.*费|参与费/],['redemption_fee',/赎回费|退出费/],['opening',/开放|预约/],['minimum_amount',/起购|认购规则|参与金额|首参|起点|认购金额/],['duration',/存续期|存续期限|产品期限/],['lockup',/封闭期|锁定期|持有期/],['warning_stop',/预警|止损/],['investment_objective',/投资目标/],['investment_scope',/投资范围/],['investment_ratios',/投资比例/],['investment_restrictions',/投资限制/],['investment_strategy',/投资策略/]];
 return mappings.find(([,re])=>re.test(l))?.[0]||`custom:${l}`;
}
function cleanName(s){return s.replace(/^\s*(?:产品名称|产品全称|计划名称|基金名称|信托(?:计划)?名称|名称)\s*[：:|]?\s*/,'').replace(/^(?:指|即)[“"]?/,'').replace(/[。；;“”"].*$/,'').replace(/(?:资产管理合同|管理合同|信托合同|基金合同|合同)$/,'').trim();}
function identity(contract){
 const candidates=[];const add=(s,b,weight)=>{s=cleanName(s);if(s.length<5||s.length>110||!new RegExp(PRODUCT_END+'$').test(s)||/[。；;：:|]/.test(s)||/本(?:集合|资产管理|信托)?计划/.test(s))return;candidates.push({name:s,b,weight});};
 const bs=contract.blocks.filter(b=>b.text?.trim());
 bs.forEach((b,i)=>{
  const t=b.text.trim();const explicit=t.match(new RegExp('(?:产品名称|产品全称|计划名称|基金名称|信托(?:计划)?名称|名称)\\s*[：:|]\\s*([^。；;\\n]+?'+PRODUCT_END+')'));
  if(explicit)add(explicit[1],b,10);
  if(t.length<130&&new RegExp('^.{2,100}'+PRODUCT_END+'(?:资产管理合同|管理合同|信托合同|基金合同|合同)?$').test(t))add(t,b,i<40?6:3);
  if(i<40&&t.length<70&&bs[i+1]?.text.length<80&&new RegExp(PRODUCT_END).test(bs[i+1].text))add(t+bs[i+1].text.trim(),{...b,text:t+'\n'+bs[i+1].text},5);
 });
 const grouped=new Map();for(const c of candidates){const k=norm(c.name);const x=grouped.get(k)||{name:c.name,evidence:[],weight:0};x.evidence.push(quote(c.b));x.weight+=c.weight;grouped.set(k,x);}
 const ranked=[...grouped.values()].filter(c=>![...grouped.values()].some(other=>other!==c&&norm(other.name).endsWith(norm(c.name))&&other.weight>=5)).sort((a,b)=>b.weight-a.weight);const winner=ranked[0];
 // A split legal suffix is not itself a competing name; complete explicit names dominate.
 if(!winner)return{name:'',evidence:[],status:'pending',reason:'合同正文尚未识别到可靠产品名称。'};
 const conflicting=ranked.filter(c=>c!==winner&&c.weight>=10&&!norm(winner.name).endsWith(norm(c.name)));
 if(conflicting.length)return{name:'',evidence:winner.evidence,status:'conflict',reason:'合同存在多个明确产品名称，需核定主体。',candidates:ranked.map(c=>c.name)};
 return{name:winner.name,shortName:winner.name.replace(new RegExp(PRODUCT_END+'$'),''),evidence:winner.evidence,status:'verified'};
}
function records(contract){
 const bs=contract.blocks.filter(b=>b.text?.trim());const out=[];
 bs.forEach((b,i)=>{
  out.push({...b,text:b.text.replace(/\u00a0/g,' '),base:b});
  if(b.text.length<70&&!/[。；;%％]/.test(b.text)){const next=bs.slice(i+1,i+4);if(next.length)out.push({...b,text:[b.text,...next.map(x=>x.text)].join('\n'),base:b,joined:true});}
 });return out;
}
function scopeFor(label,text,position=0){
 const head=(label+' '+text.slice(0,position)).split(/[。；;，,]/).at(-1);
 return {issuer:/代销|销售机构|销售银行/.test(head)?'distributor':'manager',share:(label+' '+text).match(/([A-Z])类(?:份额)?/)?.[1]||'',channel:(label+' '+text).match(/(?:渠道|份额)[：:]\s*([^，。；;]{1,20})/)?.[1]||'',period:'contract',condition:''};
}
function pick(candidates,reason='自动核对未能确定此参数，请查看合同原文确认。'){
 const map=new Map();for(const c of candidates){const k=typeof c.value==='string'&&/^\d+(?:\.\d+)?%$/.test(c.value)?String(number(c.value))+'%':norm(c.value);if(!map.has(k))map.set(k,c);else map.get(k).evidence.push(...c.evidence);}
 const arr=[...map.values()];return arr.length===1?{status:'verified',reason:'合同正文提供明确依据。',...arr[0]}:{status:arr.length?'conflict':'pending',value:null,evidence:arr.flatMap(c=>c.evidence).slice(0,8),reasonType:'needs_review',reason:arr.length?'合同存在多个不同候选参数，需核对适用条件。':reason};
}
function findMatches(rs,re,transform=(m)=>m[1],filter=()=>true){const out=[];for(const b of rs){if(!filter(b))continue;for(const m of b.text.matchAll(new RegExp(re.source,re.flags.includes('g')?re.flags:re.flags+'g'))){const v=transform(m,b);if(v!==null&&v!==undefined&&v!=='')out.push({value:String(v),evidence:[quote(b.base||b,b.text)]});}}return out;}
function filtered(rs,scope){return rs.filter(b=>!/(?:假设示例|例如.{0,5}本产品|举例说明)/.test(b.text)&&(!scope.share||!/[A-Z]类/.test(b.text)||b.text.includes(scope.share+'类')));}
function monthlyOpeningContext(text,pattern){
 const before=text.slice(0,pattern.start),start=Math.max(...['。','；',';','\n'].map(c=>before.lastIndexOf(c)))+1,tail=text.slice(pattern.start+pattern.before.length),end=tail.search(/[。；;\n]/),right=end<0?tail:tail.slice(0,end),local=text.slice(start,pattern.start)+pattern.before+right;
 const action=right.search(/开放|办理(?:参与|退出|申购|赎回)|(?:参与|申购)(?:和|及)(?:退出|赎回)/),disclosure=right.search(/(?:披露|公布|公告).{0,12}(?:净值|报告)|(?:估值日|净值计算日)/);
 if(disclosure>=0&&(action<0||disclosure<action))return false;
 if(/开放|办理(?:参与|退出|申购|赎回)|(?:参与|申购)(?:和|及)(?:退出|赎回)/.test(local))return true;
 const previous=text.slice(0,start).trim().split(/[。；;\n]/).at(-1)||'';
 return /^(?:\d+[.、]\s*)?(?:开放日|开放期|开放安排|参与和退出|参与|申购|赎回|退出)(?:安排|日期)?\s*[:：]?\s*$/.test(previous);
}
const FEE_LABELS=[['trust_total_fee',/固定信托管理费|固定管理费/],['trust_manager_fee',/受托人信托报酬|(?:收取)?信托管理费|受托人报酬/],['floating_advisor_fee',/浮动投资顾问费/],['floating_sales_fee',/浮动销售服务费/],['advisor_fee',/固定投资顾问费|投资顾问费/],['sales_fee',/固定销售服务费|销售服务费/],['custody_fee',/托管费(?:率)?|保管费(?:率)?/],['management_fee',/管理费(?:率)?/],['redemption_fee',/赎回费(?:率)?|退出费(?:率)?/],['subscription_fee',/(?:认[、/／]?申购|认购|申购|参与)(?:(?:费(?:用|率)?)?\s*[/／、及和与]\s*(?:认购|申购|参与))*费(?:用|率)?/]];
const ACTION_FEE_LABEL=/(?:认[、/／]?申购|认购|申购|参与|赎回|退出)(?:(?:费(?:用|率)?)?\s*[/／、及和与]\s*(?:认购|申购|参与|赎回|退出))*费(?:用|率)?/g;
function feeActionSet(label){
 const source=String(label).replace(/认[、/／]?申购/g,'认购/申购'),actions=[...source.matchAll(/认购|申购|参与|赎回|退出/g)].map(m=>m[0]==='申购'?'参与':m[0]==='赎回'?'退出':m[0]);
 return ['认购','参与','退出'].filter(a=>actions.includes(a)).join('|');
}
function feeActionKey(actions,fallback=''){
 const list=actions.split('|');return list.includes('退出')?(list.length>1?'subscription_redemption_fee':'redemption_fee'):actions?'subscription_fee':fallback;
}
const AGGREGATE_LAYERS=/(?:两层|双层|上下两层).{0,8}(?:合计|总计|总费|综合)|(?:合计|总计).{0,5}(?:两层|双层)/;
function feeTokens(text,baseKey,trust,baseActions=''){
 const tokens=[];for(const m of text.matchAll(PERCENT)){
  const before=text.slice(Math.max(0,m.index-100),m.index);const hits=[];
  for(const [key,re] of FEE_LABELS)for(const a of before.matchAll(new RegExp(re.source,'g')))hits.push({key,index:a.index,length:a[0].length});
  for(const a of before.matchAll(ACTION_FEE_LABEL)){const feeActions=feeActionSet(a[0]);hits.push({key:feeActionKey(feeActions),index:a.index,length:a[0].length,feeActions});}
  // A longer semantic label owns its nested generic label (信托管理费 contains 管理费).
  const specific=hits.filter(h=>!hits.some(other=>other!==h&&other.index<=h.index&&other.index+other.length>=h.index+h.length&&other.length>h.length));
  specific.sort((a,b)=>b.index-a.index||b.length-a.length||Number(!!b.feeActions)-Number(!!a.feeActions));let key=specific[0]?.key||baseKey;
  const feeActions=specific[0]?.feeActions||(!specific.length?baseActions:'');
  if(key==='trust_total_fee'&&!trust)key='management_fee';
  if(trust&&key==='management_fee')key='trust_manager_fee';
  const raw=m[0];tokens.push({key,start:m.index,before:raw,value:`${number(raw)}%`,...(feeActions?{feeActions}:{})});
 }return tokens;
}
function resolveFee(key,rs,scope,trust){
 if(['subscription_fee','redemption_fee','subscription_redemption_fee'].includes(key)){
  const required=(scope.feeActions||(key==='redemption_fee'?'退出':key==='subscription_redemption_fee'?'参与|退出':'')).split('|').filter(Boolean);
  if(required.length){
   const results=required.map(action=>{const candidates=[];for(const b of filtered(rs,scope)){
    for(const t of feeTokens(b.text,'',trust))if(t.feeActions?.split('|').includes(action))candidates.push({value:t.value,evidence:[quote(b.base||b,b.text)]});
    for(const a of b.text.matchAll(ACTION_FEE_LABEL))if(feeActionSet(a[0]).split('|').includes(action)&&/(?:不收取|不收|免收|免除)\s*$/.test(b.text.slice(Math.max(0,a.index-12),a.index)))candidates.push({value:'0%',evidence:[quote(b.base||b,b.text)]});
   }return pick(candidates);});
   if(results.length===1)return results[0];
   return results.every(r=>r.status==='verified'&&equivalent(r.value,results[0].value))?{...results[0],evidence:results.flatMap(r=>r.evidence)}:{status:'pending',value:null,evidence:results.flatMap(r=>r.evidence),reasonType:'needs_review',reason:required.join('、')+'费率需分别取得同口径依据，且数值一致，才能共用一个费率。'};
  }
 }
 const candidates=[];for(const b of filtered(rs,scope)){
  // Fee context is explicitly anchored locally; percentages from investment restrictions do not qualify.
  const tokens=feeTokens(b.text,'',trust);
  for(const t of tokens)if(t.key===key&&t.value!=='NaN%')candidates.push({value:t.value,evidence:[quote(b.base||b,b.text)]});
  const aliases=key==='subscription_fee'?/(?:不收取|不收|免收).{0,8}(?:认购|申购|参与)(?:[、/／](?:认购|申购|参与))?(?:费|费用)/:key==='redemption_fee'?/(?:不收取|不收|免收).{0,8}(?:赎回|退出)(?:费|费用)/:null;
  if(aliases?.test(b.text))candidates.push({value:'0%',evidence:[quote(b.base||b,b.text)]});
 }
 return pick(candidates);
}
function performanceTokens(text){
 const tokens=[];
 for(const m of text.matchAll(PERCENT)){
  const left=text.slice(Math.max(0,m.index-150),m.index),right=text.slice(m.index+m[0].length,m.index+m[0].length+40);
  const column=left.split('\n').at(-1),tablePayout=/[\t|]\s*$/.test(column)&&/[RＲF]|收益率/.test(column)&&/[＜≤<>＞≥]/.test(column);
  const payout=tablePayout||/(?:提取|计提|收取|提成|×|乘以|按)[【\s]*$/.test(left)||/^(?:的|比例的).{0,6}(?:业绩报酬|浮动服务费)|^(?:计提|提取)/.test(right);
  const hurdle=!payout&&(/(?:[＜≤<>＞≥=＝]|高于|低于|超过|超出|大于|小于|(?:基准|标准)(?:为|是|[:：])?|R[12](?:为|是)|收益率(?:为|是)?|年化)[【\s]*$/.test(left)||/^\s*[＜≤<>＞≥]/.test(right)||/^\s*(?:为|作为)\s*(?:(?:本期|本次|首期|当期)\s*)?(?:业绩报酬\s*)?(?:计提\s*)?(?:基准|标准)/.test(right)||/\d+(?:\.\d+)?%\s*[-—~～]\s*$/.test(left)||/^\s*[-—~～]\s*\d/.test(right)||/(?:[RＲF]|收益率)[-－]\s*$/.test(left));
  tokens.push({start:m.index,before:m[0],value:`${number(m[0])}%`,role:payout?'ratio':hurdle?'threshold':'unknown'});
 }
 return tokens;
}
function payoutTier(text,token,thresholds){
 let prefix=text.slice(0,token.start).replace(/\s+/g,'');
 // Formula multipliers belong to the immediately preceding return segment.
 const formula=prefix.match(/[（(]([^()（）]+)[）)]×[【]?$/);
 const segment=formula?.[1]||prefix.split(/[。；;，,]/).at(-1);
 const prior=prefix.slice(0,prefix.length-segment.length).split(/[。；;\n]/).at(-1);
 const context=segment||prior;
 const low=thresholds[0],high=thresholds.at(-1),two=thresholds.length>1;
 const symbolHigh=/R2/.test(context),symbolLow=/R1/.test(context);
 if(formula){
  if(symbolHigh&&symbolLow)return 1;
  if(symbolHigh&&/[RＲF]-R2/.test(context))return 2;
  const values=performanceTokens(context).filter(t=>t.role==='threshold'||context.includes('-')).map(t=>number(t.value));
  if(two&&values.includes(low)&&values.includes(high))return 1;
  if(two&&values.includes(high)&&/[RＲF]|收益率/.test(context))return 2;
  if(symbolLow||values.includes(low))return 1;
 }
 // A second extraction in the same sentence may refer back to the lower band.
 const local=prefix.split(/提取[^提取]*?%|计提[^计提]*?%/).at(-1);
 const band=local.split(/[。；;]/).at(-1);
 if(/R1.{0,20}(?:R|年化收益率).{0,8}[≤<＜].{0,2}R2|(?:不高于|不超过|≤|＜)R2/.test(band)&&/R1/.test(band))return 1;
 if(/(?:[RＲF]|收益率)[>＞]R2|(?:高于|超过|大于)R2/.test(band))return 2;
 const values=performanceTokens(band).filter(t=>t.role==='threshold').map(t=>number(t.value));
 if(two&&values.includes(low)&&values.includes(high)&&/(?:不高于|不超过|≤|＜|区间|[-—~～])/.test(band))return 1;
 if(two&&values.includes(high)&&/(?:高于|超过|超出|大于|＞|>)/.test(band))return 2;
 if(two&&values.includes(high)&&/(?:≤|＜)/.test(band))return 1;
 if(thresholds.length===1&&values.includes(low))return 1;
 return null;
}
function performanceData(rs){
 const relevant=rs.filter(b=>!b.joined&&/业绩报酬|浮动服务费|浮动信托管理费|年化收益率|R1|R2|[FＲR][≤＜<>＞]/.test(b.text));
 const thresholds=[],ratios=[];
 const all=relevant.flatMap(b=>performanceTokens(b.text).map(t=>({...t,b})));
 for(const t of all)if(t.role==='threshold')thresholds.push({value:t.value,evidence:[quote(t.b.base||t.b,t.b.text)]});
 const boundaries=unique(thresholds.map(t=>number(t.value))).sort((a,b)=>a-b);
 for(const t of all)if(t.role==='ratio'&&number(t.value)>0&&number(t.value)<100)ratios.push({value:t.value,tier:payoutTier(t.b.text,t,boundaries),evidence:[quote(t.b.base||t.b,t.b.text)]});
 const byValue=arr=>[...new Map(arr.map(a=>[a.value,a])).values()].sort((a,b)=>number(a.value)-number(b.value));
 return{thresholds:byValue(thresholds),ratios:byValue(ratios),tierRatios:ratios,symbolic:relevant.some(b=>/R1|R2|R[’′']/.test(b.text)),evidence:relevant.filter(b=>/R1|R2|R[’′']/.test(b.text)).slice(-3).map(b=>quote(b.base||b,b.text))};
}
const PERFORMANCE_POLICY_CACHE=new WeakMap();
const PERFORMANCE_CONDITIONAL=/暂不|目前|现阶段|首期|首次|首个|本期|当期|本次|该期|某期|特定|优惠|减免|免收|豁免|可不|如果|假设|例如|举例|若|当[^。；;]*时|不超过|不高于|低于|小于|[<≤＜≥>＞]|除非|另有约定|但|根据公告/;
const OTHER_PERFORMANCE_SUBJECT=/底层(?:基金|产品|资产管理计划|信托)?|投资标的|标的产品|其他产品|产品举例|假设|例如|举例/;
function performanceScope(label,text){
 const scope=scopeFor(label,text),subject=label+' '+text;
 if(/(?:投资顾问|顾问机构).{0,16}(?:业绩报酬|不收取|不提取|提取|收取)/.test(subject))scope.issuer='advisor';
 if(/(?:销售机构|代销机构|销售银行).{0,16}(?:业绩报酬|不收取|不提取|提取|收取)/.test(subject))scope.issuer='distributor';
 return scope;
}
function performanceDenial(text,heading=false){
 const t=String(text||'').trim();
 return /不(?:提取|收取|计提)(?:管理人)?(?:的)?业绩报酬|业绩报酬(?:为|是|[:：])?\s*(?:不提取|不收取|不计提|无)/.test(t)||heading&&/^(?:不提取|不收取|不计提|无)[。；;]?\s*$/.test(t);
}
function noPerformanceFee(text,heading=false){return !PERFORMANCE_CONDITIONAL.test(text)&&!OTHER_PERFORMANCE_SUBJECT.test(text)&&performanceDenial(text,heading);}
function chargesPerformanceFee(text,heading=false){
 const t=String(text||'');if(/如有|如适用|若有|可能涉及|可能产生/.test(t))return false;
 if(!/业绩报酬/.test(t))return heading&&/[RＲ]|收益率/.test(t)&&/[＜≤<>＞≥]/.test(t)&&[...t.matchAll(PERCENT)].some(m=>number(m[0])>0);
 // A payout only below/above a hurdle is still a charging policy. A noun such
 // as 计提标准 or 计提日, or a statement of no charge below a hurdle, is not.
 if([...(t.matchAll(/提取|收取|计提/g))].some(m=>!/(?:不|未|无须|无需|免于)$/.test(t.slice(Math.max(0,m.index-3),m.index))&&!/^(?:基准|标准|日|时点|比例|方式|金额|公式|起点|原则)/.test(t.slice(m.index+m[0].length))))return true;
 return heading&&/[RＲ]|收益率/.test(t)&&/[＜≤<>＞≥]/.test(t)&&[...t.matchAll(PERCENT)].some(m=>number(m[0])>0);
}
function contractPerformancePolicy(rs,scope,target){
 let cache=PERFORMANCE_POLICY_CACHE.get(rs);if(!cache){cache=new Map();PERFORMANCE_POLICY_CACHE.set(rs,cache);}const cacheKey=JSON.stringify(scope);if(cache.has(cacheKey))return cache.get(cacheKey);
 const none=[],charged=[],ambiguous=[],conditional=[],blocks=rs.filter(b=>!b.joined);let inherited={share:'',channel:''},other=false,feeHeading=false,headingText='',headingIndex=-2;
 for(const [index,b]of blocks.entries()){
  const text=b.text.trim(),bare=text.replace(/^\s*(?:第[一二三四五六七八九十百\d]+[章节]|[（(]?[一二三四五六七八九十百\d]+[）).、])\s*/,''),short=bare.length<65&&!/[。；;%％]/.test(bare);
  if(short&&/^(?:第[一二三四五六七八九十百\d]+[章节]|[一二三四五六七八九十百]+、)/.test(text)){inherited={share:'',channel:''};other=false;feeHeading=false;}
  if(short&&OTHER_PERFORMANCE_SUBJECT.test(text))other=true;
  if(short&&/^(?:本(?:集合|信托)?计划|本产品)/.test(bare))other=false;
  if(short&&/^[A-Z]类(?:份额)?/.test(bare))inherited={share:performanceScope('',bare).share,channel:''};
  if(short&&/^渠道[:：]/.test(bare))inherited={share:'',channel:performanceScope('',bare).channel};
  if(other||OTHER_PERFORMANCE_SUBJECT.test(text))continue;
  if(short&&/^(?:(?:管理人)?(?:提取|收取)?(?:业绩报酬|浮动业绩报酬)(?:计提)?|(?:管理人)?(?:提取|收取)?业绩报酬的?(?:原则|计提方法|计提方式|计提规则|计提日))[:：]?$/.test(bare)){feeHeading=true;headingText=text;headingIndex=index;continue;}
  if(short&&!/业绩报酬|[RＲ][1-9]?|收益率|不提取|不收取|不计提/.test(bare)&&/(?:管理费|托管费|税费|税收|投资范围|风险揭示|申购|赎回)/.test(bare))feeHeading=false;
  const clauses=text.match(/[^。；;\n]+[。；;]?/g)||[];
  for(const raw of clauses){
   const clause=raw.trim(),previous=blocks[index-1]?.text||'',bandContext=/[RＲ][1-9]?\s*[＜≤<>＞≥]|(?:年化)?收益率[^。；;\n]{0,15}(?:低于|高于|超过|不超过)/.test(previous);
   // A condition split into the preceding paragraph still qualifies this
   // denial. Only immediate local context is used; a new fee heading resets it.
   const localCondition=previous.length<120&&/^(?:\s*[（(]?[一二三四五六七八九十\d]+[）).、])?\s*(?:首期|首次|首个|本期|当期|本次|该期|某期|特定期间|优惠|减免|豁免|当.{1,90}时|若.{1,90}|如果.{1,90})/.test(previous);
   const denial=performanceDenial(clause,feeHeading&&headingIndex===index-1),conditionalDenial=denial&&(bandContext||localCondition||PERFORMANCE_CONDITIONAL.test(clause));
   const isNone=denial&&!conditionalDenial,isCharged=!denial&&chargesPerformanceFee(clause,feeHeading);if(!isNone&&!isCharged&&!conditionalDenial)continue;
   if(new RegExp(PRODUCT_END).test(clause)&&!clause.includes(target.name)&&!/(?:本|该)(?:集合|信托)?计划/.test(clause))continue;
   const local=performanceScope('',clause),share=local.share||inherited.share,channel=local.channel||inherited.channel;
   const evidence=quote(b.base||b,(conditionalDenial&&(bandContext||localCondition)?previous+'\n':feeHeading&&!/业绩报酬/.test(clause)?headingText+'\n':'')+text);
   if(local.issuer!==scope.issuer)continue;
   if(share&&share!==scope.share||channel&&channel!==scope.channel){if(!scope.share&&share||!scope.channel&&channel)ambiguous.push(evidence);continue;}
   if(conditionalDenial){conditional.push(evidence);continue;}
   const value=isNone?(/业绩报酬/.test(clause)?clause.replace(/^\s*[（(]?[一二三四五六七八九十\d]+[）).、]\s*/,'').replace(/[；;]$/,'。'):'本集合计划不提取业绩报酬。'):null;
   (isNone?none:charged).push({value,evidence});
  }
 }
 const evidence=unique([...none,...charged].map(x=>x.evidence).map(e=>JSON.stringify(e))).map(e=>JSON.parse(e));let result;
 if(none.length&&charged.length)result={policyMode:'conflict',status:'conflict',value:null,evidence,reason:'合同同一口径同时出现收取和不收取业绩报酬的条款，请核对适用条件；原模板暂保留。'};
 else if(ambiguous.length&&none.length||!none.length&&ambiguous.length)result={policyMode:'unresolved',status:'pending',value:null,evidence:[...evidence,...ambiguous],reason:'合同业绩报酬条款涉及不同份额或渠道，尚未确定本模板适用的收费状态。'};
 else if(none.length)result={policyMode:'none',status:'verified',value:none[0].value,evidence,source:'contract',reason:'合同明确不提取业绩报酬，以该收费状态替换旧计提话术及参数。'};
 else if(conditional.length&&!charged.length)result={policyMode:'unresolved',status:'pending',value:null,evidence:conditional,needsPolicyReview:true,reason:'合同中的不提取结论附带期间或收益条件，尚不能认定本产品始终不收取业绩报酬；保留原计提话术待核。'};
 else result={policyMode:charged.length?'charged':'unresolved',status:'pending',value:null,evidence,reason:charged.length?'合同明确收取业绩报酬，与模板不收取的状态不同；现有依据尚不能安全生成完整收费话术，请核对。':'合同尚未找到本产品同口径的明确业绩报酬收费状态，模板原结论保留待核。'};
 cache.set(cacheKey,result);return result;
}
function reviewPerformancePolicy(task,m,field,p,rs){
 if(field.manualReviewReason||keyFor(field.label)!=='performance')return false;
 const scope=performanceScope(field.label,field.before),policy=contractPerformancePolicy(rs,scope,task.target),templateNone=(field.before.match(/[^。；;\n]+[。；;]?/g)||[]).some(s=>noPerformanceFee(s,true));
 if(policy.policyMode!=='none'&&policy.policyMode!=='conflict'&&!policy.needsPolicyReview&&!templateNone)return false;
 const mixed=/管理费|托管费|销售服务费|认购费|申购费|赎回费/.test(p.text);
 const spans=mixed?[...p.text.matchAll(/[^。；;\n]+[。；;]?/g)].filter(s=>/业绩报酬/.test(s[0])).map(s=>({start:s.index,before:s[0]})):[{start:0,before:p.text}];
 for(const span of spans){
  const prior=task._records.some(r=>r.m===m&&r.field===field&&r.f.key==='performance_policy'&&r.f.policyMode==='none'&&JSON.stringify(r.atom.scope)===JSON.stringify(scope));
  const atom={key:'performance_policy',...span,value:span.before,scope,...(policy.policyMode==='none'?{renderValue:prior?'':policy.value}:{})};
  addOccurrence(task,m,field,p,atom,policy);
 }
 return true;
}
function resolveAtom(key,label,rs,scope,target,extra={}){
 const allowed=filtered(rs,scope);
 if(key==='benchmark')return pick(findMatches(allowed,/(?:业绩比较基准|业绩参考基准|业绩基准|参考业绩基准)(?:为|是|[:：])?\s*(\d+(?:\.\d+)?\s*[%％])/g,m=>`${number(m[1])}%`),'合同未给出同口径业绩比较基准，保留模板原值待核。');
 if(key==='product_name')return{status:target.status,value:target.name||null,evidence:target.evidence,reason:target.reason||'按合同产品名称核对。'};
 if(key.endsWith('_fee'))return resolveFee(key,rs,scope,/信托/.test(target.name));
 if(key.startsWith('performance_')){
  const pd=performanceData(allowed),isRatio=key.includes('ratio');const idx=Number(key.match(/\d+$/)?.[0])-1;const list=isRatio?pd.ratios:pd.thresholds;
  if(isRatio&&pd.tierRatios.some(r=>r.tier))return pick(pd.tierRatios.filter(r=>r.tier===idx+1),'合同未给出可对应的该档提成比例。');
  if(!isRatio&&pd.symbolic&&!list.length)return{status:'pending',value:null,evidence:pd.evidence,reasonType:'contract_missing',reason:'合同仅以符号定义业绩报酬基准并指向公告，未给出该档实际基准；保留模板各自原值待核。'};
  const val=list[idx];return val?{status:'verified',reason:'按合同分段计提条款核对同档参数。',...val}:{status:'pending',value:null,evidence:pd.evidence,reason:'合同未给出可对应的该档参数。'};
 }
 if(key==='risk_rating'){
 const re=/(?:风险等级|风险评级|产品评级|评级)[^。；;\nR]{0,45}(R[1-5])(?:[（(]([^）)]*风险)[）)])?/g;
  const issuerMatches=(m,b)=>{const start=Math.max(b.text.lastIndexOf('\n',m.index-1),b.text.lastIndexOf('。',m.index-1),b.text.lastIndexOf('；',m.index-1),b.text.lastIndexOf(';',m.index-1))+1;const context=b.text.slice(start,m.index+m[0].length),distributor=/(?:销售机构|代销机构|销售银行)[^。；;\n]{0,30}(?:风险等级|风险评级|评级|评定)/.test(context);return scope.issuer==='distributor'?distributor:!distributor;};
  const result=findMatches(allowed,re,(m,b)=>issuerMatches(m,b)?m[1]:null);
  if(!result.length)result.push(...findMatches(allowed,/(?:风险等级|风险评级)[^。；;\n]{0,15}?(中低风险|中高风险|中风险|低风险|高风险)/g,(m,b)=>issuerMatches(m,b)?({'低风险':'R1','中低风险':'R2','中风险':'R3','中高风险':'R4','高风险':'R5'})[m[1]]:null));
  return pick(result,'合同未提供该评级主体的明确评级，保留待核。');
 }
 if(['manager','custodian','advisor','distributor'].includes(key)){
  const a={manager:'(?:资产管理人|基金管理人|集合管理人|产品管理人|管理人|受托人)',custodian:'(?:集合托管人|资产托管人|托管人|保管人|保管银行|托管行|保管行)',advisor:'投资顾问',distributor:'(?:销售机构|代销机构|代理推介机构)'}[key];
  const org='([\\u4e00-\\u9fffA-Za-z]{2,65}?(?:有限责任公司|股份有限公司|有限公司)(?:[\\u4e00-\\u9fff]{0,20}分行)?)';
  let cs=findMatches(allowed,new RegExp(a+'(?:名称)?(?:(?:[：:|\\s—/]+)(?:指)?|指)[【“]?'+org),m=>m[1].trim(),b=>b.text.length<600);
  cs=cs.filter(c=>!/(?:履行|应当|负责|根据|聘请|担任|作为|指定)/.test(c.value));
  return pick(cs);
 }
 if(key==='product_type')return pick(findMatches(allowed,/(?:归属于|类型为|类别为|产品类型[：:|]?|本计划为|集合计划为|类型\s*[:：]?)[【\s]*((?:固定收益|混合|权益|商品及金融衍生品)类)/g));
 if(key==='product_code')return pick(findMatches(allowed,/(?:产品代码|产品编码)[：:|\s【（(]*([A-Z0-9]{5,25})/g));
 if(key==='opening_monthly_weekday'){
  const candidates=[];let unsupported=false;
  for(const b of allowed){
   for(const pattern of monthlyWeekdayPatterns(b.text)){if(!monthlyOpeningContext(b.text,pattern))continue;if(pattern.warning)unsupported=true;else candidates.push({value:pattern.value,evidence:[quote(b.base||b,b.text)]});}
  }
  return unsupported?{status:'pending',value:null,evidence:candidates.flatMap(c=>c.evidence),reason:'合同存在尚无法完整对应的月内开放安排，保留模板原值待核。'}:pick(candidates,'合同未明确同口径的月内第几个星期开放安排，保留模板原值待核。');
 }
 if(key==='opening_weekday')return pick(findMatches(allowed,/(?:每周|每星期)([一二三四五六日天])(?:[、和及][一二三四五六日天])?/g,m=>m[1],b=>/开放日|开放安排|开放期/.test(b.text)&&!/披露|估值/.test(b.text)));
 if(key==='opening_monthday')return pick(findMatches(allowed,/(?:每个?自然月|每月)(?:的)?[【\s]*(\d{1,2})[】\s]*日/g,m=>m[1],b=>/开放/.test(b.text)));
 if(key==='opening_cycle')return pick(findMatches(allowed,/(?:每满|每隔|每)(\d+)个?月(?:之对日|的?对应日|开放)/g));
 if(key==='opening_days')return pick(findMatches(allowed,/(?:每个开放期开放|连续开放)[【\s]*([一二两三四五六七八九十\d]+)[】\s]*(?:个)?交易日/g,m=>String(cnum(m[1]))));
 if(key==='lockup_days')return pick(findMatches(allowed,/(?:持有时间不得少于|持有期(?:不少于|为)|之后的第|封闭期(?:为|是))[【\s]*(\d+)[】\s]*(?:个)?(?:自然日|天)/g));
 if(key==='duration_years')return pick(findMatches(allowed,/(?:存续期|存续期限|存续时间|计划期限|信托期限|产品期限)[^。；;\n\d]{0,18}[【\s]*(\d+)[】\s]*年/g));
 if(key==='warning_line'||key==='stop_line')return pick(findMatches(allowed,new RegExp((key==='warning_line'?'预警线':'止损线')+'(?:为|为：|[：:])?[【\\s]*(0?\\.\\d+)', 'g')));
 if(key==='minimum_amount'||key==='additional_amount'){
  const re=key==='minimum_amount'?/(?:首次参与金额|首次认[/／]申购金额|参与起点|认购限额下限|最低认购金额|认购起点|最低参与金额|初始认购金额)[^。；;\n\d]{0,25}[【\s]*(\d+(?:\.\d+)?)[】\s]*万/g:/(?:追加参与|追加认购|追加申购|递增)[^。；;\n\d]{0,20}[【\s]*(\d+(?:\.\d+)?)[】\s]*万/g;
  return pick(findMatches(allowed,re));
 }
 const wanted=labelNorm(label);const candidates=[];
 for(const b of allowed){
  const lines=b.text.split(/[\n|]/).filter(Boolean);for(let i=0;i<lines.length;i++){
   const l=lines[i];const ex=l.match(/^\s*([^：:]{1,40})[：:]\s*(.+)$/);if(ex&&labelNorm(ex[1])===wanted)candidates.push({value:ex[2].trim(),evidence:[quote(b.base||b,b.text)]});
   if(labelNorm(l)===wanted&&lines[i+1])candidates.push({value:lines[i+1].trim(),evidence:[quote(b.base||b,b.text)]});
  }
 }
 if(extra.original){const text=clauseNorm(extra.original);const exact=allowed.find(b=>text.length>=4&&clauseNorm(b.text).includes(text));if(exact)return{status:'verified',value:extra.original,evidence:[quote(exact.base||exact,exact.text)],reason:'忽略排版空白和段落编号后，模板原句在合同正文中找到。'};
  if(key==='investment_ratios'){const normalized=equivalentClause(extra.original),same=allowed.find(b=>normalized.length>12&&equivalentClause(b.text)===normalized);if(same)return{status:'verified',value:extra.original,evidence:[quote(same.base||same,same.text)],reason:'同义规则仅统一“本计划”称谓、正回购融入资金余额用语；完整限制及数值一致，保留原表述。'};}
 }
 const result=pick(candidates,'自动核对未能确定对应的完整表述；请查看合同候选条款，原文保留待核。');if(result.status!=='verified'&&!result.evidence.length&&extra.original)result.evidence=nearbyEvidence(extra.original,allowed);return result;
}
function knownCellLabel(text){const value=text.trim();return value.length>0&&value.length<24&&!/[\d%％。；;]/.test(value)&&!keyFor(value).startsWith('custom:');}
function fieldRows(row){
 // Office merges retain empty physical cells. Pair only evidenced short labels
 // with their value cells. A value can resemble a label (e.g. 开放式), so only
 // the label positions establish pairing. Ambiguous grids stay pending.
 let cells=(row.cells||[]).filter(c=>c.text.trim());
 if(cells.length%2===1&&keyFor(cells[0]?.text)==='fees'&&knownCellLabel(cells[1]?.text||''))cells=cells.slice(1);
 if(cells.length<2||cells.length%2||!cells.every((c,i)=>i%2||knownCellLabel(c.text)))return cells.length>2?[{...row,manualReviewReason:'多栏要素的标签与内容尚不能可靠配对，保留各单元格原值待核。'}]:[row];
 const rows=[];for(let i=0;i<cells.length;i+=2){const pair=cells.slice(i,i+2);rows.push({...row,id:row.id+'-field-'+i,label:pair[0].text.trim(),value:pair[1].text,cells:pair,paragraphIds:pair.flatMap(c=>c.paragraphIds||[])});}return rows;
}
function performanceColumns(rows){
 // Column headings identify a performance matrix. Words such as 成立日 or
 // 收益率 inside ordinary terms do not make a multicolumn element table one.
 return rows.some(row=>{const cells=(row.cells||[]).map(c=>c.text.trim()).filter(Boolean);return cells.length>=3&&cells.filter(t=>t.length<=24&&!/[。；;，,%％]/.test(t)&&/成立(?:日期|时间|以来)|(?:单位|累计)?净值|收益率|回报率|夏普|波动率|回撤/.test(t)).length>=2;});
}
function getRows(template,target){
 const byTable=new Map();for(const r of template.rows||[]){const k=r.part+'|'+r.tableId;if(!byTable.has(k))byTable.set(k,[]);byTable.get(k).push(r);}
 const groups=[...byTable.values()].map(raw=>{const rows=raw.flatMap(fieldRows);return{rows,page:rows[0].page,part:rows[0].part,name:rows.find(r=>NAME_LABEL.test(labelNorm(r.label)))?.value.trim()||'',fields:rows.filter(r=>!keyFor(r.label).startsWith('custom:')).length,performance:performanceColumns(raw)};});
 const legalType=/信托/.test(target.name)?'trust':'asset';
 const good=groups.filter(g=>(g.fields>=2||g.name&&g.rows.length>=2)&&!MARKETING.test(template.slides?.find(s=>s.page===g.page)?.title||'')&&!g.performance&&(!g.name||(legalType==='trust'?/信托/.test(g.name):!/信托计划/.test(g.name))));
 let first=good.find(g=>g.name&&ELEMENT.test(template.slides?.find(s=>s.page===g.page)?.title||'')&&!UNDERLYING.test(template.slides?.find(s=>s.page===g.page)?.title||''))||good.find(g=>g.name&&!/底层|标的/.test(g.name));
 if(!first)first=good[0];
 const selected=[];let lastPage=null;
 for(const g of good){const slide=template.slides?.find(s=>s.page===g.page);const heading=(slide?.title||'')+(slide?.text||'').slice(0,100);
  if(g.name){if(first&&norm(g.name)!==norm(first.name))continue;if(UNDERLYING.test(heading)&&legalType==='trust')continue;selected.push(g);lastPage=g.page;}
  else if(first&&selected.length>0){
   let continuation=false;
   if(template.kind==='docx'){
    const previous=selected.at(-1);const previousIds=new Set(previous.rows.flatMap(r=>r.paragraphIds||[])),currentIds=new Set(g.rows.flatMap(r=>r.paragraphIds||[]));
    const end=Math.max(...template.paragraphs.filter(p=>previousIds.has(p.id)).map(p=>p.index));const start=Math.min(...template.paragraphs.filter(p=>currentIds.has(p.id)).map(p=>p.index));
    const between=template.paragraphs.filter(p=>p.part===g.part&&p.index>end&&p.index<start&&p.text.trim()).map(p=>p.text).join('\n');
    continuation=!!between&&/续表|接上页|要素.*续|续.*要素/.test(between)&&!MARKETING.test(between);
   }else continuation=((g.page===lastPage+1||g.page===lastPage)&&g.fields>=2||ELEMENT.test(heading)&&/续/.test(heading));
   if(continuation){selected.push(g);lastPage=g.page;}
  }
 }
 return{rows:selected.flatMap(g=>g.rows),name:first?.name||'',pages:unique(selected.map(g=>g.page).filter(Number.isInteger)).sort((a,b)=>a-b)};
}
const ANNOUNCEMENT=/开售|发售|募集|开放|预约|开放申购|开始认购/;
const REFERENCE_TITLE=/底层|投资标的|标的产品|历史|过往|同策略|代表产品|产品举例/;
const NONCURRENT_DATE=/历史|过往|上次|上一期|去年|成立(?:日期|时间|日|于)|设立于|数据(?:来源|截至)|截至.{0,25}(?:业绩|收益|净值|规模)|底层|投资标的|同策略/;
function pendingAnnouncement(){return{status:'pending',value:null,evidence:[],reasonType:'contract_missing',reason:'该具体开售/开放日期需当期公告支持，合同周期不能确认；保留原日期，数字标黄待核。'};}
function announcedDates(text,label=''){
 // Dates are spans in the template, never dates inferred from a contract cycle.
 const point='(?:\\d{4}\\s*[年/.-]\\s*)?\\d{1,2}\\s*[月/.-]\\s*\\d{1,2}日?';
 const re=new RegExp('(?<![\\d.])'+point+'(?:\\s*[-—–~～至到]\\s*(?:'+point+'|\\d{1,2}日?))?(?![\\d.%％])','g');
 const atoms=[];
 for(const m of text.matchAll(re)){
  const left=text.slice(0,m.index).split(/[。；;\n]/).at(-1),right=text.slice(m.index+m[0].length).split(/[。；;\n]/)[0];
  const context=label+' '+left+m[0]+right;
  if(NONCURRENT_DATE.test(context)||/每年\s*$/.test(left))continue;
  const localLeft=text.slice(0,m.index).split(/[。；;，,：:（(\n]/).at(-1),localRight=text.slice(m.index+m[0].length).split(/[。；;，,：:（()）\n]/)[0],localContext=localLeft+m[0]+localRight;
  const roleText=localLeft+' '+label+' '+localRight;
  const role=/截止|结束|终止|最后/.test(roleText)?'end':/起始|开始|首个|首日/.test(roleText)?'start':'date';
  const event=/预约/.test(localContext)||/预约/.test(label)?'reservation':'opening';
  const dateWindow=/仅(?:开放)?参与|仅(?:开放)?申购/.test(localRight)?'participation_only':/(?:参与|申购)(?:和|及|与)(?:退出|赎回)/.test(localRight)?'participation_exit':'';
  let parsed;try{parsed=parseDateInput(m[0],2000);}catch{continue;}
  if(parsed?.end){
   atoms.push({key:'opening_announced_date',start:m.index,before:parsed.startText,value:parsed.startText,dateRole:'start',dateEvent:event,dateWindow,rangeEndpoint:true});
   atoms.push({key:'opening_announced_date',start:m.index+parsed.endOffset,before:parsed.endText,value:parsed.endText,dateRole:'end',dateEvent:event,dateWindow,rangeEndpoint:true});
  }else atoms.push({key:'opening_announced_date',start:m.index,before:m[0],value:m[0],dateRole:role,dateEvent:event,dateWindow});
 }
 return atoms;
}
function posterHeadings(template,knownName){
 const ps=template.paragraphs.filter(p=>p.text.trim()),names=new Map(),dates=new Map();
 const reference=p=>{
  const group=p.shapeId?ps.filter(q=>q.part===p.part&&q.shapeId===p.shapeId).map(q=>q.text).join('\n'):p.text;
  const previous=ps[ps.indexOf(p)-1];
  return REFERENCE_TITLE.test(group)||!!(previous&&previous.part===p.part&&previous.text.length<24&&REFERENCE_TITLE.test(previous.text)&&!new RegExp(PRODUCT_END).test(previous.text));
 };
 const candidates=[];
 for(const p of ps){
  if(p.tableId||p.text.length>180||reference(p))continue;
  // A title may include line breaks and a separate code line in the same paragraph.
  const match=p.text.match(new RegExp('^\\s*(?:产品(?:名称|全称)\\s*[：:]\\s*)?([^。；;：:,，]{4,110}?'+PRODUCT_END+')(?=\\s*(?:$|产品(?:代码|编码)))'));
  if(match&&!/^(?:本产品|本计划|本集合计划|投向|投资于)/.test(match[1])&&/^\s*(?:产品(?:代码|编码)\s*[：:（(【\s]*[A-Z0-9]{5,25}[）)】\s]*)?$/.test(p.text.slice(match.index+match[0].length)))candidates.push({p,before:match[1],start:p.text.indexOf(match[1])});
 }
 const distinct=unique(candidates.map(c=>norm(c.before)));
 const chosen=knownName?norm(knownName):distinct.length===1?distinct[0]:distinct.find(n=>norm(template.name).includes(n.replace(new RegExp(PRODUCT_END+'$'),'')));
 for(const c of candidates)if(norm(c.before)===chosen)names.set(c.p.id,c);
 for(let i=0;i<ps.length;i++){
  const p=ps[i];if(p.tableId||p.text.length>110||reference(p)||NONCURRENT_DATE.test(p.text))continue;
  const atoms=announcedDates(p.text);if(!atoms.length)continue;
  let context=p.text;
  if(!ANNOUNCEMENT.test(context)){
   // Only an isolated date can borrow the immediately adjacent short sale cue.
   let rest=p.text;for(const a of [...atoms].reverse())rest=rest.slice(0,a.start)+rest.slice(a.start+a.before.length);
   if(!/^[\s：:（()）至到—–~-]*$/.test(rest))continue;
   const neighbor=[ps[i-1],ps[i+1]].find(q=>q&&q.part===p.part&&q.page===p.page&&!q.tableId&&q.text.length<30&&ANNOUNCEMENT.test(q.text)&&!NONCURRENT_DATE.test(q.text)&&!/[\d。；;，,]/.test(q.text));
   if(!neighbor)continue;context=neighbor.text+' '+p.text;
  }
  dates.set(p.id,{context,atoms:announcedDates(p.text,context)});
 }
 return{names,dates,excluded:new Set(ps.filter(reference).map(p=>p.id)),name:candidates.find(c=>norm(c.before)===chosen)?.before.replace(/\s/g,'')||knownName||''};
}
function posterRows(template,headings){
 const rows=[];let label='',ids=[];const ps=template.paragraphs||[];
 const flush=()=>{if(label&&ids.length){const p=ps.find(p=>p.id===ids[0]);rows.push({id:'poster-'+rows.length,label,value:ids.map(id=>ps.find(p=>p.id===id).text).join('\n'),part:p.part,page:p.page,paragraphIds:ids,cells:[{text:label,paragraphIds:[]},{text:ids.map(id=>ps.find(p=>p.id===id).text).join('\n'),paragraphIds:ids}]});}ids=[];};
 for(const p of ps){const text=p.text.trim();if(!text)continue;const clean=text.replace(/[：:]$/,'');
  if(headings?.names.has(p.id)){flush();label='';continue;}
  if(headings?.dates.has(p.id)){if(keyFor(label)==='opening')ids.push(p.id);else{flush();label='';}continue;}
  if(/^(?:历史业绩|过往业绩|业绩表现|产品表现|同策略产品|产品特点|产品亮点|管理人介绍|公司介绍|团队介绍|个人履历|风险提示|风险揭示)/.test(text)||/(?:过去.{0,8}(?:年化|收益|业绩)|成立以来|数据来源)/.test(text)){flush();label='';continue;}
  if(clean.length<22&&(!keyFor(clean).startsWith('custom:')||/[：:]$/.test(text)&&!/产品特点|产品亮点|团队|介绍|履历|姓名/.test(clean))){flush();label=clean;continue;}if(label){if(/履历|个人简介|：.{0,20}大学|现任|历任|擅长|本产品由|本资料仅/.test(text)){flush();label='';}else ids.push(p.id);}}
 flush();return rows;
}
function minimalChange(before,after){let start=0;while(start<before.length&&start<after.length&&before[start]===after[start])start++;let endA=before.length,endB=after.length;while(endA>start&&endB>start&&before[endA-1]===after[endB-1]){endA--;endB--;}return{start,before:before.slice(start,endA),after:after.slice(start,endB)};}
function alignedBoundaries(before,after){
 const n=before.length,m=after.length,d=Array.from({length:n+1},()=>new Uint16Array(m+1));
 for(let i=n;i>=0;i--)for(let j=m;j>=0;j--)d[i][j]=i===n?m-j:j===m?n-i:Math.min(d[i+1][j+1]+(before[i]===after[j]?0:1),d[i+1][j]+1,d[i][j+1]+1);
 const map=new Map();let i=0,j=0;
 while(i<n||j<m){map.set(i,j);if(i<n&&j<m&&d[i][j]===d[i+1][j+1]+(before[i]===after[j]?0:1)){i++;j++;}else if(j<m&&d[i][j]===d[i][j+1]+1)j++;else i++;}
 map.set(n,m);return map;
}
function nameStyle(canonical,before){
 let s=canonical;if(before.includes('('))s=s.replaceAll('（','(').replaceAll('）',')');
 const dot=before.indexOf('·');if(dot>0){const prefix=before.slice(0,dot);if(s.startsWith(prefix)&&!s.includes('·'))s=s.slice(0,prefix.length)+'·'+s.slice(prefix.length);}
 const boundaries=alignedBoundaries(before.replace(/\s/g,''),s);
 const spaces=[...before.matchAll(/\s+/g)].map(m=>({position:boundaries.get(before.slice(0,m.index).replace(/\s/g,'').length),text:m[0]}));for(const x of spaces.reverse())s=s.slice(0,x.position)+x.text+s.slice(x.position);
 return s;
}
function monthNow(now){return new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit'}).format(now||new Date()).replace('-','');}
function atomsFor(label,text,target,fieldText=text){
 const key=keyFor(label),atoms=[];const add=(k,m,before=m[0],value=before,extra={})=>atoms.push({key:k,start:m.index+(m[0].indexOf(before)),before,value,...extra});
 const trust=/信托/.test(target.name);
 if(key==='benchmark')return [...text.matchAll(PERCENT)].map(m=>({key,start:m.index,before:m[0],value:`${number(m[0])}%` }));
 if(key==='product_name')return[{key,start:0,before:text,value:text}];
 if(['manager','custodian','advisor','distributor','product_code'].includes(key))return[{key,start:0,before:text,value:text}];
 if(key==='product_type'){const m=text.match(/(?:固定收益|混合|权益|商品及金融衍生品)类/);return m?[{key,start:m.index,before:m[0],value:m[0]}]:[{key:'custom:product_type_wording',start:0,before:text,value:text,manualReviewReason:'模板此处未明确资产类别，可能描述运作方式或其他分类；保留原表述，不能直接替换为合同资产类别。'}];}
 if(key==='risk_rating'){
  for(const m of text.matchAll(/R[1-5]/g))add(key,m,m[0],m[0],{scope:scopeFor(label,text,m.index)});
  // Verbal ratings are separately tied to the same issuer, preserving layout and punctuation.
  for(const m of text.matchAll(/中低风险|中高风险|中风险|低风险|高风险/g))add('risk_description',m,m[0],m[0],{scope:scopeFor(label,text,m.index)});
  return atoms.length?atoms:[{key,start:0,before:text,value:text}];
 }
 if(['fees','management_fee','custody_fee','subscription_fee','redemption_fee'].includes(key)){const labeled=FEE_LABELS.find(([,re])=>new RegExp('^(?:'+re.source+')$').test(labelNorm(label)))?.[0],actions=feeActionSet(label),feeKey=feeActionKey(actions,labeled||key);return feeTokens(text,feeKey,trust,actions).map(t=>({...t,scope:{...scopeFor(label,text,t.start),...(t.feeActions&&t.key!=='redemption_fee'?{feeActions:t.feeActions}:{})}}));}
 if(key==='performance'){
  const matches=performanceTokens(text),all=performanceTokens(fieldText),thresholds=unique(all.filter(t=>t.role==='threshold').map(t=>number(t.value))).sort((a,b)=>a-b);
  const ratios=unique(all.filter(t=>t.role==='ratio').map(t=>number(t.value))).sort((a,b)=>a-b);
  for(const token of matches){
   const n=number(token.value),tier=token.role==='ratio'?(payoutTier(text,token,thresholds)||ratios.indexOf(n)+1):thresholds.indexOf(n)+1;
   if(token.role==='unknown')atoms.push({...token,key:'custom:performance_parameter',manualReviewReason:'尚无法确认该百分比是基准还是提成比例，保留模板原值待核。'});
   else atoms.push({...token,key:token.role==='ratio'?'performance_ratio_'+tier:'performance_R'+tier});
  }
  return atoms;
 }
 if(['opening','lockup','duration'].includes(key)){
  for(const pattern of monthlyWeekdayPatterns(text))atoms.push({key:'opening_monthly_weekday',start:pattern.start,before:pattern.before,value:pattern.value,monthlyPattern:pattern,...(pattern.warning?{manualReviewReason:pattern.warning}:{})});
  const rules=[['opening_weekday',/(?:每周|每星期)([一二三四五六日天])/g],['opening_monthday',/(?:每个?自然月|每月)(?:的)?[【\s]*(\d{1,2})[】\s]*日/g],['opening_cycle',/(?:每满|每隔|每)(\d+)个?月(?:之对日|的?对应日|开放)/g],['opening_days',/(?:每个开放期开放|连续开放)[【\s]*([一二两三四五六七八九十\d]+)[】\s]*(?:个)?交易日/g],['lockup_days',/(?:持有时间不得少于|持有期(?:不少于|为)|之后的第|封闭期(?:为|是))[【\s]*(\d+)[】\s]*(?:个)?(?:自然日|天)/g]];
  for(const [k,re]of rules)for(const m of text.matchAll(re))add(k,m,m[1],m[1]);
  if(key==='duration')for(const m of text.matchAll(/(\d+)年/g))add('duration_years',m,m[1]);
  const day=atoms.find(a=>a.key==='opening_monthday');if(day)for(const m of text.matchAll(/含(\d{1,2})日/g))add('opening_monthday',m,m[1]);
  atoms.push(...announcedDates(text,label));
  return atoms.length?atoms:[{key,start:0,before:text,value:text}];
 }
 if(key==='minimum_amount'){
  const ms=[...text.matchAll(/(\d+(?:\.\d+)?)\s*万元?/g)];for(const [i,m] of ms.entries())add(i===0?'minimum_amount':'additional_amount',m,m[1]);return atoms;
 }
 if(key==='warning_stop')return[{key,start:0,before:text,value:text}];
 return[{key,start:0,before:text,value:text}];
}
function materialKind(t){return t.kind==='docx'?'word':/海报/.test(t.name)||t.slides?.length===1?'poster':'ppt';}
function wordCoverNames(template,target){
 const names=new Map(),body=[];
 for(const p of template.paragraphs.filter(p=>p.part==='word/document.xml')){
  const nodes=p.node?elements(p.node,NS.w):[];
  if(p.page>1||body.length&&nodes.some(n=>n.localName==='pageBreakBefore'&&!['0','false','off'].includes(n.getAttributeNS(NS.w,'val'))))break;
  if(p.text.trim())body.push(p);
  if(nodes.some(n=>n.localName==='sectPr'||n.localName==='br'&&n.getAttributeNS(NS.w,'type')==='page'))break;
  if(body.length>=8)break;
 }
 const legal=new RegExp('^'+PRODUCT_END+'$'),full=new RegExp('^([^。；;：:,，]{4,110}?'+PRODUCT_END+')$');
 // The main cover has its own identity. A stale title need not equal the table's alias.
 for(let i=0;i<Math.min(body.length,8);i++){
  const p=body[i],text=p.text.trim();
  if(p.tableId||REFERENCE_TITLE.test(text)||MARKETING.test(text)||/产品要素|风险提示|公司介绍|管理人介绍/.test(text)||text.length>130)break;
  if(full.test(text)&&!legal.test(text)){
   names.set(p.id,{before:text,after:target.name});break;
  }
  const next=body[i+1];
  if(next&&!next.tableId&&legal.test(next.text.trim())&&text.length>=4&&!/[。；;：:,，]/.test(text)){
   names.set(p.id,{before:text,after:target.shortName});
   names.set(next.id,{before:next.text.trim(),after:target.name.slice(target.shortName.length)});break;
  }
 }
 return names;
}
function wordPerformanceParagraphs(template,selection,coverNames){
 const names=[selection.name,...[...coverNames.values()].map(n=>n.before)].filter(n=>n&&!new RegExp('^'+PRODUCT_END+'$').test(n)).map(norm);
 if(!names.length)return[];
 const full=new RegExp('^\\s*(?:产品名称\\s*[:：]\\s*)?([^。；;：:,，]{4,110}?'+PRODUCT_END+')\\s*$'),out=[];let excluded=false;
 for(const p of template.paragraphs){
  if(p.part!=='word/document.xml')continue;const text=p.text.trim();if(!text)continue;
  const title=text.match(full);
  if(coverNames.has(p.id)||title&&names.includes(norm(title[1])))excluded=false;
  else if(title||text.length<80&&/(?:底层|投资标的|标的产品|历史|过往|同策略|代表产品|产品举例|公司介绍|管理人介绍|团队介绍)/.test(text))excluded=true;
  if(excluded||p.tableId||text.length>1800||REFERENCE_TITLE.test(text)||/捐赠|假设|举例|示例/.test(text))continue;
  if(!/(?:管理人|本(?:集合|信托)?计划|本产品)/.test(text)||!/业绩报酬/.test(text)||!/提取|计提|收取|提成/.test(text))continue;
  if(new RegExp(PRODUCT_END).test(text)&&!names.some(n=>norm(text).includes(n)))continue;
  const tokens=performanceTokens(text);if(tokens.some(t=>t.role==='threshold')&&tokens.some(t=>t.role==='ratio'))out.push(p);
 }
 return out;
}
function fieldParagraphs(row,map){const ids=row.cells?.slice(1).flatMap(c=>c.paragraphIds||[])||row.paragraphIds;return unique(ids||[]).map(id=>map.get(id)).filter(Boolean);}
function addOccurrence(task,m,field,p,atom,resolution){
 const scope={product:task.target.name,...(atom.scope||scopeFor(field.label,p.text,atom.start))};
 if(atom.key==='opening_announced_date'){scope.period='announcement';scope.condition=[atom.dateEvent,atom.dateRole,norm(atom.before),atom.dateWindow||''].join('|');}
 if(atom.key.startsWith('investment_')||atom.key.startsWith('custom:')||['opening','lockup','duration'].includes(atom.key))scope.condition=norm(atom.before);
 const factKey=(atom.key==='risk_description'?'risk_rating':atom.key)+'|'+JSON.stringify(scope);
 const factLabel=atom.redemptionTier?'退出费率（'+atom.redemptionTier.label+'）':scope.feeActions?scope.feeActions.replaceAll('|','/')+'费率':atom.key==='redemption_fee'&&keyFor(field.label)!=='redemption_fee'?'退出费率':atom.key==='performance_policy'?'业绩报酬':atom.key.startsWith('performance_ratio_')?'业绩报酬提成比例（第'+atom.key.match(/\d+$/)[0]+'档）':/^performance_R\d/.test(atom.key)?'业绩报酬基准 R'+atom.key.match(/\d+$/)[0]:field.label;
 let f=task._factMap.get(factKey);if(!f){f={id:'fact-'+(task.facts.length+1),key:atom.key,label:factLabel,value:resolution.value,status:resolution.status,evidence:resolution.evidence||[],scope,occurrences:[],reason:resolution.reason,reasonType:resolution.reasonType||(resolution.status==='verified'?'confirmed':'needs_review')};task.facts.push(f);task._factMap.set(factKey,f);}
 else if((atom.key==='opening_monthly_weekday'||atom.redemptionTier)&&resolution.status!=='verified'&&f.status==='verified'){
  Object.assign(f,{value:null,status:resolution.status,reason:resolution.reason,reasonType:'needs_review'});
  for(const material of task.materials)material.operations=material.operations.filter(o=>!o.factIds?.includes(f.id));
  for(const record of task._records.filter(r=>r.f===f))renderRecord(record);
 }
 if(atom.key==='performance_policy'){f.policyMode=resolution.policyMode;if(resolution.source)f.source=resolution.source;}
 if(atom.redemptionTier)f.redemptionTier=atom.redemptionTier;
 f.occurrences.push({materialId:m.materialId,paragraphId:p.id,fieldId:field.id});field.factIds.push(f.id);
 const record={m,field,p,atom,f};task._records.push(record);renderRecord(record);
}
function operationId(m){m._operationSequence=(m._operationSequence||0)+1;return 'op-'+m._operationSequence;}
function renderRecord({m,field,p,atom,f,userRenderValue}){
 if(field.sectionOverrideId||field.feeTextOverrideId||field.customSectionOverrideId)return;
 const common={paragraphId:p.id,part:p.part,index:p.index,reason:f.reason,factIds:[f.id]};
 if(f.userConfirmed)return;
 if(atom.key==='opening_fixed_monthdays'&&f.status==='verified'&&userRenderValue===undefined){
  for(const change of atom.monthdayChanges||[])m.operations.push({...common,...change,id:operationId(m),kind:'replace'});
  return;
 }
 if(atom.key==='opening_monthly_weekday'){
  const result=f.status==='verified'?monthlyWeekdayChanges(atom.monthlyPattern,String(f.value)):{changes:[]};
  if(f.status==='verified'&&!result.warning){for(const diff of result.changes)m.operations.push({...common,id:operationId(m),kind:'replace',start:atom.start+diff.start,before:diff.before,after:diff.after});}
  else{const anchors=atom.monthlyPattern.tokens||[{start:0,before:atom.before}];for(const a of anchors){m.operations.push({...common,id:operationId(m),kind:'highlight',start:atom.start+a.start,before:a.before,color:'yellow'});m.operations.push({...common,id:operationId(m),kind:'comment',start:atom.start+a.start,before:a.before});}}
  return;
 }
 if(f.status==='verified'){
  let after=String(userRenderValue??atom.renderValue??f.value??'');if(atom.key==='product_name')after=nameStyle(after,atom.before);if(atom.key==='risk_description')after=({R1:'低风险',R2:'中低风险',R3:'中风险',R4:'中高风险',R5:'高风险'})[after]||after;
  if(atom.key==='opening_days'&&/^[一二两三四五六七八九十]$/.test(atom.before))after=Object.entries({一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10}).find(([,v])=>v===Number(after))?.[0]||after;
  if(equivalent(atom.value,after))return;
  const percentage=atom.before.match(/^(\d+(?:\.(\d+))?)(\s*[】\]]?\s*[%％])$/);
  if(/^\d+(?:\.\d+)?%$/.test(after)&&percentage){const decimals=percentage[2]?.length||0,needed=after.split('.')[1]?.replace('%','').length||0;after=number(after).toFixed(Math.max(decimals,needed))+percentage[3];}
  const lines=atom.key==='product_name'&&atom.before.includes('\n')?atom.before.split('\n'):null;
  const diffs=[];
  if(lines&&lines.length===after.split('\n').length){let offset=0;const next=after.split('\n');for(let i=0;i<lines.length;i++){const diff=minimalChange(lines[i],next[i]);diffs.push({...diff,start:offset+diff.start});offset+=lines[i].length+1;}}
  else diffs.push(atom.before.length<35&&!['product_name','manager','custodian','advisor','distributor'].includes(atom.key)?{start:0,before:atom.before,after}:minimalChange(atom.before,after));
  for(const diff of diffs)if(diff.before||diff.after){
   // A pure insertion still needs a precise existing-text anchor for the
   // Office writer. Keep the anchor inside this atom, never a global match.
   let {start,before,after:replacement}=diff;
   if(!before){const next=[...atom.before.slice(start)][0];if(next){before=next;replacement+=next;}else{const previous=[...atom.before.slice(0,start)].at(-1);if(previous){start-=previous.length;before=previous;replacement=previous+replacement;}}}
   m.operations.push({...common,id:operationId(m),kind:'replace',start:atom.start+start,before,after:replacement});
  }
 }else{
  const spans=atom.before.matchAll(/\d+(?:\.\d+)?\s*[%％]?|R[1-5]|[一二三四五六日天](?=开放)/g);const list=[...spans];
  // Compound pending text is anchored to each parameter rather than painting whole cells.
  const anchors=list.length?list.map(s=>({start:atom.start+s.index,before:s[0]})):[{start:atom.start,before:atom.before}];
  for(const a of anchors)if(a.before){m.operations.push({...common,...a,id:operationId(m),kind:'highlight',color:'yellow'});m.operations.push({...common,...a,id:operationId(m),kind:'comment'});}
 }
}
function reviewFixedMonthlyOpening(task,m,field,paragraphs){
 if(keyFor(field.label)!=='opening'||field.manualReviewReason)return false;
 const scope=scopeFor(field.label,field.before),resolution=resolveMonthlyDateRule(task.contract,field.before,scope);if(!resolution)return false;
 // A distinct wording/condition retains its own review status, while repeated
 // paragraphs of this one field share one complete-rule fact.
 scope.condition=norm(field.before);let offset=0;
 for(const p of paragraphs){
  if(p.text.trim())addOccurrence(task,m,field,p,{key:'opening_fixed_monthdays',start:0,before:p.text,value:p.text,scope,monthdayChanges:resolution.changes.filter(c=>c.start>=offset&&c.start+c.before.length<=offset+p.text.length).map(c=>({...c,start:c.start-offset}))},resolution);
  offset+=p.text.length+1;
 }
 return true;
}
function reviewParagraph(task,m,field,p,rs){
 if(reviewPerformancePolicy(task,m,field,p,rs))return;
 const atoms=field.manualReviewReason?[{key:'custom:table_layout:'+labelNorm(field.label),start:0,before:p.text,value:p.text,manualReviewReason:field.manualReviewReason}]:atomsFor(field.label,p.text,task.target,keyFor(field.label)==='performance'?m._performanceText:field.before);
 for(const atom of atoms){let resolution;const aggregate=(atom.key.endsWith('_fee')||atom.key.startsWith('performance_'))&&AGGREGATE_LAYERS.test(field.label+' '+field.before);
  if(atom.key==='redemption_fee'&&!aggregate){const tier=matchTemplateRedemptionTier(p.text,atom);if(tier){atom.redemptionTier=tier;atom.scope={...(atom.scope||scopeFor(field.label,p.text,atom.start)),...(tier.scope||{}),condition:'holding_days:'+String(tier.condition||norm(p.text))};}}
  if(aggregate)atom.scope={...(atom.scope||scopeFor(field.label,p.text,atom.start)),condition:'aggregate_layers'};
  if(atom.manualReviewReason)resolution={status:'pending',value:null,evidence:[],reasonType:'needs_review',reason:atom.manualReviewReason};
  else if(atom.redemptionTier)resolution=resolveRedemptionTier(task._redemptionTierIndex,atom.redemptionTier,atom.scope);
  else if(atom.key==='warning_stop'){
   const a=resolveAtom('warning_line',field.label,rs,{},task.target),b=resolveAtom('stop_line',field.label,rs,{},task.target);
   resolution=a.status==='verified'&&b.status==='verified'?{status:'verified',value:`预警线${a.value}，止损线${b.value}`,evidence:[...a.evidence,...b.evidence],reason:'合同明确设置预警线和止损线。'}:resolveAtom(atom.key,field.label,rs,{},task.target,{original:p.text});
  }else if(atom.key==='risk_description')resolution=resolveAtom('risk_rating',field.label,rs,atom.scope,task.target);
  else if(atom.key==='opening_announced_date')resolution=pendingAnnouncement();
  else if(atom.key==='investment_strategy'&&m._strategySummaries?.get(JSON.stringify(scopeFor(field.label,field.before)))){resolution=m._strategySummaries.get(JSON.stringify(scopeFor(field.label,field.before)));if(resolution.status==='verified')resolution={...resolution,value:atom.before};}
  else resolution=resolveAtom(atom.key,field.label,rs,atom.scope||scopeFor(field.label,p.text,atom.start),task.target,{original:atom.before});
  if(atom.key==='opening_monthly_weekday'&&resolution.status==='verified'){const alignment=monthlyWeekdayChanges(atom.monthlyPattern,resolution.value);if(alignment.warning)resolution={...resolution,status:'pending',value:null,reasonType:'needs_review',reason:alignment.warning};}
  if(aggregate&&!(resolution.status==='verified'&&resolution.evidence?.length&&resolution.evidence.every(e=>AGGREGATE_LAYERS.test(e.quote))))resolution={...resolution,status:'pending',value:null,reasonType:'needs_review',reason:'模板列示跨层合计费用，尚未取得同口径的合同依据，保留原值待核。'};
  // Retain a documented successor organization alias in the template.
  if(atom.key==='manager'&&resolution.status==='verified'&&!equivalent(atom.before,resolution.value)){
   const docs=task.templates.flatMap(t=>t.paragraphs.map(p=>p.text)).join('\n');const old=esc(resolution.value.trim()),newName=atom.before.trim();
   if(new RegExp(old+'.{0,40}(?:更名为|变更为).{0,4}'+esc(newName)).test(docs)){atom.renderValue=atom.before;resolution={...resolution,reason:'模板正文明确记录管理人更名关系，保留模板现用名称。',evidence:[...resolution.evidence,{blockId:'template-alias',quote:docs.match(new RegExp(old+'.{0,40}(?:更名为|变更为).{0,4}'+esc(newName)))?.[0]||''}]};}
  }
  if((atom.key.startsWith('custom:')||atom.key.startsWith('investment_')||['opening','lockup','duration'].includes(atom.key))&&resolution.status==='verified'&&!equivalent(atom.before,resolution.value)){
   const a=[...atom.before.matchAll(/\d+(?:\.\d+)?/g)],b=[...String(resolution.value).matchAll(/\d+(?:\.\d+)?/g)];
   if(a.length===1&&b.length===1&&norm(atom.before.replace(a[0][0],'#'))===norm(String(resolution.value).replace(b[0][0],'#'))){atom.start+=a[0].index;atom.before=a[0][0];atom.value=atom.before;resolution={...resolution,value:b[0][0]};}
   else resolution={...resolution,status:'pending',value:null,reasonType:'needs_review',reason:'找到相关条款，但表述与适用条件尚无法自动一一对应，保留待核。'};
  }
  addOccurrence(task,m,field,p,atom,resolution);
 }
}
function finalize(task){
 // Rebuild only previously compacted fields so later supplements and undo use fresh facts.
 for(const m of task.materials)for(const field of m.fields.filter(f=>f.equalFeeCompacted)){
  delete field.equalFeeCompacted;const ids=new Set(field.paragraphIds);
  m.operations=m.operations.filter(o=>!ids.has(o.paragraphId));
  for(const r of task._records.filter(r=>r.m===m&&r.field===field))renderRecord(r);
 }
 refreshSectionOperations(task);
 refreshFeeTextOperations(task);
 refreshCustomSectionOperations(task);
 for(const m of task.materials){
  m.operations=m.operations.filter((o,i,arr)=>!arr.slice(0,i).some(a=>a.paragraphId===o.paragraphId&&a.kind===o.kind&&a.start===o.start&&a.before===o.before&&a.after===o.after));
  for(const field of m.fields){if(field.source==='manager')continue;field.factIds=unique(field.factIds);const fs=field.factIds.map(id=>task.facts.find(f=>f.id===id));field.status=fs.some(f=>f.status==='conflict')?'conflict':fs.some(f=>f.status==='pending')?'pending':'verified';const replaces=m.operations.filter(o=>field.paragraphIds.includes(o.paragraphId)&&o.kind==='replace');field.after=field.paragraphIds.map(id=>{let t=m.template.paragraphs.find(p=>p.id===id)?.text||'';for(const o of replaces.filter(o=>o.paragraphId===id).sort((a,b)=>b.start-a.start))t=t.slice(0,o.start)+o.after+t.slice(o.start+o.before.length);return t;}).join('\n');if(field.source!=='company')field.reason=field.status==='verified'?(replaces.length?'合同有明确依据，最小范围更新。':'原值与已取得依据一致。'):'保留模板原值，关键参数标黄并添加批注。';}
  for(const field of m.fields){
   const fs=field.factIds.map(id=>task.facts.find(f=>f.id===id));
   if(field.status!=='verified'||!fs.length||fs.some(f=>!['subscription_fee','redemption_fee','subscription_redemption_fee'].includes(f.key)||f.userConfirmed)||field.sectionOverrideId||field.feeTextOverrideId)continue;
   const after=simplifyEqualFeeTiers(field.after);if(after===field.after)continue;
   const ids=new Set(field.paragraphIds);if(m.fields.some(f=>f!==field&&f.paragraphIds.some(id=>ids.has(id))))continue;
   const ps=field.paragraphIds.map(id=>m.template.paragraphs.find(p=>p.id===id)).filter(p=>p?.text?.trim());if(!ps.length)continue;
   m.operations=m.operations.filter(o=>!ids.has(o.paragraphId));
   for(const [i,p]of ps.entries())if(p.text!==(i===0?after:''))m.operations.push({id:operationId(m),kind:'replace',source:'equal_fee_tiers',paragraphId:p.id,part:p.part,index:p.index,start:0,before:p.text,after:i===0?after:'',factIds:field.factIds,reason:'各档费率相同且金额或持有期限覆盖完整，合并为统一费率。'});
   field.after=after;field.equalFeeCompacted=true;field.reason='各档费率相同，合并为统一费率；原合同依据保留。';
  }
 }
 refreshCustomSectionFields(task);
 task.stats={facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};return task;
}
export function analyzeMaterials(contract,templates,{now,month}={}){
 const target=identity(contract),task={target,contract,templates,month:month||monthNow(now),facts:[],materials:[],warnings:[...(contract.warnings||[])],stats:{}};
 Object.defineProperties(task,{_factMap:{value:new Map()},_records:{value:[]},_redemptionTierIndex:{value:buildRedemptionTierIndex(contract)}});
 if(!target.name){task.warnings.push(target.reason);throw new Error(target.reason);}
 const rs=records(contract);
 for(const t of templates){
  const kind=materialKind(t),selection=getRows(t,target),headings=kind==='poster'?posterHeadings(t,selection.name):null,coverNames=kind==='word'?wordCoverNames(t,target):new Map(),rows=selection.rows.length?selection.rows:kind==='poster'?posterRows(t,headings):[];
  const freePerformance=kind==='word'?wordPerformanceParagraphs(t,selection,coverNames):[];
  const m={materialId:t.id,template:t,kind,templateProductName:headings?.name||selection.name,selectedPages:kind==='poster'?(t.slides||[]).map(s=>s.page):selection.pages,operations:[],fields:[],warnings:[...(t.warnings||[])]};Object.defineProperty(m,'_performanceText',{value:[...rows.filter(r=>keyFor(r.label)==='performance').map(r=>r.value),...freePerformance.map(p=>p.text)].join('\n')});task.materials.push(m);
  if(!rows.length&&!freePerformance.length)m.warnings.push('尚未识别到可可靠定位的产品要素区，需检查模板结构。');
  const map=new Map(t.paragraphs.map(p=>[p.id,p]));const selectedIds=new Set();
  const strategyGroups=new Map();for(const row of rows.filter(r=>keyFor(r.label)==='investment_strategy')){const text=fieldParagraphs(row,map).map(p=>p.text).join('\n'),scope=scopeFor(row.label,text),key=JSON.stringify(scope);const group=strategyGroups.get(key)||{scope,texts:[]};group.texts.push(text);strategyGroups.set(key,group);}
  Object.defineProperty(m,'_strategySummaries',{value:new Map([...strategyGroups].map(([key,g])=>[key,resolveStrategySummary(contract,unique(g.texts).join('\n'),g.scope)]))});
  for(const row of rows){
   const ps=fieldParagraphs(row,map);ps.forEach(p=>selectedIds.add(p.id));const field={id:`${t.id}-field-${m.fields.length+1}`,label:row.label,status:'pending',before:ps.map(p=>p.text).join('\n'),after:'',factIds:[],paragraphIds:ps.map(p=>p.id),reason:'',manualReviewReason:row.manualReviewReason};m.fields.push(field);
   if(!reviewFixedMonthlyOpening(task,m,field,ps))for(const p of ps)if(p.text.trim())reviewParagraph(task,m,field,p,rs);
  }
  for(const p of freePerformance){if(selectedIds.has(p.id))continue;selectedIds.add(p.id);const field={id:`${t.id}-field-${m.fields.length+1}`,label:'业绩报酬',status:'pending',before:p.text,after:'',factIds:[],paragraphIds:[p.id]};m.fields.push(field);reviewParagraph(task,m,field,p,rs);}
  // Direct identity repeats are safe only in title/disclosure text, never performance prose or peer tables.
  for(const [id,{atoms}] of headings?.dates||[]){const pending=atoms.filter(a=>!task._records.some(r=>r.m===m&&r.p.id===id&&r.atom.key===a.key&&r.atom.start===a.start));if(!pending.length)continue;const p=map.get(id),field={id:`${t.id}-field-${m.fields.length+1}`,label:'开售/开放日期',before:p.text,after:'',factIds:[],paragraphIds:[id],status:'pending'};m.fields.push(field);for(const atom of pending)addOccurrence(task,m,field,p,atom,pendingAnnouncement());}
  const old=m.templateProductName;const oldShort=old?.replace(new RegExp(PRODUCT_END+'$'),'');
  for(const p of t.paragraphs){if(selectedIds.has(p.id)||p.tableId||headings?.excluded.has(p.id)||kind==='ppt'&&!m.selectedPages.includes(p.page)||MARKETING.test(p.text))continue;
   const code=p.text.match(/(?:^|\n)\s*(?:产品代码|产品编码)\s*[：:（(【\s]*([A-Z0-9]{5,25})\s*[）)】]?\s*$/);
   if(code){const f={id:`${t.id}-field-${m.fields.length+1}`,label:'产品代码',before:p.text,after:'',factIds:[],paragraphIds:[p.id],status:'pending'};m.fields.push(f);const atom={key:'product_code',start:p.text.indexOf(code[1]),before:code[1],value:code[1]};addOccurrence(task,m,f,p,atom,resolveAtom('product_code',f.label,rs,scopeFor(f.label,p.text),target));}
   if(/本(?:信托|集合)?计划|本产品|本材料/.test(p.text)&&!/(?:曾在|历任|履历)/.test(p.text))for(const a of p.text.matchAll(/([\u4e00-\u9fffA-Za-z]{2,65}?(?:股份有限公司|有限责任公司|有限公司|证券|资产管理|资管|信托|银行))(?:作为|担任)(?:本产品|本信托计划|本集合计划|本计划)?的?投资顾问/g)){
    const before=a[1].split(/(?:由|聘请|委托|指定)/).at(-1);if(!before||before.length<3)continue;const field={id:`${t.id}-field-${m.fields.length+1}`,label:'投资顾问',before:p.text,after:'',factIds:[],paragraphIds:[p.id],status:'pending'};m.fields.push(field);const atom={key:'advisor',start:a.index+a[1].lastIndexOf(before),before,value:before};addOccurrence(task,m,field,p,atom,resolveAtom('advisor',field.label,rs,scopeFor(field.label,p.text),target));
   }
   let identitySource=coverNames.get(p.id)?.before||headings?.names.get(p.id)?.before||(old&&p.text.includes(old)?old:oldShort&&oldShort.length>4&&p.text.includes(oldShort)?oldShort:null);
   if(!identitySource&&old&&[norm(old),norm(oldShort)].includes(norm(p.text)))identitySource=p.text.trim();
   if(kind==='poster'&&REFERENCE_TITLE.test(p.text))identitySource=null;
   if(identitySource&&p.text.length<150){const full=new RegExp(PRODUCT_END+'$').test(identitySource);const to=coverNames.get(p.id)?.after||(full?target.name:target.shortName);const start=p.text.indexOf(identitySource);const f={id:`${t.id}-field-${m.fields.length+1}`,label:'产品名称',before:p.text,after:'',factIds:[],paragraphIds:[p.id],status:'pending'};m.fields.push(f);addOccurrence(task,m,f,p,{key:'product_name',start,before:identitySource,value:identitySource,renderValue:to},{status:'verified',value:target.name,evidence:target.evidence,reason:'同步产品标题中的名称。'});}
   if(/(?:本(?:集合|信托)?计划|本产品|管理人).{0,30}(?:风险等级|风险评级|评级)/.test(p.text)&&/R[1-5]/.test(p.text)){const f={id:`${t.id}-field-${m.fields.length+1}`,label:'风险等级',before:p.text,after:'',factIds:[],paragraphIds:[p.id],status:'pending'};m.fields.push(f);reviewParagraph(task,m,f,p,rs);}
  }
  m.productElementsLocated=m.fields.some(f=>!['company','manager'].includes(f.source));
  if(!m.productElementsLocated)m.warnings.push('未定位到可更新的产品要素，无法完成该模板的产品要素更新；其他模板可以继续处理。');
 }
 return finalize(task);
}
const wholeSupplementKey=key=>['opening','opening_fixed_monthdays','lockup','duration','warning_stop'].includes(key)||key.startsWith('investment_')||key.startsWith('custom:')&&key!=='custom:performance_parameter';
const cleanSupplementText=value=>String(value??'').replace(/[\r\n\t]+/g,' ').replace(/ {2,}/g,' ').trim();
function feeClause(text,start,length){
 const left=text.slice(0,start),tail=text.slice(start+length),a=Math.max(...['；',';','。','\n'].map(s=>left.lastIndexOf(s)))+1,b=tail.search(/[；;。\n]/);
 return text.slice(a,start)+'@'+(b<0?tail:tail.slice(0,b));
}
function feeCondition(clause){
 let text=clause.normalize('NFKC').replace(/\s/g,'').replace(/(?:持有天数|持有时间|持有期限|持有期|持有|D)(?:为|是)?/gi,'持有');
 text=text.replace(/不少于|不低于|大于等于|不小于|≥|>=/g,'≥').replace(/不超过|不高于|小于等于|≤|<=/g,'≤').replace(/小于|少于|不足|不满|未满/g,'<').replace(/大于|超过/g,'>').replace(/自然日|天|日/g,'天');
 text=text.replace(/(\d+(?:\.\d+)?)(天|个月|月|年|万元|元)(?:及以上|以上)/g,'≥$1$2').replace(/(\d+(?:\.\d+)?)(天|个月|月|年|万元|元)(?:以下)/g,'<$1$2').replace(/持有满/g,'持有≥');
 const conditions=[...text.matchAll(/[<>≥≤]\d+(?:\.\d+)?(?:天|个月|月|年|万元|元)?|\d+(?:\.\d+)?(?:天|个月|月|年|万元|元)?(?:至|到|[-~～])\d+(?:\.\d+)?(?:天|个月|月|年|万元|元)/g)].map(m=>m[0]);
 if(!conditions.length)return '';
 // Keep other qualifiers (for example early exit or a special class), while
 // normalizing only controlled synonyms for the fee and holding period.
 let rest=text;for(const condition of conditions)rest=rest.replace(condition,'');
 for(const [,label]of FEE_LABELS)rest=rest.replace(new RegExp(label.source,'g'),'');
 rest=rest.replace(/持有|费率|比例|收取|按|为|是|的|[(),，:：@。]/g,'');
 return conditions.join('|')+'|'+rest;
}
function scalarFee(fact,value){
 assertFeeActionScope(value,fact.scope);
 let raw=String(value).normalize('NFKC').trim(),label=FEE_LABELS.find(([key])=>key===fact.key)?.[1];
 if(label&&new RegExp('^(?:不收取|不收|免收|免除)(?:'+label.source+')[。.]?$').test(raw))return '0%';
 if(label)raw=raw.replace(new RegExp('^(?:'+label.source+')(?:率)?\\s*(?:为|是)?\\s*[:：]?\\s*'),'');
 raw=raw.replace(/^(?:年费率|费率)\s*(?:为|是)?\s*[:：]?\s*/,'').replace(/\s*(?:\/年|每年|年化)[。.]?$/,'').replace(/[。.]$/,'').trim();
 if(!/^\d+(?:\.\d+)?\s*%?$/.test(raw)||number(raw)>100)throw new Error('请填写0至100之间的费率数字，可附本项名称和“%/年”；分档费率请保留各档条件填写完整内容。');
 return raw.replace(/[%\s]/g,'')+'%';
}
function feeNeedsConditions(task,fact){
 const rs=task._records.filter(r=>r.f===fact),byParagraph=new Map(),conditions=new Set();
 for(const r of rs){const key=r.m.materialId+'|'+r.p.id,list=byParagraph.get(key)||[];list.push(r);byParagraph.set(key,list);const c=feeCondition(feeClause(r.p.text,r.atom.start,r.atom.before.length));if(c)conditions.add(c);}
 return conditions.size>1||[...byParagraph.values()].some(rs=>rs.length>1&&new Set(rs.map(r=>norm(r.atom.before))).size>1);
}
function occurrenceValue(r,value){return{materialId:r.m.materialId,fieldId:r.field.id,paragraphId:r.p.id,start:r.atom.start,before:r.atom.before,context:r.p.text,state:JSON.stringify([r.f.status,r.f.value,r.userRenderValue??null]),value};}
function feeOccurrencePreview(task,fact,raw,groupId){
 const rs=task._records.filter(r=>r.f===fact),tokens=feeTokens(raw,fact.key,/信托/.test(task.target.name)).filter(t=>t.key===fact.key);
 if(!tokens.length)throw new Error('此项包含多个费率或分档条件，请按各档原条件填写完整费率，不能用一个数字覆盖所有档。');
 const candidates=tokens.map(t=>({condition:feeCondition(feeClause(raw,t.start,t.before.length)),value:scalarFee(fact,t.value)}));
 if(candidates.some(c=>!c.condition)||new Set(candidates.map(c=>c.condition)).size!==candidates.length)throw new Error('分档费率的条件不完整或重复，无法逐档对应；请保留各档持有期限等条件。');
 const used=new Set(),occurrenceValues=rs.map(r=>{const condition=feeCondition(feeClause(r.p.text,r.atom.start,r.atom.before.length)),matches=candidates.filter(c=>c.condition===condition);if(matches.length!==1)throw new Error('补充内容与模板的分档条件无法一一对应，请保留原持有期限和适用条件后修改费率。');used.add(condition);return occurrenceValue(r,matches[0].value);});
 if(used.size!==candidates.length)throw new Error('补充中有模板未包含的分档条件；请核对完整条件后再应用。');
 return{groupId,text:raw,updates:[{factId:fact.id,key:fact.key,label:fact.label,value:raw,occurrenceValues}],warnings:[]};
}
function redemptionTierSupplement(task,group,facts,raw){
 if(facts.length===1&&!facts[0].redemptionTier.warning&&/^\s*(?:(?:退出|赎回)费(?:费率|率)?\s*(?:为|是)?\s*[:：]?\s*)?\d+(?:\.\d+)?\s*[%％]?[。.]?\s*$/.test(raw.normalize('NFKC'))){
  const fact=facts[0],preview=previewUpdates(task,facts,{redemption_fee:scalarFee(fact,raw)},raw,group.id);
  for(const update of preview.updates)update.occurrenceValues=task._records.filter(r=>r.f===fact).map(r=>occurrenceValue(r,update.value));return supplementChanges(task,preview);
 }
 const candidates=feeTokens(raw,'redemption_fee',false).filter(t=>t.key==='redemption_fee').map(token=>({token,tier:matchTemplateRedemptionTier(raw,token)}));
 if(!candidates.length||candidates.some(c=>!c.tier?.condition||c.tier.warning))throw new Error('请填写完整分档文字，包括每档持有天数和退出费率，例如“持有天数＜180自然日，退出费率1%；持有天数≥180自然日，退出费率0%”。');
 if(new Set(candidates.map(c=>c.tier.condition)).size!==candidates.length)throw new Error('补充中存在重复或不同口径的持有期限，请保留各档条件后填写。');
 const updates=[],used=new Set();
 for(const fact of facts){
  const matches=candidates.filter(c=>c.tier.condition===fact.redemptionTier.condition&&['share','channel','issuer'].every(k=>!c.tier.scope?.[k]||c.tier.scope[k]===fact.scope[k]));
  if(matches.length!==1)throw new Error('补充内容与本卡片的持有期限无法逐档对应，请填写各档完整条件和费率；本次未修改。');
  const {token}=matches[0];used.add(matches[0]);for(const update of previewUpdates(task,[fact],{redemption_fee:token.value},raw,group.id).updates){update.occurrenceValues=task._records.filter(r=>r.f===fact).map(r=>occurrenceValue(r,update.value));updates.push(update);}
 }
 if(used.size!==candidates.length)throw new Error('补充中有本卡片未包含的持有期限或份额，请核对后再应用。');
 return supplementChanges(task,{groupId:group.id,text:raw,updates,warnings:[]});
}
function wholeOccurrencePreview(task,fact,raw,groupId){
 validateSupplement(fact,raw,task);const rs=task._records.filter(r=>r.f===fact),fields=new Map();
 for(const r of rs){const key=r.m.materialId+'|'+r.field.id;if(!fields.has(key))fields.set(key,[]);fields.get(key).push(r);}
 const lines=String(raw).split(/\r?\n/).filter(s=>s.trim()),occurrenceValues=[];
 for(const records of fields.values()){
  if(records.length>1&&lines.length!==records.length)throw new Error('此项原文分为'+records.length+'段，请按原段落分成'+records.length+'行填写完整内容后预览。');
  records.forEach((r,i)=>occurrenceValues.push(occurrenceValue(r,cleanSupplementText(records.length===1?raw:lines[i]))));
 }
 return{groupId,text:raw,updates:[{factId:fact.id,key:fact.key,label:fact.label,value:cleanSupplementText(raw),occurrenceValues}],warnings:[],notes:/[\r\n\t]/.test(raw)?['已将同一段内的换行或制表符转为空格，保留模板原有段落排版。']:[]};
}
function supplementChanges(task,preview){
 const changes=new Map(),ids=new Set((preview.updates||[]).map(u=>u.factId));
 const render=(text,replacements)=>{for(const r of [...replacements].sort((a,b)=>b.start-a.start))text=text.slice(0,r.start)+r.after+text.slice(r.start+r.before.length);return text;};
 for(const update of preview.updates||[])for(const r of task._records.filter(r=>r.f.id===update.factId)){
  const item=update.occurrenceValues?.find(x=>x.materialId===r.m.materialId&&x.fieldId===r.field.id&&x.paragraphId===r.p.id&&x.start===r.atom.start);
  let value=item?.value??update.value;
  if(r.atom.key==='risk_description')value=({R1:'低风险',R2:'中低风险',R3:'中风险',R4:'中高风险',R5:'高风险'})[value]||value;
  const percentage=r.atom.before.match(/^(\d+(?:\.(\d+))?)(\s*[】\]]?\s*[%％])$/);if(/^\d+(?:\.\d+)?%$/.test(value)&&percentage)value=number(value).toFixed(Math.max(percentage[2]?.length||0,value.split('.')[1]?.replace('%','').length||0))+percentage[3];
  const operations=r.m.operations.filter(o=>o.paragraphId===r.p.id&&o.kind==='replace'),key=r.m.materialId+'|'+r.p.id,itemChange=changes.get(key)||{label:r.field.label,original:r.p.text,before:render(r.p.text,operations),after:'',materialName:r.m.template.name,materialId:r.m.materialId,paragraphId:r.p.id,replacements:operations.filter(o=>!o.factIds?.some(id=>ids.has(id)))};
  itemChange.replacements.push({start:r.atom.start,before:r.atom.before,after:String(value)});changes.set(key,itemChange);
 }
 for(const c of changes.values()){c.after=render(c.original,c.replacements);delete c.replacements;delete c.original;}
 // A compacted field owns one whole-paragraph operation. Rebuild all its fees
 // for preview, including untouched facts, then apply the same simplification.
 for(const m of task.materials)for(const field of m.fields.filter(f=>f.equalFeeCompacted&&f.factIds.some(id=>ids.has(id)))){
  const scratch={operations:[]};
  for(const r of task._records.filter(r=>r.m===m&&r.field===field)){
   const u=preview.updates.find(u=>u.factId===r.f.id);
   const item=u?.occurrenceValues?.find(x=>x.materialId===m.materialId&&x.fieldId===field.id&&x.paragraphId===r.p.id&&x.start===r.atom.start);
   renderRecord({...r,m:scratch,...(u?{f:{...r.f,value:u.value,status:'verified',userConfirmed:false},userRenderValue:item?.value??u.value}:{})});
  }
  const after=field.paragraphIds.map(id=>render(m.template.paragraphs.find(p=>p.id===id)?.text||'',scratch.operations.filter(o=>o.paragraphId===id&&o.kind==='replace'))).join('\n');
  for(const id of field.paragraphIds)changes.delete(m.materialId+'|'+id);
  changes.set(m.materialId+'|'+field.id,{label:field.label,before:field.after,after:simplifyEqualFeeTiers(after),materialName:m.template.name,materialId:m.materialId});
 }
 preview.changes=[...changes.values()];return preview;
}
function validateSupplement(fact,value,task,perOccurrence=false){
 if(fact.customSectionOverrideId)throw new Error('此项已自定义整项替换，请继续在自定义替换中编辑，或撤销后分别补充原文。');
 if(fact.feeTextOverrideId)throw new Error('此项已按整段文字修改，请在本卡片重新填写完整内容，或撤销整段修改后调整数字。');
 if(fact.sectionOverrideId)throw new Error('此项已采用合同内容；请先撤销采用，再单独补充。');
 const raw=String(value??'').trim();if(!raw)throw new Error('请输入补充值。');let accepted=raw;
 if(fact.key.endsWith('_fee')){if(!perOccurrence&&!(fact.redemptionTier?.condition&&!fact.redemptionTier.warning)&&feeNeedsConditions(task,fact))throw new Error('此项包含多个分档费率，请保留各档适用条件填写完整内容，不能填写一个数字更新所有档。');accepted=scalarFee(fact,raw);}
 else if(fact.key.startsWith('performance_')&&fact.key!=='performance_policy'||fact.key==='benchmark'){if(!/^\d+(?:\.\d+)?\s*[%％]?$/.test(raw)||number(raw)>100)throw new Error('请填写0至100之间的费率数字，可带百分号。');accepted=raw.replace(/[％\s]/g,'').replace(/%$/,'')+'%';}
 else if(/^(?:minimum_amount|additional_amount|opening_cycle|opening_days|opening_monthday|lockup_days|duration_years)$/.test(fact.key)){
  if(['minimum_amount','additional_amount','lockup_days','duration_years'].includes(fact.key)){
   const fields=new Map();for(const r of task._records.filter(r=>r.f===fact)){const key=r.m.materialId+'|'+r.field.id;fields.set(key,(fields.get(key)||0)+1);}if([...fields.values()].some(n=>n>1))throw new Error('本项原文包含多个不同条件的参数，不能用一个数字同时更新；请核对各处原文和适用条件。');
  }
  const types={minimum_amount:['首次认购|认购起点|最低认购|起购金额|参与起点|首次参与','万元?'],additional_amount:['追加认购|追加参与|追加金额','万元?'],opening_cycle:['开放周期|每满|每隔','个月|月'],opening_days:['开放天数|连续开放|开放期|开放','个交易日|交易日'],opening_monthday:['每月开放日|每月','日|号'],lockup_days:['持有期|锁定期|封闭期','自然日|天|日'],duration_years:['存续期限|存续期|产品期限','年']},[labels,units]=types[fact.key];
  const input=raw.normalize('NFKC').replace(new RegExp('^(?:'+labels+')\\s*(?:为|是)?\\s*[:：]?\\s*'),'').replace(/[。.]$/,'').trim(),m=input.match(new RegExp('^(\\d+(?:\\.\\d+)?)\\s*(?:'+units+')?$'));
  if(!m)throw new Error('请填写有效数字及本项相应单位，不能混用天、月、年或金额单位。');if(!/amount/.test(fact.key)&&!Number.isInteger(Number(m[1])))throw new Error('日期、天数及周期请填写整数。');accepted=m[1];if(fact.key==='opening_monthday'&&(Number(accepted)<1||Number(accepted)>31))throw new Error('每月日期应为1至31日。');
 }
 else if(fact.key==='opening_weekday'){if(!/^(?:周|星期|每周)?[一二三四五六日天]$/.test(raw))throw new Error('请填写星期几，例如周四。');accepted=raw.at(-1);}
 else if(fact.key==='opening_monthly_weekday'){accepted=parseMonthlyWeekdayInput(raw);for(const r of task._records.filter(r=>r.f===fact)){const alignment=monthlyWeekdayChanges(r.atom.monthlyPattern,accepted);if(alignment.warning)throw new Error(alignment.warning);}}
 else if(fact.key==='product_code'){accepted=raw.normalize('NFKC').replace(/^(?:产品代码|产品编码)\s*(?:为|是)?\s*[:：]?\s*/,'').trim();if(!/^[A-Za-z0-9][A-Za-z0-9_.-]{4,24}$/.test(accepted))throw new Error('请填写5至25位产品代码，仅含字母、数字或代码连接符。');}
 else if(['risk_rating','risk_description'].includes(fact.key)){
  const map={'低风险':'R1','中低风险':'R2','中风险':'R3','中高风险':'R4','高风险':'R5'},text=raw.normalize('NFKC').toUpperCase().replace(/^(?:风险等级|风险评级|产品评级)\s*(?:为|是)?\s*[:：]?\s*/,'').replace(/[。.]$/,'').trim(),codes=unique(text.match(/R\d/g)||[]),words=unique(text.match(/中低风险|中高风险|中风险|低风险|高风险/g)||[]);
  const values=unique([...codes,...words.map(w=>map[w])]),rest=text.replace(/R[1-5]|中低风险|中高风险|中风险|低风险|高风险|[()\s]/g,'');if(values.length!==1||!/^R[1-5]$/.test(values[0])||rest)throw new Error('请填写R1至R5或相应风险等级名称，代码与中文风险等级须一致。');accepted=values[0];
 }
 else if(fact.key==='opening_announced_date')accepted=parseDateInput(raw,Number(task.month.slice(0,4))).value;
 else if(wholeSupplementKey(fact.key)){
  const records=task._records.filter(r=>r.f===fact),local=records.some(r=>r.atom.start!==0||r.atom.before!==r.p.text);
  if(local){
   if(records.some(r=>!/^\d+(?:\.\d+)?$/.test(r.atom.before)))throw new Error('此项仅定位到局部内容，无法将完整表述直接写入，请核对原文位置。');
   const scalar=raw.normalize('NFKC').trim();
   if(/^\d+(?:\.\d+)?$/.test(scalar))accepted=scalar;
   else{
    const numbers=[...scalar.matchAll(/\d+(?:\.\d+)?/g)],skeleton=numbers.length===1?norm(scalar.replace(numbers[0][0],'#')):'';
    if(numbers.length!==1||records.some(r=>norm(r.p.text.slice(0,r.atom.start)+'#'+r.p.text.slice(r.atom.start+r.atom.before.length))!==skeleton))throw new Error('此项已定位为局部数字；请只填写该数字，或保留原话术及条件、仅调整这个数字后再预览。');
    accepted=numbers[0][0];
   }
  }else{
   if(/^\d+(?:\.\d+)?\s*(?:%|％|天|日|个月|月|年)?$/.test(raw)&&records.some(r=>/[\u4e00-\u9fff]/.test(r.atom.before)))throw new Error('此项按完整原文补充，请填写完整表述，不能用单个数字替换整段。');accepted=cleanSupplementText(raw);
  }
 }
 return accepted;
}
function rebuildFacts(task,ids){
 const selected=new Set(ids);for(const m of task.materials)m.operations=m.operations.filter(o=>!o.factIds?.some(id=>selected.has(id)));
 for(const r of task._records)if(selected.has(r.f.id))renderRecord(r);return finalize(task);
}
export function previewContractSection(task,factIds){if(task.facts.some(f=>factIds?.includes(f.id)&&f.customSectionOverrideId))throw new Error('此项已自定义替换，请先撤销，再采用合同内容。');return previewSectionAdoption(task,factIds);}
export function applyContractSection(task,preview){previewContractSection(task,preview?.factIds);adoptSection(task,preview);return finalize(task);}
export function undoContractSection(task,id){const ids=releaseSection(task,id);return rebuildFacts(task,ids);}
export const previewCustomSection=previewCustomSectionText;
export function applyCustomSection(task,preview){applyCustomSectionOverride(task,preview);return finalize(task);}
export function undoCustomSection(task,id){return rebuildFacts(task,releaseCustomSection(task,id));}
const GROUP_KIND=f=>f.key==='performance_policy'?'performance_policy':f.key==='opening_announced_date'?'date':f.key.startsWith('performance_')||f.key==='custom:performance_parameter'?'performance':f.key==='benchmark'?'benchmark':f.key==='product_code'?'code':/^opening(?:_|$)/.test(f.key)?'opening':'text';
export function getSupplementGroups(task){
 const groups=new Map();
 for(const f of task.facts){
  const kind=f.redemptionTier?'redemption_tiers':GROUP_KIND(f),scope={...f.scope};let identity=f.id;
  if(kind==='redemption_tiers'){delete scope.condition;identity=kind+'|'+JSON.stringify(scope);}
  if(['performance','benchmark','code','opening','date'].includes(kind)){
   delete scope.condition;
   if(kind==='date'){scope.event=(f.scope.condition||'').split('|')[0]||'opening';if(scope.event==='sale')scope.event='opening';}
   // Whole unmatched prose has its own meaning and must not share a replacement.
   identity=kind==='opening'&&f.key==='opening'?f.id:kind+'|'+JSON.stringify(scope);
  }
  const group=groups.get(identity)||{id:'supplement-'+f.id,label:({redemption_tiers:'退出费率',performance:'业绩报酬',benchmark:'业绩比较基准',code:'产品代码',opening:'开放期',date:({sale:'本次开售日',reservation:'本次预约日',opening:'本次开放日'})[scope.event]||'本次开放日'})[kind]||f.label,kind,factIds:[],scope};
  group.factIds.push(f.id);groups.set(identity,group);
 }
 return [...groups.values()].map(g=>{const fs=g.factIds.map(id=>task.facts.find(f=>f.id===id));return{...g,status:fs.some(f=>f.status==='conflict')?'conflict':fs.some(f=>f.status!=='verified')?'pending':'verified',userConfirmed:fs.filter(f=>f.userConfirmed).length>0};});
}
function previewUpdates(task,facts,values,text,groupId){
 const updates=[],warnings=[];
 if(facts.some(f=>f.key==='custom:performance_parameter'&&f.status!=='verified'))warnings.push('业绩报酬中仍有含义未明确的百分比，本次只更新已识别的计提基准和提取比例，其余原值保留待核。');
 for(const [key,value]of Object.entries(values)){
  const matches=facts.filter(f=>f.key===key);
  if(!matches.length){warnings.push('模板未识别到'+({benchmark:'业绩比较基准',product_code:'产品代码',opening_announced_date:'本次开放日期'}[key]||key)+'对应位置，未写入。');continue;}
  for(const f of matches){const accepted=validateSupplement(f,value,task);
   if(f.key==='opening_announced_date')for(const r of task._records.filter(r=>r.f===f))formatDateLike(accepted,r.atom.before,r.atom.dateRole,Number(task.month.slice(0,4)));
   updates.push({factId:f.id,key:f.key,label:f.label,value:accepted});
  }
 }
 return{groupId,text,updates,warnings};
}
const dayString=d=>`${d.year}/${d.month}/${d.day}`;
const sameDays=(a,b)=>a.length===b.length&&a.every((d,i)=>d===b[i]);
function splitNumber(value){if(/^\d+$/.test(value))return Number(value);if(value.includes('十')){const [a,b]=value.split('十');return (a?cnum(a):1)*10+(b?cnum(b):0);}return cnum(value);}
function segmentRule(text){
 const marker=String(text).match(/(?:分段规则|开放分段|开放安排)\s*[:：]\s*([^\n]+)/),source=marker?.[1]||String(text),n='([一二两三四五六七八九十\\d]+)';
 const both=source.match(new RegExp('(?:前|首)\\s*'+n+'\\s*个?\\s*(?:开放日|交易日)[^，,;；。\\n]{0,20}(?:参与|申购)(?:和|及|与|、)?(?:退出|赎回)'));
 const only=source.match(new RegExp('后\\s*'+n+'\\s*个?\\s*(?:开放日|交易日)[^，,;；。\\n]{0,20}仅(?:开放|可|办理)?(?:参与|申购)'));
 // Excel opening remarks share this label even when they only say to follow an announcement.
 const explicit=!!both||!!only||/(?:前|首|后)\s*(?:[一二两三四五六七八九十百\d几若干]+\s*个?\s*)?(?:开放日|交易日)/.test(source);
 if(!explicit)return null;
 const first=both?splitNumber(both[1]):0,last=only?splitNumber(only[1]):0;
 if(!Number.isInteger(first)||!Number.isInteger(last)||first<=0||last<=0)return{warning:'分段规则未同时明确前段参与和退出、后段仅参与的交易日天数，分段日期保留待核。'};
 return{first,last,rule:`前${first}个开放日参与和退出，后${last}个开放日仅参与`};
}
function oldWindowInterval(records,window,year){
 const rs=records.filter(r=>(r.atom.dateWindow||'')===window),starts=unique(rs.filter(r=>r.atom.dateRole==='start').map(r=>r.atom.before)),ends=unique(rs.filter(r=>r.atom.dateRole==='end').map(r=>r.atom.before));
 if(starts.length!==1||ends.length!==1)return null;
 try{const start=parseDateInput(starts[0],year).start,range=parseDateInput(dayString(start)+'至'+ends[0],start.year);return range.end?range:null;}catch{return null;}
}
function inferSegmentRule(records,year){
 const fields=new Map();for(const r of records){const key=r.m.materialId+'|'+r.field.id;if(!fields.has(key))fields.set(key,[]);fields.get(key).push(r);}
 const counts=[],sources=[];
 for(const rs of fields.values()){
  if(!rs.some(r=>r.atom.dateWindow))continue;
  const text=rs[0].field.before||'',rule=segmentRule(text);
  if(rule){if(rule.warning)return rule;counts.push({...rule,ruleSource:'template_rule'});continue;}
  const inferredYear=Number(text.match(/(?:^|\D)(20\d{2})[年/.-]/)?.[1]||year),first=oldWindowInterval(rs,'participation_exit',inferredYear),last=oldWindowInterval(rs,'participation_only',inferredYear);
  if(!first||!last)return{warning:'模板的分段日期不完整，无法确定各段交易日天数；请补充完整分段规则。'};
  let a,b,span;try{a=getTradingDays(first.start,first.end);b=getTradingDays(last.start,last.end);span=getTradingDays(first.start,last.end);}catch{return{warning:'模板原分段日期先后顺序有误，无法可靠推算，分段日期保留待核。'};}
  if(a.warning||b.warning||span.warning)return{warning:a.warning||b.warning||span.warning};
  if(!a.dates.length||!b.dates.length||!sameDays([...a.dates,...b.dates],span.dates))return{warning:'模板原分段日期存在空档或重叠，无法可靠确定交易日划分，分段日期保留待核。'};
  if(a.dates[0]!==dayString(first.start)||a.dates.at(-1)!==dayString(first.end)||b.dates[0]!==dayString(last.start)||b.dates.at(-1)!==dayString(last.end))return{warning:'模板原分段边界不是沪深交易日，无法据此推算，分段日期保留待核。'};
  const overall=oldWindowInterval(rs,'',inferredYear);if(overall){const all=getTradingDays(overall.start,overall.end);if(all.warning)return{warning:all.warning};if(!sameDays(all.dates,span.dates))return{warning:'模板原分段未完整覆盖开放区间，分段日期保留待核。'};}
  counts.push({first:a.dates.length,last:b.dates.length,ruleSource:'template_dates',rule:`前${a.dates.length}个开放日参与和退出，后${b.dates.length}个开放日仅参与`});sources.push(...a.sources,...b.sources,...span.sources);
 }
 if(!counts.length)return{warning:'未找到可确认的完整分段规则，分段日期保留待核。'};
 if(new Set(counts.map(c=>c.first+'|'+c.last)).size>1)return{warning:'不同模板的开放分段天数不一致，请明确本次分段规则；分段日期保留待核。'};
 return{...counts[0],sources};
}
// End-only calendars preserve endpoint meaning. Infer the other endpoint only
// from a same-scope total duration, never from a default number of weekdays.
function openingDayCounts(text){
 const n='([一二两三四五六七八九十\\d]+)',patterns=[
  new RegExp('(?:每个|每次|本次|各个)?开放期(?:均|连续)?(?:为期|为|是|持续|共|开放)?\\s*[:：]?\\s*(?:连续)?\\s*[【]?'+n+'[】]?\\s*个?(?:连续)?交易日','g'),
  new RegExp('(?:连续开放|开放)\\s*[【]?'+n+'[】]?\\s*个?(?:连续)?交易日','g'),
  new RegExp('连续\\s*[【]?'+n+'[】]?\\s*个?交易日(?:为|作为)(?:一个)?开放期','g'),
 ];
 const counts=[];for(const [index,re]of patterns.entries())for(const m of String(text).matchAll(re)){
  const before=String(text).slice(0,m.index).split(/[。；;，,\n]/).at(-1)||'';
  if(index===1&&/(?:仅|赎回|退出|申购|参与|前段|后段)/.test(before))continue;
  counts.push(splitNumber(m[1]));
 }
 return unique(counts.filter(n=>Number.isInteger(n)&&n>0&&n<=366));
}
function sameOpeningScope(a,b){return ['share','channel','issuer'].every(k=>(a?.[k]||'')===(b?.[k]||''));}
function scopedOpeningBlocks(task,scope){
 const allowed=[],ambiguous=[];let inherited={share:'',channel:''},underlying=false;
 for(const b of task.contract.blocks||[]){
  const text=String(b.text||'').trim(),short=text.length<70&&!/[。；;\d%％]/.test(text);
  if(short&&/^(?:第[一二三四五六七八九十百]+[章节]|[一二三四五六七八九十百]+、)/.test(text)){inherited={share:'',channel:''};underlying=false;}
  if(short&&/^(?:本计划|本产品|本集合计划|本信托计划)/.test(text))underlying=false;
  const other=/(?:底层(?:基金|产品|资产管理计划|信托)?|投资标的|标的产品|其他产品|假设|例如|举例)/.test(text);
  if(short&&other)underlying=true;
  const local=scopeFor('',text);
  if(short&&(/^[（(]?[一二三四五六七八九十\d]*[）).、]?\s*[A-Z]类(?:份额)?/.test(text)||/^渠道[:：]/.test(text)))inherited={share:local.share||'',channel:local.channel||''};
  if(other||underlying)continue;
  const share=local.share||inherited.share,channel=local.channel||inherited.channel;
  if(share&&share!==scope.share||channel&&channel!==scope.channel){if(!scope.share&&share||!scope.channel&&channel)ambiguous.push(b);continue;}
  allowed.push(b);
 }
 return{allowed,ambiguous};
}
function deadlineDuration(task,group,records,year){
 const scoped=scopedOpeningBlocks(task,group.scope),candidates=[],ambiguous=scoped.ambiguous.filter(b=>openingDayCounts(b.text).length);
 for(const b of scoped.allowed){
  const counts=openingDayCounts(b.text);if(!counts.length){if(/(?:每个|每次|本次|各个)开放期[^。；;\n]{0,22}(?:为|是|持续|共|开放)[^。；;\n]{0,18}交易日/.test(b.text))ambiguous.push(b);continue;}
  if(/假设|例如|举例|首次开放期|首个开放期|特殊开放期|临时开放期/.test(b.text)){ambiguous.push(b);continue;}
  for(const count of counts)candidates.push({count,evidence:quote(b)});
 }
 const contractCounts=unique(candidates.map(c=>c.count));
 if(contractCounts.length>1||ambiguous.length)return{warning:'合同开放天数存在不同条件或口径，尚不能确定本次适用天数，未反推起始日。'};
 if(contractCounts.length===1)return{count:contractCounts[0],source:'contract',evidence:candidates.map(c=>c.evidence)};
 const fields=new Map();for(const r of task._records)if(/^opening(?:_|$)/.test(r.f.key)&&sameOpeningScope(r.f.scope,group.scope))fields.set(r.m.materialId+'|'+r.field.id,r.field);
 const explicit=unique([...fields.values()].flatMap(f=>openingDayCounts(f.before)));
 if(explicit.length>1)return{warning:'模板明文开放天数不一致，未反推起始日，请核对本次总天数。'};
 if(explicit.length===1)return{count:explicit[0],source:'template_rule',evidence:[...fields.values()].filter(f=>openingDayCounts(f.before).length).map(f=>({quote:f.before}))};
 const groups=new Map();for(const r of records){const key=r.m.materialId+'|'+r.field.id;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);}
 const inferred=[],sources=[];
 for(const rs of groups.values()){
  const inferredYear=Number((rs[0].field.before||'').match(/(?:^|\D)(20\d{2})[年/.-]/)?.[1]||year),interval=oldWindowInterval(rs,'',inferredYear);if(!interval)continue;
  const days=getTradingDays(interval.start,interval.end);if(days.warning)return{warning:days.warning};
  if(!days.dates.length||days.dates[0]!==dayString(interval.start)||days.dates.at(-1)!==dayString(interval.end))return{warning:'模板原开放区间边界不是沪深交易日，无法据此反推本次起始日。'};
  inferred.push(days.dates.length);sources.push(...days.sources);
 }
 const counts=unique(inferred);
 if(counts.length>1)return{warning:'不同模板原开放区间的交易日天数不一致，未反推起始日。'};
 if(!counts.length)return{warning:'尚无合同明确开放天数、模板明文天数或完整旧开放区间，不能仅凭截止日反推起始日。'};
 return{count:counts[0],source:'template_dates',sources};
}
function backfillOpeningRange(task,group,facts,endText,raw){
 const year=Number(task.month.slice(0,4)),end=parseDateInput(endText,year);if(end.end)throw new Error('开放截止日请填写单个日期。');
 const records=task._records.filter(r=>facts.includes(r.f)),context=records.map(r=>r.field.before||r.p.text).join('\n');
 const contractContext=scopedOpeningBlocks(task,group.scope).allowed.filter(b=>/开放/.test(b.text)).map(b=>b.text).join('\n');
 if(/(?:香港|港股|境外|海外).{0,35}(?:交易日|工作日)|(?:交易日|工作日).{0,35}(?:香港|港股|境外|海外)/.test(context+'\n'+contractContext+'\n'+raw))return{warning:'开放安排涉及香港或境外交易日，不能仅按沪深日历反推起始日。'};
 const endpoint=getTradingDays(end.start,end.start);if(endpoint.warning)return{warning:endpoint.warning};
 if(!endpoint.dates.length)return{warning:'本次开放截止日不是沪深交易日，未静默平移或反推，请核对档期。'};
 const duration=deadlineDuration(task,group,records,year);if(duration.warning)return duration;
 const dates=[],sources=[...(duration.sources||[])];let stamp=Date.UTC(end.start.year,end.start.month-1,end.start.day);
 while(dates.length<duration.count){
  const date=new Date(stamp),day={year:date.getUTCFullYear(),month:date.getUTCMonth()+1,day:date.getUTCDate()},trading=getTradingDays(day,day);
  if(trading.warning)return{warning:trading.warning};dates.unshift(...trading.dates);sources.push(...trading.sources);stamp-=86400000;
 }
 const value=dates[0]+'至'+dates.at(-1),label=duration.source==='contract'?'合同明确的':duration.source==='template_rule'?'模板明文规定的':'模板原完整开放区间的';
 const note=`以 ${dayString(end.start)} 为开放截止日，按${label}${duration.count}个沪深交易日反推：${value}。周末及交易所休市日已排除，请确认后应用。`;
 return{value,note,basis:{kind:'trading_day_backfill',durationSource:duration.source,count:duration.count,end:dayString(end.start),range:value,durationEvidence:duration.evidence||[],calendarSources:[...new Map(sources.map(s=>[s.year,s])).values()]}};
}
function deriveOpeningWindows(task,group,facts,raw,parsed){
 const records=task._records.filter(r=>facts.includes(r.f)),warnings=[],notes=[];
 if(group.scope.event!=='opening'||!parsed?.end||!records.some(r=>r.atom.dateWindow))return{windows:{},warnings,notes};
 const context=records.map(r=>r.field.before||r.p.text).join('\n');
 if(/(?:香港|港股|境外|海外).{0,35}(?:交易日|工作日)|(?:交易日|工作日).{0,35}(?:香港|港股|境外|海外)/.test(context+'\n'+raw))return{windows:{},warnings:['此处开放安排涉及香港或境外交易日，不能仅按沪深日历推算，分段日期保留待核。'],notes};
 const trading=getTradingDays(parsed.start,parsed.end);if(trading.warning)return{windows:{},warnings:[trading.warning],notes};
 if(!trading.dates.length||trading.dates[0]!==dayString(parsed.start)||trading.dates.at(-1)!==dayString(parsed.end))return{windows:{},warnings:['补充区间的起止日期并非均为沪深交易日，未静默平移；分段日期保留待核，请核对档期。'],notes};
 const explicit=segmentRule(raw),rule=explicit?{...explicit,ruleSource:'explicit_rule'}:inferSegmentRule(records,parsed.start.year);
 if(rule.warning)return{windows:{},warnings:[rule.warning],notes};
 if(rule.first+rule.last!==trading.dates.length)return{windows:{},warnings:[`本次区间有${trading.dates.length}个沪深交易日，与分段规则的${rule.first}+${rule.last}天不一致，分段日期保留待核。`],notes};
 const windows={participation_exit:trading.dates[0]+'至'+trading.dates[rule.first-1],participation_only:trading.dates[rule.first]+'至'+trading.dates.at(-1)};
 const sources=[...new Map([...(rule.sources||[]),...trading.sources].map(s=>[s.year,s])).values()],basis={kind:'trading_day_split',ruleSource:rule.ruleSource,rule:rule.rule,range:parsed.value,calendarSources:sources};
 notes.push(`按沪深交易日及${explicit?'补充分段规则':'模板原分段规则'}计算：参与和退出 ${windows.participation_exit}；仅参与 ${windows.participation_only}。周末及交易所休市日已排除。`);
 return{windows,warnings,notes,basis};
}
function openingDependencyInvalidations(task,updates,invalidRuleGroups=[]){
 const supplied=new Map(updates.map(u=>[u.factId,u])),invalid=new Map(),year=Number(task.month.slice(0,4));
 for(const group of getSupplementGroups(task).filter(g=>g.kind==='date'&&g.scope.event==='opening')){
  const rs=task._records.filter(r=>group.factIds.includes(r.f.id)),overall=rs.filter(r=>!r.atom.dateWindow),touched=overall.filter(r=>supplied.has(r.f.id));if(!touched.length)continue;
  const changed=touched.some(r=>formatDateLike(supplied.get(r.f.id).value,r.atom.before,r.atom.dateRole,year)!==(r.userRenderValue??r.atom.before));
  const incomplete=touched.some(r=>!parseDateInput(supplied.get(r.f.id).value,year).end),ruleFailed=invalidRuleGroups.includes(group.id);if(!changed&&!incomplete&&!ruleFailed)continue;
  for(const r of rs){if(supplied.has(r.f.id)||!r.atom.dateWindow&&!(incomplete&&r.atom.dateRole==='end'))continue;
   if(r.f.status!=='verified'&&!r.f.userConfirmed&&r.userRenderValue==null)continue;
   invalid.set(r.f.id,{f:r.f,reason:r.atom.dateWindow?(ruleFailed?'本次分段规则未通过核对，旧分段确认失效，已恢复模板原日期待核。':'档期已变更，旧分段确认失效，已恢复模板原日期待核。'):'本次未补充完整区间，旧结束日确认失效，已恢复模板原日期待核。'});
  }
 }
 return [...invalid.values()];
}
function previewDates(task,group,facts,raw){
 const windows={participation_exit:raw.match(/(?:参与和退出|参与退出|申购和赎回)\s*[:：]\s*([^;；\n]+)/)?.[1],participation_only:raw.match(/(?:仅参与|仅开放参与|仅申购)\s*[:：]\s*([^;；\n]+)/)?.[1]};
 const announced=[...raw.matchAll(/(?:^|[;；\n])\s*(?:本次开放日|本次开放日期|开放日期|开放日|本次开售日|开售日期|开售日|募集日期|本次募集期|募集期)\s*[:：]\s*([^;；\n]+)/g)].map(m=>m[1].trim());
 if(new Set(announced.map(value=>parseDateInput(value,Number(task.month.slice(0,4))).value)).size>1)throw new Error('本次开放日与本次开售日为同一日期，但补充了不同值，请统一后再应用。');
 const deadlines=[...raw.matchAll(/(?:^|[;；\n])\s*(?:(?:本次)?(?:开放|开售)(?:截止日|结束日|截止日期|结束日期)|档期(?:结束|截止)日期)\s*[:：]\s*([^;；\n]+)/g)].map(m=>m[1].trim());
 if(new Set(deadlines.map(value=>parseDateInput(value,Number(task.month.slice(0,4))).value)).size>1)throw new Error('补充的开放截止日期不一致，请统一后再应用。');
 let general=announced[0]||raw.match(/(?:本次预约日|预约日期)\s*[:：]\s*([^;；\n]+)/)?.[1]||(!deadlines.length&&!Object.values(windows).some(Boolean)?raw.split(/[;；\n]/)[0]:'');
 const updates=[],warnings=[],notes=[];let generalParsed,backfilled;if(general)generalParsed=parseDateInput(general,Number(task.month.slice(0,4)));
 if(deadlines.length){
  if(group.scope.event!=='opening')return{groupId:group.id,text:raw,updates:[],warnings:['开放截止日不能用作预约日期，请在本次开放日事项中补充。'],notes:[]};
  const end=parseDateInput(deadlines[0],Number(task.month.slice(0,4)));if(end.end)throw new Error('开放截止日请填写单个日期。');
  if(generalParsed?.end&&dayString(generalParsed.end)!==dayString(end.start))throw new Error('开放区间与开放截止日期不一致，请核对后再应用。');
  if(generalParsed&&!generalParsed.end)general=dayString(generalParsed.start)+'至'+dayString(end.start);
  else if(!generalParsed){backfilled=backfillOpeningRange(task,group,facts,deadlines[0],raw);if(backfilled.warning)return{groupId:group.id,text:raw,updates:[],warnings:[backfilled.warning],notes:[]};general=backfilled.value;notes.push(backfilled.note);}
  generalParsed=parseDateInput(general,Number(task.month.slice(0,4)));
 }
 const derived=Object.values(windows).some(Boolean)?{windows:{},warnings:[],notes:[]}:deriveOpeningWindows(task,group,facts,raw,generalParsed);Object.assign(windows,derived.windows);warnings.push(...derived.warnings);notes.push(...derived.notes);
 for(const f of facts){const rs=task._records.filter(r=>r.f===f),window=rs.find(r=>r.atom.dateWindow)?.atom.dateWindow||'',label=window==='participation_exit'?'参与和退出日期':window==='participation_only'?'仅参与日期':f.label,value=window?windows[window]:general;
  if(!value){if(window)warnings.push(label+'尚未补充，保留原日期待核；可填写“'+(window==='participation_exit'?'参与和退出':'仅参与')+'：起止日期”。');continue;}
  const parsed=parseDateInput(value,Number(task.month.slice(0,4)));
  if(!parsed.end&&rs.every(r=>r.atom.dateRole==='end')&&(group.scope.event==='opening'||rs.some(r=>r.atom.rangeEndpoint))){warnings.push(label+'的结束日期尚未补充，保留原结束日期并标黄。');continue;}
  const result=previewUpdates(task,[f],{opening_announced_date:value},raw,group.id);for(const u of result.updates){u.label=label;if(window&&derived.basis)u.basis={...derived.basis,...(backfilled?.basis?{rangeDerivation:backfilled.basis}:{})};else if(backfilled?.basis)u.basis=backfilled.basis;}updates.push(...result.updates);warnings.push(...result.warnings);
 }
 if(!general&&!Object.values(windows).some(Boolean))throw new Error('请填写开放日期或日期区间。');
 const invalidRuleGroups=generalParsed?.end&&segmentRule(raw)&&derived.warnings.length?[group.id]:[];
 warnings.push(...openingDependencyInvalidations(task,updates,invalidRuleGroups).map(i=>i.reason));
 return{groupId:group.id,text:raw,updates,warnings:unique(warnings),notes:unique(notes),invalidRuleGroups};
}
function feeTextGroup(task,group){
 const first=task.facts.find(f=>group.factIds.includes(f.id));
 const scope=f=>JSON.stringify(Object.entries({...f.scope,condition:String(f.scope?.condition||'').startsWith('holding_days:')?'':f.scope?.condition||''}).sort(([a],[b])=>a.localeCompare(b)));
 return{...group,factIds:task.facts.filter(f=>f.key===first?.key&&scope(f)===scope(first)).map(f=>f.id)};
}
export function previewCardSupplement(task,groupId,text){
 const group=getSupplementGroups(task).find(g=>g.id===groupId);if(!group)throw new Error('该补充事项已变化，请重新核对。');
 const facts=group.factIds.map(id=>task.facts.find(f=>f.id===id)),raw=String(text??'').trim();
 if(facts.length&&facts.every(f=>f.key.endsWith('_fee'))&&unique(facts.map(f=>f.key)).length===1){
  let scalar;try{scalar=scalarFee(facts[0],raw);}catch{}
  const fullGroup=feeTextGroup(task,group),override=task.feeTextOverrides?.[facts[0].feeTextOverrideId];
  if(scalar&&override){
   const tokens=feeTokens(override.text,facts[0].key,false).filter(t=>t.key===facts[0].key);
   if(tokens.length!==1)throw new Error('当前完整文字包含多个分档，请填写各档完整内容，不能用一个数字覆盖所有档。');
   const token=tokens[0];return previewFeeText(task,fullGroup,override.text.slice(0,token.start)+scalar+override.text.slice(token.start+token.before.length),FEE_LABELS);
  }
  if(!scalar&&/[\u4e00-\u9fff]/.test(raw))return previewFeeText(task,fullGroup,raw,FEE_LABELS);
 }
 return previewSupplement(task,groupId,raw);
}
export function undoFeeTextSupplement(task,groupId){const group=getSupplementGroups(task).find(g=>g.id===groupId),id=task.facts.find(f=>group?.factIds.includes(f.id))?.feeTextOverrideId||groupId;return rebuildFacts(task,releaseFeeText(task,id));}
export function previewSupplement(task,groupId,text){
 const group=getSupplementGroups(task).find(g=>g.id===groupId);if(!group)throw new Error('该补充事项已变化，请重新核对。');
 const facts=task.facts.filter(f=>group.factIds.includes(f.id)),raw=String(text).trim();let values={};
 if(group.kind==='redemption_tiers')return redemptionTierSupplement(task,group,facts,raw);
 if(facts.length===1&&facts[0].key.endsWith('_fee')&&feeNeedsConditions(task,facts[0]))return supplementChanges(task,feeOccurrencePreview(task,facts[0],raw,groupId));
 if(facts.length===1&&facts[0].key==='opening_fixed_monthdays')return supplementChanges(task,wholeOccurrencePreview(task,facts[0],raw,groupId));
 if(facts.length===1&&wholeSupplementKey(facts[0].key))values[facts[0].key]=raw;
 else if(group.kind==='performance'){
  values=parsePerformanceInput(raw);const expected=unique(facts.filter(f=>/^performance_R\d/.test(f.key)).map(f=>f.key)),given=Object.keys(values).filter(k=>/^performance_R\d/.test(k));
  if(expected.length>1&&given.length>0&&given.length<expected.length&&!/R[1-9]\s*(?:为|是|=|[:：])/.test(raw))throw new Error('模板有多个业绩报酬基准，请填写完整分段，或用R1、R2和第1档、第2档明确要补充的档位。');
 }
 else if(group.kind==='date')return previewDates(task,group,facts,raw);
 else if(group.kind==='benchmark'){const input=raw.normalize('NFKC');if(/[-−－]\s*\d/.test(input))throw new Error('当前补充百分比应为0至100，不能填写负数。');const m=input.match(/^(?:(?:业绩比较基准|业绩参考基准|业绩基准)\s*(?:为|是)?\s*[:：]?\s*)?(\d+(?:\.\d+)?\s*%?)(?:\s*(?:业绩比较基准|业绩参考基准|业绩基准))?[。.]?$/);if(!m)throw new Error('请填写单个业绩比较基准，例如2.25%，不能填入范围或多个百分比。');values.benchmark=m[1];}
 else if(group.kind==='opening'){values=parseOpeningInput(raw);if(!Object.keys(values).length&&unique(facts.map(f=>f.key)).length===1)values[facts[0].key]=raw;}
 else values[facts[0].key]=raw;
 if(!Object.keys(values).length)throw new Error('请明确填写开放规则，例如每满18个月、连续开放5个交易日。');
 const preview=previewUpdates(task,facts,values,raw,groupId);
 if(wholeSupplementKey(facts[0]?.key||'')&&/[\r\n\t]/.test(raw))preview.notes=['已将同一段内的换行或制表符转为空格，保留模板原有段落排版。'];
 return group.kind==='text'||wholeSupplementKey(facts[0]?.key||'')?supplementChanges(task,preview):preview;
}
// Supplemental input rows are independent: never bulk-parse untouched rows.
export function previewSupplementField(task,kind,text){
 const labels={date:'本次开放日',performance:'业绩报酬',benchmark:'业绩比较基准',code:'产品代码',opening:'开放期'};
 if(!labels[kind])throw new Error('请选择有效的补充要素。');
 let raw=String(text??'').normalize('NFKC').trim();if(!raw)throw new Error('请填写本行需要补充的内容。');
 const aliases={date:'本次开放日|本次开放日期|开放日期|开放日|本次开售日|开售日|开售日期|募集日期|本次募集期|募集期',performance:'业绩报酬',benchmark:'业绩比较基准|业绩参考基准|业绩基准',code:'产品代码|产品编码',opening:'开放期|开放规则|开放周期'};
 // Labels copied from another row should not be interpreted as numbers for this row.
 for(const [other,names]of Object.entries(aliases))if(other!==kind&&new RegExp('(?:^|[;；\\n])\\s*(?:'+names+')\\s*[:：]').test(raw))throw new Error('本行只解析'+labels[kind]+'，请将其他要素分别填入对应行。');
 let groups=getSupplementGroups(task).filter(g=>g.kind===kind&&(kind!=='date'||g.scope.event==='opening'));const notes=[];
 if(kind==='opening'&&groups.length>1&&new Set(groups.map(g=>JSON.stringify(g.scope,Object.keys(g.scope).sort()))).size===1){
  const atomic=groups.filter(g=>g.factIds.some(id=>task.facts.find(f=>f.id===id)?.key!=='opening'));
  if(atomic.length){groups=atomic;notes.push('未定位到具体参数的整段开放期说明保留原文待核；本次只更新已识别的周期、天数或星期。');}
 }
 if(!groups.length)return{text:raw,updates:[],warnings:['模板未识别到'+labels[kind]+'对应位置，未写入。']};
 if(groups.length>1)return{text:raw,updates:[],warnings:['该要素存在多个份额或口径，请在下方对应卡片分别填写。']};
 if(kind!=='date')raw=raw.replace(new RegExp('^(?:'+aliases[kind]+')\\s*[:：]\\s*'),'').trim();
 if(kind==='benchmark'){
  const m=raw.match(/^(?:(?:业绩比较基准|业绩参考基准|业绩基准)(?:为|是)?\s*)?(\d+(?:\.\d+)?\s*%?)(?:\s*(?:业绩比较基准|业绩参考基准|业绩基准))?[。.]?$/);
  if(!m)throw new Error('请填写单个业绩比较基准，例如2.25%，不要将业绩报酬分段填入本行。');raw=m[1];
 }
 const preview=previewSupplement(task,groups[0].id,raw);preview.warnings.push(...notes);return preview;
}
export function previewPerformanceNumbers(task,input={}){
 const mode=input.mode;if(!['single','tiered'].includes(mode))throw new Error('请选择单档或分段计提方式。');
 const parse=(key,label)=>{const value=String(input[key]??'').normalize('NFKC').trim();if(!/^\d+(?:\.\d+)?\s*%?$/.test(value))throw new Error('请填写'+label+'，仅需输入0至100之间的百分比数字。');const n=Number(value.replace(/[%\s]/g,''));if(n<0||n>100)throw new Error(label+'应在0至100之间。');return n+'%';};
 const values={performance_R1:parse('threshold1','第一档计提基准'),performance_ratio_1:parse('ratio1','第一档提取比例')};
 if(mode==='tiered'){values.performance_R2=parse('threshold2','第二档计提基准');values.performance_ratio_2=parse('ratio2','第二档提取比例');if(number(values.performance_R1)>=number(values.performance_R2))throw new Error('第一档计提基准必须小于第二档计提基准。');}
 const text=mode==='single'?`年化收益率超过${values.performance_R1}的部分，提取${values.performance_ratio_1}作为业绩报酬。`:`${values.performance_R1}＜年化收益率≤${values.performance_R2}，提取${values.performance_ratio_1}；年化收益率＞${values.performance_R2}，提取${values.performance_ratio_2}。`;
 const groups=getSupplementGroups(task).filter(g=>g.kind==='performance');
 if(!groups.length)return{text,updates:[],warnings:['模板未识别到业绩报酬参数对应位置，未写入。']};
 if(groups.length>1)return{text,updates:[],warnings:['业绩报酬存在多个份额或口径，请在下方对应卡片分别填写。']};
 const group=groups[0],facts=task.facts.filter(f=>group.factIds.includes(f.id)),expected=unique(facts.filter(f=>f.key.startsWith('performance_')).map(f=>f.key)),given=Object.keys(values);
 if(!expected.length)return{groupId:group.id,text,updates:[],warnings:['模板中的业绩报酬百分比含义尚未明确，未找到可对应的计提基准和提取比例，请在下方核对原值。']};
 if(expected.length!==given.length||expected.some(key=>!given.includes(key)))throw new Error('所选'+(mode==='single'?'单档':'分段')+'计提方式与模板已识别的档位不一致，请选择匹配的计提方式；本次未写入。');
 return previewUpdates(task,facts,values,text,group.id);
}
export function previewSupplementText(task,text){
 const raw=String(text).trim(),groups=getSupplementGroups(task),requests=[],warnings=[];
 if(/[-−－]\s*\d+(?:\.\d+)?\s*[%％]/.test(raw))throw new Error('当前补充百分比应为0至100，不能填写负数。');
 const extract=label=>raw.match(new RegExp('(?:^|[;；\\n])\\s*(?:'+label+')\\s*[:：]\\s*([^;；\\n]+)'))?.[1]?.trim();
 const dates=[...raw.matchAll(/(?:^|[;；\n])\s*(?:本次开放日|本次开放日期|开放日期|开放日|本次开售日|开售日期|开售日|募集日期|本次募集期|募集期)\s*[:：]\s*([^;；\n]+)/g)].map(m=>m[1].trim());
 if(new Set(dates.map(value=>parseDateInput(value,Number(task.month.slice(0,4))).value)).size>1)throw new Error('本次开放日与本次开售日为同一日期，但补充了不同值，请统一后再应用。');
 const deadlines=[...raw.matchAll(/(?:^|[;；\n])\s*(?:(?:本次)?(?:开放|开售)(?:截止日|结束日|截止日期|结束日期)|档期(?:结束|截止)日期)\s*[:：]\s*([^;；\n]+)/g)].map(m=>m[1].trim());
 if(new Set(deadlines.map(value=>parseDateInput(value,Number(task.month.slice(0,4))).value)).size>1)throw new Error('补充的开放截止日期不一致，请统一后再应用。');
 const deadline=deadlines[0];
 const date=dates[0],windowBoth=extract('参与和退出|参与退出|申购和赎回'),windowOnly=extract('仅参与|仅开放参与|仅申购'),splitRule=raw.match(/(?:^|[;；\n])\s*(?:分段规则|开放分段|开放安排)\s*[:：]\s*([^\n]+)/)?.[1]?.trim();
 if(date||deadline||windowBoth||windowOnly)requests.push(['date',[date?'本次开放日：'+date:'',deadline?'本次开放截止日：'+deadline:'',windowBoth?'参与和退出：'+windowBoth:'',windowOnly?'仅参与：'+windowOnly:'',splitRule?'分段规则：'+splitRule:''].filter(Boolean).join('；'),'opening']);
 const reservation=extract('本次预约日|预约日期');if(reservation)requests.push(['date',reservation,'reservation']);
 const opening=extract('开放期|开放规则|开放周期');if(opening)requests.push(['opening',opening]);
 const code=extract('产品代码|产品编码');if(code)requests.push(['code',code]);
 const benchmark=extract('业绩比较基准|业绩参考基准|业绩基准')||raw.match(/(\d+(?:\.\d+)?\s*[%％])\s*(?:业绩比较基准|业绩参考基准|业绩基准)/)?.[1];if(benchmark)requests.push(['benchmark',benchmark]);
 if(/计提|提取|提成|R[1-9]\s*(?:为|是|=|:)/.test(raw))requests.push(['performance',raw]);
 const previews=[];
 for(const [kind,value,event]of requests){const matching=groups.filter(g=>g.kind===kind&&(!event||g.scope.event===event));
  if(!matching.length){warnings.push('模板未识别到'+({date:'该事件的日期',opening:'开放期',code:'产品代码',benchmark:'业绩比较基准',performance:'业绩报酬'}[kind])+'对应位置，未写入。');continue;}
  if(matching.length>1){warnings.push(matching.map(g=>g.label).join('、')+'存在多个份额或口径，请在对应卡片分别填写。');continue;}
  previews.push(previewSupplement(task,matching[0].id,value));
 }
 if(!requests.length)throw new Error('请标明补充内容，例如本次开放日：2026/11/3；产品代码：000123；或填写业绩报酬分段。');
 return{text:raw,updates:previews.flatMap(p=>p.updates),warnings:[...warnings,...previews.flatMap(p=>p.warnings)],notes:unique(previews.flatMap(p=>p.notes||[])),invalidRuleGroups:unique(previews.flatMap(p=>p.invalidRuleGroups||[]))};
}
export function applySupplementBatch(task,previews){
 if(!Array.isArray(previews)||!previews.length)throw new Error('没有可应用的补充内容。');
 if(previews.some(p=>p.feeText)){
  if(previews.length!==1)throw new Error('整段费率文字请在本卡片单独确认。');
  const p=previews[0],g=getSupplementGroups(task).find(g=>g.id===p.groupId);if(!g)throw new Error('本项已变化，请重新预览。');
  applyFeeText(task,p,previewFeeText(task,feeTextGroup(task,g),p.text,FEE_LABELS));return finalize(task);
 }
 const prepared=new Map();
 for(const preview of previews)for(const update of preview.updates||[]){
  const f=task.facts.find(f=>f.id===update.factId);if(!f||update.key&&update.key!==f.key)throw new Error('补充对应的事项已经变化，请重新解析。');
  let value;const renders=[],records=task._records.filter(r=>r.f===f);
  if(update.occurrenceValues){
   if(!(f.key.endsWith('_fee')||f.key==='opening_fixed_monthdays')||!Array.isArray(update.occurrenceValues)||update.occurrenceValues.length!==records.length)throw new Error('补充原文定位不完整，请重新预览。');
   const remaining=new Set(update.occurrenceValues);
   for(const r of records){const matches=[...remaining].filter(x=>x.materialId===r.m.materialId&&x.fieldId===r.field.id&&x.paragraphId===r.p.id&&x.start===r.atom.start&&x.before===r.atom.before&&x.context===r.p.text&&x.state===JSON.stringify([f.status,f.value,r.userRenderValue??null]));if(matches.length!==1)throw new Error('补充对应的原文或核对结果已经变化，请重新预览。');const item=matches[0];remaining.delete(item);renders.push([r,validateSupplement(f,item.value,task,true)]);}
   value=f.key.endsWith('_fee')&&unique(renders.map(([,v])=>v)).length===1?renders[0][1]:cleanSupplementText(preview.text||update.value);
  }else{
   value=validateSupplement(f,update.value,task);
   for(const r of records)renders.push([r,f.key==='opening_announced_date'?formatDateLike(value,r.atom.before,r.atom.dateRole,Number(task.month.slice(0,4))):value]);
  }
  if(prepared.has(f.id)&&prepared.get(f.id).value!==value)throw new Error('同一事项有不同补充值，请核对后再应用。');
  if(prepared.has(f.id)&&JSON.stringify(prepared.get(f.id).renders.map(([,v])=>v))!==JSON.stringify(renders.map(([,v])=>v)))throw new Error('同一事项有不同补充值，请核对后再应用。');
  prepared.set(f.id,{f,value,renders,text:preview.text||String(update.value),source:preview.source,basis:update.basis});
 }
 if(!prepared.size)throw new Error('没有可对应到模板的补充字段。');
 // Everything above is read-only: one invalid field cannot leave a partial update.
 const invalidated=openingDependencyInvalidations(task,[...prepared.values()].map(({f,value})=>({factId:f.id,value})),previews.flatMap(p=>p.invalidRuleGroups||[]));
 for(const {f,reason}of invalidated){delete f.userConfirmed;delete f._confirmationBefore;f.value=null;f.status='pending';f.source='template';f.reason=reason;f.reasonType='needs_review';f.evidence=[];for(const r of task._records.filter(r=>r.f===f))delete r.userRenderValue;}
 for(const item of prepared.values()){
  const {f,value,renders,text,source,basis}=item;delete f.userConfirmed;delete f._confirmationBefore;
  f.value=value;f.status='verified';f.source=source?'user_calendar':'user';f.reason=source?'由用户导入的产品日历补充确认。':'由用户补充确认；不是合同自动提取的依据。';f.reasonType='user_confirmed';
  const backfill=basis?.kind==='trading_day_backfill'?basis:basis?.rangeDerivation;
  let derivation=backfill?`\n按${backfill.durationSource==='contract'?'合同明确':backfill.durationSource==='template_rule'?'模板明文':'模板原完整区间'}的${backfill.count}个沪深交易日，从截止日${backfill.end}反推：${backfill.range}。\n交易日历依据：${backfill.calendarSources.map(s=>s.title+' '+s.url).join('；')}`:'';
  if(basis?.kind==='trading_day_split')derivation+=`\n按沪深交易日及${basis.ruleSource==='explicit_rule'?'补充':'模板'}分段规则计算：${basis.rule}；本段 ${value}。\n交易日历依据：${basis.calendarSources.map(s=>s.title+' '+s.url).join('；')}`;
  if(derivation)f.reason+=backfill?'日期按明确天数和沪深交易日从截止日反推。':'分段日期按沪深交易日和分段规则计算。';
  f.evidence=[{blockId:source?'user-calendar':'user-supplement',quote:text+derivation,...(source?{source:{...source}}:{}),...(basis?{basis}:{})}];for(const [r,after]of renders)r.userRenderValue=after;
 }
 return rebuildFacts(task,[...prepared.keys(),...invalidated.map(i=>i.f.id)]);
}
export function applySupplement(task,preview){return applySupplementBatch(task,[preview]);}
export function setSupplement(task,factId,value){const f=task.facts.find(f=>f.id===factId);if(!f)throw new Error('找不到待核事项。');return applySupplement(task,{text:String(value),updates:[{factId,key:f.key,value}]});}
export function confirmFacts(task,factIds,confirmed=true){
 const ids=unique(factIds||[]),facts=ids.map(id=>task.facts.find(f=>f.id===id));if(!ids.length||facts.some(f=>!f))throw new Error('请选择有效待核事项。');
 for(const f of facts){
  if(confirmed){if(f.status==='verified')continue;f._confirmationBefore={status:f.status,source:f.source,reason:f.reason,reasonType:f.reasonType,evidence:f.evidence};f.status='verified';f.source='user_confirmation';f.userConfirmed=true;f.reason='用户已确认模板原值无误，按各处原值保留。';f.reasonType='user_confirmed';f.evidence=[...(f.evidence||[]),{blockId:'user-confirmation',quote:'用户确认模板原值无误。'}];}
  else if(f.userConfirmed&&f._confirmationBefore){Object.assign(f,f._confirmationBefore);delete f._confirmationBefore;delete f.userConfirmed;}
 }
 return rebuildFacts(task,ids);
}
