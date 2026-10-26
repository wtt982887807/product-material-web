// Local evidence rules. No network calls and no product-specific answer tables.
import {NS,elements} from './office-xml.js';
import {parseDateInput,formatDateLike,parsePerformanceInput,parseOpeningInput} from './supplement-parser.js?v=20261026';
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
const FEE_LABELS=[['trust_total_fee',/固定信托管理费|固定管理费/],['trust_manager_fee',/受托人信托报酬|(?:收取)?信托管理费|受托人报酬/],['floating_advisor_fee',/浮动投资顾问费/],['floating_sales_fee',/浮动销售服务费/],['advisor_fee',/固定投资顾问费|投资顾问费/],['sales_fee',/固定销售服务费|销售服务费/],['custody_fee',/托管费(?:率)?|保管费(?:率)?/],['management_fee',/管理费(?:率)?/],['redemption_fee',/赎回费(?:率)?|退出费(?:率)?/],['subscription_fee',/认[、/／]?申购(?:费|费用)|认购(?:费|费用)|申购(?:费|费用)|参与费(?:率)?/]];
const AGGREGATE_LAYERS=/(?:两层|双层|上下两层).{0,8}(?:合计|总计|总费|综合)|(?:合计|总计).{0,5}(?:两层|双层)/;
function feeTokens(text,baseKey,trust){
 const tokens=[];for(const m of text.matchAll(PERCENT)){
  const before=text.slice(Math.max(0,m.index-100),m.index);const hits=[];
  for(const [key,re] of FEE_LABELS)for(const a of before.matchAll(new RegExp(re.source,'g')))hits.push({key,index:a.index,length:a[0].length});
  // A longer semantic label owns its nested generic label (信托管理费 contains 管理费).
  const specific=hits.filter(h=>!hits.some(other=>other!==h&&other.index<=h.index&&other.index+other.length>=h.index+h.length&&other.length>h.length));
  specific.sort((a,b)=>b.index-a.index||b.length-a.length);let key=specific[0]?.key||baseKey;
  if(key==='trust_total_fee'&&!trust)key='management_fee';
  if(trust&&key==='management_fee')key='trust_manager_fee';
  if(baseKey==='subscription_redemption_fee'||/(?:参与|认.{0,5}购)[/／、及和].{0,4}(?:退出|赎回)(?:费率|费)[：:\s]*$/.test(before))key='subscription_redemption_fee';
  const raw=m[0];tokens.push({key,start:m.index,before:raw,value:`${number(raw)}%`});
 }return tokens;
}
function resolveFee(key,rs,scope,trust){
 const candidates=[];for(const b of filtered(rs,scope)){
  // Fee context is explicitly anchored locally; percentages from investment restrictions do not qualify.
  const tokens=feeTokens(b.text,'',trust);
  for(const t of tokens)if(t.key===key&&t.value!=='NaN%')candidates.push({value:t.value,evidence:[quote(b.base||b,b.text)]});
  const aliases=key==='subscription_fee'?/(?:不收取|不收|免收).{0,8}(?:认购|申购|参与)(?:[、/／](?:认购|申购|参与))?(?:费|费用)/:key==='redemption_fee'?/(?:不收取|不收|免收).{0,8}(?:赎回|退出)(?:费|费用)/:null;
  if(aliases?.test(b.text))candidates.push({value:'0%',evidence:[quote(b.base||b,b.text)]});
 }
 if(key==='subscription_redemption_fee'){
  const a=resolveFee('subscription_fee',rs,scope,trust),b=resolveFee('redemption_fee',rs,scope,trust);return a.status==='verified'&&b.status==='verified'&&equivalent(a.value,b.value)?{...a,evidence:[...a.evidence,...b.evidence]}:{status:'pending',value:null,evidence:[...a.evidence,...b.evidence],reason:'参与及退出两个费率需分别核实。'};
 }
 return pick(candidates);
}
function performanceTokens(text){
 const tokens=[];
 for(const m of text.matchAll(PERCENT)){
  const left=text.slice(Math.max(0,m.index-150),m.index),right=text.slice(m.index+m[0].length,m.index+m[0].length+40);
  const column=left.split('\n').at(-1),tablePayout=/[\t|]\s*$/.test(column)&&/[RＲF]|收益率/.test(column)&&/[＜≤<>＞≥]/.test(column);
  const payout=tablePayout||/(?:提取|计提|收取|提成|×|乘以|按)[【\s]*$/.test(left)||/^(?:的|比例的).{0,6}(?:业绩报酬|浮动服务费)|^(?:计提|提取)/.test(right);
  const hurdle=!payout&&(/(?:[＜≤<>＞≥=＝]|高于|低于|超过|超出|大于|小于|基准(?:为|是)?|R[12](?:为|是)|收益率(?:为|是)?|年化)[【\s]*$/.test(left)||/^\s*[＜≤<>＞≥]/.test(right)||/\d+(?:\.\d+)?%\s*[-—~～]\s*$/.test(left)||/^\s*[-—~～]\s*\d/.test(right)||/(?:[RＲF]|收益率)[-－]\s*$/.test(left));
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
 const point='(?:\\d{4}\\s*[年/-]\\s*)?\\d{1,2}\\s*[月/-]\\s*\\d{1,2}日?';
 const re=new RegExp('(?<!\\d)'+point+'(?:\\s*[-—–~～至到]\\s*(?:'+point+'|\\d{1,2}日?))?(?!\\d)','g');
 const atoms=[];
 for(const m of text.matchAll(re)){
  const left=text.slice(0,m.index).split(/[。；;\n]/).at(-1),right=text.slice(m.index+m[0].length).split(/[。；;\n]/)[0];
  const context=label+' '+left+m[0]+right;
  if(NONCURRENT_DATE.test(context)||/每年\s*$/.test(left))continue;
  const localLeft=text.slice(0,m.index).split(/[。；;，,：:（(\n]/).at(-1),localRight=text.slice(m.index+m[0].length).split(/[。；;，,：:（()）\n]/)[0],localContext=localLeft+m[0]+localRight;
  const roleText=localLeft+' '+label+' '+localRight;
  const role=/截止|结束|终止|最后/.test(roleText)?'end':/起始|开始|首个|首日/.test(roleText)?'start':'date';
  const event=/预约/.test(localContext)?'reservation':/开售|发售|募集|认购/.test(localContext+' '+label)?'sale':/预约/.test(label)?'reservation':'opening';
  const dateWindow=/仅(?:开放)?参与|仅(?:开放)?申购/.test(localRight)?'participation_only':/(?:参与|申购)(?:和|及|与)(?:退出|赎回)/.test(localRight)?'participation_exit':'';
  let parsed;try{parsed=parseDateInput(m[0],2000);}catch{}
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
 if(['fees','management_fee','custody_fee','subscription_fee','redemption_fee'].includes(key)){const labeled=FEE_LABELS.find(([,re])=>new RegExp('^(?:'+re.source+')$').test(labelNorm(label)))?.[0];const feeKey=/(?:参与|认.{0,5}购)[/／、及和].{0,4}(?:退出|赎回)/.test(label)?'subscription_redemption_fee':labeled||key;return feeTokens(text,feeKey,trust).map(t=>({...t,scope:scopeFor(label,text,t.start)}));}
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
function fieldParagraphs(row,map){const ids=row.cells?.slice(1).flatMap(c=>c.paragraphIds||[])||row.paragraphIds;return unique(ids||[]).map(id=>map.get(id)).filter(Boolean);}
function addOccurrence(task,m,field,p,atom,resolution){
 const scope={product:task.target.name,...(atom.scope||scopeFor(field.label,p.text,atom.start))};
 if(atom.key==='opening_announced_date'){scope.period='announcement';scope.condition=[atom.dateEvent,atom.dateRole,norm(atom.before),atom.dateWindow||''].join('|');}
 if(atom.key.startsWith('investment_')||atom.key.startsWith('custom:')||['opening','lockup','duration'].includes(atom.key))scope.condition=norm(atom.before);
 const factKey=(atom.key==='risk_description'?'risk_rating':atom.key)+'|'+JSON.stringify(scope);
 const factLabel=atom.key.startsWith('performance_ratio_')?'业绩报酬提成比例（第'+atom.key.match(/\d+$/)[0]+'档）':/^performance_R\d/.test(atom.key)?'业绩报酬基准 R'+atom.key.match(/\d+$/)[0]:field.label;
 let f=task._factMap.get(factKey);if(!f){f={id:'fact-'+(task.facts.length+1),key:atom.key,label:factLabel,value:resolution.value,status:resolution.status,evidence:resolution.evidence||[],scope,occurrences:[],reason:resolution.reason,reasonType:resolution.reasonType||(resolution.status==='verified'?'confirmed':'needs_review')};task.facts.push(f);task._factMap.set(factKey,f);}
 f.occurrences.push({materialId:m.materialId,paragraphId:p.id,fieldId:field.id});field.factIds.push(f.id);
 const record={m,field,p,atom,f};task._records.push(record);renderRecord(record);
}
function operationId(m){m._operationSequence=(m._operationSequence||0)+1;return 'op-'+m._operationSequence;}
function renderRecord({m,field,p,atom,f,userRenderValue}){
 const common={paragraphId:p.id,part:p.part,index:p.index,reason:f.reason,factIds:[f.id]};
 if(f.userConfirmed)return;
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
  for(const diff of diffs)if(diff.before||diff.after)m.operations.push({...common,id:operationId(m),kind:'replace',start:atom.start+diff.start,before:diff.before,after:diff.after});
 }else{
  const spans=atom.before.matchAll(/\d+(?:\.\d+)?\s*[%％]?|R[1-5]|[一二三四五六日天](?=开放)/g);const list=[...spans];
  // Compound pending text is anchored to each parameter rather than painting whole cells.
  const anchors=list.length?list.map(s=>({start:atom.start+s.index,before:s[0]})):[{start:atom.start,before:atom.before}];
  for(const a of anchors)if(a.before){m.operations.push({...common,...a,id:operationId(m),kind:'highlight',color:'yellow'});m.operations.push({...common,...a,id:operationId(m),kind:'comment'});}
 }
}
function reviewParagraph(task,m,field,p,rs){
 const atoms=field.manualReviewReason?[{key:'custom:table_layout:'+labelNorm(field.label),start:0,before:p.text,value:p.text,manualReviewReason:field.manualReviewReason}]:atomsFor(field.label,p.text,task.target,keyFor(field.label)==='performance'?m._performanceText:field.before);
 for(const atom of atoms){let resolution;const aggregate=(atom.key.endsWith('_fee')||atom.key.startsWith('performance_'))&&AGGREGATE_LAYERS.test(field.label+' '+field.before);
  if(aggregate)atom.scope={...(atom.scope||scopeFor(field.label,p.text,atom.start)),condition:'aggregate_layers'};
  if(atom.manualReviewReason)resolution={status:'pending',value:null,evidence:[],reasonType:'needs_review',reason:atom.manualReviewReason};
  else if(atom.key==='warning_stop'){
   const a=resolveAtom('warning_line',field.label,rs,{},task.target),b=resolveAtom('stop_line',field.label,rs,{},task.target);
   resolution=a.status==='verified'&&b.status==='verified'?{status:'verified',value:`预警线${a.value}，止损线${b.value}`,evidence:[...a.evidence,...b.evidence],reason:'合同明确设置预警线和止损线。'}:resolveAtom(atom.key,field.label,rs,{},task.target,{original:p.text});
  }else if(atom.key==='risk_description')resolution=resolveAtom('risk_rating',field.label,rs,atom.scope,task.target);
  else if(atom.key==='opening_announced_date')resolution=pendingAnnouncement();
  else resolution=resolveAtom(atom.key,field.label,rs,atom.scope||scopeFor(field.label,p.text,atom.start),task.target,{original:atom.before});
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
 for(const m of task.materials){
  m.operations=m.operations.filter((o,i,arr)=>!arr.slice(0,i).some(a=>a.paragraphId===o.paragraphId&&a.kind===o.kind&&a.start===o.start&&a.before===o.before&&a.after===o.after));
  for(const field of m.fields){field.factIds=unique(field.factIds);const fs=field.factIds.map(id=>task.facts.find(f=>f.id===id));field.status=fs.some(f=>f.status==='conflict')?'conflict':fs.some(f=>f.status==='pending')?'pending':'verified';const replaces=m.operations.filter(o=>field.paragraphIds.includes(o.paragraphId)&&o.kind==='replace');field.after=field.paragraphIds.map(id=>{let t=m.template.paragraphs.find(p=>p.id===id)?.text||'';for(const o of replaces.filter(o=>o.paragraphId===id).sort((a,b)=>b.start-a.start))t=t.slice(0,o.start)+o.after+t.slice(o.start+o.before.length);return t;}).join('\n');if(field.source!=='company')field.reason=field.status==='verified'?(replaces.length?'合同有明确依据，最小范围更新。':'原值与已取得依据一致。'):'保留模板原值，关键参数标黄并添加批注。';}
 }
 task.stats={facts:task.facts.length,verified:task.facts.filter(f=>f.status==='verified').length,pending:task.facts.filter(f=>f.status!=='verified').length,contractMissing:task.facts.filter(f=>f.status!=='verified'&&f.reasonType==='contract_missing').length,needsReview:task.facts.filter(f=>f.status!=='verified'&&f.reasonType!=='contract_missing').length,changes:task.materials.reduce((n,m)=>n+m.operations.filter(o=>o.kind==='replace').length,0),materials:task.materials.length};return task;
}
export function analyzeMaterials(contract,templates,{now,month}={}){
 const target=identity(contract),task={target,contract,templates,month:month||monthNow(now),facts:[],materials:[],warnings:[...(contract.warnings||[])],stats:{}};
 Object.defineProperties(task,{_factMap:{value:new Map()},_records:{value:[]}});
 if(!target.name){task.warnings.push(target.reason);throw new Error(target.reason);}
 const rs=records(contract);
 for(const t of templates){
  const kind=materialKind(t),selection=getRows(t,target),headings=kind==='poster'?posterHeadings(t,selection.name):null,coverNames=kind==='word'?wordCoverNames(t,target):new Map(),rows=selection.rows.length?selection.rows:kind==='poster'?posterRows(t,headings):[];
  const m={materialId:t.id,template:t,kind,templateProductName:headings?.name||selection.name,selectedPages:kind==='poster'?(t.slides||[]).map(s=>s.page):selection.pages,operations:[],fields:[],warnings:[...(t.warnings||[])]};Object.defineProperty(m,'_performanceText',{value:rows.filter(r=>keyFor(r.label)==='performance').map(r=>r.value).join('\n')});task.materials.push(m);
  if(!rows.length)m.warnings.push('尚未识别到可可靠定位的产品要素区，需检查模板结构。');
  const map=new Map(t.paragraphs.map(p=>[p.id,p]));const selectedIds=new Set();
  for(const row of rows){
   const ps=fieldParagraphs(row,map);ps.forEach(p=>selectedIds.add(p.id));const field={id:`${t.id}-field-${m.fields.length+1}`,label:row.label,status:'pending',before:ps.map(p=>p.text).join('\n'),after:'',factIds:[],paragraphIds:ps.map(p=>p.id),reason:'',manualReviewReason:row.manualReviewReason};m.fields.push(field);
   for(const p of ps)if(p.text.trim())reviewParagraph(task,m,field,p,rs);
  }
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
  m.productElementsLocated=m.fields.some(f=>f.source!=='company');
  if(!m.productElementsLocated)m.warnings.push('未定位到可更新的产品要素，无法完成该模板的产品要素更新；其他模板可以继续处理。');
 }
 return finalize(task);
}
function validateSupplement(fact,value,task){
 const raw=String(value??'').trim();if(!raw)throw new Error('请输入补充值。');let accepted=raw;
 if(fact.key.endsWith('_fee')||fact.key.startsWith('performance_')||fact.key==='benchmark'){if(!/^\d+(?:\.\d+)?\s*[%％]?$/.test(raw)||number(raw)>100)throw new Error('请填写0至100之间的费率数字，可带百分号。');accepted=raw.replace(/[％\s]/g,'').replace(/%$/,'')+'%';}
 else if(/^(?:minimum_amount|additional_amount|opening_cycle|opening_days|opening_monthday|lockup_days|duration_years)$/.test(fact.key)){const m=raw.match(/^(\d+(?:\.\d+)?)\s*(?:万元?|年|天|个月|个交易日|交易日|日)?$/);if(!m)throw new Error('请填写有效数字，可附相应单位。');if(!/amount/.test(fact.key)&&!Number.isInteger(Number(m[1])))throw new Error('日期、天数及周期请填写整数。');accepted=m[1];if(fact.key==='opening_monthday'&&(Number(accepted)<1||Number(accepted)>31))throw new Error('每月日期应为1至31日。');}
 else if(fact.key==='opening_weekday'){if(!/^(?:周|星期|每周)?[一二三四五六日天]$/.test(raw))throw new Error('请填写星期几，例如周四。');accepted=raw.at(-1);}
 else if(fact.key==='product_code'&&!/^[A-Za-z0-9][A-Za-z0-9_.-]{4,24}$/.test(raw))throw new Error('请填写5至25位产品代码，仅含字母、数字或代码连接符。');
 else if(fact.key==='risk_rating'){const map={'低风险':'R1','中低风险':'R2','中风险':'R3','中高风险':'R4','高风险':'R5'};accepted=map[raw]||raw.toUpperCase();if(!/^R[1-5]$/.test(accepted))throw new Error('请填写R1至R5或相应风险等级名称。');}
 else if(fact.key==='opening_announced_date')accepted=parseDateInput(raw,Number(task.month.slice(0,4))).value;
 return accepted;
}
function rebuildFacts(task,ids){
 const selected=new Set(ids);for(const m of task.materials)m.operations=m.operations.filter(o=>!o.factIds?.some(id=>selected.has(id)));
 for(const r of task._records)if(selected.has(r.f.id))renderRecord(r);return finalize(task);
}
const GROUP_KIND=f=>f.key==='opening_announced_date'?'date':f.key.startsWith('performance_')?'performance':f.key==='benchmark'?'benchmark':f.key==='product_code'?'code':/^opening(?:_|$)/.test(f.key)?'opening':'text';
export function getSupplementGroups(task){
 const groups=new Map();
 for(const f of task.facts){
  const kind=GROUP_KIND(f),scope={...f.scope};let identity=f.id;
  if(['performance','benchmark','code','opening','date'].includes(kind)){
   delete scope.condition;
   if(kind==='date')scope.event=(f.scope.condition||'').split('|')[0]||'opening';
   // Whole unmatched prose has its own meaning and must not share a replacement.
   identity=kind==='opening'&&f.key==='opening'?f.id:kind+'|'+JSON.stringify(scope);
  }
  const group=groups.get(identity)||{id:'supplement-'+f.id,label:({performance:'业绩报酬分段',benchmark:'业绩比较基准',code:'产品代码',opening:'开放期',date:({sale:'本次开售日',reservation:'本次预约日',opening:'本次开放日'})[scope.event]||'本次开放日'})[kind]||f.label,kind,factIds:[],scope};
  group.factIds.push(f.id);groups.set(identity,group);
 }
 return [...groups.values()].map(g=>{const fs=g.factIds.map(id=>task.facts.find(f=>f.id===id));return{...g,status:fs.some(f=>f.status==='conflict')?'conflict':fs.some(f=>f.status!=='verified')?'pending':'verified',userConfirmed:fs.filter(f=>f.userConfirmed).length>0};});
}
function previewUpdates(task,facts,values,text,groupId){
 const updates=[],warnings=[];
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
function previewDates(task,group,facts,raw){
 const windows={participation_exit:raw.match(/(?:参与和退出|参与退出|申购和赎回)\s*[:：]\s*([^;；\n]+)/)?.[1],participation_only:raw.match(/(?:仅参与|仅开放参与|仅申购)\s*[:：]\s*([^;；\n]+)/)?.[1]};
 const general=raw.match(/(?:本次开放日|开放日期|本次开售日|本次募集期|本次预约日)\s*[:：]\s*([^;；\n]+)/)?.[1]||(!Object.values(windows).some(Boolean)?raw:'');
 const updates=[],warnings=[];let generalParsed;if(general)generalParsed=parseDateInput(general,Number(task.month.slice(0,4)));
 for(const f of facts){const rs=task._records.filter(r=>r.f===f),window=rs.find(r=>r.atom.dateWindow)?.atom.dateWindow||'',label=window==='participation_exit'?'参与和退出日期':window==='participation_only'?'仅参与日期':f.label,value=window?windows[window]:general;
  if(!value){if(window)warnings.push(label+'尚未补充，保留原日期待核；可填写“'+(window==='participation_exit'?'参与和退出':'仅参与')+'：起止日期”。');continue;}
  const parsed=parseDateInput(value,Number(task.month.slice(0,4)));
  if(!parsed.end&&rs.every(r=>r.atom.dateRole==='end')&&(group.scope.event==='opening'||rs.some(r=>r.atom.rangeEndpoint))){warnings.push(label+'的结束日期尚未补充，保留原结束日期并标黄。');continue;}
  const result=previewUpdates(task,[f],{opening_announced_date:value},raw,group.id);for(const u of result.updates)u.label=label;updates.push(...result.updates);warnings.push(...result.warnings);
 }
 if(!general&&!Object.values(windows).some(Boolean))throw new Error('请填写开放日期或日期区间。');
 return{groupId:group.id,text:raw,updates,warnings:unique(warnings)};
}
export function previewSupplement(task,groupId,text){
 const group=getSupplementGroups(task).find(g=>g.id===groupId);if(!group)throw new Error('该补充事项已变化，请重新核对。');
 const facts=task.facts.filter(f=>group.factIds.includes(f.id)),raw=String(text).trim();let values={};
 if(group.kind==='performance'){
  values=parsePerformanceInput(raw);const expected=unique(facts.filter(f=>/^performance_R\d/.test(f.key)).map(f=>f.key)),given=Object.keys(values).filter(k=>/^performance_R\d/.test(k));
  if(expected.length>1&&given.length>0&&given.length<expected.length&&!/R[1-9]\s*(?:为|是|=|[:：])/.test(raw))throw new Error('模板有多个业绩报酬基准，请填写完整分段，或用R1、R2和第1档、第2档明确要补充的档位。');
 }
 else if(group.kind==='date')return previewDates(task,group,facts,raw);
 else if(group.kind==='benchmark'){if(/[-−－]\s*\d/.test(raw))throw new Error('当前补充百分比应为0至100，不能填写负数。');const m=raw.match(/\d+(?:\.\d+)?\s*[%％]?/);if(!m)throw new Error('请填写业绩比较基准，例如2.25%。');values.benchmark=m[0];}
 else if(group.kind==='opening'){values=parseOpeningInput(raw);if(!Object.keys(values).length&&unique(facts.map(f=>f.key)).length===1)values[facts[0].key]=raw;}
 else values[facts[0].key]=raw;
 if(!Object.keys(values).length)throw new Error('请明确填写开放规则，例如每满18个月、连续开放5个交易日。');
 return previewUpdates(task,facts,values,raw,groupId);
}
export function previewSupplementText(task,text){
 const raw=String(text).trim(),groups=getSupplementGroups(task),requests=[],warnings=[];
 if(/[-−－]\s*\d+(?:\.\d+)?\s*[%％]/.test(raw))throw new Error('当前补充百分比应为0至100，不能填写负数。');
 const extract=label=>raw.match(new RegExp('(?:^|[;；\\n])\\s*(?:'+label+')\\s*[:：]\\s*([^;；\\n]+)'))?.[1]?.trim();
 const date=extract('本次开放日|本次开放日期|开放日期|开放日'),windowBoth=extract('参与和退出|参与退出|申购和赎回'),windowOnly=extract('仅参与|仅开放参与|仅申购');
 if(date||windowBoth||windowOnly)requests.push(['date',[date?'本次开放日：'+date:'',windowBoth?'参与和退出：'+windowBoth:'',windowOnly?'仅参与：'+windowOnly:''].filter(Boolean).join('；'),'opening']);
 const sale=extract('本次开售日|开售日|募集日期|本次募集期|募集期');if(sale)requests.push(['date',sale,'sale']);
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
 return{text:raw,updates:previews.flatMap(p=>p.updates),warnings:[...warnings,...previews.flatMap(p=>p.warnings)]};
}
export function applySupplementBatch(task,previews){
 if(!Array.isArray(previews)||!previews.length)throw new Error('没有可应用的补充内容。');
 const prepared=new Map();
 for(const preview of previews)for(const update of preview.updates||[]){
  const f=task.facts.find(f=>f.id===update.factId);if(!f||update.key&&update.key!==f.key)throw new Error('补充对应的事项已经变化，请重新解析。');
  const value=validateSupplement(f,update.value,task),renders=[];
  if(prepared.has(f.id)&&prepared.get(f.id).value!==value)throw new Error('同一事项有不同补充值，请核对后再应用。');
  if(f.key==='opening_announced_date')for(const r of task._records.filter(r=>r.f===f))renders.push([r,formatDateLike(value,r.atom.before,r.atom.dateRole,Number(task.month.slice(0,4)))]);
  prepared.set(f.id,{f,value,renders,text:preview.text||String(update.value),source:preview.source});
 }
 if(!prepared.size)throw new Error('没有可对应到模板的补充字段。');
 // Everything above is read-only: one invalid field cannot leave a partial update.
 for(const item of prepared.values()){
  const {f,value,renders,text,source}=item;delete f.userConfirmed;delete f._confirmationBefore;
  f.value=value;f.status='verified';f.source=source?'user_calendar':'user';f.reason=source?'由用户导入的产品日历补充确认。':'由用户补充确认；不是合同自动提取的依据。';f.reasonType='user_confirmed';
  f.evidence=[{blockId:source?'user-calendar':'user-supplement',quote:text,...(source?{source:{...source}}:{})}];for(const [r,after]of renders)r.userRenderValue=after;
 }
 return rebuildFacts(task,[...prepared.keys()]);
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
