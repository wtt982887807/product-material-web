import {monthlyWeekdayPatterns,parseMonthlyWeekdayInput} from './monthly-weekday.js?v=20261026-monthly';
const pct=String.raw`\d+(?:\.\d+)?\s*[%％]`;
const percent=s=>`${Number(String(s).replace(/[^\d.]/g,''))}%`;
const clean=s=>String(s??'').normalize('NFKC').trim();
function validDate(y,m,d){const x=new Date(Date.UTC(y,m-1,d));if(x.getUTCFullYear()!==y||x.getUTCMonth()!==m-1||x.getUTCDate()!==d)throw new Error('请填写有效日期。');return{year:y,month:m,day:d};}
export function parseDateInput(value,year){
 const text=clean(value).replace(/^(?:本次|下次|下一)?(?:开放(?:日期|日|期)|开售日期|募集日期|预约日期)\s*[:：]?\s*/,'');
 const full=String.raw`(?:(\d{4})\s*[年/.-]\s*)?(\d{1,2})\s*[月/.-]\s*(\d{1,2})日?`;
 const re=new RegExp('^'+full+'(?:\\s*(?:至|到|[-—–~～])\\s*(?:'+full+'|(\\d{1,2})日?))?$');
 const m=text.match(re);if(!m)throw new Error('请填写开放日期或区间，例如2026/11/3至2026/11/5。');
 const y=Number(m[1]||year);if(!Number.isInteger(y)||y<1900||y>2200)throw new Error('请填写包含年份的日期。');
 const start=validDate(y,Number(m[2]),Number(m[3]));let end=null;
 if(m[5]||m[7])end=validDate(Number(m[4]||y),Number(m[5]||start.month),Number(m[6]||m[7]));
 if(end&&Date.UTC(end.year,end.month-1,end.day)<Date.UTC(start.year,start.month-1,start.day))throw new Error('结束日期不能早于开始日期。');
 const iso=d=>`${d.year}/${d.month}/${d.day}`;
 const first=text.match(new RegExp('^'+full)),delimiter=end?text.slice(first[0].length).match(/^\s*(至|到|[-—–~～])/)?.[1]:null;
 const startText=first[0],afterStart=text.slice(startText.length),separatorText=end?afterStart.match(/^\s*(?:至|到|[-—–~～])\s*/)?.[0]||'':'';
 return{start,end,value:iso(start)+(end?'至'+iso(end):''),endDayOnly:!!m[7],endHasYear:!!m[4],delimiter,startText,endText:end?afterStart.slice(separatorText.length):'',endOffset:end?startText.length+separatorText.length:0};
}
export function formatDateLike(value,before,role='date',year){
 const parsed=parseDateInput(value,year),original=clean(before);
 if(/^\d{1,2}日?$/.test(original)&&role==='end'){
  const d=parsed.end||parsed.start,day=String(d.day).padStart(/^0/.test(original)?2:1,'0'),suffix=original.endsWith('日')?'日':'';
  return parsed.end&&parsed.start.month!==d.month?(suffix?d.month+'月'+day+suffix:d.month+'/'+day):day+suffix;
 }
 let old;try{old=parseDateInput(original,year);}catch{old=null;}
 const wantsRange=!!old?.end;
 const split=old?.delimiter||'-';
 const separator=/年|月/.test(original)?'cn':original.includes('/')?'/':original.includes('.')?'.':'-';
 const hasYear=/^\d{4}/.test(original),padded=/(?:^|[年/.-])0\d/.test(original);
 const n=x=>padded?String(x).padStart(2,'0'):String(x);
 const fmt=(d,withYear=hasYear,dayOnly=false)=>dayOnly?n(d.day)+(separator==='cn'?'日':''):separator==='cn'?(withYear?d.year+'年':'')+n(d.month)+'月'+n(d.day)+'日':(withYear?d.year+separator:'')+n(d.month)+separator+n(d.day);
 if(wantsRange){
  if(!parsed.end)throw new Error('模板含开放日期区间，请同时填写开始和结束日期。');
  const dayOnly=old.endDayOnly,endHasYear=old.endHasYear;
  return fmt(parsed.start)+split+fmt(parsed.end,endHasYear||parsed.start.year!==parsed.end.year,dayOnly&&parsed.start.year===parsed.end.year&&parsed.start.month===parsed.end.month);
 }
 return fmt(role==='end'?(parsed.end||parsed.start):parsed.start);
}
export function parsePerformanceInput(value){
 const text=clean(value),result={};
 if(/[-−－]\s*\d+(?:\.\d+)?\s*%/.test(text))throw new Error('当前补充百分比应为0至100，不能填写负数。');
 const assign=(key,value)=>{if(result[key]&&result[key]!==value)throw new Error('同一业绩报酬事项给出了不同数值，请核对。');result[key]=value;};
 const action=String.raw`(?:计提|提取|提成|收取)\s*(?:业绩报酬)?\s*的?\s*(?:比例|比率)?\s*(?:为|是|=|:|按)?\s*`;
 for(const m of text.matchAll(new RegExp('(R[1-9])\\s*(?:为|是|=|:)?\\s*('+pct+')','gi')))assign('performance_'+m[1].toUpperCase(),percent(m[2]));
 for(const m of text.matchAll(new RegExp('第([1-9])档\\s*'+action+'('+pct+')','g')))assign('performance_ratio_'+m[1],percent(m[2]));
 // A line break within a condition and its payout is layout, not a new tier.
 const joined=text.replace(new RegExp('\\n\\s*(?=(?:'+action+'|(?:按|以)\\s*)\\d)','g'),' ');
 const clauses=joined.split(/[;；。\n]/).filter(Boolean),boundaries=[];const bands=[];
 const namedHurdlePattern=new RegExp('(?:业绩报酬(?:计提)?基准|计提基准)\\s*(?:为|是|=|:)?\\s*('+pct+')','g');
 const namedHurdles=[...new Set([...text.matchAll(namedHurdlePattern)].map(m=>Number(m[1].replace(/[^\d.]/g,''))))];
 const payoutPattern=new RegExp('(?:'+action+'|(?:按|以)\\s*)('+pct+')|(?:部分|超额收益|收益)\\s*的?\\s*('+pct+')(?=\\s*(?:的?比例)?\\s*(?:作为|计提|提取|提成|收取))|('+pct+')\\s*的比例\\s*(?:计提|提取|提成|收取)','g');
 const segments=clauses.flatMap(clause=>{const payouts=[...clause.matchAll(payoutPattern)];if(!payouts.length)return[{conditions:clause,payout:null}];let cursor=0;const pieces=payouts.map(payout=>{const conditions=clause.slice(cursor,payout.index);cursor=payout.index+payout[0].length;return{conditions,payout};});const trailing=clause.slice(cursor);if(new RegExp(pct).test(trailing))pieces.push({conditions:trailing,payout:null});return pieces;});
 for(const {conditions,payout}of segments){
  const isCondition=/[<>＜＞≤≥]|超过|超出|高于|大于|以上|区间/.test(conditions),namedHere=[...conditions.matchAll(namedHurdlePattern)].length>0;
  if(!isCondition&&!namedHere){
   if(!payout){if(/(?:比例|比率|提取|计提)\s*[:：为是]?\s*\d+(?:\.\d+)?\s*%/.test(conditions))throw new Error('有未识别的计提比例，请明确它对应的收益率区间。');continue;}
   if(/第[1-9]档/.test(conditions))continue;
   if(namedHurdles.length!==1)throw new Error('计提比例无法对应业绩报酬档位，请明确收益率区间或第几档。');
  }
  const values=[...conditions.matchAll(new RegExp(pct,'g'))].map(m=>Number(m[0].replace(/[^\d.]/g,'')));
  const namedSingle=payout&&!values.length&&namedHurdles.length===1&&!/R[1-9]/i.test(conditions);
  if(namedSingle)values.push(namedHurdles[0]);
  const upperBound=/[<＜≤]|不超过|未超过|不高于|不大于|低于|小于|区间/.test(conditions);
  if(values.length>2||values.length>1&&!upperBound)throw new Error('有多个百分比无法区分基准和计提比例，请明确收益率区间与提取比例。');
  if(values.length===2&&upperBound&&values[0]>=values[1])throw new Error('业绩报酬区间的下限必须小于上限。');
  boundaries.push(...values);if(payout)bands.push({clause:conditions,values,namedSingle:namedSingle||namedHere,ratio:percent(payout[1]||payout[2]||payout[3])});
 }
 const thresholds=[...new Set(boundaries)].sort((a,b)=>a-b);
 thresholds.forEach((x,i)=>{const key='performance_R'+(i+1);if(result[key]&&result[key]!==`${x}%`)throw new Error('业绩报酬基准标注与区间不一致，请核对。');result[key]=`${x}%`;});
 for(const band of bands){
  let tier=null;
  if(band.values.length>=2)tier=thresholds.indexOf(Math.min(...band.values))+1;
  else if(band.values.length===1&&(/[>＞]|(?<!不|未)(?:超过|超出|高于|大于)|以上/.test(band.clause)||band.namedSingle))tier=thresholds.indexOf(band.values[0])+1;
  else if(/R2/.test(band.clause)&&/[>＞]/.test(band.clause))tier=2;
  else if(/R1/.test(band.clause))tier=1;
  if(tier){const key='performance_ratio_'+tier;if(result[key]&&result[key]!==band.ratio)throw new Error('同一档业绩报酬给出了不同计提比例。');result[key]=band.ratio;}
  else throw new Error('有一段计提比例无法对应收益率区间，请核对比较符号和分段边界。');
 }
 if(!Object.keys(result).length){const only=text.match(new RegExp('^(?:业绩报酬(?:计提)?基准[:：]?\\s*)?('+pct+')(?:\\s*(?:业绩报酬(?:计提)?基准|业绩基准))?$'));if(only)result.performance_R1=percent(only[1]);}
 if(!Object.keys(result).length)throw new Error('未识别到业绩报酬分段，请填写收益率区间及各档计提比例。');
 if(Object.values(result).some(x=>Number(x.replace('%',''))>100))throw new Error('百分比应在0至100之间。');
 return result;
}
export function parseOpeningInput(value){
 const text=clean(value),out={};
 const monthly=monthlyWeekdayPatterns(text);if(monthly.length)out.opening_monthly_weekday=parseMonthlyWeekdayInput(text);
 for(const [key,re]of [['opening_cycle',/(?:每满|每隔|每|周期(?:为|是)?[:：]?)\s*(\d+)个?月/],['opening_days',/(?:连续开放|开放(?:期)?(?:为|是|开放)?[:：]?)\s*(\d+)个?交易日/],['opening_monthday',/每月(?:的)?\s*(\d+)日/],['opening_weekday',/(?:每周|每星期|周|星期)([一二三四五六日天])/]]){if(key==='opening_weekday'&&monthly.length)continue;const m=text.match(re);if(m)out[key]=m[1];}
 return out;
}
