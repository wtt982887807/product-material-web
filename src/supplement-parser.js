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
 for(const m of text.matchAll(new RegExp('(R[1-9])\\s*(?:为|是|=|:)?\\s*('+pct+')','g')))result['performance_'+m[1]]=percent(m[2]);
 for(const m of text.matchAll(new RegExp('第([1-9])档\\s*(?:计提|提取|提成|收取)(?:比例)?\\s*[:：]?\\s*('+pct+')','g')))result['performance_ratio_'+m[1]]=percent(m[2]);
 const clauses=text.split(/[;；\n]/).filter(Boolean),boundaries=[];const bands=[];
 for(const clause of clauses){
  const payout=clause.match(new RegExp('(?:计提|提取|提成|收取|按)\\s*('+pct+')'));
  const cutoff=payout?.index??clause.length;
  const conditions=clause.slice(0,cutoff);
  if(!/[<>＜＞≤≥]|超过|超出|高于|区间/.test(conditions))continue;
  const values=[...conditions.matchAll(new RegExp(pct,'g'))].map(m=>Number(m[0].replace(/[^\d.]/g,'')));
  if(values.length===2&&/[<＜≤]|区间/.test(conditions)&&values[0]>=values[1])throw new Error('业绩报酬区间的下限必须小于上限。');
  boundaries.push(...values);if(payout)bands.push({clause:conditions,values,ratio:percent(payout[1])});
 }
 const thresholds=[...new Set(boundaries)].sort((a,b)=>a-b);
 thresholds.forEach((x,i)=>{const key='performance_R'+(i+1);if(result[key]&&result[key]!==`${x}%`)throw new Error('业绩报酬基准标注与区间不一致，请核对。');result[key]=`${x}%`;});
 for(const band of bands){
  let tier=null;const high=thresholds.at(-1);
  if(band.values.length>=2)tier=thresholds.indexOf(Math.min(...band.values))+1;
  else if(band.values.length===1&&/[>＞]|超过|超出|高于/.test(band.clause))tier=thresholds.indexOf(band.values[0])+1;
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
 for(const [key,re]of [['opening_cycle',/(?:每满|每隔|每|周期(?:为|是)?[:：]?)\s*(\d+)个?月/],['opening_days',/(?:连续开放|开放(?:期)?(?:为|是|开放)?[:：]?)\s*(\d+)个?交易日/],['opening_monthday',/每月(?:的)?\s*(\d+)日/],['opening_weekday',/(?:每周|每星期|周|星期)([一二三四五六日天])/]]){const m=text.match(re);if(m)out[key]=m[1];}
 return out;
}
