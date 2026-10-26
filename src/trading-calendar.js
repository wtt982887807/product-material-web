// Mainland A-share exchange calendar, checked against official annual notices.
// Government make-up working weekends are NOT exchange trading days.
// Keep each year explicit: unsupported years must not fall back to weekdays.
const calendarYears = {
  2025: {
    publishedAt: '2024-12-23',
    title: '关于上海证券交易所2025年部分节假日休市安排的通知',
    url: 'https://www.sse.com.cn/disclosure/announcement/general/c/c_20241223_10767108.shtml',
    closedRanges: [
      ['01-01','01-01'], ['01-28','02-04'], ['04-04','04-06'],
      ['05-01','05-05'], ['05-31','06-02'], ['10-01','10-08'],
    ],
  },
  2026: {
    publishedAt: '2025-12-22',
    title: '关于上海证券交易所2026年部分节假日休市安排的通知',
    url: 'https://www.sse.com.cn/disclosure/announcement/general/c/c_20251222_10802507.shtml',
    corroboratingUrl: 'https://www.szse.cn/disclosure/notice/general/t20251222_618087.html',
    closedRanges: [
      ['01-01','01-03'], ['02-15','02-23'], ['04-04','04-06'],
      ['05-01','05-05'], ['06-19','06-21'], ['09-25','09-27'], ['10-01','10-07'],
    ],
  },
};

export const TRADING_CALENDAR_METADATA = Object.freeze({
  market: 'CN-A',
  label: '沪深股票交易日',
  supportedYears: Object.freeze(Object.keys(calendarYears).map(Number)),
  lastVerified: '2026-10-26',
});

const DAY = 86400000;
function dateStamp(value) {
  let year,month,day;
  if(value&&typeof value==='object'&&!Array.isArray(value)) {
    ({year,month,day}=value);
  } else {
    const match=String(value??'').trim().match(/^(\d{4})\s*(?:年|[./-])\s*(\d{1,2})\s*(?:月|[./-])\s*(\d{1,2})日?$/u);
    if(!match)throw new Error('交易日计算需要包含年份的有效日期。');
    [,year,month,day]=match;
  }
  [year,month,day]=[year,month,day].map(Number);
  if(![year,month,day].every(Number.isInteger)||year<1900||year>2200)throw new Error('交易日计算需要包含年份的有效日期。');
  const stamp=Date.UTC(year,month-1,day),date=new Date(stamp);
  if(date.getUTCFullYear()!==year||date.getUTCMonth()!==month-1||date.getUTCDate()!==day)throw new Error('交易日计算需要包含年份的有效日期。');
  return stamp;
}

/**
 * Inclusive, timezone-independent enumeration of verified CN-A trading dates.
 * Returns no partial dates when any year lacks an official built-in calendar.
 * Invalid/reversed bounds are caller errors; an unsupported year is a warning.
 */
export function getTradingDays(start,end) {
  const first=dateStamp(start),last=dateStamp(end);
  if(last<first)throw new Error('结束日期不能早于开始日期。');
  const firstYear=new Date(first).getUTCFullYear(),lastYear=new Date(last).getUTCFullYear();
  const years=Array.from({length:lastYear-firstYear+1},(_,i)=>firstYear+i);
  const unsupported=years.filter(year=>!calendarYears[year]);
  if(unsupported.length)return {dates:[],sources:[],warning:`尚未内置${unsupported.join('、')}年经交易所核实的交易日历，未推算分段开放日期，请补充具体日期后核对。`};
  const dates=[];
  for(let stamp=first;stamp<=last;stamp+=DAY) {
    const date=new Date(stamp),dayOfWeek=date.getUTCDay();
    if(dayOfWeek===0||dayOfWeek===6)continue;
    const year=date.getUTCFullYear(),month=date.getUTCMonth()+1,day=date.getUTCDate();
    const monthDay=String(month).padStart(2,'0')+'-'+String(day).padStart(2,'0');
    if(calendarYears[year].closedRanges.some(([begin,finish])=>monthDay>=begin&&monthDay<=finish))continue;
    dates.push(`${year}/${month}/${day}`);
  }
  const sources=years.map(year=>{
    const {title,url,publishedAt,corroboratingUrl}=calendarYears[year];
    return {year,title,url,publishedAt,...(corroboratingUrl?{corroboratingUrl}:{})};
  });
  return {dates,sources,warning:''};
}
