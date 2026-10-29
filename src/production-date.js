const SHANGHAI_DATE=new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'});

function fullDate(value){
 if(typeof value!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(value))throw new Error('制作日期请填写完整的年、月、日（YYYY-MM-DD）。');
 const [year,month,day]=value.split('-').map(Number),date=new Date(0);
 date.setUTCFullYear(year,month-1,day);date.setUTCHours(0,0,0,0);
 if(year<1||date.getUTCFullYear()!==year||date.getUTCMonth()!==month-1||date.getUTCDate()!==day)throw new Error('日期无效，请检查年、月、日。');
 return value;
}

/** Resolve the current calendar day in China, independent of the browser zone. */
export function currentProductionDate(now=new Date()){
 const date=now instanceof Date?now:new Date(now);
 if(!Number.isFinite(date.getTime()))throw new Error('当前时间无效，无法确定制作日期。');
 const parts=Object.fromEntries(SHANGHAI_DATE.formatToParts(date).map(p=>[p.type,p.value]));
 return fullDate(`${parts.year.padStart(4,'0')}-${parts.month}-${parts.day}`);
}

export function resolveProductionDate({manualValue,now}={}){
 return manualValue===undefined||manualValue===null||manualValue===''?currentProductionDate(now):fullDate(manualValue);
}

export function productionMonth(date){return fullDate(date).slice(0,7).replace('-','');}

/** A leap-day anniversary is complete on March 1 in a non-leap year. */
export function completedYears(start,date){
 fullDate(start);fullDate(date);
 if(start>date)throw new Error('起始日期晚于制作日期，不能计算经验年限。');
 return Number(date.slice(0,4))-Number(start.slice(0,4))-(date.slice(5)<start.slice(5)?1:0);
}
