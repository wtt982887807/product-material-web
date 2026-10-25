/** Capture the production month once per task; never infer it from business dates. */
export function currentMonth(now = new Date()) {
  const date = now instanceof Date ? now : new Date(now);
  if (Number.isNaN(date.getTime())) throw new Error('制作时间无效');
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone:'Asia/Shanghai', year:'numeric', month:'2-digit' }).formatToParts(date);
  return parts.find(p=>p.type==='year').value + parts.find(p=>p.type==='month').value;
}
export function resolveProductionMonth({manualValue,now=new Date()}={}) {
  if(manualValue===undefined)return currentMonth(now);
  const value=String(manualValue).trim();
  if(!/^20\d{2}-?(?:0[1-9]|1[0-2])$/.test(value))throw new Error('请选择有效的制作月份，例如2026年10月。');
  return value.replace('-','');
}
export function monthInputValue(month){const value=resolveProductionMonth({manualValue:month});return value.slice(0,4)+'-'+value.slice(4);}
const suffix = /(?:集合资产管理计划|单一资产管理计划|资产管理计划|集合资金信托计划|资金信托计划|信托计划)[。；;\s]*$/;
const short = value => String(value||'').trim().replace(suffix,'').replace(/[。；;]+$/,'');
const omitParenthetical = value => value.replace(/[（(][^）)]*[）)]/g,'');
const ignored = /[\s·•\-—（）()【】\[\]]/u;
function normalizedRange(text, candidate) {
  let normalized='', indices=[];
  for(let i=0;i<text.length;i++) if(!ignored.test(text[i])) {normalized+=text[i].toLowerCase();indices.push(i);}
  const needle=[...candidate].filter(c=>!ignored.test(c)).join('').toLowerCase();
  const index=normalized.indexOf(needle);
  if(!needle || index<0) return null;
  return {start:indices[index],end:indices[index+needle.length-1]+1};
}
function targetWithSameStyle(target, source, omitted) {
  let result=omitted?omitParenthetical(target):target;
  if(!/[（(]/.test(source)) result=result.replace(/[（）()]/g,'');
  else if(source.includes('(')) result=result.replace(/（/g,'(').replace(/）/g,')');
  else result=result.replace(/\(/g,'（').replace(/\)/g,'）');
  return result;
}
function updateProduct(stem, targetName, templateProductName) {
  const target=short(targetName), old=short(templateProductName);
  if(!target) throw new Error('合同尚未确定产品名称，不能生成文件名');
  for(const t of [target,omitParenthetical(target)]) if(normalizedRange(stem,t)) return stem;
  for(const omitted of [false,true]) {
    const candidate=omitted?omitParenthetical(old):old;
    const match=normalizedRange(stem,candidate);
    if(match) return stem.slice(0,match.start)+targetWithSameStyle(target,stem.slice(match.start,match.end),omitted)+stem.slice(match.end);
  }
  // A generic template filename may not contain a product name at all.
  if(!/[\d一二三四五六七八九十]+号/.test(stem)) return stem+' '+target;
  throw new Error('文件名中的原产品名称无法对应所选模板，请检查模板主体：'+stem);
}
function updateMonth(stem, month) {
  const month6=/(?<!\d)20\d{2}(?:0[1-9]|1[0-2])(?!\d)/g;
  const month4=/(?<!\d)\d{2}(?:0[1-9]|1[0-2])(?!\d)/g;
  let found=false;
  let result=stem.replace(/【([^】]*)】/,(_whole,prefix)=>{
    if(month6.test(prefix)) {month6.lastIndex=0;found=true;return '【'+prefix.replace(month6,month)+'】';}
    month6.lastIndex=0;
    if(month4.test(prefix)) {month4.lastIndex=0;found=true;return '【'+prefix.replace(month4,month.slice(2))+'】';}
    month4.lastIndex=0;found=true;return '【'+prefix+' '+month+'】';
  });
  if(!found) result='【'+month+'】'+result;
  return result;
}
export function deliveryName(originalName,{targetName,templateProductName='',kind,month=currentMonth()}={}) {
  if(!/^20\d{2}(?:0[1-9]|1[0-2])$/.test(month)) throw new Error('制作月份应为有效的YYYYMM');
  const match=String(originalName).match(/^(.*)(\.(?:docx|pptx))$/i);
  if(!match) throw new Error('只支持DOCX或PPTX输出');
  let stem=match[1].replace(/(?:[（(]\d+[）)])+$/,'').trim();
  stem=updateProduct(stem,targetName,templateProductName);
  stem=updateMonth(stem,month);
  if(kind==='ppt') stem=stem.replace(/(?:产品要素)+$/,'')+'产品要素';
  return stem+match[2].toLowerCase();
}
