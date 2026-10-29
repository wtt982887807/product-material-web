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
const compact = value => [...value].filter(c=>!ignored.test(c)).join('').toLowerCase();
const productNumber = /[\d一二三四五六七八九十]+号$/u;
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
function filenameAbbreviation(stem, fullName) {
  // Only the complete product-bearing filename body can prove an abbreviation.
  // A number elsewhere in a channel heading or marketing phrase is not identity.
  const header=stem.match(/^(?:\s*【[^】]*】\s*)*/u)[0];
  const body=stem.slice(header.length),name=short(body.replace(/(?:产品要素)+$/u,''));
  if(!productNumber.test(name)||(name.replace(productNumber,'').match(/\p{L}/gu)||[]).length<2)return null;
  const match=normalizedRange(fullName,name);
  if(!match||!match.start||compact(fullName.slice(match.end)))return null;
  const start=header.length+body.indexOf(name);
  return {name,start,end:start+name.length,omittedPrefix:fullName.slice(0,match.start)};
}
function abbreviatedTarget(target, source) {
  // Reuse an evidenced omitted prefix. For an issuer rename, an unchanged full
  // series followed by its contract number also proves the display abbreviation.
  const series=source.name.replace(productNumber,''),match=normalizedRange(target,series);
  if(match&&/^[\d一二三四五六七八九十]+号$/u.test(compact(target.slice(match.end))))return target.slice(match.start);
  const prefix=normalizedRange(target,source.omittedPrefix);
  if(prefix?.start===0&&compact(target.slice(prefix.end)))return target.slice(prefix.end);
  // When both issuer and series differ, do not invent which new words to omit.
  return target;
}
function productRange(stem, candidate) {
  const match=normalizedRange(stem,candidate);
  // A matched product cannot justify a trailing old phase or placeholder.
  // Only non-product delivery suffixes can survive replacement.
  const header=stem.match(/^(?:\s*【[^】]*】\s*)*/u)[0];
  return match&&!stem.slice(header.length,match.start).trim()&&!short(stem.slice(match.end).trim().replace(/(?:产品要素)+$/u,''))?match:null;
}
function updateProduct(stem, targetName, templateProductName) {
  const target=short(targetName), old=short(templateProductName);
  if(!target) throw new Error('合同尚未确定产品名称，不能生成文件名');
  for(const t of [target,omitParenthetical(target)]) if(productRange(stem,t)) return stem;
  for(const omitted of [false,true]) {
    const candidate=omitted?omitParenthetical(old):old;
    const match=productRange(stem,candidate);
    if(match) return stem.slice(0,match.start)+targetWithSameStyle(target,stem.slice(match.start,match.end),omitted)+stem.slice(match.end);
  }
  for(const omitted of [false,true]) {
    const candidate=omitted?omitParenthetical(old):old,source=filenameAbbreviation(stem,candidate);
    if(source){
      const display=abbreviatedTarget(omitted?omitParenthetical(target):target,source);
      return stem.slice(0,source.start)+targetWithSameStyle(display,source.name,omitted)+stem.slice(source.end);
    }
  }
  for(const candidate of [target,omitParenthetical(target)])if(filenameAbbreviation(stem,candidate))return stem;
  // Filenames are presentation, not an identity gate. Unknown aliases, old
  // numbers and phase placeholders are replaced together by the contract name.
  const header=stem.match(/^(?:\s*【[^】]*】\s*)*/u)[0];
  return header+target;
}
function updateMonth(stem, month) {
  const date8=/(?<!\d)20\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])(?!\d)/g;
  const month6=/(?<!\d)20\d{2}(?:0[1-9]|1[0-2])(?!\d)/g;
  const month4=/(?<!\d)\d{2}(?:0[1-9]|1[0-2])(?!\d)/g;
  let found=false;
  // Production dates belong to leading delivery headers, never to product names.
  const header=stem.match(/^(?:\s*【[^】]*】\s*)*/u)[0];
  let result=header.replace(/【([^】]*)】/g,(_whole,prefix)=>{
    prefix=prefix.replace(date8,()=>{found=true;return month;});
    if(month6.test(prefix)) {month6.lastIndex=0;found=true;return '【'+prefix.replace(month6,month)+'】';}
    month6.lastIndex=0;
    if(month4.test(prefix)) {month4.lastIndex=0;found=true;return '【'+prefix.replace(month4,month.slice(2))+'】';}
    month4.lastIndex=0;return _whole;
  });
  if(!found) result=header?result.replace(/【([^】]*)】/,(_whole,prefix)=>'【'+prefix+' '+month+'】'):'【'+month+'】';
  result+=stem.slice(header.length);
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

/** Keep previews, individual links and ZIP entries unique on Windows as well. */
export function uniqueDeliveryNames(names) {
  const key=name=>name.normalize('NFC').toLowerCase();
  const reserved=new Set(names.map(key)),used=new Set();
  return names.map(name=>{
    let result=name;
    if(used.has(key(result))){
      const dot=name.lastIndexOf('.'),stem=dot<0?name:name.slice(0,dot),ext=dot<0?'':name.slice(dot);
      let count=2;
      do{result=`${stem}（${count++}）${ext}`;}while(used.has(key(result))||reserved.has(key(result)));
    }
    used.add(key(result));return result;
  });
}

/** Mode-aware display names; introduction-only never infers a shared product. */
export function materialDeliveryName(material,{targetName,month=currentMonth()}={}){
 if(material.updateMode==='introductions'){
  if(!/^20\d{2}(?:0[1-9]|1[0-2])$/.test(month))throw new Error('制作月份应为有效的YYYYMM');
  const match=String(material.template?.name||'').match(/^(.*)(\.(?:docx|pptx))$/i);
  if(!match)throw new Error('只支持DOCX或PPTX输出');
  let stem=match[1];
  if(material.kind==='ppt')stem=stem.replace(/(?:产品要素及介绍|公司及经理介绍|产品要素)+$/,'');
  stem=updateMonth(stem,month);
  return stem+(material.kind==='ppt'?'公司及经理介绍':'')+match[2].toLowerCase();
 }
 const name=deliveryName(material.template.name,{targetName,templateProductName:material.templateProductName,kind:material.kind,month});
 const hasProfiles=['company','manager'].some(key=>material.profilePages?.[key]?.length);
 return material.kind==='ppt'&&hasProfiles?name.replace(/产品要素\.pptx$/,'产品要素及介绍.pptx'):name;
}
