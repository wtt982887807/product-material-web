import {parseXml, elements} from './office-xml.js';

const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const norm=value=>String(value??'').normalize('NFKC').replace(/[\s:：()（）【】\[\]]/g,'');
const trim=value=>String(value??'').trim();
const child=(node,name)=>elements(node,null,name)[0];
const texts=node=>elements(node,null,'t').map(n=>n.textContent||'').join('');
const numberText=value=>Number.isFinite(value)?String(Number(value.toPrecision(12))):'';
const MAX_ROWS=20000,MAX_COLUMNS=256;
function columnIndex(address){let result=0;for(const c of String(address).match(/^[A-Z]+/i)?.[0]||'')result=result*26+c.toUpperCase().charCodeAt(0)-64;return result-1;}
function targetPath(base,target){
 if(!target||/^[a-z]+:/i.test(target))return null;
 const parts=(target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target).split('/'),out=[];
 for(const part of parts){if(!part||part==='.')continue;if(part==='..'){if(!out.length)return null;out.pop();}else out.push(part);}return out.join('/');
}
function formatKind(format,id){
 const literal=String(format||'').replace(/"[^"]*"|\\.|\[[^\]]*\]/g,'');
 return{date:[14,15,16,17,18,19,20,21,22,27,28,29,30,31,32,33,34,35,36,45,46,47,50,51,52,53,54,55,56,57,58].includes(id)||/[yd]/i.test(literal)&&/[ymd]/i.test(literal),percent:[9,10].includes(id)||literal.includes('%'),zeros:/^0+$/.test(literal)?literal.length:0};
}
function excelDate(value,date1904){
 if(!Number.isFinite(value)||value<0||(!date1904&&Math.floor(value)===60))return '';
 const day=Math.floor(value),base=date1904?Date.UTC(1904,0,1):Date.UTC(1899,11,31),offset=date1904?day:day-(day>=60?1:0),date=new Date(base+offset*86400000);
 if(!Number.isFinite(date.getTime())||date.getUTCFullYear()<1900||date.getUTCFullYear()>2200)return '';
 return `${date.getUTCFullYear()}/${date.getUTCMonth()+1}/${date.getUTCDate()}`;
}
function validDate(text){
 const m=trim(text).match(/^(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})(?:日)?(?:[ T]00:00(?::00(?:\.0+)?)?)?$/);if(!m)return '';
 const y=+m[1],mo=+m[2],d=+m[3],date=new Date(Date.UTC(y,mo-1,d));return date.getUTCFullYear()===y&&date.getUTCMonth()===mo-1&&date.getUTCDate()===d?`${y}/${mo}/${d}`:'';
}
function headerKey(value){
 const s=norm(value).replace(/[%％]/g,'');
 if(/^(?:产品全称|产品名称|资管产品名称|产品简称)$/.test(s))return 'productName';
 if(/^(?:产品代码|产品编码|产品代销代码)$/.test(s))return 'productCode';
 if(/^(?:渠道名称|销售渠道|代销机构)$/.test(s))return 'channel';
 if(/^(?:营销阶段|销售阶段)$/.test(s))return 'stage';
 if(/(?:档期|开放|申购|募集)(?:起始|开始)日期|^(?:首个开放日|开放起始日)$/.test(s))return 'openingStart';
 if(/(?:档期|开放|申购|募集)(?:结束|截止)日期|^(?:最后开放日|开放结束日)$/.test(s))return 'openingEnd';
 if(/(?:档期|开放)(?:起始|开始)备注/.test(s))return 'startNote';
 if(/(?:档期|开放)(?:结束|截止)备注/.test(s))return 'endNote';
 if(/^(?:本次开放日|下一开放日|开放日期|开放日)$/.test(s))return 'openingDate';
 if(/^(?:开放日备注|开放备注|开放期备注)$/.test(s))return 'openingNote';
 if(/^(?:开放期|开放周期|开放频率|锁定期|封闭期)$/.test(s))return 'openingPeriod';
 if(/业绩报酬(?:计提)?基准(?:下限|下界)|(?:计提基准|业报基准)R1$/i.test(s))return 'performanceLower';
 if(/业绩报酬(?:计提)?基准(?:上限|上界)|(?:计提基准|业报基准)R2$/i.test(s))return 'performanceUpper';
 if(/区间内.*(?:计提|比例)|(?:第一档|第1档)(?:计提|提成)比例/.test(s))return 'performanceLowRatio';
 if(/超过上限.*(?:计提|比例)|(?:第二档|第2档)(?:计提|提成)比例/.test(s))return 'performanceHighRatio';
 if(/^(?:业绩报酬|业绩报酬计提|业绩报酬计提方式|业绩报酬计提规则|业报|业绩报酬说明)$/.test(s))return 'performanceText';
 if(/^(?:业绩比较基准|业绩基准|收益比较基准)$/.test(s))return 'benchmark';
 return '';
}
function findHeader(rows){
 let best=null;
 for(let r=0;r<Math.min(rows.length,30);r++){
  const map=new Map();for(let c=0;c<MAX_COLUMNS;c++){
   const current=trim(rows[r]?.[c]?.text);if(!current)continue;
   let key=headerKey(current),label=current;
   if(!key)for(let back=1;back<=2&&r-back>=0;back++){const parent=trim(rows[r-back]?.[c]?.text);if(parent&&parent!==current){key=headerKey(parent+current);if(key){label=parent+' / '+current;break;}}}
   if(key&&!map.has(key))map.set(key,{column:c,label});
  }
  const score=map.size;if(map.has('productName')&&score>=2&&(!best||score>best.score))best={row:r,map,score};
 }return best;
}
function percentage(cell){
 if(!cell)return '';const text=trim(cell.text);if(!text)return '';
 if(/^-?\d+(?:\.\d+)?\s*[%％]$/.test(text))return text.replace(/[\s％]/g,m=>m==='％'?'%':'');
 if(typeof cell.number==='number'&&Number.isFinite(cell.number))return numberText(cell.number)+'%';
 if(/^\d+(?:\.\d+)?$/.test(text))return text+'%';return text;
}
function canonicalName(name){
 return String(name??'').normalize('NFKC').replace(/\s/g,'').replace(/^(?:国泰君安|国泰海通)/,'').replace(/(?:集合资产管理计划|单一资产管理计划|集合资金信托计划|资产管理计划|集合计划)$/,'').replace(/[()（）【】\[\]·]/g,'');
}
/** Exact normalized product identity only; several matching rows remain selectable. */
export function matchCalendarRows(calendar,targetName){
 const target=canonicalName(targetName);if(!target)return [];
 return(calendar?.rows||[]).filter(row=>canonicalName(row.productName)===target);
}
/** Read an XLSX locally. No formulas, links, or embedded code are executed. */
export async function readProductCalendar({name,bytes}){
 if(!/\.xlsx$/i.test(name||''))throw new Error('产品日历请选择 XLSX 文件；可在 Excel 中另存为 XLSX。');
 if(!bytes||bytes.byteLength>30*1024*1024)throw new Error('产品日历文件为空或超过30 MB。');
 let zip;try{zip=await globalThis.JSZip.loadAsync(bytes);}catch(cause){throw new Error('无法读取这份 Excel，文件可能损坏、受密码保护或并非有效的 XLSX。请重新选择可正常打开的 XLSX 文件。',{cause});}const warnings=[],result=[];
 const read=async path=>{const entry=zip.file(path);if(!entry)return null;const text=await entry.async('string');if(text.length>40*1024*1024)throw new Error('产品日历表格过大，请只保留本次需要的产品。');return parseXml(text);};
 const wb=await read('xl/workbook.xml'),rels=await read('xl/_rels/workbook.xml.rels');if(!wb||!rels)throw new Error('未找到有效的 Excel 工作簿，请重新选择 XLSX 文件。');
 const date1904=['1','true'].includes(child(wb,'workbookPr')?.getAttribute('date1904')),paths=new Map();
 for(const rel of elements(rels,null,'Relationship'))if(rel.getAttribute('TargetMode')!=='External')paths.set(rel.getAttribute('Id'),targetPath('xl/workbook.xml',rel.getAttribute('Target')));
 const strings=await read('xl/sharedStrings.xml'),shared=strings?elements(strings,null,'si').map(texts):[];
 const styleDoc=await read('xl/styles.xml'),formats=new Map(styleDoc?elements(styleDoc,null,'numFmt').map(n=>[Number(n.getAttribute('numFmtId')),n.getAttribute('formatCode')]):[]);
 const xfs=styleDoc?child(styleDoc,'cellXfs'):null,styles=xfs?elements(xfs,null,'xf').map(n=>{const id=Number(n.getAttribute('numFmtId'));return formatKind(formats.get(id)||'',id);}):[];
 for(const sheet of elements(wb,null,'sheet')){
  const sheetName=sheet.getAttribute('name')||'工作表',path=paths.get(sheet.getAttributeNS(REL,'id')||sheet.getAttribute('r:id'));if(!path)continue;
  const doc=await read(path);if(!doc)continue;const rows=[];
  for(const row of elements(doc,null,'row')){
   const ri=Number(row.getAttribute('r'))-1;if(ri<0||ri>=MAX_ROWS){warnings.push(`${sheetName}：超出支持的行数范围，未读取该行。`);continue;}
   const out=rows[ri]||=[];
   for(const c of elements(row,null,'c')){
    const ci=columnIndex(c.getAttribute('r'));if(ci<0||ci>=MAX_COLUMNS)continue;
    const type=c.getAttribute('t'),raw=child(c,'v')?.textContent??'',style=styles[Number(c.getAttribute('s')||0)]||{},formula=!!child(c,'f');
    let text='',number;
    if(type==='s')text=shared[Number(raw)]??'';else if(type==='inlineStr')text=texts(c);else if(type==='e'){warnings.push(`${sheetName}!${c.getAttribute('r')} 为 Excel 错误值，未导入。`);}else if(type==='b')text=raw==='1'?'TRUE':'FALSE';else if(type==='str'||type==='d')text=raw;else if(raw!==''&&Number.isFinite(Number(raw))){number=Number(raw);text=style.date?excelDate(number,date1904):style.percent?numberText(number*100)+'%':style.zeros&&Number.isInteger(number)?String(number).padStart(style.zeros,'0'):numberText(number);}else text=raw;
    if(formula&&raw==='')warnings.push(`${sheetName}!${c.getAttribute('r')} 的公式没有保存计算结果，未导入。`);
    out[ci]={text,number,style,ref:c.getAttribute('r')};
   }
  }
  for(const merge of elements(doc,null,'mergeCell')){
   const [a,b]=merge.getAttribute('ref').split(':'),r1=Number(a?.match(/\d+$/)?.[0])-1,r2=Number(b?.match(/\d+$/)?.[0])-1,c1=columnIndex(a),c2=columnIndex(b),source=rows[r1]?.[c1];
   if(!source||r1<0||r2>=MAX_ROWS||c1<0||c2>=MAX_COLUMNS||(r2-r1+1)*(c2-c1+1)>10000)continue;
   for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){rows[r]||=[];rows[r][c]||=source;}
  }
  const header=findHeader(rows);if(!header){warnings.push(`${sheetName}：未找到“产品名称”等产品日历表头。`);continue;}
  for(let r=header.row+1;r<rows.length;r++){
   const cells=rows[r];if(!cells)continue;const cell=key=>cells[header.map.get(key)?.column],value=key=>trim(cell(key)?.text),productName=value('productName');if(!productName||headerKey(productName)==='productName'||/^(?:合计|总计|说明|备注)[:：]?\s*$/.test(productName))continue;
   const productCode=value('productCode'),source={file:name,sheet:sheetName,row:r+1},fields={};
   for(const [key,info]of header.map)if(value(key))fields[key]={value:value(key),cell:cell(key)?.ref||'',header:info.label};
   const openingStart=validDate(value('openingStart')),openingEnd=validDate(value('openingEnd')),startNote=value('startNote'),endNote=value('endNote'),openingNote=value('openingNote'),stage=value('stage');
   const isReservation=/预约|截止预约|家族端/.test(startNote+' '+endNote),dateEvent='opening';
   // The user's calendar defines this opening day as the schedule start date.
   // Keep the end date separately; never infer the start from a lone deadline.
   const openingDate=openingStart||validDate(value('openingDate'));
   if(!openingDate&&isReservation)warnings.push(`${sheetName}第${r+1}行：档期为预约截止安排，未自动用作开放日期。`);
   if(!openingDate&&openingEnd&&!openingStart&&!isReservation)warnings.push(`${sheetName}第${r+1}行：仅有档期结束日期，未推断首个开放日。`);
   const performanceLower=percentage(cell('performanceLower')),performanceUpper=percentage(cell('performanceUpper')),performanceLowRatio=percentage(cell('performanceLowRatio')),performanceHighRatio=percentage(cell('performanceHighRatio'));
   // Explicit R assignments retain the identity of a lone bound or payout tier.
   // Omitting an empty lower column must never renumber an upper bound as R1.
   const pieces=[];if(performanceLower)pieces.push(`R1=${performanceLower}`);if(performanceUpper)pieces.push(`R2=${performanceUpper}`);if(performanceLowRatio)pieces.push(`R1＜R≤R2，计提${performanceLowRatio}`);if(performanceHighRatio)pieces.push(`R＞R2，计提${performanceHighRatio}`);
   const performanceText=value('performanceText')||pieces.join('；');
   result.push({id:`calendar-${result.length+1}`,productName,productCode,channel:value('channel'),stage,openingDate,openingStart,openingEnd,openingDateEvent:dateEvent,openingDateNote:[startNote,endNote,openingNote].filter(Boolean).join('；'),openingPeriod:value('openingPeriod'),performanceText,performanceLower,performanceUpper,performanceLowRatio,performanceHighRatio,benchmark:percentage(cell('benchmark')),source,fields});
  }
 }
 if(!result.length)warnings.push('未读取到产品数据，可继续手动补充。');
 return {name,rows:result,warnings};
}

// Only the note for the end column can change that endpoint's meaning.
// A separate general note about reservations does not redefine the opening end.
export function calendarOpeningRange(row){
 const start=row.openingStart||row.openingDate||'',rawEnd=row.openingEnd||'',endNote=String(row.fields?.endNote?.value||'');
 const reservationEnd=!!rawEnd&&/(?:预约.{0,12}(?:截止|结束)|(?:截止|结束).{0,12}预约)/.test(endNote);
 const end=reservationEnd?'':rawEnd,value=start&&end?start+'至'+end:start;
 const warning=reservationEnd?`档期结束备注明确为预约截止安排，${rawEnd}未作为开放结束日填入；请单独核对本次开放结束日。`:'';
 return{start,end,value,warning};
}

// The complete schedule is one input: do not drop its end or its window rules.
export function calendarSupplementText(row){
 const {start,value}=calendarOpeningRange(row);
 // Recurring opening rules come from the contract, not the calendar supplement.
 const pairs=[['产品代码',row.productCode],['本次开放日',value],['业绩报酬',row.performanceText],['业绩比较基准',row.benchmark]];
 if(start&&row.openingDateNote)pairs.push(['分段规则',String(row.openingDateNote).replace(/[;；\r\n]+/g,'，')]);
 return pairs.filter(([,value])=>value!==undefined&&value!==null&&value!=='').map(([label,value])=>label+'：'+value).join('\n');
}
