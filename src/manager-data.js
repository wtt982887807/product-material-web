import {parseXml,elements,sha256} from './office-xml.js';
import {completedYears,resolveProductionDate} from './production-date.js';

const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX_BYTES=30*1024*1024,MAX_XML=40*1024*1024,MAX_ROWS=20000,MAX_COLUMNS=256;
const child=(node,name)=>elements(node,null,name)[0];
const texts=node=>elements(node,null,'t').map(n=>n.textContent||'').join('');
const norm=value=>String(value??'').normalize('NFKC').replace(/[\s:：()（）【】\[\]]/g,'');
function columnIndex(address){let result=0;for(const c of String(address).match(/^[A-Z]+/i)?.[0]||'')result=result*26+c.toUpperCase().charCodeAt(0)-64;return result-1;}
function targetPath(base,target){
 if(!target||/^[a-z]+:/i.test(target))return null;
 const out=[];for(const part of(target.startsWith('/')?target.slice(1):base.slice(0,base.lastIndexOf('/')+1)+target).split('/')){if(!part||part==='.')continue;if(part==='..'){if(!out.length)return null;out.pop();}else out.push(part);}return out.join('/');
}
function headerKey(value){
 const s=norm(value);
 if(/^(?:(?:基金|投资)?经理)?姓名$|^(?:基金经理|投资经理)$/.test(s))return'name';
 if(/^(?:(?:基金|投资)?经理|个人|人员)?(?:简介|介绍|简历)$/.test(s))return'biography';
 if(/^(?:(?:证券|金融|基金)(?:行业)?)?(?:从业|工作)(?:(?:起始|开始)(?:日期|时间|日)?|日期|时间)$/.test(s))return'workStart';
 if(/^(?:投资(?:管理)?(?:(?:起始|开始)(?:日期|时间|日)?|日期|时间)|开始投资(?:日期|时间|日))$/.test(s))return'investmentStart';
 return'';
}
function headerAt(rows,r){
 const map=new Map();for(let c=0;c<MAX_COLUMNS;c++){
  const value=rows[r]?.[c]?.text;if(!value)continue;let key=headerKey(value);
  if(!key)for(let back=1;back<=2&&r-back>=0;back++){key=headerKey((rows[r-back]?.[c]?.text||'')+value);if(key)break;}
  if(key&&!map.has(key))map.set(key,c);
 }return map.has('name')&&map.has('biography')?map:null;
}
function textDate(value){
 const text=String(value??'').trim(),m=text.match(/^(\d{4})[年/.\-](\d{1,2})[月/.\-](\d{1,2})(?:日)?(?:[ T]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/);
 if(!m)return'';const date=`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
 try{return resolveProductionDate({manualValue:date});}catch{return'';}
}
function excelDate(value,date1904){
 if(!Number.isFinite(value)||value<0||(!date1904&&Math.floor(value)===60))return'';
 const day=Math.floor(value),offset=date1904?day:day-(day>=60?1:0),date=new Date((date1904?Date.UTC(1904,0,1):Date.UTC(1899,11,31))+offset*86400000);
 if(!Number.isFinite(date.getTime())||date.getUTCFullYear()<1900||date.getUTCFullYear()>2200)return'';
 return date.toISOString().slice(0,10);
}

/** Read cached values only. Formulas, external links and the 其他 column are never executed or used. */
export async function parseManagerProfile({name,bytes}={}){
 if(!/\.xlsx$/i.test(name||''))throw new Error('基金经理介绍请选择 XLSX 文件；可在 Excel 中另存为 XLSX。');
 if(!(bytes instanceof Uint8Array)&&!(bytes instanceof ArrayBuffer))throw new Error('基金经理 Excel 文件内容无效。');
 if(!bytes.byteLength||bytes.byteLength>MAX_BYTES)throw new Error('基金经理 Excel 文件为空或超过 30 MB。');
 const source=bytes instanceof Uint8Array?bytes.slice():new Uint8Array(bytes.slice(0));
 let zip;try{zip=await globalThis.JSZip.loadAsync(source);}catch(cause){throw new Error('无法读取这份 Excel，文件可能损坏、受密码保护或并非有效的 XLSX。请重新选择 XLSX 文件。',{cause});}
 const warnings=[],records=[],read=async path=>{const file=zip.file(path);if(!file)return null;const xml=await file.async('string');if(xml.length>MAX_XML)throw new Error('基金经理 Excel 内容过大，请只保留需要的资料。');return parseXml(xml);};
 const workbook=await read('xl/workbook.xml'),rels=await read('xl/_rels/workbook.xml.rels');if(!workbook||!rels)throw new Error('未找到有效的 Excel 工作簿，请重新选择 XLSX 文件。');
 const date1904=['1','true'].includes(child(workbook,'workbookPr')?.getAttribute('date1904')),paths=new Map();
 for(const rel of elements(rels,null,'Relationship'))if(rel.getAttribute('TargetMode')!=='External')paths.set(rel.getAttribute('Id'),targetPath('xl/workbook.xml',rel.getAttribute('Target')));
 const strings=await read('xl/sharedStrings.xml'),shared=strings?elements(strings,null,'si').map(texts):[];
 for(const [sheetIndex,sheet]of elements(workbook,null,'sheet').entries()){
  const sheetName=sheet.getAttribute('name')||`工作表${sheetIndex+1}`,path=paths.get(sheet.getAttributeNS(REL,'id')||sheet.getAttribute('r:id'));if(!path)continue;
  const doc=await read(path);if(!doc)continue;const rows=[];
  for(const row of elements(doc,null,'row')){
   const ri=Number(row.getAttribute('r'))-1;if(ri<0||ri>=MAX_ROWS){warnings.push(`${sheetName}：超出支持的行数范围，未读取该行。`);continue;}const out=rows[ri]||=[];
   for(const cell of elements(row,null,'c')){
    const ci=columnIndex(cell.getAttribute('r'));if(ci<0||ci>=MAX_COLUMNS)continue;
    const type=cell.getAttribute('t'),raw=child(cell,'v')?.textContent??'',formula=!!child(cell,'f');let text='',number;
    if(type==='s')text=shared[Number(raw)]??'';else if(type==='inlineStr')text=texts(cell);else if(type==='e')warnings.push(`${sheetName}!${cell.getAttribute('r')} 为 Excel 错误值，未导入。`);else if(type==='b')text=raw==='1'?'TRUE':'FALSE';else if(type==='str'||type==='d')text=raw;else if(raw!==''&&Number.isFinite(Number(raw))){number=Number(raw);text=raw;}else text=raw;
    if(formula&&raw==='')warnings.push(`${sheetName}!${cell.getAttribute('r')} 的公式没有保存计算结果，未导入。`);
    out[ci]={text,number,ref:cell.getAttribute('r')};
   }
  }
  // Only header merges are propagated; a blank biography never inherits another person's text.
  for(const merge of elements(doc,null,'mergeCell')){
   const[a,b]=merge.getAttribute('ref').split(':'),r1=Number(a?.match(/\d+$/)?.[0])-1,r2=Number(b?.match(/\d+$/)?.[0])-1,c1=columnIndex(a),c2=columnIndex(b),source=rows[r1]?.[c1];
   if(!source||!headerKey(source.text)||r1<0||r2>=50||c1<0||c2>=MAX_COLUMNS||(r2-r1+1)*(c2-c1+1)>10000)continue;
   for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){rows[r]||=[];rows[r][c]||=source;}
  }
  let start=-1,map=null;for(let r=0;r<Math.min(rows.length,50);r++){const candidate=headerAt(rows,r);if(candidate&&(!map||candidate.size>map.size)){start=r;map=candidate;}}
  if(!map){warnings.push(`${sheetName}：未找到“姓名”和“简介”表头，未导入此表。`);continue;}
  for(let r=start+1;r<rows.length;r++){
   if(!rows[r])continue;const repeated=headerAt(rows,r);if(repeated){map=repeated;continue;}
   const cell=key=>rows[r][map.get(key)],name=String(cell('name')?.text||'').trim(),biography=String(cell('biography')?.text||'');
   if(!name||!biography.trim()||/^(?:合计|总计|说明|备注)[：:]?$/.test(name))continue;
   const date=key=>{const c=cell(key);if(!c?.text?.trim())return'';const value=typeof c.number==='number'?excelDate(c.number,date1904):textDate(c.text);if(!value)warnings.push(`${sheetName}第${r+1}行${key==='workStart'?'从业日期':'投资起始日期'}不是有效的完整日期，保留简介中的原年限。`);return value;};
   records.push({id:`manager-${sheetIndex+1}-row-${r+1}`,name,biography,workStart:date('workStart'),investmentStart:date('investmentStart'),sheet:sheetName,row:r+1});
  }
 }
 if(!records.length)throw new Error('未读取到有效的基金经理简介，请检查 Excel 是否包含“姓名”和“简介”列及非空记录。');
 return{schemaVersion:1,name,sha256:await sha256(source),records,warnings};
}

// A duration must immediately describe experience; calendar/education years
// cannot match. Broad work experience and the narrower investment tenure use
// separate starts. Only the first non-subset claim of each kind is updated.
function experienceSpans(text){
 const re=/(?:超过|超|近|约)?\d+(?:\.\d+)?\s*(?:余|多)?年(?:以上)?\s*[^\d\n，,。；;：:、]{0,26}?经验/g,out=[];
 for(const match of text.matchAll(re)){
  // A four-digit calendar year can precede “投资经验” in historical prose.
  // It is never a duration, even though the surrounding words look similar.
  if(Number(match[0].match(/\d+(?:\.\d+)?/)?.[0])>=100)continue;
  const value=match[0],description=value.replace(/^(?:超过|超|近|约)?\d+(?:\.\d+)?\s*(?:余|多)?年(?:以上)?\s*/,''),kind=/从业|投研|研究投资/.test(description)?'work':/投资(?:组合)?(?:管理|及管理)?经验$/.test(description)?'investment':null;
  if(!kind)continue;const prefix=text.slice(Math.max(0,match.index-8),match.index);
  out.push({kind,start:match.index,before:value,description,subset:/其中\s*$/.test(prefix)});
 }return out;
}

export function prepareManagerBiography(record,productionDate){
 const text=String(record?.biography??''),date=resolveProductionDate({manualValue:productionDate}),spans=experienceSpans(text),calculations=[],edits=[];
 for(const kind of['work','investment']){
  const startDate=record?.[kind==='work'?'workStart':'investmentStart'];if(!startDate)continue;
  let years;try{years=completedYears(startDate,date);}catch{continue;}
  const candidates=spans.filter(s=>s.kind===kind),span=candidates.find(s=>!s.subset)||candidates[0];if(!span)continue;
  const after=`${years}年${span.description}`;
  calculations.push({kind,label:kind==='work'?'从业年限':'投资年限',startDate,productionDate:date,years,before:span.before,after});
  edits.push({...span,after});
 }
 let result=text;for(const edit of edits.sort((a,b)=>b.start-a.start))result=result.slice(0,edit.start)+edit.after+result.slice(edit.start+edit.before.length);
 return{text:result,calculations};
}
