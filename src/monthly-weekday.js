// A monthly ordinal weekday is one schedule, not a weekly weekday or month day.
const MONTH=String.raw`每(?:个)?(?:自然)?月(?:的)?\s*`;
const WEEK=String.raw`(?:周|星期|礼拜)\s*[一二三四五六日天1-7]`;
const SEPARATOR=/\s*(?:以及|[、，,和及与])\s*/g;
const NUMBERS={一:1,二:2,两:2,三:3,四:4,五:5,六:6,七:7,日:7,天:7};
const chinese=n=>['','一','二','三','四','五','六','日'][n];
const integer=s=>NUMBERS[s]??Number(s);
const sorted=entries=>entries.map((e,index)=>({...e,index})).sort((a,b)=>a.ordinal-b.ordinal||a.weekday-b.weekday);
function canonical(entries){
 const ordered=sorted(entries),same=ordered.every(e=>e.weekday===ordered[0].weekday);
 return '每月'+ordered.map(e=>`第${e.ordinal}个${same?'':'周'+chinese(e.weekday)}`).join('、')+(same?'周'+chinese(ordered[0].weekday):'');
}
function parseFragment(before,start){
 const prefix=before.match(new RegExp('^'+MONTH)),body=before.slice(prefix[0].length),pieces=[];
 let at=0;for(const m of body.matchAll(SEPARATOR)){pieces.push({text:body.slice(at,m.index),start:prefix[0].length+at});at=m.index+m[0].length;}pieces.push({text:body.slice(at),start:prefix[0].length+at});
 const entries=[],tokens=[];let awaiting=[];
 for(const piece of pieces){
  const m=piece.text.match(/^(?:第\s*)?([一二两三四五六七八九十\d]+)\s*个?\s*(?:(?:周|星期|礼拜)\s*([一二三四五六日天1-7]))?$/);
  if(!m)return{start,before,warning:'月内第几个星期的写法尚无法完整对应，保留原安排待核。'};
  const ordinal=integer(m[1]);if(!Number.isInteger(ordinal)||ordinal<1||ordinal>5)return{start,before,warning:'月内星期序号须为第1至第5个，保留原安排待核。'};
  const index=entries.length;entries.push({ordinal,weekday:null});awaiting.push(index);
  tokens.push({role:'ordinal',entryIndices:[index],start:piece.start+piece.text.indexOf(m[1]),before:m[1]});
  if(m[2]){const weekday=integer(m[2]);for(const i of awaiting)entries[i].weekday=weekday;
   tokens.push({role:'weekday',entryIndices:awaiting,start:piece.start+piece.text.lastIndexOf(m[2]),before:m[2]});awaiting=[];
  }
 }
 if(awaiting.length||!entries.length||new Set(entries.map(e=>e.ordinal+'|'+e.weekday)).size!==entries.length)return{start,before,warning:'月内开放安排包含未明确的星期或重复序号，保留原安排待核。'};
 return{start,before,entries,tokens,value:canonical(entries)};
}
export function monthlyWeekdayPatterns(text){
 // Retain unsupported ordinal wording as a pending fragment instead of silently
 // omitting it when another part of the same cell was successfully recognized.
 const re=new RegExp(MONTH+'(?:第|首|末|最后)[^。；;\\n：:]{0,70}?'+WEEK+'(?:\\s*(?:以及|[、，,和及与])\\s*第?[^。；;\\n：:]{0,25}?'+WEEK+')*','g');
 return [...String(text).matchAll(re)].map(m=>{
  const tail=String(text).slice(m.index+m[0].length).match(/^\s*(?:以及|[、，,和及与/／])\s*(?:(?:第|首|末|最后)[^。；;\n（(）)]{0,30}|(?:周|星期|礼拜)\s*[一二三四五六日天1-7])/);
  if(tail)return{start:m.index,before:m[0]+tail[0],warning:'月内开放安排还有未完整对应的枚举项，保留整个安排待核。'};
  return parseFragment(m[0],m.index);
 });
}
export function parseMonthlyWeekdayInput(text){
 const patterns=monthlyWeekdayPatterns(String(text).normalize('NFKC').trim());
 if(patterns.length!==1||patterns[0].warning)throw new Error(patterns[0]?.warning||'请填写完整月内开放安排，例如每月第1个、第3个周三。');
 return patterns[0].value;
}
export function monthlyWeekdayChanges(pattern,value){
 if(pattern.warning)return{warning:pattern.warning,changes:[]};
 const next=monthlyWeekdayPatterns(value);if(next.length!==1||next[0].warning)return{warning:'补充的月内开放安排无法完整对应，保留原值待核。',changes:[]};
 const oldOrder=sorted(pattern.entries),newOrder=sorted(next[0].entries);
 if(oldOrder.length!==newOrder.length)return{warning:'合同与模板的月内开放日项数不同，无法仅替换序号；保留原安排待核。',changes:[]};
 const targets=new Map(oldOrder.map((e,i)=>[e.index,newOrder[i]])),changes=[];
 for(const token of pattern.tokens){
  const values=[...new Set(token.entryIndices.map(i=>targets.get(i)[token.role]))];
  if(values.length!==1)return{warning:'模板共用的星期与新安排不能逐项对应，保留原安排待核。',changes:[]};
  const n=values[0];if(integer(token.before)===n)continue;
  const after=/^\d+$/.test(token.before)?String(n).padStart(token.before.length,'0'):n===7&&token.before==='天'?'天':chinese(n);
  if(after!==token.before)changes.push({start:token.start,before:token.before,after});
 }
 return{changes};
}
