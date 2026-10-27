// Simplify redundant presentation only. Contract resolution and user wording stay separate.
const quantity='(\\d+(?:\\.\\d+)?)(亿元|万元|元|自然日|交易日|工作日|天|日|个月|月|年)';
const subject='(?:参与金额|认购金额|申购金额|持有天数|持有期限|持有期|持有时间)';
const fee='((?:认购|申购|参与|赎回|退出)(?:[/、及和](?:认购|申购|参与|赎回|退出))*(?:费率|费))';
const value=(n,u)=>({value:Number(n)*({亿元:1e8,万元:1e4,元:1}[u]||1),unit:/元$/.test(u)?'amount':/^(?:自然日|天|日)$/.test(u)?'day':/月/.test(u)?'month':u});
function band(text){
 const t=text.normalize('NFKC').replace(/\s/g,'').replace(/>=/g,'≥').replace(/<=/g,'≤');
 let m=t.match(new RegExp('^'+subject+'([<≤>≥])'+quantity+'$'));
 if(m){const q=value(m[2],m[3]);return {...q,low:/[>≥]/.test(m[1])?q.value:0,high:/[<≤]/.test(m[1])?q.value:Infinity,lowClosed:m[1]!=='＞'&&m[1]!=='>',highClosed:m[1]==='≤'};}
 m=t.match(new RegExp('^'+quantity+'([<≤])'+subject+'([<≤])'+quantity+'$'));
 if(!m)return null;
 const a=value(m[1],m[2]),b=value(m[5],m[6]);if(a.unit!==b.unit)return null;
 return {unit:a.unit,low:a.value,high:b.value,lowClosed:m[3]==='≤',highClosed:m[4]==='≤'};
}
function clause(text){
 const t=text.normalize('NFKC').replace(/\s/g,'');
 const m=t.match(new RegExp('^(.+?)[,，]'+fee+'(为|是|[:：])?(\\d+(?:\\.\\d+)?)%$'));
 if(!m)return null;const range=band(m[1]);if(!range||range.high<=range.low)return null;
 const dimension=m[1].match(new RegExp(subject))?.[0]?.replace(/^持有.*$/,'持有');
 // Do not reinterpret a holding condition as a subscription amount or vice versa.
 const exit=/退出|赎回/.test(m[2]);if(exit&&/认购|申购|参与/.test(m[2])||exit===(range.unit==='amount'))return null;
 return {...range,dimension,label:m[2],join:m[3]||'',rate:Number(m[4])};
}
function complete(items){
 const sorted=[...items].sort((a,b)=>a.low-b.low);
 if(sorted[0].low!==0||!sorted[0].lowClosed||sorted.at(-1).high!==Infinity)return false;
 return sorted.every((x,i)=>!i||sorted[i-1].high===x.low&&sorted[i-1].highClosed!==x.lowClosed);
}
export function simplifyEqualFeeTiers(text){
 if(/最低|最高|类份额|渠道|优惠|折扣|另行|公告|代销|管理人|受托人/.test(text))return text;
 const chunks=[...text.matchAll(/([^；;。\n]+)([；;。\n]*|$)/g)].map(m=>({start:m.index,end:m.index+m[0].length,suffix:m[2],parsed:clause(m[1])}));
 const changes=[];
 for(let i=0;i<chunks.length;i++){
  const first=chunks[i],a=first.parsed;if(!a||a.rate<0||a.rate>100)continue;
  let j=i+1;while(j<chunks.length&&chunks[j].parsed?.label===a.label&&chunks[j].parsed.unit===a.unit&&chunks[j].parsed.dimension===a.dimension)j++;
  const group=chunks.slice(i,j),last=group.at(-1);
  if(group.length>1&&group.every(c=>c.parsed.rate===a.rate)&&complete(group.map(c=>c.parsed)))changes.push({start:first.start,end:last.end,after:a.label+a.join+a.rate+'%'+last.suffix});
  i=j-1;
 }
 let after=text;for(const c of changes.reverse())after=after.slice(0,c.start)+c.after+after.slice(c.end);
 return after;
}
