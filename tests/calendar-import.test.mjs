import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {readProductCalendar,matchCalendarRows} from '../src/calendar-import.js';
Object.assign(globalThis,{JSZip,DOMParser});
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function col(n){let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
async function workbook(rows,{date1904=false,merges=[],formats=[],shared=[]}={}){
 const z=new JSZip();
 z.file('xl/workbook.xml',`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="${date1904?1:0}"/><sheets><sheet name="产品日历" sheetId="1" r:id="r1"/></sheets></workbook>`);
 z.file('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>');
 z.file('xl/styles.xml',`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts>${formats.map((f,i)=>`<numFmt numFmtId="${164+i}" formatCode="${esc(f)}"/>`).join('')}</numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="10"/>${formats.map((_,i)=>`<xf numFmtId="${164+i}"/>`).join('')}</cellXfs></styleSheet>`);
 if(shared.length)z.file('xl/sharedStrings.xml',`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${shared.map(s=>`<si><r><t>${esc(s.slice(0,2))}</t></r><r><t>${esc(s.slice(2))}</t></r></si>`).join('')}</sst>`);
 z.file('xl/worksheets/sheet1.xml',`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.map((row,ri)=>`<row r="${ri+1}">${row.map((v,ci)=>{const r=col(ci)+(ri+1);if(v==null)return '';if(typeof v==='object')return `<c r="${r}" s="${v.s||0}" t="${v.t||'n'}">${v.formula?'<f>'+esc(v.formula)+'</f>':''}${v.v!=null?'<v>'+esc(v.v)+'</v>':''}</c>`;if(typeof v==='number')return `<c r="${r}"><v>${v}</v></c>`;return `<c r="${r}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`;}).join('')}</row>`).join('')}</sheetData><mergeCells>${merges.map(ref=>`<mergeCell ref="${ref}"/>`).join('')}</mergeCells></worksheet>`);
 return{name:'测试日历.xlsx',bytes:await z.generateAsync({type:'uint8array'})};
}
const headers=['产品名称','产品代码','渠道名称','营销阶段','档期起始日期','档期结束日期','档期结束备注','业绩报酬计提基准下限%','业绩报酬计提基准上限%','区间内部计提超额部分比例%','超过上限部分计提超额部分比例%','锁定期','业绩比较基准'];

test('calendar columns preserve lower/upper hurdles, payout ratios and independent benchmark',async()=>{
 const calendar=await readProductCalendar(await workbook([headers,['青澜定开债8号','00953794','建设银行','持营开放','2026/11/23','2026/11/27','前2个开放日参与退出，后3个仅参与',2.05,2.35,5,50,'每12个月',{v:0.0225,s:2}]]));
 assert.equal(calendar.rows.length,1);const r=calendar.rows[0];
 assert.equal(r.productCode,'00953794');assert.equal(r.openingDate,'2026/11/23');assert.equal(r.openingEnd,'2026/11/27');assert.equal(r.openingPeriod,'每12个月');assert.equal(r.performanceText,'R1=2.05%；R2=2.35%；R1＜R≤R2，计提5%；R＞R2，计提50%');assert.equal(r.benchmark,'2.25%');assert.equal(r.source.row,2);assert.equal(r.fields.performanceLower.cell,'H2');assert.equal(r.openingDateEvent,'opening');
});

test('appointment end is never silently imported as an opening date',async()=>{
 const calendar=await readProductCalendar(await workbook([headers,['青澜信用优享18个月1号','953957','建设银行','持营开放',null,'2026/11/17','家族端预约截止日',2.45,2.8,5,40,'每18个月']]));
 const r=calendar.rows[0];assert.equal(r.openingDate,'');assert.equal(r.openingEnd,'2026/11/17');assert.equal(r.openingDateEvent,'opening');assert.equal(r.benchmark,'');assert.match(calendar.warnings.join(' '),/未自动用作开放日期/);
});

test('schedule start is the opening day even when the marketing stage is fundraising',async()=>{
 const calendar=await readProductCalendar(await workbook([headers,['青澜多元稳健19号','','建设银行','首发募集','2026-11-04','2026-11-18','',3.05,3.4,5,30,'每18个月']]));
 assert.equal(calendar.rows[0].openingDateEvent,'opening');assert.equal(calendar.rows[0].openingDate,'2026/11/4');assert.equal(calendar.rows[0].stage,'首发募集');assert.equal(calendar.rows[0].openingEnd,'2026/11/18');
});

test('Excel serial dates honor 1900 and 1904 systems and percentages do not double-scale',async()=>{
 for(const [date1904,serial]of[[false,46349],[true,44887]]){
  const calendar=await readProductCalendar(await workbook([headers,['青澜8号',{v:1234,s:3},'建设银行','持营开放',{v:serial,s:1},{v:serial+4,s:1},'',{v:.02,s:2},{v:.03,s:2},{v:.2,s:2},{v:.4,s:2},'每12个月']],{date1904,formats:['000000']}));
  const r=calendar.rows[0];assert.equal(r.productCode,'001234');assert.equal(r.openingDate,'2026/11/23');assert.equal(r.openingEnd,'2026/11/27');assert.equal(r.performanceText,'R1=2%；R2=3%；R1＜R≤R2，计提20%；R＞R2，计提40%');
 }
});

test('merged two-level headers and shared rich strings are read without manufacturing empty cells',async()=>{
 const rows=[['产品名称','产品代码','业绩报酬计提基准',null,'区间内部计提超额部分比例%', '超过上限部分计提超额部分比例%'],[null,null,'下限%','上限%'],[{v:0,t:'s'},{v:1,t:'s'},2,3,20,40]];
 const c=await readProductCalendar(await workbook(rows,{merges:['A1:A2','B1:B2','C1:D1','E1:E2','F1:F2'],shared:['青澜信用优享18个月5号','000123']}));
 assert.equal(c.rows.length,1);assert.equal(c.rows[0].productName,'青澜信用优享18个月5号');assert.equal(c.rows[0].productCode,'000123');assert.equal(c.rows[0].performanceLower,'2%');assert.equal(c.rows[0].performanceUpper,'3%');assert.equal(c.rows[0].source.row,3);
});

test('matching excludes different tenor and serial number, and leaves duplicate rows selectable',()=>{
 const rows=['国泰海通青澜（信用优享）18个月5号','青澜(信用优享)18个月1号','青澜（信用优享）12个月5号','青澜（信用优享）18个月5号'].map((productName,i)=>({id:'row'+i,productName}));
 assert.deepEqual(matchCalendarRows({rows},'国泰君安青澜（信用优享）18个月5号集合资产管理计划').map(r=>r.id),['row0','row3']);
 assert.deepEqual(matchCalendarRows({rows},'国泰君安青澜（信用优享）18个月6号集合资产管理计划'),[]);
});

test('invalid dates and formula errors remain absent while valid fields still import',async()=>{
 const c=await readProductCalendar(await workbook([headers,['青澜8号','009999',null,null,'2026/2/30','2026/3/2',null,{v:'#VALUE!',t:'e'},3,{formula:'1/0'},40]]));
 assert.equal(c.rows[0].openingDate,'');assert.equal(c.rows[0].performanceLower,'');assert.equal(c.rows[0].performanceLowRatio,'');assert.match(c.warnings.join(' '),/错误值/);assert.match(c.warnings.join(' '),/没有保存计算结果/);assert.match(c.warnings.join(' '),/仅有档期结束日期/);
});

test('unsupported input produces an actionable error without executing content',async()=>{
 await assert.rejects(readProductCalendar({name:'calendar.xls',bytes:new Uint8Array([1])}),/请选择 XLSX/);
 await assert.rejects(readProductCalendar({name:'损坏.xlsx',bytes:new Uint8Array([1,2,3])}),/无法读取这份 Excel.*重新选择/);
 const z=new JSZip();z.file('xl/workbook.xml','<!DOCTYPE x [<!ENTITY e SYSTEM "file:///secret">]><workbook/>');z.file('xl/_rels/workbook.xml.rels','<Relationships/>');
 await assert.rejects(readProductCalendar({name:'calendar.xlsx',bytes:await z.generateAsync({type:'uint8array'})}),/外部声明/);
});

test('partial numeric columns preserve the R2 and second payout identities',async()=>{
 const c=await readProductCalendar(await workbook([headers,['青澜8号','001234',null,null,null,null,null,null,3,null,40]]));
 assert.equal(c.rows[0].performanceText,'R2=3%；R＞R2，计提40%');
 assert.equal(c.rows[0].performanceLower,'');assert.equal(c.rows[0].performanceLowRatio,'');
 const {parsePerformanceInput}=await import('../src/supplement-parser.js');
 const parsed=parsePerformanceInput(c.rows[0].performanceText);
 assert.equal(parsed.performance_R1,undefined);assert.equal(parsed.performance_ratio_1,undefined);
 assert.equal(parsed.performance_R2,'3%');assert.equal(parsed.performance_ratio_2,'40%');
});
