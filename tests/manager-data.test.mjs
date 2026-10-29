import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
Object.assign(globalThis,{JSZip,DOMParser});
const api=await import('../src/manager-data.js').catch(error=>{if(error.code==='ERR_MODULE_NOT_FOUND')return{};throw error;});
const call=(name,...args)=>{assert.equal(typeof api[name],'function',`${name} is implemented`);return api[name](...args);};
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const headers=['姓名','简介','其他','从业日期','投资起始日期'];
async function workbook(sheets,{date1904=false}={}){
 const z=new JSZip();
 z.file('xl/workbook.xml',`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="${+date1904}"/><sheets>${sheets.map((s,i)=>`<sheet name="${s.name||'经理'+i}" r:id="r${i}"/>`).join('')}</sheets></workbook>`);
 z.file('xl/_rels/workbook.xml.rels',`<Relationships>${sheets.map((s,i)=>`<Relationship Id="r${i}" Target="worksheets/sheet${i+1}.xml"/>`).join('')}</Relationships>`);
 z.file('xl/styles.xml','<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>');
 for(const [i,sheet]of sheets.entries())z.file(`xl/worksheets/sheet${i+1}.xml`,`<worksheet><sheetData>${sheet.rows.map((row,r)=>`<row r="${r+1}">${row.map((value,c)=>{const ref=String.fromCharCode(65+c)+(r+1);if(value==null)return'';if(typeof value==='object')return`<c r="${ref}" t="${value.type||'n'}" s="${value.style||0}">${value.formula?'<f>1+1</f>':''}${value.value!=null?'<v>'+esc(value.value)+'</v>':''}</c>`;return`<c r="${ref}" t="inlineStr"><is><t>${esc(value)}</t></is></c>`;}).join('')}</row>`).join('')}</sheetData></worksheet>`);
 return{name:'经理.xlsx',bytes:await z.generateAsync({type:'uint8array'})};
}
test('all biography text and duplicate source rows survive across sheets, notes are excluded',async()=>{
 const biography=' 10年证券从业经验。\n现任投资经理。\n投资特点\n长期研究。 ';
 const file=await workbook([{name:'公募',rows:[['说明'],headers,['甲',biography,'内部待补充','2016年10月30日','2019/10/29'],['甲','第二版简介','非简介']]},{name:'私募',rows:[headers,['乙','完整简介']]}]);
 const profile=await call('parseManagerProfile',file);
 assert.equal(profile.schemaVersion,1);assert.equal(profile.name,file.name);assert.match(profile.sha256,/^[a-f0-9]{64}$/);assert.equal(profile.records.length,3);
 assert.equal(profile.records[0].biography,biography);assert.equal(profile.records[0].workStart,'2016-10-30');assert.equal(profile.records[0].investmentStart,'2019-10-29');assert.equal(profile.records[0].sheet,'公募');assert.equal(profile.records[0].row,3);
 assert.equal(new Set(profile.records.map(r=>r.id)).size,3);assert.doesNotMatch(JSON.stringify(profile.records),/内部待补充|非简介/);
});
test('date columns accept Excel serials in both date systems and reject missing precision or invalid days',async()=>{
 for(const [date1904,serial]of[[false,46349],[true,44887]]){
  const p=await call('parseManagerProfile',await workbook([{rows:[headers,['甲','9年从业经验',{value:'unused'},{value:serial,style:1},{value:serial}]]}],{date1904}));
  assert.equal(p.records[0].workStart,'2026-11-23');assert.equal(p.records[0].investmentStart,'2026-11-23');
 }
 const p=await call('parseManagerProfile',await workbook([{rows:[headers,['甲','8年从业经验','','2020-02-30','2020年2月'],['乙','6年从业经验','',{value:60,style:1},{type:'d',value:'2020-02-29T00:00:00Z'}]]}]));
 assert.equal(p.records[0].workStart,'');assert.equal(p.records[0].investmentStart,'');assert.equal(p.records[1].workStart,'');assert.equal(p.records[1].investmentStart,'2020-02-29');assert.ok(p.warnings.length>=3);
});
test('spreadsheet errors and missing formula caches do not become start dates',async()=>{
 const p=await call('parseManagerProfile',await workbook([{rows:[headers,['甲','原简介','',{type:'e',value:'#VALUE!'},{formula:true}]]}]));
 assert.equal(p.records[0].workStart,'');assert.equal(p.records[0].investmentStart,'');assert.match(p.warnings.join(' '),/错误值/);assert.match(p.warnings.join(' '),/没有保存/);
});
test('a missing date preserves every original word and historical or educational years',()=>{
 const biography='2016年毕业，4年本科学习。12年从业经验，其中2年证券从业经验，9年投资经验。2020年加入。';
 assert.deepEqual(call('prepareManagerBiography',{biography},'2026-10-29'),{text:biography,calculations:[]});
 const p=call('prepareManagerBiography',{biography,workStart:'2013-10-30',investmentStart:'2016-10-29'},'2026-10-29');
 assert.equal(p.text,'2016年毕业，4年本科学习。12年从业经验，其中2年证券从业经验，10年投资经验。2020年加入。');
 assert.equal(p.calculations.length,2);assert.equal(p.calculations[0].years,12);assert.equal(p.calculations[1].years,10);
});
test('work and investment experience are recalculated independently on anniversaries',()=>{
 const p=call('prepareManagerBiography',{biography:'10年证券投研工作经验，其中7年中国债券市场投资经验。',workStart:'2016-10-30',investmentStart:'2019-10-29'},'2026-10-29');
 assert.equal(p.text,'9年证券投研工作经验，其中7年中国债券市场投资经验。');
 assert.deepEqual(p.calculations.map(c=>[c.kind,c.startDate,c.productionDate,c.years]),[['work','2016-10-30','2026-10-29',9],['investment','2019-10-29','2026-10-29',7]]);
 const only=call('prepareManagerBiography',{biography:'13年中国债券市场投资经验。',workStart:'2000-01-01',investmentStart:'2014-10-30'},'2026-10-29');
 assert.equal(only.text,'11年中国债券市场投资经验。');assert.equal(only.calculations.length,1);
});
test('explicit start dates replace imprecise numeric tenure without appending absent clauses',()=>{
 const p=call('prepareManagerBiography',{biography:'超22年海外及国内固收领域从业经验，其中近1.5年投资组合管理经验。',workStart:'2000-10-29',investmentStart:'2025-11-01'},'2026-10-29');
 assert.equal(p.text,'26年海外及国内固收领域从业经验，其中0年投资组合管理经验。');
 assert.equal(call('prepareManagerBiography',{biography:'2026年5月起任投资经理。',workStart:'2020-01-01',investmentStart:'2026-05-01'},'2026-10-29').text,'2026年5月起任投资经理。');
 assert.equal(call('prepareManagerBiography',{biography:'10年从业经验。',workStart:'2027-01-01'},'2026-10-29').text,'10年从业经验。');
});
test('historical calendar years leading to experience wording are never treated as durations',()=>{
 const biography='2016年开始积累投资经验。2020年起积累证券从业经验。具有10年从业经验和6年投资经验。';
 const p=call('prepareManagerBiography',{biography,workStart:'2015-01-01',investmentStart:'2019-01-01'},'2026-10-29');
 assert.equal(p.text,'2016年开始积累投资经验。2020年起积累证券从业经验。具有11年从业经验和7年投资经验。');
});
test('invalid or non-manager workbooks produce actionable errors',async()=>{
 await assert.rejects(async()=>call('parseManagerProfile',{name:'经理.xls',bytes:new Uint8Array([1])}),/XLSX/);
 await assert.rejects(async()=>call('parseManagerProfile',{name:'经理.xlsx',bytes:new Uint8Array([1])}),/Excel|XLSX/);
 await assert.rejects(async()=>call('parseManagerProfile',await workbook([{rows:[['姓名','其他'],['甲','仅备注']]}])),/姓名|简介/);
});
test('the downloadable built-in Excel exactly matches the built-in records and hash',async()=>{
 const {DEFAULT_MANAGER_PROFILE,DEFAULT_MANAGER_TEMPLATE}=await import('../src/default-manager-profile.js');
 const bytes=new Uint8Array(await fs.readFile(new URL('../assets/manager-introductions.xlsx',import.meta.url)));
 const parsed=await call('parseManagerProfile',{name:DEFAULT_MANAGER_TEMPLATE.name,bytes});
 assert.deepEqual(parsed,DEFAULT_MANAGER_PROFILE);assert.ok(parsed.records.length>50);assert.ok(parsed.records.filter(r=>r.name==='刘博洋').length>1);
 const zip=await JSZip.loadAsync(bytes);assert.ok(!zip.file('docProps/custom.xml'));
 assert.doesNotMatch(await zip.file('xl/workbook.xml').async('string'),/Users\\Administrator|absPath/);
});
