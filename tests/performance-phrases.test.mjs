import test from 'node:test';
import assert from 'node:assert/strict';
import {parsePerformanceInput} from '../src/supplement-parser.js';

const single={performance_R1:'2.5%',performance_ratio_1:'60%'};
test('a single excess-return fee keeps the hurdle distinct from the payout ratio',()=>{
 for(const text of [
  '超过业绩报酬计提基准2.50%的部分提取60%作为管理人业绩报酬',
  '超过业绩报酬计提基准2.50%的部分，提取比例为60%',
  '超过业绩报酬计提基准2.50%的部分\n提取60%作为管理人业绩报酬',
  '业绩报酬计提基准为2.50%，超过基准的部分按60%的比例计提',
  '年化收益率高于2.50%的部分，收取比例：60%',
  '超过2.50%的收益部分的60%作为管理人业绩报酬',
  '2.5%以上部分的60%作为业绩报酬',
  '年化收益率超过２．５０％的部分，以６０％的比例提取业绩报酬',
  '业绩报酬计提基准为2.50%，计提比例为60%',
  '业绩报酬计提基准：2.50%；超出基准的部分提取60%',
  '年化收益率大于2.5%的部分提取60%',
  '超过2.50%的部分，提取业绩报酬的比例为60%',
 ])assert.deepEqual(parsePerformanceInput(text),single,text);
});

test('natural-language tiered fees tolerate line wrapping and punctuation',()=>{
 const expected={performance_R1:'2%',performance_R2:'3%',performance_ratio_1:'20%',performance_ratio_2:'40%'};
 for(const text of [
  '2.0%＜年化收益率≤3.0%，计提20%；年化收益率＞3.0%，计提40%',
  '2.0%＜年化收益率≤3.0%，\n提取比例为20%\n年化收益率＞3.0%，\n提取比例为40%',
  '2%＜年化收益率≤3%，计提比例为20%。超过3%的部分收取40%作为业绩报酬。',
  '超过2%未超过3%的收益部分提取20%；超过3%的部分提取40%',
  '2%＜年化收益率≤3%，计提20%，年化收益率＞3%，计提40%',
 ])assert.deepEqual(parsePerformanceInput(text),expected,text);
});

test('explicit performance slots remain supported and conflicting values are rejected',()=>{
 assert.deepEqual(parsePerformanceInput('R2=3%；R＞R2，计提40%'),{performance_R2:'3%',performance_ratio_2:'40%'});
 assert.deepEqual(parsePerformanceInput('R2=3%；第2档提取比例为40%'),{performance_R2:'3%',performance_ratio_2:'40%'});
 assert.throws(()=>parsePerformanceInput('R1=2%；R1=3%'),/不同|不一致|冲突/);
 assert.throws(()=>parsePerformanceInput('超过2.5%的部分提取60%；超过2.5%的部分提取40%'),/不同|不一致|冲突/);
});

test('unclassified payout percentages cannot silently become hurdle values',()=>{
 assert.throws(()=>parsePerformanceInput('超过2.50%的部分，待确认60%'),/无法|未识别|比例|明确/);
 assert.throws(()=>parsePerformanceInput('超过2.50%的部分提取160%'),/0至100|0到100/);
 assert.throws(()=>parsePerformanceInput('超过2.50%的部分提取-60%'),/负数/);
 assert.throws(()=>parsePerformanceInput('R1=2.5%；R2=3%；提取60%作为业绩报酬'),/无法|明确|档/);
 assert.throws(()=>parsePerformanceInput('超过2.5%的部分提取60%；比率40%'),/无法|未识别|比例|明确/);
 assert.throws(()=>parsePerformanceInput('超过2.5%的部分提取60%，另有比率40%'),/无法|未识别|比例|明确/);
 assert.throws(()=>parsePerformanceInput('年化收益率不超过2.5%的部分提取60%'),/无法|分段/);
});
