import test from 'node:test';
import assert from 'node:assert/strict';
import {getTradingDays,TRADING_CALENDAR_METADATA} from '../src/trading-calendar.js';

test('September opening window uses five exchange days, excluding its weekend',()=>{
  const result=getTradingDays('2026/9/8','2026/9/14');
  assert.deepEqual(result.dates,['2026/9/8','2026/9/9','2026/9/10','2026/9/11','2026/9/14']);
  assert.equal(result.warning,'');
  assert.equal(result.sources[0].year,2026);
  assert.match(result.sources[0].url,/sse\.com\.cn\/disclosure\/announcement/);
  assert.equal(TRADING_CALENDAR_METADATA.market,'CN-A');
});

test('National Day and government make-up weekends remain closed',()=>{
  assert.deepEqual(getTradingDays('2026/9/30','2026/10/12').dates,['2026/9/30','2026/10/8','2026/10/9','2026/10/12']);
  assert.deepEqual(getTradingDays('2026/10/1','2026/10/7').dates,[]);
  assert.deepEqual(getTradingDays('2026/9/18','2026/9/21').dates,['2026/9/18','2026/9/21']);
});

test('all 2026 holiday boundaries match the official reopening dates',()=>{
  for(const [start,end,expected]of [
    ['2026/1/1','2026/1/5',['2026/1/5']],
    ['2026/2/13','2026/2/24',['2026/2/13','2026/2/24']],
    ['2026/4/3','2026/4/7',['2026/4/3','2026/4/7']],
    ['2026/4/30','2026/5/6',['2026/4/30','2026/5/6']],
    ['2026/6/18','2026/6/22',['2026/6/18','2026/6/22']],
    ['2026/9/24','2026/9/28',['2026/9/24','2026/9/28']],
  ])assert.deepEqual(getTradingDays(start,end).dates,expected);
});

test('2025 templates and the supported cross-year boundary use their respective calendars',()=>{
  assert.deepEqual(getTradingDays('2025/1/27','2025/2/5').dates,['2025/1/27','2025/2/5']);
  assert.deepEqual(getTradingDays('2025/9/30','2025/10/13').dates,['2025/9/30','2025/10/9','2025/10/10','2025/10/13']);
  const result=getTradingDays('2025/12/31','2026/1/5');
  assert.deepEqual(result.dates,['2025/12/31','2026/1/5']);
  assert.deepEqual(result.sources.map(s=>s.year),[2025,2026]);
});

test('unknown calendar coverage never guesses weekday dates or returns partial verified output',()=>{
  for(const [start,end]of [['2027/9/8','2027/9/14'],['2026/12/31','2027/1/4'],['2024/9/8','2024/9/14']]){
    const result=getTradingDays(start,end);assert.deepEqual(result.dates,[]);assert.deepEqual(result.sources,[]);assert.match(result.warning,/未推算分段开放日期/);
  }
});

test('explicit date formats and date objects have identical timezone-free results',()=>{
  for(const [start,end]of [['2026-09-08','2026-09-09'],['2026.9.8','2026.9.9'],['2026年9月8日','2026年9月9日'],[{year:2026,month:9,day:8},{year:2026,month:9,day:9}]]){
    assert.deepEqual(getTradingDays(start,end).dates,['2026/9/8','2026/9/9']);
  }
});

test('invalid and reversed dates fail clearly instead of rolling into another month',()=>{
  for(const date of ['2026/2/29','2026/4/31','2026/13/1','9/8','2026/9/8 extra',null])assert.throws(()=>getTradingDays(date,'2026/12/31'),/有效日期/);
  assert.throws(()=>getTradingDays('2026/9/14','2026/9/8'),/结束日期不能早于/);
  assert.deepEqual(getTradingDays('2026/9/12','2026/9/13').dates,[]);
});
