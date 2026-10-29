import test from 'node:test';
import assert from 'node:assert/strict';

const api=await import('../src/production-date.js').catch(error=>{if(error.code==='ERR_MODULE_NOT_FOUND')return{};throw error;});
const call=(name,...args)=>{assert.equal(typeof api[name],'function',`${name} is implemented`);return api[name](...args);};

test('production day follows Shanghai even before UTC midnight and can be recomputed after midnight',()=>{
 assert.equal(call('currentProductionDate',new Date('2026-10-28T15:59:59Z')),'2026-10-28');
 assert.equal(call('currentProductionDate',new Date('2026-10-28T16:00:00Z')),'2026-10-29');
 assert.equal(call('resolveProductionDate',{now:new Date('2026-12-31T16:00:00Z')}),'2027-01-01');
});
test('a manual full date wins while incomplete and impossible dates cannot silently fall back',()=>{
 assert.equal(call('resolveProductionDate',{manualValue:'2025-03-08',now:new Date('2026-10-29T00:00:00Z')}),'2025-03-08');
 for(const manualValue of ['2026-10','2026-2-01','2026-02-29','2026-04-31','2026-10-29junk'])assert.throws(()=>call('resolveProductionDate',{manualValue}),/日期/);
 assert.equal(call('productionMonth','2026-10-29'),'202610');
 assert.throws(()=>call('currentProductionDate',new Date('invalid')),/日期|时间/);
});
test('completed years use full anniversaries without rounding or inventing a future start',()=>{
 assert.equal(call('completedYears','2016-10-30','2026-10-29'),9);
 assert.equal(call('completedYears','2016-10-29','2026-10-29'),10);
 assert.equal(call('completedYears','2026-10-29','2026-10-29'),0);
 assert.equal(call('completedYears','2020-02-29','2025-02-28'),4);
 assert.equal(call('completedYears','2020-02-29','2025-03-01'),5);
 assert.throws(()=>call('completedYears','2027-01-01','2026-10-29'),/日期|晚于/);
});
