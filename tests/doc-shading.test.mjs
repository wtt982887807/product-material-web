import test from 'node:test';
import assert from 'node:assert/strict';
import {normalizeLegacyDocShading} from '../src/doc-shading.js';

// Synthetic MS-CFB / MS-DOC structures only. No real document content or colors.
const END = 0xFFFFFFFE, FREE = 0xFFFFFFFF, FAT = 0xFFFFFFFD;
const SECTOR = 512;
const signature = Uint8Array.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
const u16 = n => [n & 255, n >>> 8 & 255];
const u32 = n => [n & 255, n >>> 8 & 255, n >>> 16 & 255, n >>> 24 & 255];
const concat = (...parts) => Uint8Array.from(parts.flatMap(part => [...part]));

function shade(rgb, {nil = false, pattern = 0} = {}) {
  if (nil) return Uint8Array.from([...Array(8).fill(255), 0, 0]);
  return Uint8Array.from([0, 0, 0, 255, ...rgb, 0, ...u16(pattern)]);
}

function prl(opcode, payload) {
  return concat(u16(opcode), opcode === 0xD608 ? u16(payload.length + 1) : [payload.length], payload);
}

function tableDefinition(columns = 2) {
  // TC80 entries may be omitted: the specified columns then use default TC80.
  const edges = Array.from({length: columns + 1}, (_, i) => u16(i * 100)).flat();
  return prl(0xD608, [columns, ...edges]);
}

function rowGroup(shading, columns = 2) {
  return concat([0x16, 0x24, 1, 0x17, 0x24, 1, 0x49, 0x66, 1, 0, 0, 0], tableDefinition(columns), shading);
}

function fixture({group = rowGroup(prl(0xD670, shade([0x27, 0x49, 0x83]))), huge = false, dataGroup = group, hugeOpcode = 0x6646} = {}) {
  // Directory sector 0, FAT sector 1; three ordinary 4096-byte streams.
  const bytes = new Uint8Array(27 * SECTOR);
  const dv = new DataView(bytes.buffer);
  const put16 = (at, n) => dv.setUint16(at, n, true);
  const put32 = (at, n) => dv.setUint32(at, n, true);
  const absolute = (sector, offset = 0) => (sector + 1) * SECTOR + offset;
  bytes.set(signature);
  put16(24, 0x003E); put16(26, 3); put16(28, 0xFFFE);
  put16(30, 9); put16(32, 6);
  put32(44, 1); put32(48, 0); put32(56, 4096);
  put32(60, END); put32(64, 0); put32(68, END); put32(72, 0);
  for (let i = 0; i < 109; i++) put32(76 + i * 4, i ? FREE : 1);
  const fatAt = absolute(1);
  for (let i = 0; i < 128; i++) put32(fatAt + i * 4, FREE);
  put32(fatAt, END); put32(fatAt + 4, FAT);
  for (const start of [2, 10, 18]) {
    for (let i = 0; i < 8; i++) put32(fatAt + (start + i) * 4, i === 7 ? END : start + i + 1);
  }
  function entry(index, name, type, start, size, {left = FREE, right = FREE, child = FREE} = {}) {
    const at = absolute(0, index * 128);
    for (let i = 0; i < name.length; i++) put16(at + i * 2, name.charCodeAt(i));
    put16(at + 64, (name.length + 1) * 2);
    bytes[at + 66] = type; bytes[at + 67] = 1;
    put32(at + 68, left); put32(at + 72, right); put32(at + 76, child);
    put32(at + 116, start); put32(at + 120, size);
  }
  entry(0, 'Root Entry', 5, END, 0, {child: 2});
  entry(1, 'WordDocument', 2, 2, 4096);
  entry(2, '1Table', 2, 10, 4096, {left: 3, right: 1});
  entry(3, 'Data', 2, 18, 4096);
  const wordAt = absolute(2), tableAt = absolute(10), dataAt = absolute(18);
  put16(wordAt, 0xA5EC); put16(wordAt + 2, 0x010C);
  put16(wordAt + 10, 0x0200); // fWhichTblStm -> 1Table.
  put32(wordAt + 24, 1024); put32(wordAt + 28, 1025);
  put16(wordAt + 32, 14);
  put16(wordAt + 62, 22);
  put32(wordAt + 64 + 3 * 4, 1);
  put16(wordAt + 152, 93);
  const fibPairs = wordAt + 154;
  put32(fibPairs + 13 * 8, 128); put32(fibPairs + 13 * 8 + 4, 12);
  put32(tableAt + 128, 1024); put32(tableAt + 132, 1025); put32(tableAt + 136, 4);
  const fkpAt = wordAt + 2048, papAt = fkpAt + 128;
  put32(fkpAt, 1024); put32(fkpAt + 4, 1025);
  bytes[fkpAt + 8] = 64; bytes[fkpAt + 511] = 1;
  const grp = huge ? concat(u16(hugeOpcode), u32(128)) : group;
  // Both PapxInFkp size encodings are exercised as required by parity.
  const istdAndGroup = concat([0, 0], grp);
  let groupAt;
  if (istdAndGroup.length % 2) {
    bytes[papAt] = (istdAndGroup.length + 1) / 2;
    bytes.set(istdAndGroup, papAt + 1); groupAt = papAt + 3;
  } else {
    bytes[papAt] = 0; bytes[papAt + 1] = istdAndGroup.length / 2;
    bytes.set(istdAndGroup, papAt + 2); groupAt = papAt + 4;
  }
  assert.ok(groupAt + grp.length < fkpAt + 511, 'test property group must fit its FKP');
  if (huge) { put16(dataAt + 128, dataGroup.length); bytes.set(dataGroup, dataAt + 130); }
  return {bytes, dv, wordAt, tableAt, dataAt, fatAt, fkpAt, groupAt: huge ? dataAt + 130 : groupAt, group};
}

function changedBytes(before, after) {
  assert.equal(after.length, before.length);
  return Array.from(before.keys()).filter(i => before[i] !== after[i]);
}

function useMiniStreams(f) {
  // Put Table (3 mini sectors) and Data (8 mini sectors) into a two-sector root
  // stream. The old Data sectors become unreferenced bytes, which must stay put.
  const table = f.bytes.slice(f.tableAt, f.tableAt + 192), data = f.bytes.slice(f.dataAt, f.dataAt + 512);
  const rootAt = 11 * SECTOR, miniFatAt = 12 * SECTOR;
  const put32 = (at, n) => f.dv.setUint32(at, n, true);
  for (let sid = 10; sid <= 17; sid++) put32(f.fatAt + sid * 4, FREE);
  put32(f.fatAt + 10 * 4, 12); put32(f.fatAt + 12 * 4, END); put32(f.fatAt + 11 * 4, END);
  put32(SECTOR + 116, 10); put32(SECTOR + 120, 704);
  put32(SECTOR + 2 * 128 + 116, 0); put32(SECTOR + 2 * 128 + 120, 192);
  put32(SECTOR + 3 * 128 + 116, 3); put32(SECTOR + 3 * 128 + 120, 512);
  put32(60, 11); put32(64, 1);
  const root = concat(table, data);
  f.bytes.set(root.subarray(0, 512), rootAt);
  f.bytes.set(root.subarray(512), 13 * SECTOR);
  for (let i = 0; i < 128; i++) put32(miniFatAt + i * 4, FREE);
  for (let i = 0; i < 11; i++) put32(miniFatAt + i * 4, i === 2 || i === 10 ? END : i + 1);
  return {...f, groupAt: rootAt + 192 + 130, miniFatAt};
}

test('bridges an explicit arbitrary RGB shading opcode without mutating source or other bytes', () => {
  const f = fixture(), original = f.bytes.slice();
  const output = normalizeLegacyDocShading(f.bytes);
  assert.ok(output instanceof Uint8Array);
  assert.deepEqual(f.bytes, original);
  const offset = f.groupAt + f.group.length - 13;
  assert.deepEqual(changedBytes(original, output), [offset]);
  assert.deepEqual([...output.slice(offset, offset + 3)], [0x12, 0xD6, 10]);
  assert.deepEqual(output.slice(offset + 3, offset + 13), shade([0x27, 0x49, 0x83]));
});

test('maps all three column groups in a Data-stream HugePapx', () => {
  const raw = concat(prl(0xD670, shade([0x11, 0x32, 0x54])), prl(0xD671, shade([0x65, 0x43, 0x21])), prl(0xD672, shade([0x81, 0x29, 0x47])));
  const f = fixture({group: rowGroup(raw, 45), huge: true});
  const output = normalizeLegacyDocShading(f.bytes);
  const first = f.groupAt + f.group.length - raw.length;
  assert.deepEqual(changedBytes(f.bytes, output), [first, first + 13, first + 26]);
  assert.deepEqual([output[first], output[first + 13], output[first + 26]], [0x12, 0x16, 0x0C]);
});

test('preserves a shading property reached through mini-stream Table and Data', () => {
  const f = useMiniStreams(fixture({huge: true}));
  const original = f.bytes.slice(), output = normalizeLegacyDocShading(f.bytes);
  const offset = f.groupAt + f.group.length - 13;
  assert.deepEqual(changedBytes(original, output), [offset]);
  assert.equal(output[offset], 0x12);
  assert.deepEqual(f.bytes, original);
});

test('handles legacy HugePapx and PTableProps pointer forms', () => {
  for (const hugeOpcode of [0x6645, 0x646B]) {
    const f = fixture({huge: true, hugeOpcode});
    assert.equal(changedBytes(f.bytes, normalizeLegacyDocShading(f.bytes)).length, 1);
  }
});

test('uses property boundaries instead of replacing matching bytes elsewhere', () => {
  const f = fixture(), decoy = prl(0xD670, shade([0x33, 0x55, 0x77]));
  for (const at of [f.wordAt + 1500, f.tableAt + 900, f.dataAt + 3000]) f.bytes.set(decoy, at);
  const original = f.bytes.slice(), output = normalizeLegacyDocShading(f.bytes);
  assert.equal(changedBytes(original, output).length, 1);
  for (const at of [f.wordAt + 1500, f.tableAt + 900, f.dataAt + 3000]) assert.deepEqual(output.slice(at, at + decoy.length), decoy);
});

test('passes non-CFB DOCX/RTF bytes through unchanged', () => {
  for (const input of [new Uint8Array(), Uint8Array.from([80, 75, 3, 4, 6, 7]), new TextEncoder().encode('{\\rtf1 sample}')]) {
    assert.deepEqual(normalizeLegacyDocShading(input), input);
  }
});

test('preserves a valid DOC that contains no Raw shading', () => {
  const f = fixture({group: rowGroup(prl(0xD612, shade([0x73, 0x51, 0x29])))});
  assert.deepEqual(normalizeLegacyDocShading(f.bytes), f.bytes);
});

test('rejects a malformed Raw shading payload instead of accepting a partial color', () => {
  const f = fixture({group: rowGroup(prl(0xD670, new Uint8Array(9)))});
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects a property whose operand extends beyond the declared group', () => {
  const f = fixture();
  f.bytes[f.groupAt + f.group.length - 11] = 220;
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects a cyclic FAT chain', () => {
  const f = fixture();
  f.dv.setUint32(f.fatAt + 2 * 4, 2, true);
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects a cyclic mini-FAT chain', () => {
  const f = useMiniStreams(fixture({huge: true}));
  f.dv.setUint32(f.miniFatAt, 0, true);
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects a cyclic HugePapx property chain', () => {
  const cycle = concat(u16(0x6646), u32(128), [0x16, 0x24, 1, 0x17, 0x24, 1]);
  const f = fixture({huge: true, dataGroup: cycle});
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects an out-of-range HugePapx pointer', () => {
  const f = fixture({huge: true});
  const inlineGroup = f.fkpAt + 128 + 4;
  f.dv.setUint32(inlineGroup + 2, 4095, true);
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects Raw ShdNil inheritance rather than converting it into a flat fill', () => {
  const f = fixture({group: rowGroup(prl(0xD670, shade([], {nil: true})))});
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects ipatNil and unsupported patterned fills rather than flattening them', () => {
  for (const payload of [new Uint8Array(10).fill(255), shade([0x49, 0x71, 0x23], {pattern: 8})]) {
    const f = fixture({group: rowGroup(prl(0xD670, payload))});
    assert.throws(() => normalizeLegacyDocShading(f.bytes));
  }
});

test('rejects a later conflicting legacy shading array', () => {
  const raw = prl(0xD670, shade([0x47, 0x25, 0x63]));
  const old = prl(0xD612, shade([0x17, 0x39, 0x51]));
  const f = fixture({group: rowGroup(concat(raw, old))});
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('rejects Raw automatic shading that could expose a legacy Shd80 fallback', () => {
  const automatic = Uint8Array.from([0, 0, 0, 255, 0, 0, 0, 255, 0, 0]);
  const f = fixture({group: rowGroup(concat(prl(0xD609, [0x40, 0]), prl(0xD670, automatic)))});
  assert.throws(() => normalizeLegacyDocShading(f.bytes));
});

test('preserves an unambiguous Raw automatic shading value', () => {
  const automatic = Uint8Array.from([0, 0, 0, 255, 0, 0, 0, 255, 0, 0]);
  const f = fixture({group: rowGroup(prl(0xD670, automatic))});
  const original = f.bytes.slice(), output = normalizeLegacyDocShading(f.bytes);
  const offset = f.groupAt + f.group.length - 13;
  assert.deepEqual(changedBytes(original, output), [offset]);
  assert.deepEqual(output.slice(offset + 3, offset + 13), automatic);
  assert.deepEqual(f.bytes, original);
});
