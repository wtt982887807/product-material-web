// Binary DOC compatibility for LibreOffice's WW8 importer. This is not a DOC
// format normalization: only the disposable input copy sent to LO is changed.
// MS-DOC Table Properties: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b39a6648-501c-4361-8366-4f042f579469
// PapxFkp / PapxInFkp / PrcData describe the only property blocks visited below.
// LO ww8par2.cxx reads TDefTableShd, but omits TDefTableShdRaw.
const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const END = 0xfffffffe, FREE = 0xffffffff;
const RAW = new Map([[0xd670, 0xd612], [0xd671, 0xd616], [0xd672, 0xd60c]]);
const LEGACY = new Set([0xd609, 0xd612, 0xd616, 0xd60c]);
const fail = reason => { throw new Error('旧版 Word 底纹无法可靠保留：' + reason + '。请先在 Word 中另存为 DOCX。'); };
const require = (condition, reason = '文件结构不受支持') => { if (!condition) fail(reason); };

/** Preserve explicit cell fills without rewriting any text, RGB or file offsets. */
export function normalizeLegacyDocShading(bytes) {
  if (!(bytes instanceof Uint8Array) || !SIGNATURE.every((b, i) => bytes[i] === b)) return bytes;
  try {
    const cfb = compoundFile(bytes), word = cfb.stream('WordDocument');
    require(word, '不是受支持的二进制 Word 文档');
    const w = reader(word.bytes);
    require(w.u16(0) === 0xa5ec && [0xc1, 0xd9, 0x101, 0x10c, 0x112].includes(w.u16(2)), 'Word 版本不受支持');
    require(!(w.u16(10) & 0x8100), '文件已加密');
    const table = cfb.stream(w.u16(10) & 0x200 ? '1Table' : '0Table');
    require(table, '缺少格式索引');
    const t = reader(table.bytes);
    let pos = 32;
    pos += 2 + 2 * w.u16(pos); pos += 2 + 4 * w.u16(pos);
    const pairs = w.u16(pos); pos += 2;
    require(pairs >= 14 && pos + pairs * 8 <= w.length, '格式索引越界');
    const fc = w.u32(pos + 13 * 8), lcb = w.u32(pos + 13 * 8 + 4);
    if (!lcb) return bytes;
    require(lcb >= 4 && (lcb - 4) % 8 === 0 && fc + lcb <= t.length, '段落格式索引损坏');
    const count = (lcb - 4) / 8, patches = new Map(), visited = new Set(), visiting = new Set();
    let data;
    const group = (stream, start, end, depth = 0) => {
      require(depth < 32 && start <= end && start >= 0 && end <= stream.bytes.length, '格式属性链越界');
      const key = stream.name + ':' + start + ':' + end;
      require(!visiting.has(key), '格式属性链循环');
      if (visited.has(key)) return;
      visiting.add(key);
      const r = reader(stream.bytes), items = [];
      for (let p = start; p < end;) {
        require(p + 2 <= end, '格式属性被截断');
        const opcode = r.u16(p), spra = opcode >>> 13;
        let size = [1, 1, 2, 4, 2, 2, 0, 3][spra];
        if (spra === 6) {
          require(p + 3 <= end, '格式属性被截断');
          if (opcode === 0xd608) size = 1 + r.u16(p + 2);
          else if (opcode === 0xc615 && r.u8(p + 2) === 255) {
            const removed = r.u8(p + 3), added = r.u8(p + 4 + 4 * removed);
            size = 3 + 4 * removed + 3 * added;
          } else size = 1 + r.u8(p + 2);
        }
        const limit = p + 2 + size;
        require(limit <= end, '格式属性长度越界');
        items.push({opcode, p, limit}); p = limit;
      }
      for (const item of items) {
        const {opcode, p, limit} = item;
        if (RAW.has(opcode)) {
          const size = r.u8(p + 2), max = opcode === 0xd672 ? 190 : 220;
          require(size > 0 && size % 10 === 0 && size <= max, '底纹属性长度无效');
          let hasAuto = false;
          for (let q = p + 3; q < limit; q += 10) {
            const fore = r.u32(q), back = r.u32(q + 4), pattern = r.u16(q + 8);
            require(fore !== 0xffffffff && back !== 0xffffffff && pattern === 0, '包含样式继承或复杂图案底纹');
            require((fore >>> 24 === 0 || fore === 0xff000000) && (back >>> 24 === 0 || back === 0xff000000), '底纹颜色编码不受支持');
            hasAuto ||= back === 0xff000000;
          }
          // LO falls back to Shd80 when the modern value is automatic. Also,
          // a later legacy array would override this newly visible Raw array.
          require(!items.some(other => other.opcode === 0xd609 && hasAuto), '新旧底纹设置存在冲突');
          require(!items.some(other => LEGACY.has(other.opcode) && other.p > p), '新旧底纹优先级不一致');
          patches.set(stream.physical(p), RAW.get(opcode) & 255);
        }
        if (opcode === 0x646b || ((opcode === 0x6645 || opcode === 0x6646) && p === start)) {
          data ||= cfb.stream('Data'); require(data, '缺少扩展格式数据');
          const offset = r.u32(p + 2), dr = reader(data.bytes), length = dr.u16(offset);
          require(length >= 10 && length <= 0x3fa2, '扩展格式长度无效');
          group(data, offset + 2, offset + 2 + length, depth + 1);
          // Inspect the compatibility tail too: LO may choose it in preference
          // to PTableProps. Never search arbitrary data or embedded objects.
        }
      }
      visiting.delete(key); visited.add(key);
    };
    const pages = new Set();
    for (let i = 0; i < count; i++) {
      const base = (t.u32(fc + 4 * (count + 1) + 4 * i) & 0x3fffff) * 512;
      if (pages.has(base)) continue;
      pages.add(base); require(base + 512 <= w.length, '段落格式页越界');
      const paragraphs = w.u8(base + 511), bx = base + 4 * (paragraphs + 1);
      require(paragraphs >= 1 && paragraphs <= 29 && bx + 13 * paragraphs <= base + 511, '段落格式页无效');
      for (let j = 0; j < paragraphs; j++) {
        require(w.u32(base + 4 * j) < w.u32(base + 4 * (j + 1)), '段落位置无效');
        const offset = w.u8(bx + 13 * j); if (!offset) continue;
        const p = base + offset * 2;
        require(p >= bx + 13 * paragraphs && p < base + 511, '段落格式位置无效');
        const cb = w.u8(p), start = p + (cb ? 1 : 2), end = cb ? p + 2 * cb : p + 2 + 2 * w.u8(p + 1);
        require(start + 2 <= end && end <= base + 511, '段落格式长度无效');
        group(word, start + 2, end);
      }
    }
    if (!patches.size) return bytes;
    const copy = bytes.slice();
    for (const [offset, value] of patches) copy[offset] = value;
    return copy;
  } catch (error) {
    if (error.message?.startsWith('旧版 Word 底纹')) throw error;
    fail('格式结构读取失败');
  }
}

function reader(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const check = (p, size) => require(Number.isSafeInteger(p) && p >= 0 && p + size <= bytes.length, '文件结构越界');
  return {length: bytes.length, u8(p) { check(p, 1); return bytes[p]; }, u16(p) { check(p, 2); return dv.getUint16(p, true); }, u32(p) { check(p, 4); return dv.getUint32(p, true); }};
}

function compoundFile(bytes) {
  const r = reader(bytes), version = r.u16(26), shift = r.u16(30), sectorSize = 2 ** shift;
  require(r.u16(28) === 0xfffe && ((version === 3 && shift === 9) || (version === 4 && shift === 12)) && r.u16(32) === 6 && r.u32(56) === 4096, '复合文件版本不受支持');
  const totalSectors = Math.floor(bytes.length / sectorSize) - 1, reserved = new Set(), claimed = new Map();
  const sector = sid => { require(sid < totalSectors, '复合文件扇区越界'); return (sid + 1) * sectorSize; };
  const reserve = sid => { sector(sid); require(!reserved.has(sid), '复合文件索引重复'); reserved.add(sid); };
  const fatSectors = [];
  for (let i = 0; i < 109; i++) { const sid = r.u32(76 + i * 4); if (sid !== FREE) { reserve(sid); fatSectors.push(sid); } }
  let difat = r.u32(68); const difatCount = r.u32(72);
  require(difatCount <= totalSectors, '复合文件索引长度无效');
  for (let i = 0; i < difatCount; i++) {
    reserve(difat); const at = sector(difat);
    for (let j = 0; j < sectorSize / 4 - 1; j++) { const sid = r.u32(at + j * 4); if (sid !== FREE) { reserve(sid); fatSectors.push(sid); } }
    difat = r.u32(at + sectorSize - 4);
  }
  require(fatSectors.length === r.u32(44) && (!difatCount || difat === END), '复合文件索引长度不符');
  const fat = [];
  for (const sid of fatSectors) { const at = sector(sid); for (let j = 0; j < sectorSize / 4; j++) fat.push(r.u32(at + j * 4)); }
  const chain = (start, size, name) => {
    if (size === 0) return {bytes: new Uint8Array(), physical() { fail('空数据链'); }};
    const ids = [], seen = new Set(); let sid = start;
    while (sid !== END) {
      require(ids.length < totalSectors && sid < fat.length && !seen.has(sid) && !reserved.has(sid), '复合文件数据链无效');
      const at = sector(sid); void at;
      require(!claimed.has(sid) || claimed.get(sid) === name, '复合文件数据链重叠'); claimed.set(sid, name);
      seen.add(sid); ids.push(sid); sid = fat[sid];
    }
    const length = size ?? ids.length * sectorSize;
    require(length <= ids.length * sectorSize && (size == null || ids.length === Math.ceil(size / sectorSize)), '复合文件数据长度不符');
    const value = new Uint8Array(length);
    ids.forEach((id, i) => value.set(bytes.subarray(sector(id), sector(id) + Math.min(sectorSize, length - i * sectorSize)), i * sectorSize));
    return {bytes: value, physical(p) { require(p >= 0 && p < length, '格式写入越界'); return sector(ids[Math.floor(p / sectorSize)]) + p % sectorSize; }};
  };
  const directory = chain(r.u32(48), null, '$directory'), d = reader(directory.bytes), entries = [];
  for (let p = 0; p + 128 <= d.length; p += 128) {
    const type = d.u8(p + 66), nameLength = d.u16(p + 64); let name = '';
    if (type) {
      require(nameLength >= 2 && nameLength <= 64 && nameLength % 2 === 0, '复合文件目录名称无效');
      for (let q = p; q < p + nameLength - 2; q += 2) name += String.fromCharCode(d.u16(q));
    }
    const low = d.u32(p + 120), high = version === 4 ? d.u32(p + 124) : 0;
    entries.push({name, type, left: d.u32(p + 68), right: d.u32(p + 72), child: d.u32(p + 76), start: d.u32(p + 116), size: low + high * 0x100000000});
  }
  require(entries[0]?.type === 5, '缺少复合文件根目录');
  const byName = new Map(), treeSeen = new Set();
  const walk = id => {
    if (id === FREE) return;
    require(id < entries.length && !treeSeen.has(id), '复合文件目录循环'); treeSeen.add(id);
    const entry = entries[id]; require(entry.type === 1 || entry.type === 2, '复合文件目录类型无效');
    walk(entry.left);
    if (entry.type === 2) { require(!byName.has(entry.name), '复合文件目录名称重复'); byName.set(entry.name, entry); }
    walk(entry.right);
  };
  walk(entries[0].child);
  let miniRoot, miniFat; const miniClaimed = new Map(), cache = new Map();
  const stream = name => {
    if (cache.has(name)) return cache.get(name);
    const entry = byName.get(name); if (!entry) return null;
    require(entry.size <= bytes.length, '复合文件流大小无效');
    let result;
    if (entry.size >= 4096 || entry.size === 0) result = chain(entry.start, entry.size, name);
    else {
      miniRoot ||= chain(entries[0].start, entries[0].size, '$mini-root');
      miniFat ||= reader(chain(r.u32(60), r.u32(64) * sectorSize, '$mini-fat').bytes);
      const ids = [], seen = new Set(); let sid = entry.start;
      while (sid !== END) {
        require(ids.length < Math.ceil(miniRoot.bytes.length / 64) && sid * 64 + 64 <= miniRoot.bytes.length && !seen.has(sid) && !miniClaimed.has(sid), '复合文件小数据链无效');
        seen.add(sid); miniClaimed.set(sid, name); ids.push(sid); sid = miniFat.u32(sid * 4);
      }
      require(ids.length === Math.ceil(entry.size / 64), '复合文件小数据长度不符');
      const value = new Uint8Array(entry.size);
      ids.forEach((id, i) => value.set(miniRoot.bytes.subarray(id * 64, id * 64 + Math.min(64, entry.size - i * 64)), i * 64));
      result = {bytes: value, physical(p) { require(p >= 0 && p < entry.size, '格式写入越界'); return miniRoot.physical(ids[Math.floor(p / 64)] * 64 + p % 64); }};
    }
    result.name = name; cache.set(name, result); return result;
  };
  return {stream};
}
