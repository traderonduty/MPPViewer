/*
 * mpp-reader.js — reads Microsoft Project files in the browser, no server needed.
 *
 *   .mpp / .mpt  saved by Project 2010, 2013, 2016, 2019, 2021 and Microsoft 365 ("MPP14" format)
 *   .xml         Microsoft Project XML (any version; use File > Save As > XML in Project)
 *
 * The .mpp decoding is a JavaScript port of the relevant parts of MPXJ
 * (https://www.mpxj.org, Copyright Jon Iles, GNU LGPL 2.1). This file is
 * therefore distributed under the GNU Lesser General Public License v2.1.
 *
 * Usage:  const project = await MPPReader.read(file)   // File or Blob
 *         MPPReader.readBuffer(arrayBuffer, fileName)
 * Returns { project, tasks, resources } (see shape at the bottom of the file).
 */
(function (global) {
'use strict';

/* @@TABLES@@ */

class ReadError extends Error {}

/* ------------------------------------------------------------------ bytes */
const u8 = (b, o) => (o >= 0 && o < b.length ? b[o] : 0);
const getShort = (b, o) => u8(b, o) | (u8(b, o + 1) << 8);                         // unsigned 16
const getInt = (b, o) => (u8(b, o) | (u8(b, o + 1) << 8) | (u8(b, o + 2) << 16) | (u8(b, o + 3) << 24)); // signed 32
function getLong(b, o) {                                                           // signed 64 as BigInt
  let r = 0n;
  for (let i = 7; i >= 0; i--) r = (r << 8n) | BigInt(u8(b, o + i));
  return BigInt.asIntN(64, r);
}
function getDouble(b, o) {
  if (o < 0 || o + 8 > b.length) return 0;
  const v = new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
  return Number.isNaN(v) ? 0 : v;
}
const pad = (b, size) => { const n = new Uint8Array(size); n.set(b.subarray(0, Math.min(b.length, size))); return n; };
function unicodeString(b, o = 0) {
  if (!b || o >= b.length) return '';
  let end = b.length;
  for (let i = o; i < b.length - 1; i += 2) if (b[i] === 0 && b[i + 1] === 0) { end = i; break; }
  return new TextDecoder('utf-16le').decode(b.subarray(o, end - ((end - o) % 2)));
}
function asciiString(b, o = 0) {
  let s = '';
  for (let i = o; i < b.length && b[i] !== 0; i++) s += String.fromCharCode(b[i]);
  return s;
}

/* ------------------------------------------------- compound file (OLE2) */
function readCFB(buffer) {
  const b = new Uint8Array(buffer);
  const sig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  if (b.length < 512 || sig.some((v, i) => b[i] !== v)) return null;
  const dv = new DataView(buffer);
  const u32 = o => dv.getUint32(o, true);
  const ssz = 1 << dv.getUint16(0x1E, true), mssz = 1 << dv.getUint16(0x20, true);
  const nFat = u32(0x2C), dirStart = u32(0x30), cutoff = u32(0x38), mfStart = u32(0x3C), difStart = u32(0x44), nDif = u32(0x48);
  const END = 0xFFFFFFFA;
  const secOff = s => (s + 1) * ssz;
  const fatSecs = [];
  for (let i = 0; i < 109; i++) { const v = u32(0x4C + i * 4); if (v < END) fatSecs.push(v); }
  for (let d = difStart, k = 0; k < nDif && d < END && secOff(d) + ssz <= b.length; k++) {
    for (let j = 0; j < ssz / 4 - 1; j++) { const v = u32(secOff(d) + j * 4); if (v < END) fatSecs.push(v); }
    d = u32(secOff(d) + ssz - 4);
  }
  const fat = new Uint32Array(fatSecs.length * ssz / 4);
  fatSecs.slice(0, Math.max(nFat, 0) || fatSecs.length).forEach((s, i) => {
    const o = secOff(s); if (o + ssz <= b.length) fat.set(new Uint32Array(buffer.slice(o, o + ssz)), i * ssz / 4);
  });
  const chain = (start, table) => {
    const out = []; const seen = new Set();
    for (let s = start; s < END && s < table.length && !seen.has(s); s = table[s]) { seen.add(s); out.push(s); }
    return out;
  };
  const readBig = (start, size) => {
    const secs = chain(start, fat), out = new Uint8Array(size == null ? secs.length * ssz : size);
    let p = 0;
    for (const s of secs) { if (p >= out.length) break; const o = secOff(s); const n = Math.min(ssz, out.length - p); out.set(b.subarray(o, o + n), p); p += n; }
    return out;
  };
  const dirBytes = readBig(dirStart, null);
  const entries = [];
  for (let o = 0; o + 128 <= dirBytes.length; o += 128) {
    const e = dirBytes.subarray(o, o + 128), edv = new DataView(e.buffer, e.byteOffset, 128);
    const nl = edv.getUint16(64, true);
    entries.push({
      name: nl > 2 ? new TextDecoder('utf-16le').decode(e.subarray(0, nl - 2)) : '',
      type: e[66], left: edv.getUint32(68, true), right: edv.getUint32(72, true), child: edv.getUint32(76, true),
      start: edv.getUint32(116, true), size: edv.getUint32(120, true)
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new ReadError("The file's internal structure is damaged.");
  const mini = readBig(root.start, root.size);
  const miniFatBytes = mfStart < END ? readBig(mfStart, null) : new Uint8Array(0);
  const miniFat = new Uint32Array(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.length >> 2);
  const readMini = (start, size) => {
    const out = new Uint8Array(size); let p = 0;
    for (const s of chain(start, miniFat)) { if (p >= size) break; const o = s * mssz, n = Math.min(mssz, size - p); out.set(mini.subarray(o, o + n), p); p += n; }
    return out;
  };
  const node = idx => {
    const e = entries[idx];
    const n = { name: e.name, isDir: e.type === 1 || e.type === 5, children: new Map() };
    if (e.type === 2) n.data = () => (e.size < cutoff ? readMini(e.start, e.size) : readBig(e.start, e.size));
    if (n.isDir && e.child < entries.length) {
      const stack = [e.child], seen = new Set();
      while (stack.length) {
        const c = stack.pop();
        if (c >= entries.length || seen.has(c)) continue; seen.add(c);
        const ce = entries[c];
        if (ce.type) n.children.set(ce.name, c);
        stack.push(ce.left, ce.right);
      }
      n.get = name => (n.children.has(name) ? node(n.children.get(name)) : null);
    }
    return n;
  };
  return node(0);
}

/* ------------------------------------------------------------ structures */
function props14(b) {
  const m = new Map(); if (!b) return m;
  const count = getShort(b, 12); let p = 16, found = 0;
  while (found < count) {
    if (b.length - p < 12) break;
    const size = getInt(b, p), key = getInt(b, p + 4); p += 12;
    if (b.length - p < size || size < 1) break;
    m.set(key, b.subarray(p, p + size)); p += size; found++;
    if (size % 2) p++;
  }
  return m;
}
const P = {
  PASSWORD_FLAG: 893386752, PROTECTION_PASSWORD_HASH: 893386756, ENCRYPTION_CODE: 893386759,
  TASK_FIELD_MAP: 131092, TASK_FIELD_MAP2: 50331668, RESOURCE_FIELD_MAP: 131093, RESOURCE_FIELD_MAP2: 50331669,
  ASSIGNMENT_FIELD_MAP: 131095, ASSIGNMENT_FIELD_MAP2: 50331671,
  DURATION_UNITS: 37748757, MINUTES_PER_DAY: 37748765, MINUTES_PER_WEEK: 37748766, DAYS_PER_MONTH: 37753743,
  CRITICAL_SLACK_LIMIT: 37748756, PROJECT_START_DATE: 37748738, PROJECT_FINISH_DATE: 37748739, STATUS_DATE: 37748805,
  TITLE: 37748744, AUTHOR: 37748748, MANAGER: 37748737, COMPANY: 37748736
};
const pShort = (m, k) => (m.get(k) ? getShort(m.get(k), 0) : 0);
const pInt = (m, k) => (m.get(k) ? getInt(m.get(k), 0) : 0);

const FIXED_MAGIC = 0xFADFADBA | 0;
function fixedMeta(b, sizeProvider) {
  if (getInt(b, 0) !== FIXED_MAGIC) throw new ReadError('Unexpected data block in the file.');
  const itemCount = getInt(b, 8);
  const itemSize = typeof sizeProvider === 'number' ? sizeProvider : sizeProvider(b.length, itemCount);
  const adjusted = Math.max(0, Math.floor((b.length - 16) / itemSize));
  const items = [];
  for (let i = 0; i < adjusted; i++) items.push(b.subarray(16 + i * itemSize, 16 + (i + 1) * itemSize));
  return { itemCount, adjusted, items, get: i => (i >= 0 && i < items.length ? items[i] : null) };
}
function fixedMetaMatching(b, other, sizes) {
  return fixedMeta(b, (fileSize, itemCount) => {
    let itemSize = sizes[0], distance = -Infinity; const available = fileSize - 16;
    for (const t of sizes) {
      if (available % t === 0) {
        if (available / t === other.count) { itemSize = t; break; }
        const d = itemCount * t - available;
        if (d <= 0 && d > distance) { itemSize = t; distance = d; }
      }
    }
    return itemSize;
  });
}
function makeFixed(items, offsets) {
  return {
    items, offsets, count: items.length,
    get: i => (i >= 0 && i < items.length && items[i] ? items[i] : null),
    valid: i => i != null && i >= 0 && i < items.length,
    indexOf: off => offsets.indexOf(off)
  };
}
function fixedData(meta, buf, maxExpected = 0, minSize = 0) {
  const n = meta.adjusted, items = new Array(n).fill(null), offsets = new Array(n).fill(0);
  for (let i = 0; i < n; i++) {
    const off = getInt(meta.get(i), 4);
    if (off < 0 || off > buf.length) continue;
    let size = i + 1 === n ? buf.length - off : getInt(meta.get(i + 1), 4) - off;
    if (size === 0) size = minSize;
    const available = buf.length - off;
    if (size < 0 || size > available) size = maxExpected === 0 ? available : Math.min(maxExpected, available);
    if (maxExpected !== 0 && size > maxExpected) size = maxExpected;
    if (size > 0) { items[i] = buf.subarray(off, off + size); offsets[i] = off; }
  }
  return makeFixed(items, offsets);
}
function fixedDataItemSize(meta, itemSize, buf) {               // FixedData(meta, itemSize, stream)
  const n = meta.adjusted, items = new Array(n).fill(null), offsets = new Array(n).fill(0);
  for (let i = 0; i < n; i++) {
    const off = getInt(meta.get(i), 4);
    if (off < 0 || off > buf.length) continue;
    const available = buf.length - off;
    if (itemSize < 0) itemSize = available; else if (itemSize > available) itemSize = available;
    items[i] = buf.subarray(off, off + itemSize); offsets[i] = off;
  }
  return makeFixed(items, offsets);
}
function fixedDataBlocks(itemSize, buf) {                       // FixedData(itemSize, stream)
  const n = Math.floor(buf.length / itemSize), items = [], offsets = [];
  for (let i = 0; i < n; i++) { items.push(buf.subarray(i * itemSize, (i + 1) * itemSize)); offsets.push(i * itemSize); }
  return makeFixed(items, offsets);
}
function varMeta(b) {
  const table = new Map(), offsets = [];
  const magic = getInt(b, 0);
  if (magic !== 0 && magic !== FIXED_MAGIC) throw new ReadError('Unexpected data block in the file.');
  const count = getInt(b, 8); let p = 24;
  for (let i = 0; i < count; i++) {
    if (b.length - p < 12) break;
    const uid = getInt(b, p), off = getInt(b, p + 4), type = getShort(b, p + 8); p += 12;
    if (!table.has(uid)) table.set(uid, new Map());
    table.get(uid).set(type, off); offsets.push(off);
  }
  offsets.sort((a, c) => a - c);
  return { table, offsets, offset: (uid, type) => (table.has(uid) && type != null ? table.get(uid).get(type) : undefined) };
}
function varData(meta, b) {
  const map = new Map();
  for (const off of meta.offsets) {
    if (off >= b.length || map.has(off)) continue;
    const size = getInt(b, off);
    if (size < 0 || size > b.length - off - 4) continue;
    map.set(off, b.subarray(off + 4, off + 4 + size));
  }
  const bytes = (uid, type) => { const o = meta.offset(uid, type); return o == null ? null : map.get(o) || null; };
  return {
    meta, bytes,
    short: (u, t) => { const v = bytes(u, t); return v && v.length >= 2 ? getShort(v, 0) : 0; },
    int: (u, t) => { const v = bytes(u, t); return v && v.length >= 4 ? getInt(v, 0) : 0; },
    double: (u, t) => { const v = bytes(u, t); return v && v.length >= 8 ? getDouble(v, 0) : 0; },
    unicode: (u, t) => { const v = bytes(u, t); return v ? unicodeString(v, 0) : null; },
    ascii: (u, t) => { const v = bytes(u, t); return v ? asciiString(v, 0) : null; },
    timestamp: (u, t) => { const v = bytes(u, t); return v && v.length >= 4 ? timestamp(v, 0) : null; }
  };
}

/* ------------------------------------------------------ dates, durations */
const EPOCH = Date.UTC(1983, 11, 31);
function timestamp(b, o) {                                       // -> ms since 1970, as naive local time in UTC fields
  const days = getShort(b, o + 2);
  if (days <= 1 || days === 65535) return null;
  let time = getShort(b, o); if (time === 65535) time = 0;
  const secs = time * 6;
  if (days < 100 && secs % 60 !== 0) return null;
  return EPOCH + days * 864e5 + secs * 1000;
}
const iso = ms => { if (ms == null) return null; const s = new Date(ms).toISOString(); return s.slice(17, 19) === '00' ? s.slice(0, 16) : s.slice(0, 19); };
const UNIT_CODES = { 3: 'm', 4: 'em', 5: 'h', 6: 'eh', 7: 'd', 8: 'ed', 9: 'w', 10: 'ew', 11: 'mo', 12: 'emo', 19: '%', 20: 'e%' };
const timeUnits = (v, def) => { const k = v & 0x1F; return k === 21 ? (def || 'd') : UNIT_CODES[k] || 'd'; };
function adjustedDuration(pp, raw, unit) {
  if (raw === -1) return null;
  const per = {
    d: pp.minutesPerDay * 10, ed: 14400, w: pp.minutesPerWeek * 10, ew: 100800,
    mo: pp.minutesPerDay * pp.daysPerMonth * 10, emo: 432000, m: 10, em: 10, h: 600, eh: 600
  }[unit];
  if (per === undefined) return { v: raw, u: unit };
  return { v: per ? raw / per : 0, u: unit };
}
function toMinutes(d, pp) {
  if (!d) return null;
  const f = { m: 1, em: 1, h: 60, eh: 60, d: pp.minutesPerDay, ed: 1440, w: pp.minutesPerWeek, ew: 10080, mo: pp.minutesPerDay * pp.daysPerMonth, emo: 43200 }[d.u];
  return f == null ? null : d.v * f;
}
const durStr = d => (d ? `${Number.isInteger(d.v) ? d.v.toFixed(1) : +d.v.toFixed(6)}${d.u}` : null);

/* ------------------------------------------------------------ field maps */
function parseSizes(s) { const m = new Map(); for (const p of s.split(',')) { const [a, b] = p.split(':'); m.set(+a, +b); } return m; }
const SIZES = { task: parseSizes(TASK_SIZES), rsc: parseSizes(RSC_SIZES), assn: parseSizes(ASSN_SIZES) };
const BASES = { task: 0x0B400000, rsc: 0x0C400000, assn: 0x0F400000 };
const DEFS = { task: TASK_FIELDS, rsc: RSC_FIELDS, assn: ASSN_FIELDS };

function fieldMap(data, kind) {
  const items = new Map(), max = [0, 0];
  if (!data) return { items, max, offset: () => -1, key: () => undefined };
  let idx = 0, last = 0, block = 0;
  while (idx + 28 <= data.length) {
    const blockOffset = getShort(data, idx + 4), typeValue = getInt(data, idx + 12), category = getShort(data, idx + 20);
    const sameClass = (typeValue & 0xFFFF0000) === BASES[kind] && (typeValue & 0x8000) === 0;
    const fid = typeValue & 0xFFFF;
    const def = sameClass ? DEFS[kind][fid] : null;
    const size = sameClass ? SIZES[kind].get(fid) || 0 : 0;
    let loc = null;
    if (category !== 0x0B && category !== 0x64) {
      if (blockOffset !== 65535) {
        loc = 'fixed';
        if (blockOffset < last) ++block;
        last = blockOffset;
        if (blockOffset + size > max[block]) max[block] = blockOffset + size;
      } else if (fid !== 0) loc = 'var';
    }
    if (def && loc) items.set(def[0], { name: def[0], type: def[1], units: def[2], loc, block, offset: blockOffset, key: fid });
    else if (def) items.delete(def[0]);
    idx += 28;
  }
  return { items, max, offset: n => (items.get(n) && items.get(n).loc === 'fixed' ? items.get(n).offset : -1), key: n => (items.get(n) ? items.get(n).key : undefined) };
}

function readField(fm, item, uid, fixed, vd, pp) {
  const unitsOf = (fallback) => {
    if (!item.units || !fm.items.has(item.units)) return fallback;
    return readField(fm, fm.items.get(item.units), uid, fixed, vd, pp) || fallback;
  };
  if (item.loc === 'fixed') {
    const b = fixed[item.block], o = item.offset;
    if (!b || o >= b.length) return null;
    switch (item.type) {
      case 'DATE': return timestamp(b, o);
      case 'INTEGER': return getInt(b, o);
      case 'SHORT': return getShort(b, o);
      case 'DURATION': return adjustedDuration(pp, getInt(b, o), unitsOf(pp.durationUnits));
      case 'TIME_UNITS': return timeUnits(getShort(b, o), pp.durationUnits);
      case 'CONSTRAINT': return getShort(b, o);
      case 'PERCENTAGE': { const v = getShort(b, o); return v >= 0 && v <= 100 ? v : null; }
      case 'CURRENCY': { const a = getDouble(b, o); return (Math.abs(a) < 0.1 ? 0 : a) / 100; }
      case 'WORK': { const a = getDouble(b, o); return (Math.abs(a) < 1000 ? 0 : a) / 60000; }
      case 'BOOLEAN': return getShort(b, o) !== 0;
      default: return null;
    }
  }
  if (item.loc === 'var') {
    const k = item.key;
    const listFlag = () => { const f = vd.short(uid, k); return f === 0x0701 || f === 0x0401; };
    switch (item.type) {
      case 'STRING': return listFlag() ? null : vd.unicode(uid, k);
      case 'DATE': return listFlag() ? null : vd.timestamp(uid, k);
      case 'NOTES': { const s = vd.ascii(uid, k); return s == null ? null : rtfToText(s); }
      case 'INTEGER': return vd.int(uid, k);
      case 'SHORT': case 'PERCENTAGE': return vd.short(uid, k);
      case 'BOOLEAN': return vd.short(uid, k) !== 0;
      case 'TIME_UNITS': return timeUnits(vd.short(uid, k), pp.durationUnits);
      case 'CURRENCY': { const a = vd.double(uid, k); return (Math.abs(a) < 0.1 ? 0 : a) / 100; }
      case 'WORK': { const a = vd.double(uid, k); return (Math.abs(a) < 1000 ? 0 : a) / 60000; }
      case 'DURATION': {
        const b = vd.bytes(uid, k);
        if (!b) return null;
        if (b.length === 512) return null;              // free-text duration
        return b.length >= 4 ? adjustedDuration(pp, getInt(b, 0), unitsOf('h')) : null;
      }
      default: return null;
    }
  }
  return null;
}
function populate(fm, uid, fixed, vd, pp) {
  const o = {};
  for (const item of fm.items.values()) o[item.name] = readField(fm, item, uid, fixed, vd, pp);
  return o;
}

/* --------------------------------------------------------------- RTF text */
function rtfToText(rtf) {
  if (!rtf || !rtf.startsWith('{\\rtf')) return (rtf || '').trim() || null;
  const skipDest = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'header', 'footer', 'xmlnstbl', 'listtable', 'listoverridetable', 'rsidtbl', 'generator', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'mmathPr', 'fldinst']);
  const cp = new TextDecoder('windows-1252');
  let out = '', i = 0, depth = 0, ucSkip = 1, pendingSkip = 0;
  const stack = []; let skipping = false;
  while (i < rtf.length) {
    const c = rtf[i];
    if (c === '{') { stack.push(skipping); depth++; i++; continue; }
    if (c === '}') { skipping = stack.pop() || false; depth--; i++; continue; }
    if (c === '\\') {
      const m = /^\\([a-zA-Z]+)(-?\d+)? ?|^\\'([0-9a-fA-F]{2})|^\\([^a-zA-Z])/.exec(rtf.slice(i, i + 40));
      if (!m) { i++; continue; }
      i += m[0].length;
      if (m[3]) { if (pendingSkip) { pendingSkip--; continue; } if (!skipping) out += cp.decode(new Uint8Array([parseInt(m[3], 16)])); continue; }
      if (m[4]) { if (m[4] === '*') skipping = true; else if (!skipping && '\\{}'.includes(m[4])) out += m[4]; else if (!skipping && m[4] === '~') out += ' '; continue; }
      const word = m[1], arg = m[2];
      if (skipDest.has(word)) { skipping = true; continue; }
      if (skipping) continue;
      if (word === 'par' || word === 'line' || word === 'sect' || word === 'row') out += '\n';
      else if (word === 'tab' || word === 'cell') out += '\t';
      else if (word === 'uc') ucSkip = +arg;
      else if (word === 'u') { let n = +arg; if (n < 0) n += 65536; out += String.fromCharCode(n); pendingSkip = ucSkip; }
      else if (word === 'emdash') out += '—'; else if (word === 'endash') out += '–';
      else if (word === 'bullet') out += '•'; else if (word === 'lquote') out += '‘'; else if (word === 'rquote') out += '’';
      else if (word === 'ldblquote') out += '“'; else if (word === 'rdblquote') out += '”';
      continue;
    }
    if (c === '\r' || c === '\n') { i++; continue; }
    if (pendingSkip) { pendingSkip--; i++; continue; }
    if (!skipping) out += c;
    i++;
  }
  out = out.replace(/[ \t]+\n/g, '\n').trim();
  return out || null;
}

/* ------------------------------------------------- OLE property sets */
function propertySet(b) {
  const out = {};
  if (!b || b.length < 48) return out;
  const dv = new DataView(b.buffer, b.byteOffset, b.length);
  const sec = dv.getUint32(44, true);
  if (sec + 8 > b.length) return out;
  const count = dv.getUint32(sec + 4, true);
  let codepage = 1252;
  const props = [];
  for (let i = 0; i < count && sec + 8 + i * 8 + 8 <= b.length; i++) props.push([dv.getUint32(sec + 8 + i * 8, true), dv.getUint32(sec + 12 + i * 8, true)]);
  const cpProp = props.find(p => p[0] === 1); if (cpProp && sec + cpProp[1] + 6 <= b.length) codepage = dv.getUint16(sec + cpProp[1] + 4, true);
  let dec; try { dec = new TextDecoder(codepage === 65001 ? 'utf-8' : codepage === 1200 ? 'utf-16le' : 'windows-' + codepage); } catch { dec = new TextDecoder('windows-1252'); }
  for (const [id, off] of props) {
    const p = sec + off; if (p + 8 > b.length) continue;
    const type = dv.getUint32(p, true);
    try {
      if (type === 0x1E) { const n = dv.getUint32(p + 4, true); out[id] = dec.decode(b.subarray(p + 8, p + 8 + n)).replace(/\0+$/, ''); }
      else if (type === 0x1F) { const n = dv.getUint32(p + 4, true); out[id] = new TextDecoder('utf-16le').decode(b.subarray(p + 8, p + 8 + n * 2)).replace(/\0+$/, ''); }
      else if (type === 0x40) { const lo = dv.getUint32(p + 4, true), hi = dv.getUint32(p + 8, true); const t = (hi * 4294967296 + lo) / 1e4 - 11644473600000; if (t > 0) out[id] = new Date(t); }
    } catch { /* ignore a malformed property */ }
  }
  return out;
}
const localIso = d => (d ? new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16) : null);

/* ------------------------------------------------------------ MPP14 read */
const TASK_BIT_FLAGS = {
  2010: { meta: [['MILESTONE', 8, 0x20]], meta2: [['ACTIVE', 8, 0x04], ['TASK_MODE', 8, 0x08]] },
  2013: { meta: [['MILESTONE', 10, 0x02]], meta2: [['ACTIVE', 8, 0x40], ['TASK_MODE', 8, 0x80]] },
  2016: { meta: [['MILESTONE', 10, 0x02]], meta2: [['ACTIVE', 8, 0x40], ['TASK_MODE', 8, 0x80]] }
};
const RELATION_TYPES = ['FF', 'FS', 'SF', 'SS'];

function readMPP(root, fileName) {
  const compObj = root.get('\u0001CompObj');
  let appName = '', format = '';
  if (compObj) {
    const c = compObj.data();
    const len1 = getInt(c, 28); appName = new TextDecoder('windows-1252').decode(c.subarray(32, 32 + Math.max(0, len1 - 1)));
    const p2 = 32 + len1, len2 = getInt(c, p2);
    if (len2 > 0) format = new TextDecoder('windows-1252').decode(c.subarray(p2 + 4, p2 + 4 + len2 - 1));
  }
  const vm = /Microsoft.Project.(\d+)\.0/.exec(appName);
  const appVersion = vm ? +vm[1] : 0;
  if (!/^MSProject\.(MPP|MPT|GLOBAL)14$/.test(format)) {
    if (/^MSProject\.(MPP|MPT|GLOBAL)(8|9|12)$/.test(format) || appName === 'Microsoft Project 4.0')
      throw new ReadError('This file was saved in an older Project format (Project 2007 or earlier). Open it in Project and use Save As, choosing "Project" (.mpp) or "XML", then upload that file.');
    throw new ReadError("This doesn't look like a Microsoft Project file.");
  }

  const rootProps = props14(root.get('Props14') && root.get('Props14').data());
  const pwFlag = rootProps.get(P.PASSWORD_FLAG) ? rootProps.get(P.PASSWORD_FLAG)[0] : 0;
  if ((pwFlag & 1) && rootProps.get(P.PROTECTION_PASSWORD_HASH))
    throw new ReadError('This file is password-protected. Remove the password in Project (File > Info > Protect), save, and upload it again.');
  const encCode = rootProps.get(P.ENCRYPTION_CODE) ? rootProps.get(P.ENCRYPTION_CODE)[0] : 0;
  const mask = pwFlag !== 0 ? (encCode === 0 ? 0 : (0xFF - encCode) & 0xFF) : 0;
  const decoded = n => { const d = n.data(); if (!mask) return d; const o = new Uint8Array(d.length); for (let i = 0; i < d.length; i++) o[i] = d[i] ^ mask; return o; };

  const projDir = root.get('   114');
  if (!projDir) throw new ReadError("The file's project data couldn't be found.");
  const need = (dir, name) => { const n = dir && dir.get(name); if (!n) throw new ReadError(`The file is missing its ${name} data.`); return n; };
  const props = props14(decoded(need(projDir, 'Props')));

  let durationUnits = timeUnits(pShort(props, P.DURATION_UNITS)); if (durationUnits === 'd' && (pShort(props, P.DURATION_UNITS) & 0x1F) === 21) durationUnits = 'd';
  const pp = {
    durationUnits,
    minutesPerDay: pInt(props, P.MINUTES_PER_DAY) || 480,
    minutesPerWeek: pInt(props, P.MINUTES_PER_WEEK) || 2400,
    daysPerMonth: pShort(props, P.DAYS_PER_MONTH) || 20,
    criticalSlackLimit: pInt(props, P.CRITICAL_SLACK_LIMIT)
  };
  const pDate = k => (props.get(k) ? timestamp(props.get(k), 0) : null);
  const pStr = k => (props.get(k) ? unicodeString(props.get(k), 0) : null);
  const getFM = (kind, ...keys) => fieldMap(keys.map(k => props.get(k)).find(Boolean), kind);

  /* ---- tasks */
  const fm = getFM('task', P.TASK_FIELD_MAP, P.TASK_FIELD_MAP2);
  const taskDir = need(projDir, 'TBkndTask');
  const tVarMeta = varMeta(need(taskDir, 'VarMeta').data());
  const tVar = varData(tVarMeta, need(taskDir, 'Var2Data').data());
  const tFixedMeta = fixedMeta(need(taskDir, 'FixedMeta').data(), 47);
  const tFixed = fixedData(tFixedMeta, need(taskDir, 'FixedData').data(), fm.max[0]);
  const tFixed2Meta = fixedMetaMatching(need(taskDir, 'Fixed2Meta').data(), tFixed, [92, 93, 94, 95, 96]);
  const tFixed2 = fixedData(tFixed2Meta, need(taskDir, 'Fixed2Data').data());

  // createTaskMap
  const taskMap = new Map();
  const uidOff = fm.offset('UNIQUE_ID');
  for (let loop = tFixedMeta.adjusted - 1; loop > 2; loop--) {
    const data = tFixed.get(loop), data2 = tFixed2.get(loop);
    if (!data || !data2) continue;
    const flags = getInt(tFixedMeta.get(loop), 0);
    if (flags & 0x02) { const u = getShort(data, 0); if (!taskMap.has(u)) taskMap.set(u, null); continue; }
    if (data.length === 16) { const u = getInt(data, 0); if (!taskMap.has(u)) taskMap.set(u, loop); continue; }
    const maxSize = fm.max[0];
    if (maxSize === 0 || Math.trunc(data.length * 100 / maxSize) > 75) {
      const u = getInt(data, uidOff);
      if (!taskMap.has(u) || (tVarMeta.table.get(u) && tVarMeta.table.get(u).size)) {
        if (!taskMap.has(u) || (flags & 0x04) === 0) taskMap.set(u, loop);
      }
    }
  }

  const ver = appVersion <= 14 ? 2010 : appVersion === 15 ? 2013 : 2016;
  const flags = TASK_BIT_FLAGS[ver];
  const tasks = new Map();            // uid -> task (insertion order)
  const byId = new Map();
  const taskOrder = [], nullOrder = new Map();
  let autoWBS = true;
  const addNull = (data) => {
    const id = getInt(data, 4);
    if (nullOrder.has(id)) return;
    const t = { uid: getInt(data, 0), id, isNull: true };
    tasks.set(t.uid, t); byId.set(id, t); nullOrder.set(id, t.uid);
  };
  const idOff = fm.offset('ID');

  for (const uid of [...taskMap.keys()].sort((a, c) => a - c)) {
    const index = taskMap.get(uid);
    if (!tFixed.valid(index)) continue;
    let data = tFixed.get(index);
    if (!data) continue;
    const meta = tFixedMeta.get(index);
    const id = getInt(data, idOff);
    if (data.length === 16) { addNull(data); continue; }
    if (data.length < fm.max[0]) data = pad(data, fm.max[0] + 8);
    const meta2 = tFixed2Meta.get(index), data2 = tFixed2.get(index);
    const existing = byId.get(id);
    if (existing) {
      if (!tVarMeta.table.has(uid)) continue;
      if (existing.NAME == null) { tasks.delete(existing.uid); byId.delete(id); if (existing.isNull) nullOrder.delete(id); }
    }
    const t = populate(fm, uid, [data, data2], tVar, pp);
    for (const [name, off, m] of flags.meta) t[name] = !!(meta && (getInt(meta, off) & m));
    for (const [name, off, m] of flags.meta2) if (meta2) t[name] = !!(getInt(meta2, off) & m);
    t.uid = uid; t.id = id;
    t.manual = !!t.TASK_MODE;
    { const o = fm.offset('ACTUAL_DURATION_UNITS'); t.estimated = o >= 0 && (getShort(data, o) & 0x20) !== 0; }
    if (t.START == null || (t.SCHEDULED_START != null && !t.manual)) t.START = t.SCHEDULED_START;
    if (t.FINISH == null || (t.SCHEDULED_FINISH != null && !t.manual)) t.FINISH = t.SCHEDULED_FINISH;
    t.DURATION = t.SCHEDULED_DURATION;
    if (t.CONSTRAINT_TYPE === 1) {                                   // as late as possible
      if (t.LATE_START != null && (t.START == null || t.START < t.LATE_START)) t.START = t.LATE_START;
      if (t.LATE_FINISH != null && (t.FINISH == null || t.FINISH < t.LATE_FINISH)) t.FINISH = t.LATE_FINISH;
    } else if (t.CONSTRAINT_TYPE === 5 || t.CONSTRAINT_TYPE === 7) {
      if (t.FINISH != null && t.START != null && t.FINISH < t.START) t.FINISH = t.LATE_FINISH;
    }
    if (t.WBS != null) autoWBS = false;
    if (t.manual) t.DURATION = t.MANUAL_DURATION; else t.MANUAL_DURATION = null;
    if (t.NAME == null && (t.START == null || t.FINISH == null || t.CREATED == null)) { addNull(data); continue; }
    tasks.set(uid, t); byId.set(id, t);
    if (!data2 || data2.length < 24) nullOrder.set(id, uid);
    else taskOrder.push([getLong(data2, 16), uid]);
    if (!t.manual || t.ACTUAL_START != null) t.START_TEXT = t.FINISH_TEXT = t.DURATION_TEXT = null;
    t.external = !!t.SUBPROJECT_TASK_ID;
    if (t.BASELINE_START == null) t.BASELINE_START = t.BASELINE_ESTIMATED_START;
    if (t.BASELINE_FINISH == null) t.BASELINE_FINISH = t.BASELINE_ESTIMATED_FINISH;
  }

  // postProcessTasks: rebuild ID order
  taskOrder.sort((a, c) => (a[0] < c[0] ? -1 : a[0] > c[0] ? 1 : 0));
  const inc = (Math.floor(nullOrder.size / 1000) + 1) * 2000;
  const order = new Map();
  let nextID = tasks.has(0) ? 0 : inc;
  for (const [, u] of taskOrder) { order.set(nextID, u); nextID += inc; }
  let insertions = 0; const offsetMap = new Map();
  for (const id of [...nullOrder.keys()].sort((a, c) => a - c)) {
    const base = (id - insertions) * inc; let target = base;
    const prev = offsetMap.get(base); let offset = prev == null ? 0 : prev + 1;
    ++insertions;
    while (order.has(target)) { ++offset; if (offset === inc) throw new ReadError("The task order in this file couldn't be worked out."); target = base - (inc - offset); }
    offsetMap.set(base, offset); order.set(target, nullOrder.get(id));
  }
  nextID = tasks.has(0) ? 0 : 1;
  for (const key of [...order.keys()].sort((a, c) => a - c)) { const t = tasks.get(order.get(key)); if (t) t.id = nextID; nextID++; }

  /* ---- relations */
  const consDir = projDir.get('TBkndCons');
  if (consDir) {
    const cMeta = fixedMeta(need(consDir, 'FixedMeta').data(), 10);
    const cData = fixedDataItemSize(cMeta, 20, decoded(need(consDir, 'FixedData')));
    const p15 = appVersion > 14, unitsOff = p15 ? 18 : 14, durOff = p15 ? 14 : 16;
    for (let i = 0; i < cMeta.itemCount; i++) {
      const m = cMeta.get(i); if (!m || getShort(m, 0) !== 0) continue;
      const idx = cData.indexOf(getInt(m, 4)); if (idx === -1) continue;
      const d = cData.get(idx); if (!d || d.length < 14) continue;
      const a = getInt(d, 4), b = getInt(d, 8);
      if (!a || !b || a === b) continue;
      const t1 = tasks.get(a), t2 = tasks.get(b);
      if (!t1 || !t2 || t1.isNull || t2.isNull) continue;
      const typeVal = getShort(d, 12);
      const lag = adjustedDuration(pp, getInt(d, durOff), timeUnits(getShort(d, unitsOff)));
      (t2.preds = t2.preds || []).push({ uid: a, type: RELATION_TYPES[typeVal] || 'FS', lag });
    }
  }

  /* ---- resources and assignments */
  const resources = new Map();
  try {
    const rfm = getFM('rsc', P.RESOURCE_FIELD_MAP, P.RESOURCE_FIELD_MAP2);
    const rDir = projDir.get('TBkndRsc');
    if (rDir) {
      const rVarMeta = varMeta(need(rDir, 'VarMeta').data());
      const rVar = varData(rVarMeta, need(rDir, 'Var2Data').data());
      const rMeta = fixedMeta(need(rDir, 'FixedMeta').data(), 37);
      const rData = fixedData(rMeta, decoded(need(rDir, 'FixedData')));
      const rUidOff = rfm.offset('UNIQUE_ID'), nameKey = rfm.key('NAME');
      const rMap = new Map();
      for (let i = 0; i < rMeta.adjusted; i++) {
        const d = rData.get(i); if (!d || d.length < rfm.max[0]) continue;
        const u = getShort(d, rUidOff); if (!rMap.has(u)) rMap.set(u, i);
      }
      for (const u of [...rVarMeta.table.keys()].sort((a, c) => a - c)) {
        if (!rMap.has(u)) continue;
        const name = rVar.unicode(u, nameKey);
        resources.set(u, { uid: u, name: name || null });
      }
    }
    const afm = getFM('assn', P.ASSIGNMENT_FIELD_MAP, P.ASSIGNMENT_FIELD_MAP2);
    const aDir = projDir.get('TBkndAssn');
    if (aDir) {
      const aVarMeta = varMeta(need(aDir, 'VarMeta').data());
      const aMeta = fixedMeta(need(aDir, 'FixedMeta').data(), 34);
      const aData = fixedDataBlocks(110, decoded(need(aDir, 'FixedData')));
      const oU = afm.offset('UNIQUE_ID'), oT = afm.offset('TASK_UNIQUE_ID'), oR = afm.offset('RESOURCE_UNIQUE_ID');
      for (let i = 0; i < aMeta.itemCount; i++) {
        const m = aMeta.get(i); if (!m || m[0] !== 0) continue;
        let d = aData.get(aData.indexOf(getInt(m, 4))); if (!d) continue;
        if (d.length < afm.max[0]) d = pad(d, afm.max[0]);
        if (!aVarMeta.table.has(getInt(d, oU))) continue;
        const t = tasks.get(getInt(d, oT)); if (!t || t.isNull) continue;
        const r = resources.get(getInt(d, oR));
        t.res = t.res || new Set();
        if (r && r.name) t.res.add(r.name);
      }
    }
  } catch (e) { if (!(e instanceof ReadError)) console.warn('Resources could not be read', e); }

  /* ---- project properties */
  const si = propertySet(root.get('\u0005SummaryInformation') && root.get('\u0005SummaryInformation').data());
  const dsi = propertySet(root.get('\u0005DocumentSummaryInformation') && root.get('\u0005DocumentSummaryInformation').data());
  const project = {
    title: si[2] || pStr(P.TITLE), author: si[4] || pStr(P.AUTHOR),
    manager: dsi[14] || pStr(P.MANAGER), company: dsi[15] || pStr(P.COMPANY),
    start: iso(pDate(P.PROJECT_START_DATE)), finish: iso(pDate(P.PROJECT_FINISH_DATE)), statusDate: iso(pDate(P.STATUS_DATE)),
    lastSaved: localIso(si[13]), application: appName || null, fileType: 'MPP', fileName
  };

  return finish(project, [...tasks.values()], [...resources.values()], pp, autoWBS);
}

/* ------------------------------------------ structure, slack, output */
function finish(project, list, resources, pp, autoWBS) {
  list.sort((a, c) => a.id - c.id);
  const structured = list.length > 1;
  const rootKids = [];
  let last = null, lastLevel = -1;
  const kids = new Map();
  for (const t of list) {
    let parent = null;
    if (!t.isNull && structured) {
      let level = t.OUTLINE_LEVEL || 0;
      if (last) {
        if (level === lastLevel) { parent = last.parentTask; level = lastLevel; }
        else if (level > lastLevel) parent = last;
        else {
          while (level <= lastLevel) {
            parent = last.parentTask;
            if (!parent) break;
            lastLevel = parent.OUTLINE_LEVEL || 0; last = parent;
          }
        }
      }
      last = t; lastLevel = level;
      const siblings = parent ? (kids.get(parent.uid) || []) : rootKids;
      const childCount = siblings.length + 1;
      const gen = (pv) => (parent == null ? (t.uid === 0 ? '0' : String(rootKids.length + 1)) : (pv === '0' ? String(childCount) : pv + '.' + childCount));
      if (autoWBS || t.WBS == null) t.WBS = gen(parent && parent.WBS);
      let pon = parent && parent.outline; if (pon) { const k = pon.lastIndexOf('.0'); if (k !== -1) pon = pon.slice(0, k); }
      t.outline = gen(pon);
    }
    t.parentTask = parent;
    if (parent) { if (!kids.has(parent.uid)) kids.set(parent.uid, []); kids.get(parent.uid).push(t); }
    else rootKids.push(t);
  }
  for (const t of list) {
    t.summary = kids.has(t.uid) || (!!t.SUBPROJECT_FILE && !t.external);
    // drop links from a task to its own summary (as Project does)
    if (t.preds) t.preds = t.preds.filter(p => { const pt = list.find(x => x.uid === p.uid); return !(t.outline && pt && pt.outline && t.outline.startsWith(pt.outline + '.')); });
  }

  const out = list.map(t => {
    if (t.isNull) return { uid: t.uid, id: t.id, parent: null, wbs: null, outline: null, level: 0, name: null, blank: true, summary: false, milestone: false, critical: false, manual: false, active: true, start: null, finish: null, duration: null, durationDays: null, pct: 0, actualStart: null, actualFinish: null, baselineStart: null, baselineFinish: null, totalSlack: null, work: null, cost: null, notes: null, resources: [], preds: [] };
    const dur = t.DURATION;
    const durMin = toMinutes(dur, pp);
    // total slack = min(start slack, finish slack), or finish slack once started
    let slack = null;
    if (t.START_SLACK && t.FINISH_SLACK && dur) {
      slack = t.ACTUAL_START != null ? t.FINISH_SLACK : (toMinutes(t.START_SLACK, pp) < toMinutes(t.FINISH_SLACK, pp) ? t.START_SLACK : t.FINISH_SLACK);
    } else if (t.FINISH_SLACK && t.ACTUAL_START != null) slack = t.FINISH_SLACK;
    let slackMin = slack ? toMinutes(slack, pp) : null;
    if (slack && dur && slack.u !== dur.u) { const f = toMinutes({ v: 1, u: dur.u }, pp); if (f) slack = { v: slackMin / f, u: dur.u }; }
    const pct = t.PERCENT_COMPLETE || 0;
    const critical = t.ACTUAL_FINISH == null && slackMin != null && slackMin / pp.minutesPerDay <= pp.criticalSlackLimit && pct !== 100 && (!t.manual || (t.DURATION_TEXT == null && t.START_TEXT == null && t.FINISH_TEXT == null));
    return {
      uid: t.uid, id: t.id, parent: t.parentTask && !t.parentTask.isNull ? t.parentTask.uid : null,
      wbs: t.WBS || null, outline: t.outline || null, level: t.OUTLINE_LEVEL || 0,
      name: t.NAME || null, blank: false, summary: t.summary, milestone: !!t.MILESTONE, critical, estimated: !!t.estimated,
      manual: t.manual, active: t.ACTIVE !== false, external: !!t.external,
      start: iso(t.START), finish: iso(t.FINISH),
      duration: durStr(dur), durationDays: durMin == null ? null : durMin / pp.minutesPerDay,
      pct, actualStart: iso(t.ACTUAL_START), actualFinish: iso(t.ACTUAL_FINISH),
      baselineStart: iso(t.BASELINE_START), baselineFinish: iso(t.BASELINE_FINISH),
      totalSlack: durStr(slack), work: t.WORK != null ? `${t.WORK}h` : null, cost: t.COST,
      notes: t.NOTES || null, resources: t.res ? [...t.res].sort() : [],
      preds: (t.preds || []).map(p => ({ uid: p.uid, type: p.type, lag: p.lag && p.lag.v ? durStr(p.lag) : null }))
    };
  });
  return { project, tasks: out, resources: resources.filter(r => r.name).map(r => ({ uid: r.uid, name: r.name })) };
}

/* ---------------------------------------------------- Project XML read */
function readXML(text, fileName) {
  if (typeof DOMParser === 'undefined') throw new ReadError('XML files can only be read in a browser.');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new ReadError("The XML file couldn't be read. It may be damaged.");
  const root = doc.documentElement;
  if (!root || root.localName !== 'Project') throw new ReadError("This XML file isn't a Microsoft Project XML file.");
  const kid = (el, name) => { if (!el) return null; for (const c of el.children) if (c.localName === name) return c; return null; };
  const kids = (el, name) => (el ? [...el.children].filter(c => c.localName === name) : []);
  const txt = (el, name) => { const c = kid(el, name); return c ? c.textContent.trim() : null; };
  const num = (el, name) => { const v = txt(el, name); return v == null || v === '' ? null : +v; };
  const date = (el, name) => { const v = txt(el, name); return v ? v.slice(0, 16) : null; };
  const mpd = num(root, 'MinutesPerDay') || 480;
  const pDur = s => { // PT8H30M0S
    if (!s) return null;
    const m = /^-?P(?:(\d+(?:\.\d+)?)Y)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)D)?(?:T(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(s);
    if (!m) return null;
    const mins = (+(m[3] || 0)) * 1440 + (+(m[4] || 0)) * 60 + (+(m[5] || 0)) + (+(m[6] || 0)) / 60;
    return (s.startsWith('-') ? -1 : 1) * mins;
  };
  const days = mins => (mins == null ? null : mins / mpd);
  const ds = d => (d == null ? null : `${Number.isInteger(+d.toFixed(6)) ? d.toFixed(1) : +d.toFixed(6)}d`);

  const resources = new Map();
  for (const r of kids(kid(root, 'Resources'), 'Resource')) { const n = txt(r, 'Name'); if (n) resources.set(num(r, 'UID'), n); }
  const resByTask = new Map();
  for (const a of kids(kid(root, 'Assignments'), 'Assignment')) {
    const n = resources.get(num(a, 'ResourceUID')); if (!n) continue;
    const u = num(a, 'TaskUID'); if (!resByTask.has(u)) resByTask.set(u, new Set()); resByTask.get(u).add(n);
  }
  const TYPES = { 0: 'FF', 1: 'FS', 2: 'SF', 3: 'SS' };
  const fmtLagUnits = { 3: 'm', 4: 'em', 5: 'h', 6: 'eh', 7: 'd', 8: 'ed', 9: 'w', 10: 'ew', 11: 'mo', 12: 'emo', 19: '%', 20: 'e%' };
  const list = [];
  for (const t of kids(kid(root, 'Tasks'), 'Task')) {
    if (txt(t, 'IsNull') === '1') { list.push({ uid: num(t, 'UID'), id: num(t, 'ID'), parent: null, blank: true, level: 0, resources: [], preds: [], pct: 0, summary: false, milestone: false, critical: false, manual: false, active: true }); continue; }
    const base0 = kids(t, 'Baseline').find(b => txt(b, 'Number') === '0');
    const durMin = pDur(txt(t, 'Duration'));
    const slackTenths = num(t, 'TotalSlack');
    const preds = kids(t, 'PredecessorLink').map(l => {
      const lag = num(l, 'LinkLag') || 0, fmt = num(l, 'LagFormat');
      const unit = fmtLagUnits[fmt] || 'd';
      let lagStr = null;
      if (lag) lagStr = unit === '%' || unit === 'e%' ? `${lag}%` : ds(lag / 10 / mpd);
      return { uid: num(l, 'PredecessorUID'), type: TYPES[num(l, 'Type')] || 'FS', lag: lagStr };
    });
    list.push({
      uid: num(t, 'UID'), id: num(t, 'ID'), wbs: txt(t, 'WBS'), outline: txt(t, 'OutlineNumber'), level: num(t, 'OutlineLevel') || 0,
      name: txt(t, 'Name'), blank: false, summary: txt(t, 'Summary') === '1', milestone: txt(t, 'Milestone') === '1',
      critical: txt(t, 'Critical') === '1', manual: txt(t, 'Manual') === '1', active: txt(t, 'Active') !== '0', estimated: txt(t, 'Estimated') === '1',
      start: date(t, 'Start'), finish: date(t, 'Finish'), duration: ds(days(durMin)), durationDays: days(durMin),
      pct: num(t, 'PercentComplete') || 0, actualStart: date(t, 'ActualStart'), actualFinish: date(t, 'ActualFinish'),
      baselineStart: date(base0, 'Start'), baselineFinish: date(base0, 'Finish'),
      totalSlack: slackTenths == null ? null : ds(slackTenths / 10 / mpd), work: null, cost: num(t, 'Cost'),
      notes: txt(t, 'Notes'), resources: [...(resByTask.get(num(t, 'UID')) || [])].sort(), preds
    });
  }
  list.sort((a, c) => a.id - c.id);
  // parents from outline level
  const stack = [];
  for (const t of list) {
    if (t.blank) { t.parent = null; continue; }
    while (stack.length && stack[stack.length - 1].level >= t.level) stack.pop();
    t.parent = stack.length ? stack[stack.length - 1].uid : null;
    stack.push(t);
  }
  const project = {
    title: txt(root, 'Title') || txt(root, 'Name'), author: txt(root, 'Author'), manager: txt(root, 'Manager'), company: txt(root, 'Company'),
    start: date(root, 'StartDate'), finish: date(root, 'FinishDate'), statusDate: date(root, 'StatusDate'), lastSaved: date(root, 'LastSaved'),
    application: 'Microsoft Project XML', fileType: 'MSPDI', fileName
  };
  return { project, tasks: list, resources: [...resources].map(([uid, name]) => ({ uid, name })) };
}

/* ---------------------------------------------------------------- entry */
function readBuffer(buffer, fileName = '') {
  const b = new Uint8Array(buffer);
  const root = readCFB(buffer);
  if (root) return readMPP(root, fileName);
  let start = 0; if (b[0] === 0xEF && b[1] === 0xBB && b[2] === 0xBF) start = 3;
  const head = new TextDecoder().decode(b.subarray(start, start + 400)).trimStart();
  if (head.startsWith('<')) return readXML(new TextDecoder().decode(b.subarray(start)), fileName);
  if (b[0] === 0x50 && b[1] === 0x4B) throw new ReadError("This is a ZIP-based file, not a Microsoft Project file.");
  throw new ReadError("This doesn't look like a Microsoft Project file. Upload an .mpp file, or an .xml file saved from Project.");
}
async function read(file) {
  const buf = await file.arrayBuffer();
  try { return readBuffer(buf, file.name || ''); }
  catch (e) {
    if (e instanceof ReadError) throw e;
    console.error(e);
    throw new ReadError("The file couldn't be read. It may be damaged, or saved in a format this viewer doesn't support.");
  }
}

global.MPPReader = { read, readBuffer, ReadError };
})(typeof window !== 'undefined' ? window : globalThis);
