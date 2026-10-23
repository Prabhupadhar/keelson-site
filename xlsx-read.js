/* Keelson xlsx reader: .xlsx -> [{name, rows: string[][], hidden}] in the browser, no library.
   Zip central directory + DecompressionStream("deflate-raw"), then a small XML tokenizer (node has no DOMParser,
   so the same code runs in tests/web/xlsx.test.js). Dates come back as ISO text, numbers as plain decimal text. */
(function (root) {
  // MAX_ROWS / MAX_COLS: cells beyond them are ignored (a plan never needs them; reports read the first 1,000 rows)
  const MAX_BYTES = 50 * 1024 * 1024, MAX_CELLS = 2000000, MAX_ROWS = 50000, MAX_COLS = 1024;
  const NOT_XLSX = "This file isn't an .xlsx workbook. If it's an older .xls or password-protected file, save it as .xlsx or CSV and try again.";
  // every error we throw on purpose is marked plain: its message is written for the user; anything else is a bug or a crafted file
  const fail = msg => Object.assign(new Error(msg), { plain: true });

  // ---- zip ----
  function entries(buf) {
    const b = new Uint8Array(buf), dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let eocd = -1;
    const need = (o, n) => { if (!(o >= 0 && o + n <= b.length)) throw fail(NOT_XLSX); };
    for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw fail(NOT_XLSX);
    const n = dv.getUint16(eocd + 10, true), out = new Map();
    let p = dv.getUint32(eocd + 16, true), declared = 0;
    const dec = new TextDecoder();
    for (let i = 0; i < n; i++) {
      need(p, 46);
      if (dv.getUint32(p, true) !== 0x02014b50) throw fail(NOT_XLSX);
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
      need(p + 46, nl + xl + cl);
      const name = dec.decode(b.subarray(p + 46, p + 46 + nl));
      if (csize === 0xffffffff || usize === 0xffffffff) throw fail("This workbook is too large to open here (over 4 GB).");
      declared += usize;
      out.set(name, { method, csize, usize, off });
      p += 46 + nl + xl + cl;
    }
    if (declared > MAX_BYTES) throw fail("This workbook unpacks to more than 50 MB, too large to open here. Save the plan sheet as CSV and try that.");
    return { b, dv, map: out, need };
  }

  async function readEntry(z, name, budget) {
    const e = z.map.get(name);
    if (!e) return null;
    const { b, dv, need } = z, o = e.off;
    need(o, 30);
    if (dv.getUint32(o, true) !== 0x04034b50) throw fail(NOT_XLSX);
    const start = o + 30 + dv.getUint16(o + 26, true) + dv.getUint16(o + 28, true);
    need(start, e.csize);
    const raw = b.subarray(start, start + e.csize);
    if (e.method === 0) { budget.left -= raw.length; if (budget.left < 0) throw bomb(); return new TextDecoder().decode(raw); }
    if (e.method !== 8) throw fail(NOT_XLSX);
    // never trust the declared size: count real bytes as they come out and stop at the limit
    const reader = new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw")).getReader();
    const parts = []; let len = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      len += value.length; budget.left -= value.length;
      if (budget.left < 0) { reader.cancel().catch(() => {}); throw bomb(); }
      parts.push(value);
    }
    const all = new Uint8Array(len); let q = 0;
    for (const v of parts) { all.set(v, q); q += v.length; }
    return new TextDecoder().decode(all);
  }
  const bomb = () => fail("This workbook unpacks to more than 50 MB, too large to open here. Save the plan sheet as CSV and try that.");

  // ---- xml: a callback tokenizer; names lose their namespace prefix ----
  const ENT = Object.assign(Object.create(null), { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" });
  const cp = n => n <= 0x10ffff ? String.fromCodePoint(n) : "\ufffd";
  const unesc = s => s.indexOf("&") < 0 ? s : s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) =>
    e[0] === "#" ? cp(e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e] ?? m);
  const local = n => { const i = n.indexOf(":"); return i < 0 ? n : n.slice(i + 1); };
  function attr(raw, key) { // value of attribute `key` (prefix ignored) in a raw attribute string
    const re = /([\w.:-]+)\s*=\s*("([^"]*)"|'([^']*)')/g; let m;
    while ((m = re.exec(raw))) if (local(m[1]) === key) return unesc(m[3] ?? m[4]);
    return null;
  }
  function walk(xml, open, close, text) {
    let i = 0;
    const n = xml.length;
    while (i < n) {
      const lt = xml.indexOf("<", i);
      if (lt < 0) { text && text(xml.slice(i)); break; }
      if (lt > i && text) text(xml.slice(i, lt));
      const c = xml[lt + 1];
      if (c === "!") {
        if (xml.startsWith("<![CDATA[", lt)) { const e = xml.indexOf("]]>", lt); if (text) text(null, xml.slice(lt + 9, e < 0 ? n : e)); i = e < 0 ? n : e + 3; }
        else if (xml.startsWith("<!--", lt)) { const e = xml.indexOf("-->", lt); i = e < 0 ? n : e + 3; }
        else { const e = xml.indexOf(">", lt); i = e < 0 ? n : e + 1; }
        continue;
      }
      if (c === "?") { const e = xml.indexOf("?>", lt); i = e < 0 ? n : e + 2; continue; }
      let j = lt + 1, q = "";
      for (; j < n; j++) { const ch = xml[j]; if (q) { if (ch === q) q = ""; } else if (ch === '"' || ch === "'") q = ch; else if (ch === ">") break; }
      const body = xml.slice(lt + 1, j);
      i = j + 1;
      if (body[0] === "/") { close(local(body.slice(1).trim())); continue; }
      const self = body.endsWith("/"), inner = self ? body.slice(0, -1) : body;
      const sp = inner.search(/\s/), name = local(sp < 0 ? inner : inner.slice(0, sp)), raw = sp < 0 ? "" : inner.slice(sp);
      open(name, raw);
      if (self) close(name);
    }
  }
  // text() gets (escapedText) or (null, cdataText)
  const plain = (esc, cdata) => esc == null ? cdata : unesc(esc);
  // Excel's _xHHHH_ escapes; control characters (which Word and XML reject) become a space
  const CTRL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
  const xstr = s => s.indexOf("_x") < 0 ? s : s.replace(/_x([0-9a-fA-F]{4})_/g, (m, h) => { const c = String.fromCharCode(parseInt(h, 16)); return CTRL.test(c) ? " " : c; });

  // ---- workbook parts ----
  function rels(xml, base) {
    const out = {};
    if (!xml) return out;
    walk(xml, (n, raw) => {
      if (n !== "Relationship") return;
      let t = attr(raw, "Target") || "";
      if (attr(raw, "TargetMode") === "External") return;
      t = t[0] === "/" ? t.slice(1) : norm(base + t);
      out[attr(raw, "Id")] = { target: t, type: attr(raw, "Type") || "" };
    }, () => {});
    return out;
  }
  const norm = p => { const s = []; for (const x of p.split("/")) { if (x === "..") s.pop(); else if (x && x !== ".") s.push(x); } return s.join("/"); };
  const dir = p => p.slice(0, p.lastIndexOf("/") + 1);

  function sharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    let si = false, t = false, rph = 0, buf = "";
    walk(xml, n => {
      if (n === "si") { si = true; buf = ""; } else if (n === "t") t = true; else if (n === "rPh") rph++;
    }, n => {
      if (n === "si") { out.push(xstr(buf)); si = false; } else if (n === "t") t = false; else if (n === "rPh") rph--;
    }, (e, cd) => { if (si && t && !rph) buf += plain(e, cd); });
    return out;
  }

  // numFmtId of each cell style -> "date" | "datetime" | "time" | null
  const BUILTIN = { 14: "date", 15: "date", 16: "date", 17: "date", 18: "time", 19: "time", 20: "time", 21: "time", 22: "datetime", 45: "time", 46: "time", 47: "time" };
  function fmtKind(code) {
    if (/\[(h+|m+|s+)\]/i.test(code)) return null; // elapsed time such as [h]:mm is a duration, kept as a number
    const s = code.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "").replace(/\\./g, "").replace(/[_*]./g, "");
    if (!/[dmyhs]/i.test(s) || /^general$/i.test(s.trim())) return null;
    const time = /[hs]/i.test(s), date = /[dy]/i.test(s) || (!time && /m/i.test(s));
    return date && time ? "datetime" : date ? "date" : "time";
  }
  function styles(xml) {
    const custom = {}, xfs = [];
    if (!xml) return xfs;
    let inXfs = false;
    walk(xml, (n, raw) => {
      if (n === "numFmt") custom[attr(raw, "numFmtId")] = attr(raw, "formatCode") || "";
      else if (n === "cellXfs") inXfs = true;
      else if (n === "xf" && inXfs) xfs.push(+attr(raw, "numFmtId") || 0);
    }, n => { if (n === "cellXfs") inXfs = false; });
    return xfs.map(id => custom[id] != null ? fmtKind(custom[id]) : BUILTIN[id] || null);
  }

  // ---- values ----
  const DAY = 864e5;
  const pad = (v, n = 2) => String(v).padStart(n, "0");
  function serialText(num, kind, d1904) {
    let whole = Math.floor(num), mins = Math.round((num - whole) * 1440);
    if (mins >= 1440) { whole++; mins -= 1440; }
    const hm = `${pad(Math.floor(mins / 60))}:${pad(mins % 60)}`;
    if (kind === "time" && num < 1) return hm;
    // 1900 system: serial 25569 = 1970-01-01, and Excel counts a 29 Feb 1900 that never was (serials below 60 shift by one)
    const days = d1904 ? whole - 24107 : whole < 60 ? whole - 25568 : whole - 25569;
    const d = new Date(days * DAY), iso = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    return kind !== "date" && mins ? `${iso} ${hm}` : iso;
  }
  const numText = v => { const x = Number(v); return Number.isFinite(x) ? String(+x.toPrecision(15)) : v; };
  // "B12" -> [1, 11]; Excel's limits are 3 letters / column XFD (16384) and row 1048576
  const BAD_REF = "This workbook has a cell outside Excel's grid, so it looks damaged. Open it in Excel and save it again, or export it as CSV.";
  const colIndex = ref => {
    let c = 0, i = 0;
    for (; i < ref.length; i++) { const k = ref.charCodeAt(i); if (k < 65 || k > 90) break; c = c * 26 + k - 64; }
    const r = ref.slice(i);
    if (i > 3 || c > 16384 || (r && !(/^\d{1,7}$/.test(r) && +r >= 1 && +r <= 1048576))) throw fail(BAD_REF);
    return [c - 1, r ? +r - 1 : -1];
  };
  const rowIndex = r => { if (!/^\d{1,7}$/.test(r) || +r < 1 || +r > 1048576) throw fail(BAD_REF); return +r - 1; };

  const tooMany = () => fail("This workbook has more than 2,000,000 cells, too large to open here. Save the plan sheet as CSV and try that.");
  function sheetRows(xml, ss, kinds, d1904, count) {
    const rows = [];
    let row = -1, col = -1, ref = null, type = "", style = 0, inV = false, inIs = false, inT = false, rph = 0, v = "", isv = "", has = false;
    walk(xml, (n, raw) => {
      if (n === "row") { const r = attr(raw, "r"); row = r ? rowIndex(r) : row + 1; col = -1; }
      else if (n === "c") {
        if (++count.n > MAX_CELLS) throw tooMany();
        ref = attr(raw, "r"); type = attr(raw, "t") || "n"; style = +(attr(raw, "s") || 0); v = ""; isv = ""; has = false;
        if (ref) { const [c, r] = colIndex(ref); col = c; if (r >= 0) row = r; } else col++;
      }
      else if (n === "v") { inV = true; has = true; }
      else if (n === "is") { inIs = true; has = true; }
      else if (n === "t") inT = true;
      else if (n === "rPh") rph++;
    }, n => {
      if (n === "v") inV = false;
      else if (n === "is") inIs = false;
      else if (n === "t") inT = false;
      else if (n === "rPh") rph--;
      else if (n === "c") {
        if (!has || row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) return;
        let out;
        if (type === "s") out = ss[+v] ?? "";
        else if (type === "inlineStr") out = xstr(isv);
        else if (type === "str" || type === "e") out = xstr(v);
        else if (type === "b") out = v.trim() === "1" ? "TRUE" : "FALSE";
        else if (type === "d") { const m = /^(\d{4}-\d\d-\d\d)(?:T(\d\d:\d\d))?/.exec(v.trim()); out = !m ? v : m[2] && m[2] !== "00:00" ? `${m[1]} ${m[2]}` : m[1]; }
        else { const t = v.trim(); out = t === "" ? "" : kinds[style] && Number.isFinite(+t) ? serialText(+t, kinds[style], d1904) : numText(t); }
        if (out === "") return;
        const r = rows[row] || (rows[row] = []);
        if (r.length < col) { count.n += col - r.length; if (count.n > MAX_CELLS) throw tooMany(); while (r.length < col) r.push(""); } // backfill counts as cells
        r[col] = out;
      }
    }, (e, cd) => { if (inV) v += plain(e, cd); else if (inIs && inT && !rph) isv += plain(e, cd); });
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  async function sheets(buf) {
    const z = entries(buf), budget = { left: MAX_BYTES }, count = { n: 0 };
    const top = rels(await readEntry(z, "_rels/.rels", budget), "");
    const wbPath = Object.values(top).find(r => /officeDocument$/.test(r.type))?.target || "xl/workbook.xml";
    const wbXml = await readEntry(z, wbPath, budget);
    if (wbXml == null) throw fail(NOT_XLSX);
    const wr = rels(await readEntry(z, dir(wbPath) + "_rels/" + wbPath.slice(dir(wbPath).length) + ".rels", budget), dir(wbPath));
    const part = (suffix, fallback) => Object.values(wr).find(r => r.type.endsWith(suffix))?.target || fallback;
    let d1904 = false; const list = [];
    walk(wbXml, (n, raw) => {
      if (n === "workbookPr") { const d = attr(raw, "date1904"); d1904 = d === "1" || d === "true"; }
      else if (n === "sheet") list.push({ name: attr(raw, "name") || "Sheet", id: attr(raw, "id"), hidden: !!attr(raw, "state") && attr(raw, "state") !== "visible" });
    }, () => {});
    const ss = sharedStrings(await readEntry(z, part("/sharedStrings", dir(wbPath) + "sharedStrings.xml"), budget));
    const kinds = styles(await readEntry(z, part("/styles", dir(wbPath) + "styles.xml"), budget));
    const out = [];
    for (const s of list) {
      const r = wr[s.id];
      if (!r || !/\/worksheet$/.test(r.type)) continue; // chart sheets and dialog sheets hold no table
      const xml = await readEntry(z, r.target, budget);
      if (xml == null) continue;
      out.push({ name: s.name, rows: sheetRows(xml, ss, kinds, d1904, count), hidden: s.hidden });
    }
    return [...out.filter(s => !s.hidden), ...out.filter(s => s.hidden)];
  }

  root.KeelsonXlsx = { sheets, _: { fmtKind, serialText, walk, attr, colIndex } };
  if (typeof module !== "undefined") module.exports = root.KeelsonXlsx;
})(typeof window !== "undefined" ? window : globalThis);
