/* Keelson xlsx writer: [{name, rows, ...}] -> .xlsx bytes in the browser, no library. The other half of xlsx-read.js.
   Stored (uncompressed) zip, inline strings, one date format, input cells unlocked under sheet protection (no password:
   it stops accidental edits, not a determined user), dropdown / number / date checks, hidden tabs.
   Cells: string, number, {d: days since 1970} for a date, null for empty. Tests: tests/web/xlsx-write.test.js. */
(function (root) {
  "use strict";
  const enc = new TextEncoder();
  // XML 1.0 can't hold most control characters; a plan cell can (pasted from anywhere)
  const esc = s => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/g, "").replace(/[\ud800-\udbff](?![\udc00-\udfff])|(^|[^\ud800-\udbff])[\udc00-\udfff]/g, "$1")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const colName = i => { let s = ""; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s; return s; };
  // Excel's sheet-name rules: 1-31 chars, none of []:*?/\, not starting or ending with '
  const sheetName = (s, used) => { let n = String(s || "Sheet").replace(/[[\]:*?/\\]/g, " ").replace(/^'+|'+$/g, "").replace(/\s+/g, " ").trim().slice(0, 31) || "Sheet";
    for (let i = 2, b = n; used.has(n.toLowerCase()); i++) n = b.slice(0, 31 - String(i).length - 1) + " " + i; used.add(n.toLowerCase()); return n; };

  // styles: 0 plain, 1 header, 2 date, 3 input text, 4 input date, 5 input number, 6 wrapped text
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="d mmm yyyy"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEFB"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFB4C3EA"/></left><right style="thin"><color rgb="FFB4C3EA"/></right><top style="thin"><color rgb="FFB4C3EA"/></top><bottom style="thin"><color rgb="FFB4C3EA"/></bottom></border></borders>
<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="7"><xf/>
<xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="164" applyNumberFormat="1"><alignment vertical="top"/></xf>
<xf fillId="3" borderId="1" applyFill="1" applyBorder="1" applyProtection="1"><alignment wrapText="1" vertical="top"/><protection locked="0"/></xf>
<xf numFmtId="164" fillId="3" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyProtection="1"><alignment vertical="top"/><protection locked="0"/></xf>
<xf fillId="3" borderId="1" applyFill="1" applyBorder="1" applyProtection="1"><alignment vertical="top"/><protection locked="0"/></xf>
<xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  // sheet: {name, rows: cell[][] (row 0 is the header), widths: number[], input: {col: "text"|"date"|"number"},
  //   checks: [{col, list: string[]} | {col, min, max} | {col, date: true}], protect, hidden}
  function sheetXml(s) {
    const input = s.input || {}, n = Math.max(1, s.rows.length), last = n;
    const cell = (v, r, c) => {
      const ref = colName(c) + (r + 1), kind = r ? input[c] : null;
      const st = r === 0 ? 1 : kind === "date" ? 4 : kind === "number" ? 5 : kind ? 3 : v && typeof v === "object" ? 2 : typeof v === "string" && v.length > 40 ? 6 : 0;
      if (v == null || v === "") return kind || r === 0 ? `<c r="${ref}" s="${st}"/>` : "";
      if (typeof v === "number" && Number.isFinite(v)) return `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
      if (typeof v === "object" && Number.isFinite(v.d)) return `<c r="${ref}" s="${st}"><v>${v.d + 25569}</v></c>`;
      return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    };
    const width = Math.max(0, ...s.rows.map(r => r.length), ...Object.keys(input).map(c => +c + 1));
    const rows = s.rows.map((r, i) => `<row r="${i + 1}">${Array.from({ length: width }, (_, c) => cell(r[c], i, c)).join("")}</row>`).join("");
    const cols = (s.widths || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
    const dv = (s.checks || []).map(k => {
      const sq = `${colName(k.col)}2:${colName(k.col)}${Math.max(2, last)}`, msg = `showErrorMessage="1" errorTitle="Check this value"`;
      if (k.list) return `<dataValidation type="list" allowBlank="1" ${msg} error="Pick one from the list." sqref="${sq}"><formula1>"${esc(k.list.join(","))}"</formula1></dataValidation>`;
      if (k.date) return `<dataValidation type="date" operator="greaterThan" allowBlank="1" ${msg} error="Type a date, for example 14 Nov 2026." sqref="${sq}"><formula1>1</formula1></dataValidation>`;
      return `<dataValidation type="whole" operator="between" allowBlank="1" ${msg} error="Type a whole number from ${k.min} to ${k.max}." sqref="${sq}"><formula1>${k.min}</formula1><formula2>${k.max}</formula2></dataValidation>`;
    }).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0">${s.rows.length > 1 ? `<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>` : ""}</sheetView></sheetViews>`
      + (cols ? `<cols>${cols}</cols>` : "") + `<sheetData>${rows}</sheetData>`
      + (s.protect ? `<sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" sort="0" autoFilter="0"/>` : "")
      + (dv ? `<dataValidations count="${(s.checks || []).length}">${dv}</dataValidations>` : "") + `</worksheet>`;
  }

  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  function zip(files) { // [[name, string]] -> Uint8Array, stored entries with UTF-8 names
    const parts = [], dir = []; let off = 0;
    for (const [name, text] of files) {
      const nb = enc.encode(name), data = enc.encode(text), crc = crc32(data), h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x800, true); h.setUint32(14, crc, true);
      h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nb.length, true);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x800, true); c.setUint32(16, crc, true);
      c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, nb.length, true); c.setUint32(42, off, true);
      parts.push(new Uint8Array(h.buffer), nb, data); dir.push(new Uint8Array(c.buffer), nb); off += 30 + nb.length + data.length;
    }
    const size = dir.reduce((a, b) => a + b.length, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, size, true); e.setUint32(16, off, true);
    const all = [...parts, ...dir, new Uint8Array(e.buffer)], out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let p = 0; for (const b of all) { out.set(b, p); p += b.length; }
    return out;
  }

  function book(sheets) {
    const used = new Set(), names = sheets.map(s => sheetName(s.name, used));
    const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    return zip([
      ["[Content_Types].xml", X + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`],
      ["_rels/.rels", X + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ["xl/workbook.xml", X + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><bookViews><workbookView activeTab="${Math.max(0, sheets.findIndex(s => !s.hidden))}"/></bookViews><sheets>${sheets.map((s, i) => `<sheet name="${esc(names[i])}" sheetId="${i + 1}"${s.hidden ? ' state="hidden"' : ""} r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`],
      ["xl/_rels/workbook.xml.rels", X + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="${R}/styles" Target="styles.xml"/></Relationships>`],
      ["xl/styles.xml", STYLES],
      ...sheets.map((s, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s)]),
    ]);
  }

  const api = { book, MIME: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", _: { crc32, colName, sheetName, esc } };
  root.KeelsonXlsxWrite = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
