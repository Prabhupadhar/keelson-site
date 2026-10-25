// Keelson Pro: ReportDoc -> editable Word (.docx). Contract: docs/PRO-ARCHITECTURE.md section 2-3.
// KeelsonDocx.load()  -> Promise<docx>  injects site/vendor/docx.umd.js once (our own origin), resolves when window.docx exists
// KeelsonDocx.build(doc, {brand:{color, logo}, footer, rasterize}) -> Promise<Blob (browser) | Uint8Array (node)>
//   rasterize(svg, w, h) -> Promise<Uint8Array PNG>; defaults to svg -> canvas -> PNG at 2x in the browser.
// Everything but svg blocks and the logo is real, editable Word text. Placeholder blocks print their exportText, or are left out.
// Control characters (invalid in Word's XML) are dropped from every string.
(function (root) {
  "use strict";
  const script = typeof document !== "undefined" && document.currentScript && document.currentScript.src;
  let loading = null;
  function load(src) {
    if (root.docx) return Promise.resolve(root.docx);
    if (typeof document === "undefined") return Promise.reject(new Error("KeelsonDocx: set globalThis.docx before build() outside the browser"));
    return loading || (loading = new Promise((ok, no) => {
      const s = document.createElement("script");
      s.src = src || new URL("vendor/docx.umd.js", script || document.baseURI).href;
      s.onload = () => root.docx ? ok(root.docx) : (loading = null, no(new Error("Word library did not load")));
      s.onerror = () => { loading = null; s.remove(); no(new Error("Word library could not be loaded")); };
      document.head.appendChild(s);
    }));
  }

  // same palette as render-html.js
  const INK = "1D2939", MUT = "667085", LINE = "EAECF0", SOFT = "F2F4F7", RULE = "D0D5DD";
  const TEXT = { red: "B42318", amber: "93370D", green: "067647", grey: MUT, teal: "0F5E63" };
  const CHIP = { red: "FEE4E2", amber: "FEF0C7", green: "DCFAE6", grey: SOFT, teal: "D5F0F1" };
  const SOLID = { red: ["D92D20", "FFFFFF"], amber: ["F79009", "2B1700"], green: ["067647", "FFFFFF"] };
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const fmtDate = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ""); return m ? `${+m[3]} ${MON[m[2] - 1]} ${m[1]}` : String(s || ""); };
  const CTRL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
  const str = v => String(v == null ? "" : v).replace(CTRL, "");
  const forExport = b => !b || !b.placeholder ? b : b.exportText == null ? null
    : Object.assign({}, b, { placeholder: false }, b.type === "list" ? { items: [].concat(b.exportText) } : { text: b.exportText });
  // brand: reuse the renderer's contrast fix when it is loaded, else a local copy of the same rule
  function brandHex(c) {
    const R = root.KeelsonRender || (typeof require === "function" ? (() => { try { return require("./render-html.js"); } catch (e) { return null; } })() : null);
    if (R) return R.brandColors(c).ink.slice(1).toUpperCase();
    return /^#[0-9a-f]{6}$/i.test(c || "") ? c.slice(1).toUpperCase() : "1A7F86";
  }

  function b64(s) {
    if (typeof atob === "function") { const t = atob(s), u = new Uint8Array(t.length); for (let i = 0; i < t.length; i++) u[i] = t.charCodeAt(i); return u; }
    return new Uint8Array(Buffer.from(s, "base64"));
  }
  // natural size of a png/jpeg/gif, else null
  function dims(u) {
    if (u[0] === 0x89 && u[1] === 0x50) return [(u[16] << 24 | u[17] << 16 | u[18] << 8 | u[19]) >>> 0, (u[20] << 24 | u[21] << 16 | u[22] << 8 | u[23]) >>> 0];
    if (u[0] === 0x47 && u[1] === 0x49) return [u[6] | u[7] << 8, u[8] | u[9] << 8];
    if (u[0] === 0xff && u[1] === 0xd8) for (let i = 2; i + 9 < u.length;) {
      if (u[i] !== 0xff) { i++; continue; }
      const m = u[i + 1], len = u[i + 2] << 8 | u[i + 3];
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return [u[i + 7] << 8 | u[i + 8], u[i + 5] << 8 | u[i + 6]];
      i += 2 + len;
    }
    return null;
  }
  async function rasterizeSvg(svg, w, h) {
    if (!/xmlns=/.test(svg)) svg = svg.replace(/<svg\b/, '<svg xmlns="http://www.w3.org/2000/svg"');
    const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" })), img = new Image();
    try {
      await new Promise((ok, no) => { img.onload = ok; img.onerror = () => no(new Error("chart could not be drawn")); img.src = url; });
      const c = document.createElement("canvas"); c.width = Math.round(w * 2); c.height = Math.round(h * 2);
      const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
      const blob = await new Promise(r => c.toBlob(r, "image/png"));
      return new Uint8Array(await blob.arrayBuffer());
    } finally { URL.revokeObjectURL(url); }
  }
  async function logoImage(logo, rasterize) {
    const m = /^data:image\/(png|jpe?g|gif|svg\+xml);base64,([a-z0-9+/=\s]+)$/i.exec(logo || "");
    if (!m) return null;
    let bytes = b64(m[2].replace(/\s/g, "")), type = m[1].toLowerCase().replace("jpeg", "jpg"), wh;
    if (type === "svg+xml") {
      const svg = new TextDecoder().decode(bytes), vb = /viewBox="[\d.\s-]*?([\d.]+)\s+([\d.]+)"/.exec(svg);
      const W = /\bwidth="([\d.]+)/.exec(svg), H = /\bheight="([\d.]+)/.exec(svg);
      wh = W && H ? [+W[1], +H[1]] : vb ? [+vb[1], +vb[2]] : [300, 100];
      if (!rasterize) return null;
      bytes = await rasterize(svg, wh[0], wh[1]); type = "png";
    } else wh = dims(bytes) || [300, 100];
    const h = 30, w = Math.min(150, Math.round(h * wh[0] / wh[1]));
    return { type, data: bytes, w, h: Math.round(w * wh[1] / wh[0]) };
  }

  async function build(doc, opts) {
    const D = await load(), o = opts || {};
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType, HeadingLevel,
      ImageRun, Header, Footer, LineRuleType, PageOrientation, PageNumber, LevelFormat, VerticalAlign, TableLayoutType } = D;
    const rasterize = o.rasterize || (typeof document !== "undefined" ? rasterizeSvg : null);
    const brand = brandHex(o.brand && o.brand.color), land = doc.orientation === "landscape";
    const MARGIN = 1134, PW = land ? 16838 : 11906, CW = PW - 2 * MARGIN; // twips; 2 cm margins
    const asOf = fmtDate(doc.asOf), rag = str(doc.rag).toLowerCase();
    const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE };
    const run = (text, x) => new TextRun(Object.assign({ text: str(text) }, x));
    // a string with \n\n paragraphs and \n line breaks -> paragraphs
    const paras = (text, p, r) => str(text).split(/\n{2,}/).map(chunk => new Paragraph(Object.assign({
      children: chunk.split("\n").map((ln, i) => run(ln, Object.assign({ break: i ? 1 : 0 }, r)))
    }, p)));
    const label = (t, inCell) => new Paragraph({ keepNext: !inCell, spacing: { before: 60, after: 40 }, children: [run(t, { size: 15, bold: true, color: MUT, allCaps: true, characterSpacing: 16 })] });
    const spacer = () => new Paragraph({ spacing: { before: 0, after: 0, line: 160, lineRule: LineRuleType.AUTO }, children: [] });

    function cellRuns(c) {
      if (c && typeof c === "object") {
        const t = TEXT[c.tone] ? c.tone : null;
        return { runs: [run(c.chip ? ` ${str(c.text)} ` : c.text, { bold: !!(c.bold || c.chip), color: t ? TEXT[t] : INK, size: c.chip ? 16 : 18 })], fill: c.chip && t ? CHIP[t] : null };
      }
      return { runs: [run(c, { size: 18 })], fill: null };
    }
    function widths(cols, total) {
      const fr = cols.map(c => { const m = /^(\d+(?:\.\d+)?)%$/.exec(c.width || ""); return m ? +m[1] / 100 : null; });
      const used = fr.reduce((a, v) => a + (v || 0), 0), free = fr.filter(v => v == null).length;
      return fr.map(v => Math.round(total * (v != null ? v : Math.max(0, 1 - used) / (free || 1))));
    }
    const ALIGN = { right: AlignmentType.RIGHT, center: AlignmentType.CENTER };

    async function blocks(list, width, inCell) {
      const out = [], bs = (list || []).map(forExport).filter(Boolean);
      const kept = bs.filter((b, i) => b.type !== "heading" || (bs[i + 1] && bs[i + 1].type !== "heading"));
      for (const b of kept) out.push(...await block(b, width, inCell));
      return out;
    }
    async function block(b, width, inCell) {
      b = forExport(b);
      if (!b || typeof b !== "object") return [];
      switch (b.type) {
        case "heading": return [new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: !inCell, children: [run(b.text)] })];
        case "text": return [...(b.label ? [label(b.label, inCell)] : []), ...paras(b.text, { spacing: { after: 120 } })];
        case "note": return paras(b.text, { spacing: { before: 80, after: 80 } }, { italics: true, color: MUT, size: 16 });
        case "list": return (b.items || []).map(t => new Paragraph({ numbering: { reference: "kr-bullet", level: 0 }, spacing: { after: 60 }, children: [run(t)] }));
        case "tiles": {
          const items = b.items || [];
          if (!items.length) return [];
          const w = Math.floor(width / items.length), GAP = { style: BorderStyle.SINGLE, size: 36, color: "FFFFFF" };
          return [new Table({
            width: { size: width, type: WidthType.DXA }, columnWidths: items.map(() => w), layout: TableLayoutType.FIXED, borders: noBorders,
            rows: [new TableRow({ cantSplit: true, children: items.map((t, i) => {
              const tn = TEXT[t.tone] ? t.tone : null, solid = tn && SOLID[tn] && (typeof t.solid === "boolean" ? t.solid : /^(red|amber|green)$/i.test(str(t.value))) ? SOLID[tn] : null;
              const fg = solid ? solid[1] : null;
              return new TableCell({
                width: { size: w, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: solid ? solid[0] : SOFT },
                margins: { top: 150, bottom: 150, left: 170, right: 170 },
                borders: { top: tn && !solid ? { style: BorderStyle.SINGLE, size: 24, color: { red: "D92D20", amber: "F79009", green: "067647", grey: "98A2B3", teal: "1A7F86" }[tn] } : GAP, bottom: GAP, left: i ? GAP : NONE, right: i < items.length - 1 ? GAP : NONE },
                children: [
                  new Paragraph({ spacing: { after: 40 }, children: [run(t.label, { size: 14, bold: true, allCaps: true, characterSpacing: 14, color: fg || "5D6679" })] }),
                  new Paragraph({ spacing: { after: 40 }, children: [run(t.value, { size: w < 2100 ? 30 : 34, bold: true, color: fg || (tn && tn !== "grey" ? TEXT[tn] : INK) })] }),
                  ...(t.sub ? [new Paragraph({ spacing: { after: 0 }, children: [run(t.sub, { size: 15, color: fg || "5D6679" })] })] : [])
                ]
              });
            }) })]
          }), spacer()];
        }
        case "table": {
          const cols = b.columns || [], rows = b.rows || [], cw = widths(cols, width), n = cols.length || 1;
          const line = { style: BorderStyle.SINGLE, size: 4, color: LINE };
          const head = new TableRow({ tableHeader: true, cantSplit: true, children: cols.map((c, i) => new TableCell({
            width: { size: cw[i], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: SOFT }, verticalAlign: VerticalAlign.BOTTOM,
            margins: { top: 90, bottom: 90, left: 110, right: 110 }, borders: { bottom: { style: BorderStyle.SINGLE, size: 10, color: INK } },
            children: [new Paragraph({ alignment: ALIGN[c.align], spacing: { after: 0 }, children: [run(c.label, { size: 14, bold: true, allCaps: true, characterSpacing: 12, color: "475467" })] })]
          })) });
          const body = rows.length ? rows.map(r => new TableRow({ cantSplit: true, children: cols.map((c, i) => {
            const { runs, fill } = cellRuns((r || [])[i]);
            return new TableCell({
              width: { size: cw[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, margins: { top: 80, bottom: 80, left: 110, right: 110 }, borders: { bottom: line },
              children: [new Paragraph({ alignment: ALIGN[c.align], spacing: { after: 0 }, children: runs })],
              ...(fill ? { shading: { type: ShadingType.CLEAR, color: "auto", fill } } : {})
            });
          }) })) : [new TableRow({ children: [new TableCell({
            columnSpan: n, width: { size: width, type: WidthType.DXA }, margins: { top: 140, bottom: 140, left: 110, right: 110 }, borders: { bottom: line },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [run(b.empty || "Nothing to report.", { italics: true, color: MUT, size: 18 })] })]
          })] })];
          return [new Table({ width: { size: width, type: WidthType.DXA }, columnWidths: cw, layout: TableLayoutType.FIXED, borders: noBorders, rows: [head, ...body] }), spacer()];
        }
        case "svg": {
          if (!rasterize || !b.svg) return [];
          const w0 = +b.width || 680, h0 = +b.height || 300, px = Math.min(w0, width / 15, (land ? 7600 : 12600) / 15 * w0 / h0), png = await rasterize(b.svg, w0, h0);
          return [new Paragraph({ keepLines: true, spacing: { before: 60, after: 120 }, children: [new ImageRun({
            type: "png", data: png, transformation: { width: Math.round(px), height: Math.round(px * h0 / w0) },
            altText: { name: "chart", title: str(b.alt).slice(0, 120), description: str(b.alt) }
          })] })];
        }
        case "columns": {
          const cs = (b.columns || []).filter(Boolean);
          if (!cs.length) return [];
          // Word paginates side-by-side tables badly (the whole row jumps pages): only text and lists sit side by side
          const flat = c => c.flatMap(x => x && x.type === "columns" ? (x.columns || []).flatMap(flat) : [x]);
          if (cs.some(c => flat(c).some(x => x && /^(table|svg|tiles)$/.test(x.type)))) return blocks(cs.flat(), width, inCell);
          const GUT = 340, w = Math.floor(width / cs.length);
          const cells = [];
          for (let i = 0; i < cs.length; i++) cells.push(new TableCell({
            width: { size: w, type: WidthType.DXA }, borders: noBorders,
            margins: { left: i ? GUT / 2 : 0, right: i < cs.length - 1 ? GUT / 2 : 0, top: 0, bottom: 0 },
            children: [...await blocks(cs[i], w - GUT / 2, true), new Paragraph({ children: [] })]
          }));
          return [new Table({ width: { size: width, type: WidthType.DXA }, columnWidths: cs.map(() => w), layout: TableLayoutType.FIXED, borders: noBorders, rows: [new TableRow({ children: cells })] })];
        }
        default: return [];
      }
    }

    const logo = await logoImage(o.brand && o.brand.logo, rasterize);
    const logoRun = () => new ImageRun({ type: logo.type, data: logo.data, transformation: { width: logo.w, height: logo.h }, altText: { name: "logo", title: "Logo", description: "Logo" } });
    const headTable = first => new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: [Math.round(CW * .7), CW - Math.round(CW * .7)], layout: TableLayoutType.FIXED,
      borders: Object.assign({}, noBorders, first ? {} : { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE } }),
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: Math.round(CW * .7), type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: noBorders, margins: { bottom: 80, left: 0, right: 0, top: 0 },
          children: [new Paragraph({ children: first ? [] : [run(doc.title, { size: 16, bold: true, color: INK }), run(`   ${[doc.name, asOf && "as of " + asOf].filter(Boolean).join(" · ")}`, { size: 16, color: MUT })] })] }),
        new TableCell({ width: { size: CW - Math.round(CW * .7), type: WidthType.DXA }, verticalAlign: VerticalAlign.BOTTOM, borders: noBorders, margins: { bottom: 80, left: 0, right: 0, top: 0 },
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: logo ? [logoRun()] : [] })] })
      ] })]
    });
    const foot = () => new Footer({ children: [new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 6 } },
      tabStops: [{ type: "right", position: CW }],
      children: [run(o.footer ? "Made with Keelson · keelson.co.in" : str(doc.title), { size: 14, color: MUT }), run("\t", { size: 14 }),
        new TextRun({ size: 14, color: MUT, children: ["Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES] })]
    })] });

    const ragRun = SOLID[rag] ? [run("     ", { size: 18 }), run(` ${str(doc.rag).toUpperCase()} `, { size: 18, bold: true, characterSpacing: 20, color: SOLID[rag][1], shading: { type: ShadingType.CLEAR, color: "auto", fill: SOLID[rag][0] } })] : [];
    const top = [
      ...(doc.name ? [new Paragraph({ spacing: { after: 60 }, children: [run(doc.name, { size: 16, bold: true, allCaps: true, characterSpacing: 30, color: brand })] })] : []),
      new Paragraph({ heading: HeadingLevel.TITLE, children: [run(doc.title)] }),
      ...(doc.subtitle ? [new Paragraph({ spacing: { after: 80 }, children: [run(doc.subtitle, { size: 20, color: MUT })] })] : []),
      new Paragraph({ spacing: { before: 60, after: 280 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: brand, space: 8 } },
        children: [...(asOf ? [run("AS OF  ", { size: 15, bold: true, color: MUT, characterSpacing: 16 }), run(asOf, { size: 19, bold: true, color: INK })] : []), ...ragRun] })
    ];

    const d = new Document({
      creator: "Keelson", title: str(doc.title), description: str(doc.name),
      styles: {
        default: {
          document: { run: { font: "Calibri", size: 20, color: INK }, paragraph: { spacing: { after: 100, line: 276, lineRule: LineRuleType.AUTO } } },
          title: { run: { font: "Calibri", size: 46, bold: true, color: INK }, paragraph: { spacing: { before: 0, after: 60, line: 240, lineRule: LineRuleType.AUTO } } },
          heading1: { run: { font: "Calibri", size: 24, bold: true, color: brand }, paragraph: { spacing: { before: 300, after: 100 }, keepNext: true, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: LINE, space: 4 } } } }
        }
      },
      numbering: { config: [{ reference: "kr-bullet", levels: [{ level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 240 } }, run: { color: brand } } }] }] },
      sections: [{
        properties: { titlePage: true, page: { size: { width: 11906, height: 16838, orientation: land ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: 567, footer: 567 } } },
        headers: { first: new Header({ children: [headTable(true)] }), default: new Header({ children: [headTable(false)] }) },
        footers: { first: foot(), default: foot() },
        children: [...top, ...await blocks(doc.blocks, CW)]
      }]
    });
    return typeof document !== "undefined" ? Packer.toBlob(d) : new Uint8Array(await Packer.toBuffer(d));
  }

  // shared with export-pptx.js so both exports clean text, pick colours and read logos the same way
  const util = { str, forExport, fmtDate, brandHex, b64, logoImage, rasterizeSvg, INK, MUT, LINE, SOFT, RULE, TEXT, CHIP, SOLID };
  const api = { load, build, util };
  root.KeelsonDocx = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
