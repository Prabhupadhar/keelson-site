// Keelson Pro: ReportDoc -> editable PowerPoint (.pptx), 16:9. Contract: docs/PRO-ARCHITECTURE.md section 3.
// KeelsonPptx.load()  -> Promise<PptxGenJS>  injects site/vendor/pptxgen.bundle.js once (our own origin), resolves when window.PptxGenJS exists
// KeelsonPptx.build(doc, {brand:{color, logo}, footer, rasterize}) -> Promise<Blob (browser) | Uint8Array (node)>
//   rasterize(svg, w, h) -> Promise<Uint8Array PNG>, as for Word (export-docx.js, whose helpers this reuses via KeelsonDocx.util).
// Slide 1: title, subtitle, as-of date, RAG chip, logo, and the KPI tiles when the report starts with them. The rest flows across
// slides titled by the nearest heading ("(cont.)" when it runs on): text, lists and notes as text boxes, tables as native tables
// split by rows (header repeated), charts as 2x PNGs. Heights are estimated (no text measuring), so the budgets are conservative.
(function (root) {
  "use strict";
  const script = typeof document !== "undefined" && document.currentScript && document.currentScript.src;
  let loading = null;
  function load(src) {
    if (root.PptxGenJS) return Promise.resolve(root.PptxGenJS);
    if (typeof document === "undefined") return Promise.reject(new Error("KeelsonPptx: set globalThis.PptxGenJS before build() outside the browser"));
    return loading || (loading = new Promise((ok, no) => {
      const s = document.createElement("script");
      s.src = src || new URL("vendor/pptxgen.bundle.js", script || document.baseURI).href;
      s.onload = () => root.PptxGenJS ? ok(root.PptxGenJS) : (loading = null, no(new Error("PowerPoint library did not load")));
      s.onerror = () => { loading = null; s.remove(); no(new Error("PowerPoint library could not be loaded")); };
      document.head.appendChild(s);
    }));
  }
  const U = () => (root.KeelsonDocx || (typeof require === "function" ? require("./export-docx.js") : {})).util;

  const W = 13.333, H = 7.5, MX = 0.5, CW = W - 2 * MX, TOP = 1.2, BOTTOM = 6.8, FONT = "Calibri";
  const TOPBAR = { red: "D92D20", amber: "F79009", green: "067647", grey: "98A2B3", teal: "1A7F86" };
  const lh = pt => pt * 1.2 / 72; // line height, inches
  // estimated wrapped lines of text at pt in a box w inches wide (avg glyph ~0.52em; caps wider)
  const lines = (t, w, pt, k = 0.52) => String(t).split("\n").reduce((n, l) => n + Math.max(1, Math.ceil(l.length * pt * k / 72 / Math.max(w, 0.3))), 0);
  function toB64(u) {
    if (typeof Buffer !== "undefined") return Buffer.from(u).toString("base64");
    let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  }

  async function build(doc, opts) {
    const P = await load(), o = opts || {}, u = U();
    const { str, forExport, fmtDate, brandHex, logoImage, rasterizeSvg, INK, MUT, LINE, SOFT, TEXT, CHIP, SOLID } = u;
    const rasterize = o.rasterize || (typeof document !== "undefined" ? rasterizeSvg : null);
    doc = doc || {};
    const brand = brandHex(o.brand && o.brand.color), asOf = str(fmtDate(doc.asOf)), rag = str(doc.rag).toLowerCase();
    const pres = new P();
    pres.layout = "LAYOUT_WIDE"; pres.author = "Keelson"; pres.company = "Keelson"; pres.title = str(doc.title).slice(0, 200); pres.subject = str(doc.name).slice(0, 200);
    const T = (x, extra) => Object.assign({ fontFace: FONT, color: INK, margin: 0, valign: "top", fit: "none" }, x, extra);
    const footText = [doc.title, doc.name, asOf && "as of " + asOf].map(str).filter(Boolean).join(" · ");

    function slide(title) {
      const s = pres.addSlide();
      s.background = { color: "FFFFFF" };
      s.addShape("line", { x: MX, y: 6.95, w: CW, h: 0, line: { color: LINE, width: 0.75 } });
      s.addText(footText, T({ x: MX, y: 7.02, w: o.footer ? 7.3 : 10.8, h: 0.3, fontSize: 9, color: MUT }));
      if (o.footer) s.addText("Made with Keelson · keelson.co.in", T({ x: 7.9, y: 7.02, w: 3.9, h: 0.3, fontSize: 9, color: MUT, align: "right" }));
      s.slideNumber = { x: W - MX - 0.8, y: 7.02, w: 0.8, h: 0.3, fontFace: FONT, fontSize: 9, color: MUT, align: "right", margin: 0 };
      if (title != null) {
        s.addText(str(title), T({ x: MX, y: 0.4, w: CW - (logo ? 2.2 : 0), h: 0.6, fontSize: 24, bold: true, color: brand, valign: "bottom" }));
        s.addShape("rect", { x: MX, y: 1.05, w: 0.9, h: 0.05, fill: { color: brand }, line: { type: "none" } });
        if (logo) s.addImage(Object.assign({}, logo, { y: 0.4 }));
      }
      return s;
    }

    // ---- text-like blocks -> paragraph units [{runs, h}], each small enough for one slide
    const MAXH = BOTTOM - TOP - 0.5;
    function para(text, w, pt, ro, po) {
      const t = str(text).replace(/\r\n?/g, "\n").replace(/\n{2,}/g, "\n").trim();
      if (!t) return [];
      const maxLines = Math.floor(MAXH / lh(pt)), perLine = Math.max(10, Math.floor(w * 72 / (pt * 0.52)));
      // ponytail: a paragraph longer than a slide is cut into line-sized character chunks, not at sentence ends
      const chunks = [];
      let cur = "";
      for (const ln of t.split("\n")) {
        const cand = cur ? cur + "\n" + ln : ln;
        if (lines(cand, w, pt) <= maxLines) { cur = cand; continue; }
        if (cur) chunks.push(cur);
        cur = ln;
        while (lines(cur, w, pt) > maxLines) { const cut = maxLines * perLine; chunks.push(cur.slice(0, cut)); cur = cur.slice(cut); }
      }
      if (cur) chunks.push(cur);
      return chunks.map(c => ({ runs: [{ text: c, options: Object.assign({ fontSize: pt, breakLine: true, paraSpaceAfter: 6 }, ro, po) }], h: lines(c, w - (po && po.bullet ? 0.3 : 0), pt) * lh(pt) + 0.09 }));
    }
    function units(b, w) {
      switch (b.type) {
        case "heading": return [{ runs: [{ text: str(b.text), options: { fontSize: 15, bold: true, color: brand, breakLine: true, paraSpaceAfter: 4 } }], h: lines(str(b.text), w, 15) * lh(15) + 0.1, heading: true }];
        case "text": {
          const us = para(b.text, w, 14);
          if (b.label && str(b.label).trim()) {
            const lab = { text: str(b.label).toUpperCase(), options: { fontSize: 10, bold: true, color: MUT, charSpacing: 1, breakLine: true, paraSpaceAfter: 3 } };
            const lu = { runs: [lab], h: lines(lab.text, w, 10, 0.62) * lh(10) + 0.06 };
            if (us.length && us[0].h + lu.h <= MAXH) { us[0] = { runs: [lab, ...us[0].runs], h: us[0].h + lu.h }; } else us.unshift(lu);
          }
          return us;
        }
        case "note": return para(b.text, w, 11, { italic: true, color: MUT });
        case "list": return (b.items || []).flatMap(t => para(t, w - 0.3, 14, null, { bullet: { indent: 18 }, paraSpaceAfter: 4 }));
        default: return [];
      }
    }
    const textBox = (s, us, x, y, w) => {
      const h = us.reduce((a, v) => a + v.h, 0);
      const runs = us.flatMap(v => v.runs);
      runs[runs.length - 1] = { text: runs[runs.length - 1].text, options: Object.assign({}, runs[runs.length - 1].options, { breakLine: false }) };
      s.addText(runs, T({ x, y, w, h: Math.max(h, 0.3) }));
      return h;
    };

    // ---- flow state: one cursor across slides
    let s = null, y = TOP, used = false, section = str(doc.name || doc.title) || "Report", sectionSlides = 0;
    const room = () => s ? BOTTOM - y : 0;
    function fresh() { s = slide(sectionSlides++ ? `${section} (cont.)` : section); y = TOP; used = false; }
    function need(h) { if (!s || (used && h > room())) fresh(); }

    // KPI tiles, at most 6 a row; returns the height used
    const tilesH = items => Math.ceil(items.length / 6) * 1.65 - 0.2;
    function tiles(sl, items, y0) {
      const per = Math.min(6, items.length), GAP = 0.18, w = (CW - GAP * (per - 1)) / per, h = 1.45;
      items.forEach((t, i) => {
        const x = MX + (i % per) * (w + GAP), y0r = y0 + Math.floor(i / per) * 1.65, tn = TEXT[t.tone] ? t.tone : null;
        const solid = tn && SOLID[tn] && (typeof t.solid === "boolean" ? t.solid : /^(red|amber|green)$/i.test(str(t.value))) ? SOLID[tn] : null, fg = solid ? solid[1] : null;
        sl.addShape("rect", { x, y: y0r, w, h, fill: { color: solid ? solid[0] : SOFT }, line: { type: "none" } });
        if (tn && !solid) sl.addShape("rect", { x, y: y0r, w, h: 0.06, fill: { color: TOPBAR[tn] }, line: { type: "none" } });
        const vpt = w < 1.9 ? 20 : 26;
        sl.addText([
          { text: str(t.label).toUpperCase(), options: { fontSize: 9, bold: true, color: fg || "5D6679", charSpacing: 1, breakLine: true, paraSpaceAfter: 2 } },
          { text: str(t.value), options: { fontSize: vpt, bold: true, color: fg || (tn && tn !== "grey" ? TEXT[tn] : INK), breakLine: !!t.sub, paraSpaceAfter: 2 } },
          ...(t.sub ? [{ text: str(t.sub), options: { fontSize: 9.5, color: fg || "5D6679" } }] : [])
        ], T({ x: x + 0.14, y: y0r + 0.14, w: w - 0.28, h: h - 0.22 }));
      });
      return tilesH(items);
    }

    function widths(cols) {
      const fr = cols.map(c => { const m = /^(\d+(?:\.\d+)?)%$/.exec(str(c && c.width)); return m ? +m[1] / 100 : null; });
      const used = fr.reduce((a, v) => a + (v || 0), 0), free = fr.filter(v => v == null).length;
      const ws = fr.map(v => v != null ? v : Math.max(0, 1 - used) / (free || 1)), sum = ws.reduce((a, v) => a + v, 0) || 1;
      return ws.map(v => CW * (v || 0.02) / sum);
    }
    const ALIGN = { right: "right", center: "center" };
    const BL = { type: "solid", pt: 0.75, color: LINE }, BN = { type: "none" };
    function table(b) {
      const cols = (b.columns || []).filter(c => c && typeof c === "object"), cw = widths(cols), n = cw.length;
      if (!n) return;
      const head = cols.map((c, i) => ({ text: str(c.label).toUpperCase(), options: {
        bold: true, fontSize: 9, color: "475467", fill: { color: SOFT }, align: ALIGN[c.align] || "left", valign: "bottom", charSpacing: 0.5,
        border: [BN, BN, { type: "solid", pt: 1.25, color: INK }, BN] } }));
      const headH = Math.max(...head.map((c, i) => lines(c.text, cw[i] - 0.15, 9, 0.66))) * lh(9) + 0.14;
      const rowsOf = (b.rows || []).map(r => cols.map((c, i) => {
        const v = (r || [])[i], obj = v && typeof v === "object", t = obj ? (TEXT[v.tone] ? v.tone : null) : null;
        return { text: str(obj ? v.text : v), options: Object.assign({ fontSize: 11, color: t ? TEXT[t] : INK, bold: !!(obj && (v.bold || v.chip)), align: ALIGN[c.align] || "left", valign: "middle",
          border: [BN, BN, BL, BN] }, obj && v.chip && t ? { fill: { color: CHIP[t] } } : {}) };
      }));
      const rowH = r => Math.max(...r.map((c, i) => lines(c.text || " ", cw[i] - 0.15, 11))) * lh(11) + 0.12;
      const body = rowsOf.length ? rowsOf.map(r => ({ r, h: rowH(r) })) : [{ r: [{ text: str(b.empty || "Nothing to report."), options: { colspan: n, italic: true, color: MUT, fontSize: 11, align: "center", border: [BN, BN, BL, BN] } }], h: 0.4 }];
      const tableOpts = { x: MX, w: CW, colW: cw, fontFace: FONT, fontSize: 11, color: INK, margin: [4, 6, 4, 6], autoPage: false };
      let i = 0;
      while (i < body.length) {
        need(headH + body[i].h); // a row taller than a slide still goes on (and runs off) its own slide
        let h = headH, j = i;
        while (j < body.length && (h + body[j].h <= room() || j === i)) h += body[j++].h;
        s.addTable([head, ...body.slice(i, j).map(x => x.r)], Object.assign({ y }, tableOpts));
        y += h + 0.2; used = true; i = j;
        if (i < body.length) fresh();
      }
    }

    async function svg(b) {
      if (!rasterize || !b.svg) return;
      const w0 = +b.width > 0 ? +b.width : 680, h0 = +b.height > 0 ? +b.height : 300, ar = w0 / h0;
      const fit = (maxH) => { let w = CW, h = w / ar; if (h > maxH) { h = maxH; w = h * ar; } return { w, h }; };
      const ideal = fit(BOTTOM - TOP);
      if (!s || (used && fit(room() - 0.05).h < Math.min(ideal.h, 2.6))) fresh(); // shrink to fit what is left, down to 2.6in tall
      const { w, h } = fit(room() - 0.05), png = await rasterize(str(b.svg), w0, h0);
      s.addImage({ data: "image/png;base64," + toB64(png), x: MX + (CW - w) / 2, y, w, h, altText: str(b.alt).slice(0, 1000) });
      y += h + 0.2; used = true;
    }

    function flowUnits(us, gap = 0.12) {
      let k = 0;
      while (k < us.length) {
        need(us[k].h);
        let h = 0, j = k;
        while (j < us.length && (h + us[j].h <= room() || j === k)) h += us[j++].h;
        textBox(s, us.slice(k, j), MX, y, CW);
        y += h + gap; used = true; k = j;
        if (k < us.length) fresh();
      }
    }

    const flat = c => (c || []).flatMap(x => x && x.type === "columns" ? (x.columns || []).flatMap(flat) : [x]);
    async function block(b) {
      b = forExport(b);
      if (!b || typeof b !== "object") return;
      switch (b.type) {
        case "heading": {
          section = str(b.text); sectionSlides = 0;
          // stays on this slide as a sub-heading when it and a few lines of what follows fit; else starts a slide named after it
          if (s && used && room() > 1.6) { const us = units(b, CW); textBox(s, us, MX, y + 0.08, CW); y += us[0].h + 0.12; sectionSlides = 1; }
          else s = null;
          return;
        }
        case "text": case "note": case "list": return flowUnits(units(b, CW));
        case "tiles": { const it = b.items || []; if (!it.length) return; need(tilesH(it)); y += tiles(s, it, y) + 0.25; used = true; return; }
        case "table": return table(b);
        case "svg": return svg(b);
        case "columns": {
          const cs = (b.columns || []).filter(Boolean);
          if (!cs.length) return;
          const GUT = 0.4, w = (CW - GUT * (cs.length - 1)) / cs.length;
          const per = cs.map(c => flat(c).map(forExport).filter(Boolean));
          if (per.some(c => c.some(x => /^(table|svg|tiles)$/.test(x.type)))) { for (const x of per.flat()) await block(x); return; }
          const us = per.map(c => c.flatMap(x => units(x, w)));
          const hs = us.map(c => c.reduce((a, v) => a + v.h, 0)), h = Math.max(0, ...hs);
          if (!h) return;
          if (h > BOTTOM - TOP) { for (const x of per.flat()) await block(x); return; } // too tall to sit side by side: one after another
          need(h);
          us.forEach((c, i) => { if (c.length) textBox(s, c, MX + i * (w + GUT), y, w); });
          y += h + 0.2; used = true; return;
        }
        default: return;
      }
    }

    // ---- slide 1
    const lg = await logoImage(o.brand && o.brand.logo, rasterize).catch(() => null);
    const logo = lg ? (() => { const r = lg.w / lg.h, h = Math.min(0.7, 2 / r), w = h * r; return { data: `image/${lg.type === "jpg" ? "jpeg" : lg.type};base64,${toB64(lg.data)}`, x: W - MX - w, w, h, altText: "Logo" }; })() : null;
    const cover = slide(null);
    if (logo) cover.addImage(Object.assign({}, logo, { y: 0.45 }));
    let cy = 0.75;
    if (doc.name) { cover.addText(str(doc.name).toUpperCase(), T({ x: MX, y: cy, w: CW - 2.4, h: 0.35, fontSize: 12, bold: true, color: brand, charSpacing: 2 })); cy += 0.45; }
    const titleH = Math.min(2.2, lines(str(doc.title) || " ", CW - 2.4, 36) * lh(36) + 0.1);
    cover.addText(str(doc.title), T({ x: MX, y: cy, w: CW - 2.4, h: titleH, fontSize: 36, bold: true }));
    cy += titleH + 0.1;
    if (doc.subtitle) { const sh = Math.min(1, lines(str(doc.subtitle), CW, 16) * lh(16) + 0.05); cover.addText(str(doc.subtitle), T({ x: MX, y: cy, w: CW, h: sh, fontSize: 16, color: MUT })); cy += sh + 0.15; }
    if (asOf) cover.addText([{ text: "AS OF  ", options: { fontSize: 11, bold: true, color: MUT, charSpacing: 1 } }, { text: asOf, options: { fontSize: 16, bold: true } }], T({ x: MX, y: cy, w: 4, h: 0.4, valign: "middle" }));
    if (SOLID[rag]) cover.addText(str(doc.rag).toUpperCase(), T({ x: asOf ? MX + 2.5 : MX, y: cy + 0.02, w: 1.3, h: 0.36, fontSize: 12, bold: true, charSpacing: 1.5, align: "center", valign: "middle",
      color: SOLID[rag][1], fill: { color: SOLID[rag][0] }, shape: "roundRect", rectRadius: 0.08 }));
    cy += 0.55;
    cover.addShape("rect", { x: MX, y: cy, w: CW, h: 0.04, fill: { color: brand }, line: { type: "none" } });
    cy += 0.4;

    const bs = (doc.blocks || []).map(forExport).filter(b => b && typeof b === "object");
    const kept = bs.filter((b, i) => b.type !== "heading" || (bs[i + 1] && bs[i + 1].type !== "heading"));
    let start = 0;
    if (kept[0] && kept[0].type === "tiles" && (kept[0].items || []).length && cy + tilesH(kept[0].items) <= BOTTOM) { tiles(cover, kept[0].items, cy); start = 1; }
    for (const b of kept.slice(start)) await block(b);
    return typeof document !== "undefined" ? pres.write({ outputType: "blob" }) : pres.write({ outputType: "uint8array" });
  }

  const api = { load, build };
  root.KeelsonPptx = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
