// Keelson Pro: ReportDoc -> HTML (screen + print/PDF). Contract: docs/PRO-ARCHITECTURE.md section 2-3.
// KeelsonRender.css  : stylesheet string, inject once (every class is prefixed kr-)
// KeelsonRender.html(doc, {brand:{color, logo}, watermark, footer, editable}) -> one <article class="kr-page"> string
// Plan text is untrusted: every string is escaped (control characters dropped), except svg blocks (trusted output of reports.js).
// Placeholder blocks show their greyed prompt while editable; otherwise (print, thumbnails) their exportText, when given.
(function (root) {
  "use strict";
  const DEF = "#1a7f86";
  const CTRL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
  const esc = v => String(v == null ? "" : v).replace(CTRL, "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

  // colour helpers: brand must read as a rule/graphic (3:1) and, for text, 4.5:1 on white
  const hex = c => { c = String(c || "").trim(); if (/^#[0-9a-f]{3}$/i.test(c)) c = "#" + c.slice(1).replace(/./g, "$&$&"); return /^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : null; };
  const rgb = c => [1, 3, 5].map(i => parseInt(c.slice(i, i + 2), 16));
  const lum = c => { const [r, g, b] = rgb(c).map(v => { v /= 255; return v <= .03928 ? v / 12.92 : Math.pow((v + .055) / 1.055, 2.4); }); return .2126 * r + .7152 * g + .0722 * b; };
  const contrast = (a, b) => { const x = lum(a), y = lum(b); return (Math.max(x, y) + .05) / (Math.min(x, y) + .05); };
  const darken = (c, min) => { let [r, g, b] = rgb(c); for (let i = 0; i < 40 && contrast(c, "#ffffff") < min; i++) { r *= .94; g *= .94; b *= .94; c = "#" + [r, g, b].map(v => Math.round(v).toString(16).padStart(2, "0")).join(""); } return c; };
  const brandColors = color => { const c = hex(color) || DEF; return { rule: darken(c, 3), ink: darken(c, 4.5) }; };

  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const fmtDate = s => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || ""); return m ? `${+m[3]} ${MON[m[2] - 1]} ${m[1]}` : String(s || ""); };
  const TONES = ["red", "amber", "green", "grey", "teal"];
  const tone = t => TONES.includes(t) ? t : null;
  const RAG = { red: 1, amber: 1, green: 1 };
  const safeLogo = s => /^data:image\/(png|jpe?g|gif|webp|svg\+xml);base64,[a-z0-9+/=\s]+$/i.test(s || "") ? s : null;
  const num = (v, d) => (typeof v === "number" && isFinite(v) && v > 0 ? Math.round(v) : d);

  const wmSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="260"><g font-family="Arial,Helvetica,sans-serif" font-size="21" font-weight="700" fill="#1d2939" fill-opacity=".1" text-anchor="middle"><text transform="translate(105 75) rotate(-28)">PREVIEW · keelson.co.in</text><text transform="translate(315 205) rotate(-28)">PREVIEW · keelson.co.in</text></g></svg>`;
  const wmUrl = "data:image/svg+xml," + encodeURIComponent(wmSvg).replace(/'/g, "%27");

  const F = `-apple-system, BlinkMacSystemFont, "Segoe UI Variable Text", "Segoe UI", Inter, system-ui, "Helvetica Neue", Arial, sans-serif`;
  const css = `
.kr-page { --kr-ink: #1d2939; --kr-mut: #667085; --kr-line: #eaecf0; --kr-soft: #f2f4f7; --kr-navy: #1f3a5f; --kr-brand: ${DEF}; --kr-brand-ink: ${DEF};
  position: relative; box-sizing: border-box; width: 100%; max-width: 794px; margin: 0 auto; padding: 46px 48px 26px; display: flex; flex-direction: column;
  background: #fff; color: var(--kr-ink); color-scheme: light; text-align: left; border-radius: 3px;
  box-shadow: 0 1px 2px rgba(16, 24, 40, .06), 0 18px 40px -18px rgba(16, 24, 40, .28);
  font: 400 12.5px/1.55 ${F}; letter-spacing: -.004em; -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility }
.kr-page.kr-landscape { max-width: 1123px }
.kr-page *, .kr-page *::before, .kr-page *::after { box-sizing: border-box; margin: 0; padding: 0 }
:where(.kr-page) :where(h1, h2, p, ul, table, figure) { font: inherit; color: inherit }
.kr-tile > b, .kr-tb .kr-r, .kr-asof b { font-variant-numeric: tabular-nums }

.kr-head { display: flex; justify-content: space-between; align-items: flex-end; gap: 28px; padding-bottom: 18px; border-bottom: 2px solid var(--kr-brand) }
.kr-head-l { min-width: 0 }
.kr-eyebrow { font-size: 10.5px; line-height: 1.2; font-weight: 600; letter-spacing: .12em; text-transform: uppercase; color: var(--kr-brand-ink) }
.kr-page .kr-title { margin-top: 9px; font-size: 27px; line-height: 1.12; font-weight: 650; letter-spacing: -.028em; color: var(--kr-ink); overflow-wrap: anywhere }
.kr-sub { margin-top: 6px; color: var(--kr-mut); font-size: 12.5px; line-height: 1.4 }
.kr-meta { flex: none; display: flex; flex-direction: column; align-items: flex-end; gap: 12px; text-align: right }
.kr-logo { display: block; max-height: 34px; max-width: 160px; width: auto; height: auto; object-fit: contain }
.kr-stamp { display: flex; align-items: center; gap: 14px }
.kr-asof { font-size: 10px; line-height: 1.2; font-weight: 600; letter-spacing: .1em; text-transform: uppercase; color: var(--kr-mut) }
.kr-asof b { display: block; margin-top: 3px; font-size: 14px; letter-spacing: -.01em; text-transform: none; color: var(--kr-ink); font-weight: 600; white-space: nowrap }
.kr-rag { display: inline-flex; align-items: center; gap: 7px; padding: 7px 12px 7px 10px; border-radius: 999px; font-size: 11px; line-height: 1; font-weight: 700; letter-spacing: .06em; text-transform: uppercase; white-space: nowrap }
.kr-rag::before { content: ""; width: 7px; height: 7px; border-radius: 50%; background: currentColor; opacity: .85 }
.kr-rag-red { background: #d92d20; color: #fff } .kr-rag-amber { background: #f79009; color: #2b1700 } .kr-rag-green { background: #067647; color: #fff }

.kr-body { flex: 1 0 auto; padding-top: 20px; display: flex; flex-direction: column }
.kr-col { min-width: 0; display: flex; flex-direction: column }
.kr-blk + .kr-blk { margin-top: 12px }
.kr-blk + .kr-h { margin-top: 22px }
.kr-h + .kr-blk { margin-top: 10px }
.kr-col > .kr-h:first-child { margin-top: 0 }

.kr-tiles { display: grid; grid-template-columns: repeat(var(--kr-n, 4), minmax(0, 1fr)); gap: 10px }
.kr-tile { position: relative; overflow: hidden; background: var(--kr-soft); border-radius: 8px; padding: 13px 14px 12px; display: flex; flex-direction: column; gap: 5px; break-inside: avoid }
.kr-tile > span { font-size: 10px; line-height: 1.2; font-weight: 600; letter-spacing: .09em; text-transform: uppercase; color: #5d6679 }
.kr-tile > b { font-size: 25px; line-height: 1.05; font-weight: 650; letter-spacing: -.035em; color: var(--kr-ink); overflow-wrap: anywhere }
.kr-tile > small { font-size: 11px; line-height: 1.35; color: #5d6679 }
.kr-portrait .kr-tile > b { font-size: calc(31px - var(--kr-n, 4) * 2px) }
.kr-tile.kr-tn::before { content: ""; position: absolute; left: 0; top: 0; right: 0; height: 3px; background: var(--kr-tc) }
.kr-tile.kr-tn-red { --kr-tc: #d92d20 } .kr-tile.kr-tn-red > b { color: #b42318 }
.kr-tile.kr-tn-amber { --kr-tc: #f79009 } .kr-tile.kr-tn-amber > b { color: #93370d }
.kr-tile.kr-tn-green { --kr-tc: #067647 } .kr-tile.kr-tn-green > b { color: #067647 }
.kr-tile.kr-tn-grey { --kr-tc: #98a2b3 } .kr-tile.kr-tn-teal { --kr-tc: #1a7f86 } .kr-tile.kr-tn-teal > b { color: #0f5e63 }
.kr-tile.kr-solid::before { display: none }
.kr-tile.kr-solid.kr-tn-red { background: #d92d20 } .kr-tile.kr-solid.kr-tn-green { background: #067647 }
.kr-tile.kr-solid.kr-tn-red > *, .kr-tile.kr-solid.kr-tn-green > * { color: #fff }
.kr-tile.kr-solid.kr-tn-red > span, .kr-tile.kr-solid.kr-tn-red > small, .kr-tile.kr-solid.kr-tn-green > span, .kr-tile.kr-solid.kr-tn-green > small { color: rgba(255, 255, 255, .9) }
.kr-tile.kr-solid.kr-tn-amber { background: #f79009 } .kr-tile.kr-solid.kr-tn-amber > * { color: #2b1700 }

.kr-h { position: relative; padding-bottom: 7px; border-bottom: 1px solid var(--kr-line); font-size: 13.5px; line-height: 1.3; font-weight: 650; letter-spacing: -.014em; color: var(--kr-ink); break-after: avoid; break-inside: avoid }
.kr-h::after { content: ""; position: absolute; left: 0; bottom: -1px; width: 28px; height: 2px; background: var(--kr-brand) }

.kr-label { margin-bottom: 4px; font-size: 10px; line-height: 1.3; font-weight: 600; letter-spacing: .09em; text-transform: uppercase; color: var(--kr-mut); break-after: avoid }
.kr-txt { white-space: pre-line; overflow-wrap: anywhere; max-width: 82ch; font-size: 12.5px; line-height: 1.62; color: var(--kr-ink); orphans: 3; widows: 3 }
.kr-ph .kr-txt, .kr-ph.kr-ul { color: var(--kr-mut); font-style: italic; background: #fcfcfd; border: 1px dashed #d0d5dd; border-radius: 6px; padding: 8px 10px; max-width: none }
.kr-edit { border-radius: 4px; outline: 1px dashed transparent; outline-offset: 4px; cursor: text; transition: outline-color .15s, background .15s }
.kr-edit:hover { outline-color: #d0d5dd }
.kr-edit:focus { outline: 2px solid var(--kr-brand); outline-offset: 4px; background: #fcfcfd }

.kr-ul { list-style: none; display: grid; gap: 6px }
.kr-ul li { position: relative; padding-left: 17px; line-height: 1.5; overflow-wrap: anywhere; break-inside: avoid }
.kr-ul li::before { content: ""; position: absolute; left: 1px; top: .58em; width: 6px; height: 6px; border-radius: 1.5px; background: var(--kr-brand) }

.kr-tw { overflow-x: auto; -webkit-overflow-scrolling: touch }
.kr-tb { width: 100%; border-collapse: collapse; font-size: 12px; line-height: 1.4 }
.kr-tb th { padding: 0 10px 8px; border-bottom: 1.5px solid var(--kr-ink); font-size: 10px; line-height: 1.25; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; color: var(--kr-mut); text-align: left; vertical-align: bottom }
.kr-tb td { padding: 7px 10px; border-bottom: 1px solid var(--kr-line); vertical-align: middle; overflow-wrap: break-word }
.kr-tb th:first-child, .kr-tb td:first-child { padding-left: 0 } .kr-tb th:last-child, .kr-tb td:last-child { padding-right: 0 }
.kr-tb tr { break-inside: avoid } .kr-tb thead { display: table-header-group }
.kr-tb .kr-r { text-align: right; white-space: nowrap } .kr-tb .kr-c { text-align: center; white-space: nowrap }
.kr-tb .kr-empty td { padding: 18px 0; text-align: center; color: var(--kr-mut); font-style: italic }
.kr-b { font-weight: 650 }
.kr-tx.kr-tn-red { color: #b42318 } .kr-tx.kr-tn-amber { color: #93370d } .kr-tx.kr-tn-green { color: #067647 } .kr-tx.kr-tn-grey { color: var(--kr-mut) } .kr-tx.kr-tn-teal { color: #0f5e63 }
.kr-chip { display: inline-block; padding: 4px 9px; border-radius: 999px; font-size: 10.5px; line-height: 1; font-weight: 600; letter-spacing: .01em; white-space: nowrap; background: var(--kr-soft); color: #475467 }
.kr-chip.kr-tn-red { background: #fee4e2; color: #b42318 } .kr-chip.kr-tn-amber { background: #fef0c7; color: #93370d } .kr-chip.kr-tn-green { background: #dcfae6; color: #067647 }
.kr-chip.kr-tn-grey { background: #f2f4f7; color: #475467 } .kr-chip.kr-tn-teal { background: #d5f0f1; color: #0f5e63 }

.kr-fig { overflow-x: auto; break-inside: avoid }
.kr-fig svg { display: block; width: 100%; height: auto }

.kr-cols { display: grid; grid-template-columns: repeat(var(--kr-n, 2), minmax(0, 1fr)); gap: 32px }
.kr-note { font-size: 11px; line-height: 1.5; color: var(--kr-mut); padding-left: 10px; border-left: 2px solid var(--kr-line); break-inside: avoid }

.kr-foot { margin-top: 26px; padding-top: 10px; border-top: 1px solid var(--kr-line); display: flex; flex-wrap: wrap; justify-content: space-between; gap: 4px 16px; font-size: 10px; line-height: 1.3; color: var(--kr-mut) }
.kr-foot b { font-weight: 600; color: var(--kr-ink) }

.kr-wm { position: absolute; inset: 0; z-index: 5; pointer-events: none; user-select: none; -webkit-user-select: none; background: url("${wmUrl}") 0 0 / 420px 260px repeat }

@media screen and (max-width: 760px) {
  .kr-page { padding: 24px 18px 18px; border-radius: 0 }
  .kr-head { flex-direction: column-reverse; align-items: stretch; gap: 14px }
  .kr-meta { flex-direction: row-reverse; justify-content: space-between; align-items: center; text-align: left }
  .kr-stamp { flex-direction: row-reverse }
  .kr-page .kr-title { font-size: 22px }
  .kr-tiles { grid-template-columns: repeat(2, minmax(0, 1fr)) }
  .kr-tiles > .kr-solid:first-child { grid-column: 1 / -1 }
  .kr-tile > b { font-size: 21px }
  .kr-cols { grid-template-columns: minmax(0, 1fr); gap: 0 } .kr-cols > .kr-col + .kr-col { margin-top: 26px }
  .kr-tb { min-width: 560px } .kr-tb.kr-tb-empty { min-width: 0 }
  .kr-tw { background: linear-gradient(90deg, #fff 40%, rgba(255, 255, 255, 0)) left / 28px 100% no-repeat local, linear-gradient(270deg, #fff 40%, rgba(255, 255, 255, 0)) right / 28px 100% no-repeat local,
    radial-gradient(farthest-side at 0 50%, rgba(16, 24, 40, .16), transparent) left / 10px 100% no-repeat scroll, radial-gradient(farthest-side at 100% 50%, rgba(16, 24, 40, .16), transparent) right / 10px 100% no-repeat scroll }
  .kr-fig svg { min-width: 520px }
}
@media (prefers-reduced-motion: reduce) { .kr-edit { transition: none } }

@page kr-portrait { size: A4 portrait; margin: 12mm }
@page kr-landscape { size: A4 landscape; margin: 12mm }
@media print {
  .kr-page { max-width: none; width: auto; margin: 0; padding: 0; box-shadow: none; border-radius: 0; min-height: 270mm; page: kr-portrait; font-size: 12px }
  .kr-page.kr-landscape { min-height: 183mm; page: kr-landscape }
  .kr-page + .kr-page { break-before: page }
  .kr-foot { break-before: avoid; break-inside: avoid } /* never a page holding only the footer line */
  .kr-page, .kr-page * { -webkit-print-color-adjust: exact; print-color-adjust: exact }
  .kr-tw, .kr-fig { overflow: visible }
  .kr-fig svg { width: auto; max-width: 100%; max-height: 225mm; margin: 0 auto } .kr-landscape .kr-fig svg { max-height: 140mm }
  .kr-edit, .kr-edit:hover, .kr-edit:focus { outline: none; background: none }
  .kr-tile, .kr-tb tr, .kr-h, .kr-note, .kr-fig, .kr-ul li { break-inside: avoid }
}`;

  function cell(c, col) {
    const al = col && (col.align === "right" ? "kr-r" : col.align === "center" ? "kr-c" : "");
    let inner;
    if (c && typeof c === "object") {
      const t = tone(c.tone), cls = [c.chip ? "kr-chip" : t ? "kr-tx" : "", t ? "kr-tn-" + t : "", c.bold ? "kr-b" : ""].filter(Boolean).join(" ");
      inner = cls ? `<span class="${cls}">${esc(c.text)}</span>` : esc(c.text);
    } else inner = esc(c);
    return `<td${al ? ` class="${al}"` : ""}>${inner}</td>`;
  }

  // a placeholder block as it should print: its exportText in place of the prompt (text, or list items)
  const forExport = b => !b.placeholder || b.exportText == null ? b
    : Object.assign({}, b, { placeholder: false }, b.type === "list" ? { items: [].concat(b.exportText) } : { text: b.exportText });

  function block(b, o) {
    if (!b || typeof b !== "object") return "";
    if (!o.editable) b = forExport(b);
    const ph = b.placeholder ? " kr-ph" : "";
    const edit = o.editable && b.editable && b.id != null;
    const ed = kind => edit ? ` class="${kind} kr-edit" data-kr-edit="${esc(b.id)}" contenteditable="plaintext-only" spellcheck="true"` : ` class="${kind}"`;
    switch (b.type) {
      case "tiles": {
        const items = Array.isArray(b.items) ? b.items : [];
        return `<div class="kr-blk kr-tiles" style="--kr-n:${Math.min(Math.max(items.length, 1), 6)}">` + items.map(t => {
          const tn = tone(t.tone), solid = tn && RAG[tn] && (typeof t.solid === "boolean" ? t.solid : RAG[String(t.value).toLowerCase()]);
          return `<div class="kr-tile${tn ? " kr-tn kr-tn-" + tn : ""}${solid ? " kr-solid" : ""}"><span>${esc(t.label)}</span><b>${esc(t.value)}</b>${t.sub ? `<small>${esc(t.sub)}</small>` : ""}</div>`;
        }).join("") + `</div>`;
      }
      case "heading": return `<h2 class="kr-blk kr-h">${esc(b.text)}</h2>`;
      case "text": return `<div class="kr-blk kr-text${ph}">${b.label ? `<p class="kr-label">${esc(b.label)}</p>` : ""}<div${ed("kr-txt")}${edit && b.label ? ` aria-label="${esc(b.label)}"` : ""}>${esc(b.text)}</div></div>`;
      case "list": return `<ul${ed("kr-blk kr-ul" + ph)}${edit ? ` data-kr-kind="list"` : ""}>${(b.items || []).map(i => `<li>${esc(i)}</li>`).join("")}</ul>`;
      case "table": {
        const cols = Array.isArray(b.columns) ? b.columns : [], rows = Array.isArray(b.rows) ? b.rows : [];
        const al = c => c.align === "right" ? ` class="kr-r"` : c.align === "center" ? ` class="kr-c"` : "";
        const w = c => /^\d{1,3}(\.\d+)?%$/.test(c.width || "") ? ` style="width:${c.width}"` : "";
        return `<div class="kr-blk kr-tw${ph}"><table class="kr-tb${rows.length ? "" : " kr-tb-empty"}"><thead><tr>${cols.map(c => `<th scope="col"${al(c)}${w(c)}>${esc(c.label)}</th>`).join("")}</tr></thead><tbody>` +
          (rows.length ? rows.map(r => `<tr>${(r || []).map((c, i) => cell(c, cols[i])).join("")}</tr>`).join("")
            : `<tr class="kr-empty"><td colspan="${cols.length || 1}">${esc(b.empty || "Nothing to report.")}</td></tr>`) + `</tbody></table></div>`;
      }
      case "svg": return `<figure class="kr-blk kr-fig${ph}" role="img" aria-label="${esc(b.alt)}" style="max-width:${num(b.width, 680)}px">${String(b.svg || "")}</figure>`;
      case "columns": {
        const cs = Array.isArray(b.columns) ? b.columns : [];
        return `<div class="kr-blk kr-cols" style="--kr-n:${Math.min(Math.max(cs.length, 1), 3)}">${cs.map(c => `<div class="kr-col">${(c || []).map(x => block(x, o)).join("")}</div>`).join("")}</div>`;
      }
      case "note": return `<p class="kr-blk kr-note${ph}">${esc(b.text)}</p>`;
      default: return "";
    }
  }

  function html(doc, opts) {
    doc = doc || {}; const o = opts || {};
    const br = brandColors(o.brand && o.brand.color), logo = safeLogo(o.brand && o.brand.logo);
    const land = doc.orientation === "landscape", rag = String(doc.rag || "").toLowerCase(), asOf = fmtDate(doc.asOf);
    return `<article class="kr-page ${land ? "kr-landscape" : "kr-portrait"}" lang="en" style="--kr-brand:${br.rule};--kr-brand-ink:${br.ink}">` +
      `<header class="kr-head"><div class="kr-head-l">${doc.name ? `<p class="kr-eyebrow">${esc(doc.name)}</p>` : ""}<h1 class="kr-title">${esc(doc.title)}</h1>${doc.subtitle ? `<p class="kr-sub">${esc(doc.subtitle)}</p>` : ""}</div>` +
      `<div class="kr-meta">${logo ? `<img class="kr-logo" src="${esc(logo)}" alt="">` : ""}<div class="kr-stamp">${RAG[rag] ? `<span class="kr-rag kr-rag-${rag}">${esc(doc.rag)}</span>` : ""}${asOf ? `<p class="kr-asof">As of<b>${esc(asOf)}</b></p>` : ""}</div></div></header>` +
      `<div class="kr-body">${(doc.blocks || []).map(b => block(b, o)).join("")}</div>` +
      `<footer class="kr-foot"><span>${esc([doc.title, doc.name, asOf && "as of " + asOf].filter(Boolean).join(" · "))}</span>${o.footer ? `<span>Made with <b>Keelson</b> · keelson.co.in</span>` : ""}</footer>` +
      (o.watermark ? `<div class="kr-wm" aria-hidden="true"></div>` : "") + `</article>`;
  }

  const api = { css, html, brandColors, contrast, esc, forExport, CTRL };
  root.KeelsonRender = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
