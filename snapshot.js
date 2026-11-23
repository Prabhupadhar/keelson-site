/* Keelson snapshot: one self-contained .html file to send to a sponsor, client or team, who open it in any browser.
   Tabs (Overview, Reports, Look-ahead, Plan checks) are CSS only: the file holds no script at all, and its own
   Content-Security-Policy stops it loading anything (no network, no fonts, images only as data: URIs). It is a copy as of
   the day it was made: it does not update. Plan text is untrusted and escaped here; report pages come in as HTML from
   render-html.js, which escapes its own text.
     verdict(k)  the status, why and four facts, shared with the app's verdict block
     build(o)    -> the whole file as a string
   Pure: no DOM, no clock, no network. Tests: tests/web/snapshot.test.js. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const esc = v => String(v == null ? "" : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const D = v => E.fmt(v, true), pc = x => Math.round(x * 100) + "%";
  const LABEL = { Green: "On track", Amber: "At risk", Red: "Off track" }, TONE = { Green: "ok", Amber: "warn", Red: "bad" };

  function verdict(k) {
    if (!k) return null;
    const why = k.ragReason || (k.why || []).join(" · "); if (!TONE[k.rag] && !why) return null;
    const facts = [];
    if (k.finish != null) facts.push(["Finish", D(k.finish), k.bfinish == null ? "No baseline in this plan" : k.var > 0 ? `${k.var} working days after the baseline` : k.var < 0 ? `${-k.var} working days before the baseline` : "On the baseline"]);
    if (k.pct != null) facts.push(["Done", pc(k.pct), k.exp != null ? `${pc(k.exp)} expected by now` : `${k.nTaskDone} of ${k.nTask} tasks`]);
    if (k.nextMs) facts.push(["Next milestone", k.nextMs.name, D(k.nextMs.finish)]);
    facts.push(["Overdue", String(k.overT ?? 0), k.overT === 1 ? "task" : "tasks"]);
    return { rag: TONE[k.rag] ? k.rag : "", tone: TONE[k.rag] || "", label: LABEL[k.rag] || "Not enough dates to judge", why, facts };
  }

  const CSS = `
:root { color-scheme: light; --ink: #1d2939; --mut: #5d6779; --line: #e4e7ec; --soft: #f2f4f7; --acc: #1f4fd8; --ok: #067647; --warn: #9a5b00; --bad: #b42318 }
* { box-sizing: border-box } body { margin: 0; background: var(--soft); color: var(--ink); font: 15px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif }
.wrap { max-width: 1180px; margin: 0 auto; padding: 24px 16px 48px } h1 { margin: 0; font-size: 26px; line-height: 1.2 } h2 { margin: 0 0 12px; font-size: 19px }
.top p { margin: 4px 0 0; color: var(--mut); font-size: 14px }
.tabs > input { position: absolute; opacity: 0; pointer-events: none }
.bar { display: flex; flex-wrap: wrap; gap: 4px; margin: 20px 0; padding: 4px; border-radius: 999px; background: #e4e7ec; width: fit-content; max-width: 100% }
.bar label { padding: 8px 16px; border-radius: 999px; font-weight: 700; color: #344054; cursor: pointer }
.panel { display: none } .card { background: #fff; border: 1px solid var(--line); border-radius: 12px; padding: 20px; margin-bottom: 16px }
${[1, 2, 3, 4].map(i => `#t${i}:checked ~ .bar label[for="t${i}"] { background: #fff; color: var(--ink); box-shadow: 0 1px 2px rgb(0 0 0 / .12) } #t${i}:checked ~ #p${i} { display: block } #t${i}:focus-visible ~ .bar label[for="t${i}"] { outline: 3px solid var(--acc); outline-offset: 2px }`).join("\n")}
.verdict { display: grid; grid-template-columns: minmax(200px, 1fr) 2fr; gap: 16px 32px; align-items: center; border-left: 6px solid #98a2b3 }
.verdict.ok { border-left-color: var(--ok) } .verdict.warn { border-left-color: #dc8a00 } .verdict.bad { border-left-color: var(--bad) }
.chip { display: inline-block; padding: 1px 10px; border-radius: 999px; font-size: 13px; font-weight: 800; background: var(--soft) }
.chip.ok { background: #dcfae6; color: var(--ok) } .chip.warn { background: #fef0c7; color: var(--warn) } .chip.bad { background: #fee4e2; color: var(--bad) }
.v-main b { display: block; margin: 8px 0 2px; font-size: 28px; line-height: 1.15 } .v-main p { margin: 0; color: #344054 }
.facts { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 16px } .facts span, .facts small { display: block; font-size: 13px; color: var(--mut) } .facts b { display: block; font-size: 18px; overflow-wrap: anywhere }
.toc { display: flex; flex-wrap: wrap; gap: 8px; margin: 0 0 16px; padding: 0; list-style: none } .toc a { display: inline-block; padding: 6px 12px; border-radius: 999px; background: #fff; border: 1px solid var(--line); color: var(--ink); text-decoration: none; font-weight: 600; font-size: 14px }
.rep { margin: 0 0 28px; overflow-x: auto } .rep > .kr-page { margin: 0 auto }
table { width: 100%; border-collapse: collapse; font-size: 14px } th, td { padding: 7px 10px 7px 0; text-align: left; border-bottom: 1px solid var(--line); vertical-align: top } th { color: var(--mut); font-size: 13px }
td.n, th.n { text-align: center; padding: 7px 6px; font-variant-numeric: tabular-nums } .x { overflow-x: auto }
.st { display: inline-block; min-width: 84px; padding: 1px 8px; border-radius: 999px; font-size: 12px; font-weight: 800; background: var(--soft); color: #344054; text-align: center }
.st.pass { background: #dcfae6; color: var(--ok) } .st.fail { background: #fef0c7; color: var(--warn) }
.mut { color: var(--mut) } .note { font-size: 13px; color: var(--mut); margin: 24px 0 0 }
@media (max-width: 700px) { .verdict { grid-template-columns: 1fr } h1 { font-size: 22px } }
@media print { body { background: #fff } .bar, .toc { display: none } .panel { display: block !important; break-before: page } .tabs > .panel:first-of-type { break-before: auto } .card { break-inside: avoid } }`;

  // o: { title, asOf (day number), made (day number), verdict, reports: [{id, name, html}], reportCss, lookahead: {weeks, groups} | null,
  //      checks: {quality, drivers, workload} | null }
  function build(o) {
    const v = o.verdict, tabs = [];
    tabs.push(["Overview", (v ? `<section class="card verdict ${v.tone}" aria-label="Where the plan stands"><div class="v-main">${v.rag ? `<span class="chip ${v.tone}">${esc(v.rag)}</span>` : ""}<b>${esc(v.label)}</b><p>${esc(v.why)}</p></div>` +
      `<div class="facts">${v.facts.map(([l, b, s]) => `<div><span>${esc(l)}</span><b>${esc(b)}</b><small>${esc(s)}</small></div>`).join("")}</div></section>` : "") +
      `<section class="card"><h2>In this snapshot</h2><ul>${o.reports.length ? `<li>${o.reports.length === 1 ? "One report" : `${o.reports.length} reports`}: ${esc(o.reports.map(r => r.name).join(", "))}</li>` : ""}` +
      `${o.lookahead ? `<li>Every open task in the next ${o.lookahead.weeks} weeks, by person</li>` : ""}${o.checks ? "<li>Plan checks: how well the plan is built, what drives the finish date and who is busiest</li>" : ""}</ul></section>`]);
    if (o.reports.length) tabs.push(["Reports", `<ul class="toc">${o.reports.map(r => `<li><a href="#r-${esc(r.id)}">${esc(r.name)}</a></li>`).join("")}</ul>` +
      o.reports.map(r => `<section class="rep" id="r-${esc(r.id)}" aria-label="${esc(r.name)}">${r.html}</section>`).join("")]);
    if (o.lookahead) { const L = o.lookahead, n = L.groups.reduce((t, g) => t + g.items.length, 0);
      tabs.push(["Look-ahead", `<section class="card"><h2>Next ${L.weeks} weeks: ${n} open ${n === 1 ? "task" : "tasks"}</h2>${n ? `<div class="x"><table><thead><tr><th>Owner</th><th>Task</th><th>Start</th><th>Finish</th><th>Done</th><th>In this window</th></tr></thead><tbody>` +
        L.groups.flatMap(g => g.items.map((i, j) => `<tr><td>${j ? "" : `<b>${esc(g.owner)}</b>`}</td><td>${esc(i.name)}</td><td>${i.start != null ? esc(D(i.start)) : ""}</td><td>${esc(D(i.finish))}</td><td>${i.pct != null ? pc(i.pct) : ""}</td><td>${esc(i.note)}${i.note === "Overdue" ? ` by ${i.late} wd` : ""}</td></tr>`)).join("") +
        `</tbody></table></div>` : `<p class="mut">No open tasks in this window.</p>`}</section>`]); }
    if (o.checks) { const { quality: q, drivers: d, workload: w } = o.checks, lbl = { pass: "Pass", fail: "Look at", na: "Not checked" };
      tabs.push(["Plan checks", `<section class="card"><h2>Plan quality: ${q.pass} of ${q.of} checks passed</h2><p class="mut">After the DCMA 14-point assessment. The limits are a starting point, not a verdict.</p><div class="x"><table><thead><tr><th>Result</th><th>Check</th><th>Value</th><th>Limit</th></tr></thead><tbody>` +
        q.checks.map(c => `<tr><td><span class="st ${c.state}">${lbl[c.state]}</span></td><td>${c.n}. ${esc(c.name)}${c.state === "fail" && c.rows.length ? `<br><small class="mut">${esc(c.rows.slice(0, 5).map(r => r.name).join(", "))}${c.rows.length > 5 ? ` and ${c.rows.length - 5} more` : ""}</small>` : ""}</td><td>${c.state === "na" ? `<small class="mut">${esc(c.why)}</small>` : esc(c.value)}</td><td>${c.state === "na" ? "" : esc(c.limit)}</td></tr>`).join("") + `</tbody></table></div></section>` +
        `<section class="card"><h2>What drives the finish date</h2>${!d.has ? `<p class="mut">The plan has no float or critical column, so critical tasks are not shown.</p>` : !d.critical.length && !d.near.length ? `<p class="mut">No critical or near-critical open tasks.</p>` :
          `<div class="x"><table><thead><tr><th>Task</th><th>Owner</th><th>Finish</th><th>Why it matters</th></tr></thead><tbody>${[...d.critical.map(x => ({ ...x, c: 1 })), ...d.near].slice(0, 40).map(x => `<tr><td>${esc(x.name)}</td><td>${esc(x.owner)}</td><td>${esc(D(x.finish))}</td><td>${x.c ? "Critical" : "Near-critical"}${x.float != null ? ` · ${Math.round(x.float * 10) / 10} wd float` : ""}</td></tr>`).join("")}</tbody></table></div>`}</section>` +
        `<section class="card"><h2>Workload by person</h2>${w.rows.length ? `<p class="mut">Open tasks each person has in each week. Milestones are not counted.</p><div class="x"><table><thead><tr><th>Owner</th>${w.weeks.map(s => `<th class="n">${esc(E.fmt(s))}</th>`).join("")}<th class="n">Busiest</th></tr></thead><tbody>` +
          w.rows.map(r => `<tr><td>${esc(r.owner)}</td>${r.counts.map(n => `<td class="n">${n || ""}</td>`).join("")}<td class="n"><b>${r.peak}</b></td></tr>`).join("") + `</tbody></table></div>` : `<p class="mut">No open tasks in the coming weeks.</p>`}</section>`]); }
    const title = o.title || "Project";
    return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; form-action 'none'; base-uri 'none'">
<meta name="referrer" content="no-referrer"><title>${esc(title)}: status as of ${esc(D(o.asOf))}</title>
<style>${CSS}\n${o.reportCss || ""}</style></head><body><main class="wrap">
<header class="top"><h1>${esc(title)}</h1><p>Status as of ${esc(D(o.asOf))}. A snapshot made on ${esc(D(o.made))}: a copy that does not change when the plan does.</p></header>
<div class="tabs">${tabs.map((t, i) => `<input type="radio" name="tab" id="t${i + 1}"${i ? "" : " checked"}>`).join("")}
<nav class="bar" aria-label="Sections">${tabs.map((t, i) => `<label for="t${i + 1}">${t[0]}</label>`).join("")}</nav>
${tabs.map((t, i) => `<section class="panel" id="p${i + 1}" aria-label="${t[0]}">${t[1]}</section>`).join("\n")}
</div>
<p class="note">This file opens in any browser, works without a connection and holds no script.</p>
</main></body></html>`;
  }
  const fileName = (title, asOf) => `${String(title || "Project").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80) || "Project"} - status ${E.iso(asOf)}.html`;

  const api = { verdict, build, fileName, esc };
  root.KeelsonSnapshot = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
