// day/night: light by default, follows the device until the visitor chooses; remember the choice when storage is available
const root = document.documentElement, themeBtn = document.getElementById("theme");
const current = () => root.dataset.theme || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
const paintBtn = () => themeBtn.setAttribute("aria-label", current() === "dark" ? "Switch to light theme" : "Switch to dark theme");
themeBtn.addEventListener("click", () => {
  const next = current() === "dark" ? "light" : "dark";
  root.dataset.theme = next;
  try { localStorage.setItem("keelson-theme", next); } catch (e) {}
  paintBtn();
});
matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change", paintBtn);
paintBtn();

// The free dashboard downloads straight from this site. Pro and the packs are mailto waitlists in the HTML (the no-JS
// default); a checkout link in store-config.js turns its [data-buy] button into a Lemon Squeezy checkout and swaps the
// [data-when="presale"] notes for the [data-when="sale"] ones (data-for names the link; none = any link).
document.querySelectorAll('[data-store="dashboard"]').forEach(a => { a.href = "downloads/Keelson-dashboard-v1.0.0.zip"; a.setAttribute("download", ""); });
const CO = (window.KEELSON_STORE || {}).checkout || {};
const buyUrl = k => typeof CO[k] === "string" && /^https:\/\/[a-z0-9-]+\.lemonsqueezy\.com\/(checkout\/)?buy\/[0-9a-f-]{36}(\?.*)?$/.test(CO[k]) ? CO[k] : "";
const BUY_LABEL = { proMonthly: "Get Keelson Pro", proYearly: "Get Keelson Pro", packSaas: "Buy the pack · $9" };
function paintBuy(bill = "m") {
  document.querySelectorAll("[data-buy]").forEach(a => {
    if (a.dataset.buyM) a.dataset.buy = [a.dataset["buy" + bill.toUpperCase()], a.dataset.buyM, a.dataset.buyY].find(buyUrl) || a.dataset.buyM; // billing toggle picks the variant
    const url = buyUrl(a.dataset.buy); if (!url) return;
    a.href = url; a.target = "_blank"; a.rel = "noopener";
    const t = [...a.childNodes].find(n => n.nodeType === 3 && n.nodeValue.trim()); // keep any icon after the label
    if (t) t.nodeValue = BUY_LABEL[a.dataset.buy] + (t.nextSibling ? " " : "");
  });
}
paintBuy();
const onSale = k => k ? !!buyUrl(k) : Object.keys(BUY_LABEL).some(buyUrl);
document.querySelectorAll("[data-when]").forEach(el => { el.hidden = (el.dataset.when === "sale") !== !!onSale(el.dataset.for); });

// popovers: product menu and the small-screen menu. Escape or a click outside closes and returns focus.
function popover(btn, pop) {
  const set = open => { btn.setAttribute("aria-expanded", open); pop.hidden = !open; };
  btn.addEventListener("click", () => set(btn.getAttribute("aria-expanded") !== "true"));
  pop.addEventListener("click", e => { if (e.target.closest("a")) set(false); });
  document.addEventListener("keydown", e => { if (e.key === "Escape" && !pop.hidden) { set(false); btn.focus(); } });
  document.addEventListener("click", e => { if (!pop.hidden && !pop.contains(e.target) && !btn.contains(e.target)) set(false); });
  btn.parentElement.addEventListener("focusout", e => { if (e.relatedTarget && !btn.parentElement.contains(e.relatedTarget) && !pop.contains(e.relatedTarget)) set(false); });
  pop.addEventListener("keydown", e => { // arrow keys move between the menu's links
    if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    const ls = [...pop.querySelectorAll("a")], i = ls.indexOf(document.activeElement);
    e.preventDefault(); ls[(i + (e.key === "ArrowDown" ? 1 : -1) + ls.length) % ls.length].focus();
  });
  btn.addEventListener("keydown", e => { if (e.key === "ArrowDown") { e.preventDefault(); set(true); pop.querySelector("a").focus(); } });
}
popover(document.getElementById("prodbtn"), document.getElementById("prodmenu"));
popover(document.getElementById("menubtn"), document.getElementById("mnav"));

// WAI-ARIA tabs with roving tabindex and automatic activation, for every tablist on the page
document.querySelectorAll('[role="tablist"]').forEach(list => {
  const tabs = [...list.querySelectorAll('[role="tab"]')];
  const sel = t => tabs.forEach(x => { const on = x === t; x.setAttribute("aria-selected", on); x.tabIndex = on ? 0 : -1; document.getElementById(x.getAttribute("aria-controls")).hidden = !on; });
  tabs.forEach((t, i) => {
    t.addEventListener("click", () => sel(t));
    t.addEventListener("keydown", e => {
      const j = e.key === "ArrowRight" ? i + 1 : e.key === "ArrowLeft" ? i - 1 : e.key === "Home" ? 0 : e.key === "End" ? tabs.length - 1 : null;
      if (j === null) return; e.preventDefault(); const n = tabs[(j + tabs.length) % tabs.length]; sel(n); n.focus(); n.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
  });
});

// the sections below live on one page each (index, templates, tools): each block runs only where its markup is
const $ = id => document.getElementById(id);

// billing toggle
const bills = [...document.querySelectorAll("[data-bill]")];
bills.forEach(b => b.addEventListener("click", () => {
  const k = b.dataset.bill; bills.forEach(x => x.setAttribute("aria-pressed", x === b));
  document.querySelectorAll("[data-m][data-y]").forEach(el => { el.textContent = el.dataset[k]; });
  paintBuy(k);
}));

// the CSV picker is a label: make Enter and Space open it like a button
$("pg-file-l")?.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); document.getElementById("pg-file").click(); } });

// slider demo: every value precomputed by the Keelson scheduler on the sample plan
const SCHED = {"6":{"rows":[[5.7,11.4,"22 Jan–29 Jan"],[35.7,1.4,"Fri 12 Feb"],[41.4,20,"16 Feb–1 Mar"],[61.4,20,"2 Mar–15 Mar"],[80,1.4,"Mon 15 Mar"],[84.3,1.4,"Thu 18 Mar"]],"golive":"Thu 18 Mar","var":0},"7":{"rows":[[5.7,15.7,"22 Jan–1 Feb"],[41.4,1.4,"Tue 16 Feb"],[42.9,20,"17 Feb–2 Mar"],[62.9,20,"3 Mar–16 Mar"],[81.4,1.4,"Tue 16 Mar"],[85.7,1.4,"Fri 19 Mar"]],"golive":"Fri 19 Mar","var":1},"8":{"rows":[[5.7,17.1,"22 Jan–2 Feb"],[42.9,1.4,"Wed 17 Feb"],[44.3,20,"18 Feb–3 Mar"],[64.3,20,"4 Mar–17 Mar"],[82.9,1.4,"Wed 17 Mar"],[90,1.4,"Mon 22 Mar"]],"golive":"Mon 22 Mar","var":2},"9":{"rows":[[5.7,18.6,"22 Jan–3 Feb"],[44.3,1.4,"Thu 18 Feb"],[45.7,20,"19 Feb–4 Mar"],[65.7,20,"5 Mar–18 Mar"],[84.3,1.4,"Thu 18 Mar"],[91.4,1.4,"Tue 23 Mar"]],"golive":"Tue 23 Mar","var":3},"10":{"rows":[[5.7,20,"22 Jan–4 Feb"],[45.7,1.4,"Fri 19 Feb"],[50,17.1,"22 Feb–5 Mar"],[70,17.1,"8 Mar–19 Mar"],[85.7,1.4,"Fri 19 Mar"],[92.9,1.4,"Wed 24 Mar"]],"golive":"Wed 24 Mar","var":4},"11":{"rows":[[5.7,21.4,"22 Jan–5 Feb"],[50,1.4,"Mon 22 Feb"],[51.4,20,"23 Feb–8 Mar"],[71.4,20,"9 Mar–22 Mar"],[90,1.4,"Mon 22 Mar"],[94.3,1.4,"Thu 25 Mar"]],"golive":"Thu 25 Mar","var":5},"12":{"rows":[[5.7,25.7,"22 Jan–8 Feb"],[51.4,1.4,"Tue 23 Feb"],[52.9,20,"24 Feb–9 Mar"],[72.9,20,"10 Mar–23 Mar"],[91.4,1.4,"Tue 23 Mar"],[95.7,1.4,"Fri 26 Mar"]],"golive":"Fri 26 Mar","var":6}};
const NAMES = ["Customer portal build", "Build complete", "System testing", "User acceptance testing", "Go / No-Go", "Go-live"];
const MS = [0, 1, 0, 0, 1, 1];
const host = $("rows"), slider = $("slider");
if (slider) {
host.innerHTML = NAMES.map((n, i) => `<div class="dr"><div class="n"><span>${n}</span><time></time></div><div class="dt"><i class="${MS[i] ? "ms" : ""}"></i></div></div>`).join("");
const render = () => {
  const v = Number(slider.value), r = SCHED[v], base = SCHED[6];
  slider.style.setProperty("--p", ((v - 6) / 6 * 100) + "%");
  document.getElementById("slv").textContent = v + " days";
  [...host.children].forEach((row, i) => {
    const [l, w, t] = r.rows[i], bar = row.querySelector("i");
    bar.style.left = l + "%"; bar.style.width = w + "%"; row.querySelector("time").textContent = t;
    row.classList.toggle("moved", t !== base.rows[i][2]);
  });
  const gl = document.getElementById("gl");
  gl.textContent = "Go‑live " + r.golive; gl.classList.toggle("bad", r.var > 0);
  document.getElementById("gv").textContent = r.var ? `+${r.var} working day${r.var > 1 ? "s" : ""} vs baseline. Every date after it moved on its own.` : "On baseline";
};
slider.addEventListener("input", render);
render();
}

const E = window.KeelsonEngine;
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const today = () => Math.floor((Date.now() - new Date().getTimezoneOffset() * 6e4) / E.DAY);
const fromInput = v => { const p = (v || "").split("-").map(Number); return p.length === 3 && p.every(Number.isFinite) ? E.fromYMD(...p) : null; };
const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
async function copy(text, out) {
  try { await navigator.clipboard.writeText(text); }
  catch (e) { const t = document.createElement("textarea"); t.value = text; document.body.append(t); t.select(); document.execCommand("copy"); t.remove(); }
  out.textContent = "Copied. Paste it into email, Teams or Slack."; setTimeout(() => out.textContent = "", 3000);
}

// Quick check: samples or the visitor's own plan, recalculated on every change, never sent anywhere.
// It opens on the latest weekly export of the FieldFlow sample, the same plan every mockup on the page and the app's sample use.
if ($("pg")) {
const SAMPLES = window.KEELSON_SAMPLES, OWN = "own", HIST = window.KEELSON_SAMPLE_HISTORY, LATEST = HIST.versions[HIST.versions.length - 1];
SAMPLES.plans.unshift({ id: "weekly", label: "Weekly export", note: "FieldFlow, exported " + E.fmt(fromInput(LATEST.asOf), true), text: LATEST.text, msKeys: HIST.msKeys });
const pg = { src: SAMPLES.plans[0].id, text: {}, cols: {}, sd: {}, title: {}, ms: {}, range: [0, 1], last: null, sliding: false };
SAMPLES.plans.forEach(p => { pg.text[p.id] = p.text; pg.sd[p.id] = fromInput(SAMPLES.statusDate); pg.title[p.id] = "FieldFlow rollout (sample)"; pg.ms[p.id] = p.msKeys.join(""); });
pg.text[OWN] = ""; pg.sd[OWN] = today(); pg.title[OWN] = "My project"; pg.ms[OWN] = "";
const txt = $("pg-text"), sheet = $("sheet"), slide = $("pg-slide");
$("pg-src").innerHTML = SAMPLES.plans.map(p => `<button type="button" aria-pressed="false" data-src="${p.id}">${esc(p.label)} sample</button>`).join("") +
  `<button type="button" aria-pressed="false" data-src="${OWN}">Your plan</button>`;
const srcTabs = [...$("pg-src").children];
srcTabs.forEach(b => b.addEventListener("click", () => use(b.dataset.src)));

function use(src) {
  pg.src = src; pg.last = null;
  srcTabs.forEach(b => b.setAttribute("aria-pressed", b.dataset.src === src));
  const s = SAMPLES.plans.find(p => p.id === src);
  $("pg-note").textContent = s ? s.note + " · fictional" : "Paste or open a file";
  txt.value = pg.text[src]; $("pg-title").value = pg.title[src]; $("pg-mskey").value = pg.ms[src];
  run(true);
}
txt.addEventListener("paste", e => { // pasting over a sample starts "Your plan" instead of mixing the two
  if (pg.src === OWN) return;
  const t = e.clipboardData?.getData("text"); if (!t) return;
  e.preventDefault(); pg.text[OWN] = t; pg.cols[OWN] = {}; pg.sd[OWN] = today(); use(OWN);
});
txt.addEventListener("input", () => { pg.text[pg.src] = txt.value; run(true); });
$("pg-file").addEventListener("change", e => {
  const f = e.target.files[0]; if (!f) return;
  const r = new FileReader();
  r.onload = () => { pg.text[OWN] = String(r.result); pg.cols[OWN] = {}; pg.sd[OWN] = today(); pg.title[OWN] = f.name.replace(/\.[^.]+$/, ""); use(OWN); };
  r.readAsText(f); e.target.value = "";
});
$("pg-clear").addEventListener("click", () => { pg.text[OWN] = ""; pg.cols[OWN] = {}; use(OWN); txt.focus(); });
$("pg-title").addEventListener("input", e => { pg.title[pg.src] = e.target.value; run(); });
$("pg-mskey").addEventListener("input", e => { pg.ms[pg.src] = e.target.value; run(); });
$("pg-date").addEventListener("change", e => { const v = fromInput(e.target.value); if (v != null) { pg.sd[pg.src] = v; run(); } });
slide.addEventListener("input", () => { pg.sliding = true; pg.sd[pg.src] = pg.range[0] + Number(slide.value); run(); });
slide.addEventListener("change", () => { pg.sliding = false; run(); });

const PLAY = '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M4 2.5v11l9-5.5z"/></svg>';
const PAUSE = '<svg viewBox="0 0 16 16" fill="currentColor" aria-hidden="true"><path d="M3.5 2.5h3v11h-3zM9.5 2.5h3v11h-3z"/></svg>';
let playing = 0;
const playBtn = $("pg-play"); playBtn.innerHTML = PLAY;
const stop = () => { cancelAnimationFrame(playing); playing = 0; pg.sliding = false; playBtn.innerHTML = PLAY; playBtn.setAttribute("aria-label", "Play the project from start to finish"); run(); };
playBtn.addEventListener("click", () => {
  if (playing) return stop();
  const [a, b] = pg.range, from = pg.sd[pg.src] >= b ? a : pg.sd[pg.src], t0 = performance.now(), ms = 7000 * (b - from) / Math.max(1, b - a);
  playBtn.innerHTML = PAUSE; playBtn.setAttribute("aria-label", "Pause");
  const tick = now => {
    const v = Math.min(b, Math.round(from + (b - from) * Math.min(1, (now - t0) / ms)));
    pg.sliding = true; pg.sd[pg.src] = v; run();
    if (v < b) playing = requestAnimationFrame(tick); else stop();
  };
  playing = requestAnimationFrame(tick);
});

const pctTxt = v => Math.round(v * 100) + "%";
const chip = s => `<span class="sc ${/^\+/.test(s) ? "Slip" : s.replace(" ", "-")}">${esc(s)}</span>`;
function run(newPlan) {
  const k = E.compute(pg.text[pg.src], { statusDate: pg.sd[pg.src], cols: pg.cols[pg.src], msKeys: [pg.ms[pg.src]] });
  if (newPlan) { // timeline range for the date slider: project start to the later of forecast and baseline finish
    const end = Math.max(k.finish ?? 0, k.bfinish ?? 0) || null, a = k.start ?? today(), b = (end ?? a + 90) + 7;
    pg.range = [Math.min(a - 7, b - 14), b]; slide.max = pg.range[1] - pg.range[0];
  }
  const sd = pg.sd[pg.src];
  slide.value = Math.max(0, Math.min(slide.max, sd - pg.range[0]));
  slide.style.setProperty("--p", (slide.value / slide.max * 100) + "%");
  $("pg-slide-out").textContent = E.fmt(sd, true);
  $("pg-date").value = E.iso(sd);
  colMap(k);
  sheet.innerHTML = draw(k);
  const rag = sheet.querySelector(".t-rag");
  if (rag && pg.last && pg.last !== k.rag && !reduce) rag.classList.add("flash");
  pg.last = k.rag; pg.k = k;
  if (!pg.sliding) $("pg-say").innerHTML = k.rows && k.cols.name >= 0 && k.cols.finish >= 0
    ? `<b>${k.rag}</b> on ${E.fmt(sd, true)}: ${esc(k.why.join(", "))}. ${pctTxt(k.pct)} complete, forecast finish ${E.fmt(k.finish, true)}.` : "";
}

function colMap(k) {
  const found = E.FIELDS.filter(([key]) => k.cols[key] >= 0).length;
  $("pg-colsum").textContent = k.heads.length ? `${found} of ${E.FIELDS.length}` : "";
  $("pg-colinfo").textContent = k.heads.length ? `Headers found in row ${k.hdrRow}${k.evidence ? `. Dates read ${k.dayFirst ? "day first (31/12)" : "month first (12/31)"}` : ""}. Change any match below.` : "";
  const opts = ['<option value="-1">— not used</option>'].concat(k.heads.map((h, i) => h ? `<option value="${i}">${esc(h)}</option>` : "")).join("");
  $("pg-colmap").innerHTML = E.FIELDS.map(([key, label, req]) =>
    `<label for="cm-${key}" class="${k.cols[key] >= 0 ? "ok" : ""}">${req ? "<b>" + label + "</b>" : label}</label><select id="cm-${key}" data-k="${key}">${opts}</select>`).join("");
  $("pg-colmap").querySelectorAll("select").forEach(s => {
    s.value = String(k.cols[s.dataset.k]);
    s.addEventListener("change", () => { (pg.cols[pg.src] ||= {})[s.dataset.k] = Number(s.value); run(true); });
  });
}

function draw(k) {
  const title = esc(pg.title[pg.src] || "Project status");
  if (!pg.text[pg.src].trim()) return `<div class="sh-empty"><b>Paste your plan to begin</b><span>Copy it from Excel, MS Project, Smartsheet or Jira, headers included. Or pick a sample above.</span></div>`;
  if (k.cols.name < 0 || k.cols.finish < 0) {
    document.querySelector(".pg-cols").open = true;
    return `<div class="sh-empty"><b>Almost there</b><span>We couldn't find a ${k.cols.name < 0 ? "task-name" : "finish or due date"} column. Pick it under <i>Columns found</i>, on the left.</span></div>`;
  }
  const sd = k.statusDate, f = v => E.fmt(v), F = v => E.fmt(v, true);
  const fin = k.finish == null ? "—" : F(k.finish);
  const finSub = k.finish == null ? "" : k.bfinish == null ? "No baseline in plan" : `Baseline ${f(k.bfinish)} · ${k.var === 0 ? "on time" : k.var > 0 ? `+${k.var} days late` : `${-k.var} days early`}`;
  const msSub = !k.nMs ? "" : k.nextMs ? `Next: ${esc(k.nextMs.name.slice(0, 24))} · ${f(k.nextMs.finish)}` : k.nMsDone === k.nMs ? "All met" : "None upcoming";
  const tiles = `<div class="sh-tiles">
    <div class="sh-t t-rag ${k.rag}"><span>Overall status</span><b>${k.rag}</b><small>${esc(k.why.join(" · "))}</small></div>
    <div class="sh-t"><span>Complete</span><b>${pctTxt(k.pct)}</b><small>${k.nTask ? `Expected ${pctTxt(k.exp)} by now · ${k.nTaskDone} of ${k.nTask} tasks done` : ""}</small></div>
    <div class="sh-t"><span>Forecast finish</span><b>${fin}</b><small>${finSub}</small></div>
    <div class="sh-t"><span>Milestones met</span><b>${k.ms}</b><small>${msSub}</small></div>
    <div class="sh-t"><span>Overdue tasks</span><b>${k.overT}</b><small>${k.overMs ? `${k.overMs} overdue milestone${k.overMs > 1 ? "s" : ""} · ` : ""}of ${k.nTask - k.nTaskDone} open tasks</small></div></div>`;
  const phNote = !k.nPhases ? "No phases found: add a Phase column, WBS numbers or outline levels."
    : [k.phOff ? `${k.phOff} earlier completed phase${k.phOff > 1 ? "s" : ""} not shown` : "", k.nPhases - k.phOff > 8 ? `+${k.nPhases - k.phOff - 8} more phases not shown` : ""].filter(Boolean).join(" · ");
  const phases = `<div><h4>Progress by phase</h4><table class="sh-tb"><colgroup><col style="width:38%"><col><col class="hide-sm" style="width:62px"><col style="width:84px"></colgroup>
    <thead><tr><th>Phase</th><th>Progress</th><th class="c hide-sm">Finish</th><th class="c">Status</th></tr></thead><tbody>${k.phases.map(p =>
      `<tr><td title="${esc(p.name)}">${esc(p.name)}</td><td><div style="display:flex;gap:8px;align-items:center"><div class="pbar" style="flex:1"><i style="width:${(p.pct * 100).toFixed(1)}%"></i></div><span class="r" style="width:34px">${pctTxt(p.pct)}</span></div></td><td class="c hide-sm">${f(p.finish)}</td><td class="c">${chip(p.status)}</td></tr>`).join("")}</tbody></table>
    ${phNote ? `<p class="sh-note">${phNote}</p>` : ""}</div>`;
  const msNote = !k.nMs ? "No milestones found: add a 0-day task, a Milestone column, or a keyword on the left." : k.msOpen > 8 ? `+${k.msOpen - 8} more open milestones` : k.msOpen === 0 ? "All milestones met" : "";
  const miles = `<div><h4>Milestones <span>${k.nMs ? k.ms + " met" : ""}</span></h4><table class="sh-tb"><colgroup><col><col style="width:62px"><col style="width:78px"></colgroup>
    <thead><tr><th>Milestone</th><th class="c">Due</th><th class="c">Status</th></tr></thead><tbody>${k.msList.map(m =>
      `<tr><td title="${esc(m.name)}">${esc(m.name)}</td><td class="c">${f(m.finish)}</td><td class="c">${chip(m.status)}</td></tr>`).join("")}</tbody></table>
    ${msNote ? `<p class="sh-note">${msNote}</p>` : ""}</div>`;
  // timeline: phase bars across the project, milestone diamonds, today line
  const t0 = k.start, t1 = Math.max(k.finish ?? t0, k.bfinish ?? t0), L = Math.max(1, t1 - t0 + 1), pos = v => ((v - t0) / L * 100).toFixed(2);
  const rows = k.phases.length ? k.phases : [{ name: "Whole plan", start: k.start, finish: k.finish, status: k.overT ? "Late" : "On track", ms: k.msList.map(m => m.finish) }];
  const months = []; if (t0 != null) { const [y, m] = E.iso(t0).split("-").map(Number); for (let i = 1; i < 60; i++) { const v = E.fromYMD(y, m + i, 1); if (v > t1) break; months.push(v); } }
  const every = Math.ceil(months.length / 8);
  const tl = t0 == null ? "" : `<div class="sh-tl"><h4>Timeline <span class="tl-leg">${["Done", "On track", "Planned", "At risk", "Late"].map(chip).join("")}<span class="now-key">Status date</span></span></h4><div class="tl-wrap"><div class="tl">
    <div class="tl-r tl-ax"><span></span><div class="tl-t">${months.filter((_, i) => i % every === 0).map(v => `<b style="left:${pos(v)}%">${E.fmt(v).split(" ")[1]}</b>`).join("")}</div></div>
    ${rows.map(p => { const a = p.start ?? p.finish, b = p.finish ?? a; return a == null ? "" :
      `<div class="tl-r"><span title="${esc(p.name)}">${esc(p.name)}</span><div class="tl-t"><i class="${p.status.replace(" ", "-")}" style="left:${pos(a)}%;width:${Math.max(.8, (b - a + 1) / L * 100).toFixed(2)}%"></i>${p.ms.map(v => `<em style="left:${pos(v + .5)}%"></em>`).join("")}${sd >= t0 && sd <= t1 ? `<span class="tl-now" style="left:${pos(sd)}%"></span>` : ""}</div></div>`; }).join("")}
    </div></div></div>`;
  const list = (items, late) => items.map(o => `<tr><td title="${esc(o.name)}">${esc(o.name)}</td><td>${esc(o.owner)}</td><td class="c">${f(o.finish)}</td>${late ? `<td class="r late">${o.late}d</td>` : ""}</tr>`).join("");
  const over = `<div><h4>Overdue tasks <span>${k.overT || ""}</span></h4><table class="sh-tb"><colgroup><col><col style="width:28%"><col style="width:56px"><col style="width:44px"></colgroup>
    <thead><tr><th>Task</th><th>Owner</th><th class="c">Due</th><th class="r">Late</th></tr></thead><tbody>${list(k.overdue, true)}</tbody></table>
    <p class="sh-note">${!k.overT ? "Nothing overdue." : k.overdueMore ? `+${k.overdueMore} more overdue tasks` : ""}</p></div>`;
  const up = `<div><h4>Due in the next ${k.upcomingDays} days</h4><table class="sh-tb"><colgroup><col><col style="width:30%"><col style="width:56px"></colgroup>
    <thead><tr><th>Task</th><th>Owner</th><th class="c">Due</th></tr></thead><tbody>${list(k.upcoming)}</tbody></table>
    <p class="sh-note">${!k.upcoming.length ? "Nothing due in this window." : k.upcomingMore ? `+${k.upcomingMore} more` : ""}</p></div>`;
  const c = k.checks, issues = [c.noFinish && `${c.noFinish} without a finish date`, c.badDates && `${c.badDates} with unreadable dates`, c.order && `${c.order} finish before start`,
    c.mismatch && `${c.mismatch} status/% mismatches`, c.capped && "only the first 1,000 rows read"].filter(Boolean);
  const check = issues.length ? `<span class="warn">Data check: ${k.rows} rows read · ${issues.join(" · ")}</span>` : `<span>Data check: ✓ ${k.rows} rows read cleanly${k.nCancel ? ` · ${k.nCancel} cancelled left out` : ""}</span>`;
  return `<div class="sh-head"><div><h3>${title}</h3><p>Executive summary</p></div><p class="asof">As of ${F(sd)}</p></div>${tiles}
    <div class="sh-two">${phases}${miles}</div>${tl}<div class="sh-two">${over}${up}</div>
    <div class="sh-foot">${check}<span>Made with Keelson · keelson.co.in</span></div>`;
}

$("pg-print").addEventListener("click", () => { // print a copy of the sheet alone, one landscape page
  const c = sheet.cloneNode(true); c.id = "print-sheet"; document.body.append(c); document.body.classList.add("printing");
  addEventListener("afterprint", () => { c.remove(); document.body.classList.remove("printing"); }, { once: true });
  print();
});
$("pg-copy").addEventListener("click", () => {
  const k = pg.k; if (!k || !k.rows || k.cols.finish < 0) return;
  const F = v => E.fmt(v, true), lines = [
    `${pg.title[pg.src] || "Project"}: status ${k.rag.toUpperCase()} as of ${F(k.statusDate)}`,
    `Why: ${k.why.join("; ")}`,
    `Progress: ${pctTxt(k.pct)} complete${k.nTask ? ` (expected ${pctTxt(k.exp)}), ${k.nTaskDone} of ${k.nTask} tasks done` : ""}`,
    `Forecast finish: ${F(k.finish)}${k.bfinish != null ? ` (baseline ${F(k.bfinish)}, ${k.var === 0 ? "on time" : k.var > 0 ? `+${k.var} working days` : `${-k.var} working days early`})` : ""}`,
    k.nMs ? `Milestones: ${k.ms} met${k.nextMs ? `, next: ${k.nextMs.name} on ${F(k.nextMs.finish)}` : ""}` : "",
    k.overT ? `Overdue: ${k.overdue.map(o => `${o.name} (${o.late} days)`).join(", ")}${k.overdueMore ? ` and ${k.overdueMore} more` : ""}` : "Overdue: none",
    k.upcoming.length ? `Due in the next ${k.upcomingDays} days: ${k.upcoming.map(o => `${o.name} (${E.fmt(o.finish)})`).join(", ")}` : "",
    "", "Made with Keelson, free at keelson.co.in"];
  copy(lines.filter((l, i) => l || i === lines.length - 2).join("\n"), $("pg-copied"));
});
use(pg.src);
}

// working-day calculator: Monday to Friday minus the holidays typed in
if ($("wd")) {
const wdState = { m: "add" };
const holidays = () => { const ls = $("wd-h").value.split(/[\n,;]+/).map(s => s.trim()).filter(Boolean), df = ls.some(s => /^\d{1,2}[\/.\-]/.test(s) && +s.split(/[\/.\-]/)[0] > 12);
  return new Set(ls.map(s => E.parseDate(s, df)).filter(v => v != null)); };
const working = (v, H) => E.wd(v) < 5 && !H.has(v);
function wdRun() {
  const H = holidays(), a = fromInput($("wd-a").value), out = $("wd-out"), sub = $("wd-sub"), xl = $("wd-xl");
  if (a == null) { out.textContent = "—"; sub.textContent = "Pick a start date."; xl.textContent = ""; return; }
  if (wdState.m === "add") {
    const n = Math.round(Number($("wd-n").value) || 0);
    let v = a - 1, c = 0; const dir = n < 0 ? -1 : 1;
    if (n === 0) v = a; else { if (dir < 0) v = a + 1; while (c < Math.abs(n)) { v += dir; if (working(v, H)) c++; } }
    out.textContent = (n < 0 ? "Starts " : "Finishes ") + ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"][E.wd(v)] + " " + E.fmt(v, true);
    const begin = !working(a, H) && n > 0 ? (() => { let b = a; while (!working(b, H)) b++; return b; })() : null;
    sub.textContent = n === 0 ? "A 0-day task is a milestone on its start date." : n < 0 ? `Work back ${-n} working days to hit ${E.fmt(a, true)}.` :
      `A ${n}-day task starting ${E.fmt(a, true)}${begin ? ` (a non-working day, so work begins ${E.fmt(begin)})` : ""} spans ${v - a + 1} calendar days.`;
    xl.textContent = n > 0 ? `=WORKDAY(A2-1, ${n}, Holidays)` : n < 0 ? `=WORKDAY(A2+1, ${n}, Holidays)` : "=A2";
  } else {
    const b = fromInput($("wd-b").value);
    if (b == null) { out.textContent = "—"; sub.textContent = "Pick an end date."; xl.textContent = ""; return; }
    const [lo, hi] = a <= b ? [a, b] : [b, a];
    let n = 0, hol = 0; for (let v = lo; v <= hi; v++) { if (E.wd(v) < 5) { if (H.has(v)) hol++; else n++; } }
    out.textContent = `${a <= b ? "" : "−"}${n} working day${n === 1 ? "" : "s"}`;
    sub.textContent = `${hi - lo + 1} calendar days, counting both ends · ${hi - lo + 1 - n - hol} weekend days${hol ? ` · ${hol} holiday${hol > 1 ? "s" : ""}` : ""}`;
    xl.textContent = "=NETWORKDAYS(A2, B2, Holidays)";
  }
}
document.querySelectorAll("#wd [data-m]").forEach(b => b.addEventListener("click", () => {
  wdState.m = b.dataset.m;
  document.querySelectorAll("#wd [data-m]").forEach(o => o.setAttribute("aria-pressed", o === b));
  $("wd-n-l").hidden = wdState.m !== "add"; $("wd-b-l").hidden = wdState.m === "add"; wdRun();
}));
["wd-a", "wd-n", "wd-b", "wd-h"].forEach(id => $(id).addEventListener("input", wdRun));
$("wd-a").value = E.iso(today()); $("wd-b").value = E.iso(today() + 30); wdRun();
}

// Go / No-Go: the runbook's criteria and verdict rules; marks are remembered in this browser only
if ($("gng")) {
const GNG = [["Testing", "SIT complete with no open Severity 1/2 defects", 1], ["Testing", "UAT signed off by business process owners", 1],
  ["Testing", "Performance test passed at expected peak load", 1], ["Testing", "Penetration test: no open critical or high findings", 1],
  ["Data", "Dress rehearsal completed inside the cutover window", 1], ["Data", "Reconciliation within agreed tolerance", 1], ["Data", "Data owners signed off cleansed data", 1],
  ["Business readiness", "End users trained to the agreed level", 1], ["Business readiness", "Go-live communications sent", 0], ["Business readiness", "Manual fallback procedures agreed", 1],
  ["Technical", "Production configured and smoke tested", 1], ["Technical", "Integrations verified in production", 1], ["Technical", "Monitoring and alerting enabled", 0],
  ["Support", "Hypercare rota and command centre staffed", 1], ["Support", "Knowledge-base articles published", 0],
  ["Rollback", "Rollback tested during dress rehearsal", 1], ["Rollback", "Rollback decision point and owner agreed", 1]];
const STATES = ["Met", "Waived", "Pending", "Not met"];
let marks = {}; try { marks = JSON.parse(localStorage.getItem("keelson-gng") || "{}"); } catch (e) {}
$("gng-list").innerHTML = GNG.map(([g, c, m], i) => (i === 0 || GNG[i - 1][0] !== g ? `<p class="gng-g">${g}</p>` : "") +
  `<div class="gng-r"><span>${esc(c)}${m ? "" : '<span class="m">OPTIONAL</span>'}</span><div class="st4" role="group" aria-label="${esc(c)}">${STATES.map(s => `<button type="button" data-i="${i}" data-v="${s}" aria-pressed="false">${s}</button>`).join("")}</div></div>`).join("");
function gngRun() {
  document.querySelectorAll("#gng-list [data-v]").forEach(b => b.setAttribute("aria-pressed", marks[b.dataset.i] === b.dataset.v));
  const st = i => marks[i] || "";
  const mandOpen = GNG.map((r, i) => i).filter(i => GNG[i][2] && ["Not met", "Pending", ""].includes(st(i)));
  const cond = GNG.map((r, i) => i).filter(i => st(i) === "Waived" || (!GNG[i][2] && ["Not met", "Pending"].includes(st(i))));
  const v = mandOpen.length ? ["nogo", "NO-GO", `${mandOpen.length} mandatory criteri${mandOpen.length > 1 ? "a" : "on"} not met yet`]
    : cond.length ? ["cond", "GO WITH CONDITIONS", `${cond.length} waived or optional item${cond.length > 1 ? "s" : ""} open`] : ["go", "GO", "All criteria met"];
  $("gng-v").className = "verdict " + v[0]; $("gng-v").querySelector("b").textContent = v[1]; $("gng-v").querySelector("span").textContent = v[2];
  gngRun.text = [`Go / No-Go: ${v[1]} (${v[2]})`, ...GNG.map(([g, c], i) => `${(st(i) || "Not checked").padEnd(11)} ${c}`), "", "Checked with Keelson, free at keelson.co.in"].join("\n");
  try { localStorage.setItem("keelson-gng", JSON.stringify(marks)); } catch (e) {}
}
$("gng-list").addEventListener("click", e => { const b = e.target.closest("[data-v]"); if (!b) return; marks[b.dataset.i] = marks[b.dataset.i] === b.dataset.v ? "" : b.dataset.v; gngRun(); });
$("gng-reset").addEventListener("click", () => { marks = {}; gngRun(); });
$("gng-copy").addEventListener("click", () => copy(gngRun.text, $("gng-copied")));
gngRun();
}
