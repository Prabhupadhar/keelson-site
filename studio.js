/* Keelson Studio: the Dashboard tab. Eight dark/light "studio" screens drawn from the same computed plan (KeelsonEngine k)
   as the reports, with motion from Pts (site/vendor/pts.min.js, loaded on first open). Nothing here is sample data:
   every figure comes from k, the saved versions (for "since last version") or the engine's rules.
   KeelsonStudio.load() -> Promise (Pts, optional: screens still render without motion if it fails)
   KeelsonStudio.mount(el, { k, prev, hist, title, client, screen }) -> { destroy() }
   KeelsonStudio.model(k, opts) -> plain figures, for tests */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine;
  const SRC = (typeof document !== "undefined" && document.currentScript && document.currentScript.src) || (typeof document !== "undefined" ? document.baseURI : "");
  const SCREENS = [["glance", "At a glance"], ["now", "Now line"], ["executive", "Executive"], ["overview", "Overview"], ["health", "Workstreams"], ["timeline", "Timeline"], ["team", "Team"], ["risks", "Risks"],
    ["journey", "Delivery"], ["milestones", "Milestones"], ["report", "Status report"]];
  // two palettes: neon on dark, deeper tones on light (AA as text on white); TH holds the canvas colours of the current look
  const DARK_C = { violet: "#8b5cf6", blue: "#3b82f6", cyan: "#22d3ee", teal: "#2dd4bf", green: "#34d399", amber: "#fbbf24", orange: "#fb923c", red: "#f43f5e", pink: "#e879f9", slate: "#94a3b8" };
  const LIGHT_C = { violet: "#6d28d9", blue: "#1d4ed8", cyan: "#0e7490", teal: "#0f766e", green: "#047857", amber: "#b45309", orange: "#c2410c", red: "#be123c", pink: "#a21caf", slate: "#64748b" };
  const C = { ...DARK_C }, ST = {}, TH = {};
  function theme(look) {
    const light = look.mode === "light";
    Object.assign(C, light ? LIGHT_C : DARK_C);
    Object.assign(ST, { "Done": C.green, "On track": C.teal, "At risk": C.amber, "Late": C.red, "Planned": C.slate });
    Object.assign(TH, { light, glow: { gloss: 1, matte: .35, flat: 0 }[look.finish] ?? 1, flat: look.finish === "flat",
      ink: light ? "#0f172a" : "#eef2ff", mute: light ? "#5b6475" : "#a5b4d8", card: light ? "rgba(255,255,255,.94)" : "rgba(15,22,55,.82)",
      line: light ? "rgba(15,23,42,.08)" : "rgba(148,163,255,.12)", trail: light ? "rgba(15,23,42,.3)" : "rgba(200,215,255,.45)" });
  }
  theme({ mode: "dark", finish: "gloss" });
  const ST_LABEL = { "Planned": "Not started" };
  const RAG = { Green: ["On track", "ok"], Amber: ["At risk", "warn"], Red: ["Off track", "bad"] };
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const HORIZON = 28; // calendar days of "load": the next 4 weeks

  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const F = v => v == null ? "—" : E.fmt(v, true);
  const pc = v => Math.round((v || 0) * 100);
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const ymd = v => E.iso(v).split("-").map(Number);
  const rgba = (hex, a) => { const n = parseInt(hex.slice(1), 16); return `rgba(${n >> 16},${n >> 8 & 255},${n & 255},${a})`; };
  const owners = s => String(s || "").split(/[;,]/).map(x => x.replace(/\[[^\]]*\]/g, "").trim()).filter(Boolean);
  const plural = (n, w) => `${n} ${w}${n === 1 ? "" : "s"}`;
  const nwd = (a, b) => b < a ? 0 : E.nwd(a, b);

  // ---------- figures ----------
  function model(k, o = {}) {
    const sd = k.statusDate, T = k.tasks || [], rules = k.rules || E.RULES;
    const work = T.filter(t => !t.ms), ms = T.filter(t => t.ms && t.finish != null).sort((a, b) => a.finish - b.finish);
    const prev = o.prev || null;
    const m = { title: o.title || "Project", client: o.client || "", sd, start: k.start, finish: k.finish, bfinish: k.bfinish,
      pct: pc(k.pct), exp: pc(k.exp), behind: k.behind, var: k.var, rag: RAG[k.rag] || ["No status", "none"],
      nTask: k.nTask, nTaskDone: k.nTaskDone, nMs: k.nMs, nMsDone: k.nMsDone, overT: k.overT, overMs: k.overMs, nPhases: k.nPhases,
      tasks: T, fc: o.fc || null, cols: k.cols || {}, heads: k.heads || [], dayFirst: !!k.dayFirst, extra: k.extra || [], guessed: k.guessed || [], rowsRead: k.rows,
      prev: prev ? { asOf: prev.asOf, pct: pc(prev.pct), overT: prev.overT } : null, rules, ragReason: k.ragReason || "" };
    m.phases = (k.phases || []).map((p, i) => {
      const pl = T.filter(t => t.phase === p.name);
      return { ...p, pct: pc(p.pct), i, n: pl.length, late: pl.filter(t => t.od).length, w: pl.filter(t => !t.ms).reduce((s, t) => s + t.w, 0),
        owners: [...new Set(pl.flatMap(t => owners(t.owner)))].slice(0, 3), ms: pl.filter(t => t.ms && t.finish != null) };
    });
    // load over the next 4 weeks: open working days that fall in the window, against the window's working days (one person, one task at a time)
    const end = sd + HORIZON, cap = Math.max(1, nwd(sd, end)), team = new Map();
    for (const t of work) for (const n of owners(t.owner)) {
      const r = team.get(n) || team.set(n, { name: n, open: 0, done: 0, late: 0, days: 0 }).get(n);
      if (t.done) { r.done++; continue; }
      r.open++; if (t.od) r.late++;
      const s = Math.max(t.start ?? t.finish ?? sd, sd), f = Math.min(t.finish ?? s, end);
      if (t.finish != null && f >= s) r.days += nwd(s, f) * (1 - (t.pct || 0));
      else if (t.od) r.days += 1; // overdue work still has to happen this month
    }
    m.team = [...team.values()].map(r => ({ ...r, load: Math.round(r.days / cap * 100) })).sort((a, b) => b.load - a.load || b.open - a.open);
    m.hasOwner = !!(k.cols && k.cols.owner >= 0) && m.team.length > 0;
    // risks the plan itself shows
    const R = [];
    for (const t of ms) if (t.od) R.push({ name: t.name, why: `Milestone ${plural(t.late, "working day")} overdue`, impact: "High", like: "High", status: "Late", owner: t.owner, score: 90 + Math.min(9, t.late / 5) });
    for (const t of ms) if (!t.done && !t.od && (t.var ?? 0) > 0) R.push({ name: t.name, why: `Milestone ${t.var} days later than baseline`, impact: "High", like: "Medium", status: "Slipping", owner: t.owner, score: 60 + Math.min(9, t.var / 5) });
    for (const t of work) if (t.od) R.push({ name: t.name, why: `${plural(t.late, "working day")} overdue`, impact: t.late > 10 ? "High" : "Medium", like: "High", status: "Overdue", owner: t.owner, score: (t.late > 10 ? 70 : 40) + Math.min(9, t.late / 5) });
    for (const p of m.phases) if (p.status === "At risk") R.push({ name: p.name, why: (p.var ?? 0) > rules.amberVar ? `Workstream finish ${p.var} days late` : `${p.behind} pts behind plan`, impact: "Medium", like: "Medium", status: "Watch", owner: p.owners.join(", "), score: 30 + Math.min(9, Math.max(p.var || 0, p.behind) / 5) });
    m.risks = R.sort((a, b) => b.score - a.score);
    // milestones for the galaxy: the last two done, then what is next
    const done = ms.filter(t => t.done), open = ms.filter(t => !t.done);
    m.ms = [...done.slice(-Math.max(1, 6 - open.length)), ...open].slice(0, 6).map(t => ({ name: t.name, finish: t.finish, done: t.done, od: t.od, late: t.late, var: t.var }));
    m.nextMs = open.filter(t => t.finish >= sd).slice(0, 4).map(t => ({ name: t.name, finish: t.finish }));
    // progress journey: planned % over time from the task dates, actual from saved versions, pace forecast
    const ww = work.reduce((s, t) => s + t.w, 0);
    const planAt = d => ww ? work.reduce((s, t) => { const a = t.start ?? t.finish, b = t.finish; if (a == null || b == null) return s;
      return s + t.w * (d < a ? 0 : d >= b ? 1 : Math.min(1, nwd(a, d) / t.w)); }, 0) / ww : 0;
    const pace = k.start != null && sd > k.start ? k.pct / (sd - k.start) : 0;
    m.paceFinish = k.pct >= 1 ? null : pace > 0 ? Math.round(sd + (1 - k.pct) / pace) : null;
    m.paceVar = m.paceFinish != null && k.finish != null ? m.paceFinish - k.finish : null;
    const t0 = k.start ?? sd, t1 = Math.max(k.finish ?? sd, sd, Math.min(m.paceFinish ?? 0, (k.finish ?? sd) + 730));
    const N = 40;
    m.curve = { t0, t1, plan: Array.from({ length: N + 1 }, (_, i) => { const d = Math.round(t0 + (t1 - t0) * i / N); return [d, planAt(d)]; }),
      actual: [...(o.hist || []).filter(v => v.day < sd && v.day >= t0).map(v => [v.day, v.pct]), [t0, 0], [sd, k.pct]].sort((a, b) => a[0] - b[0]) };
    m.work = { total: Math.round(ww), done: Math.round(work.reduce((s, t) => s + t.w * (t.done ? 1 : t.pct || 0), 0)) };
    const byPh = m.phases.filter(p => p.w > 0).sort((a, b) => b.w - a.w);
    m.mix = [...byPh.slice(0, 4).map(p => ({ name: p.name, v: p.w })), ...(byPh.length > 4 ? [{ name: "Others", v: byPh.slice(4).reduce((s, p) => s + p.w, 0) }] : [])];
    const lvl = (v, a, r) => v > r ? "Off track" : v > a ? "At risk" : "On track";
    m.health = [["Schedule", k.bfinish != null ? lvl(k.var, rules.amberVar, rules.redVar) : "No baseline"],
      ["Milestones", k.overMs >= rules.redMs ? "Off track" : k.overMs >= 1 ? "At risk" : "On track"],
      ["Tasks", lvl(k.overT, rules.amberOver, Infinity)], ["Progress", lvl(k.behind, rules.amberBehind, rules.redBehind)]];
    m.trend = [...(o.hist || []).filter(v => v.day < sd).map(v => [v.day, v.pct]), [sd, k.pct]].sort((a, b) => a[0] - b[0]);
    return m;
  }

  // ---------- small html helpers ----------
  const pill = (txt, cls) => `<span class="st-pill ${cls}">${esc(txt)}</span>`;
  const stCls = s => ({ "Done": "ok", "On track": "ok", "At risk": "warn", "Late": "bad", "Off track": "bad", "Overdue": "warn", "Slipping": "warn", "Watch": "info", "High": "bad", "Medium": "warn", "Low": "info", "No baseline": "none" })[s] || "none";
  const ic = (d, cls = "") => `<svg class="st-ic ${cls}" viewBox="0 0 24 24" aria-hidden="true">${d}</svg>`;
  const I = { cal: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4m6-4v4"/>', dl: '<path d="M12 4v11m0 0-4.5-4.5M12 15l4.5-4.5"/><path d="M4 16v1a3 3 0 0 0 3 3h10a3 3 0 0 0 3-3v-1"/>' };
  const fromIso = s => { const x = /^(\d{4})-(\d\d)-(\d\d)/.exec(s || ""); return x ? E.fromYMD(+x[1], +x[2], +x[3]) : null; };

  // ---------- screens ----------
  const S = {};
  // ---------- shared pieces of the original screens: a sentence, a field, a few figures ----------
  const head = (m, say) => `<div class="sx-head"><p class="sx-meta">As of ${esc(F(m.sd))}${m.client ? ` · ${esc(m.client)}` : ""}</p><h2>${esc(m.title)}</h2><p class="sx-say">${say}</p></div>`;
  const field = (id, h, cls = "") => `<div class="sx-field ${cls}" data-scene="${id}" style="height:${h}px" aria-hidden="true"></div>`;
  const stat = (k, v, sub, cls = "") => `<div class="sx-stat ${cls}"><span>${k}</span><b>${v}</b><small>${sub}</small></div>`;
  const stats = a => `<div class="sx-stats" style="--n:${a.length}">${a.join("")}</div>`;
  const keyL = a => `<ul class="sx-key">${a.map(([shape, c, t]) => `<li><i class="${shape}" style="--c:${c}"></i>${t}</li>`).join("")}</ul>`;
  const srT = (cap, hd, rows) => `<table class="st-sr"><caption>${cap}</caption><tr>${hd.map(h => `<th>${h}</th>`).join("")}</tr>${rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</table>`;
  const planAt = (m, d) => { const pl = m.curve.plan; if (!pl.length) return 0; if (d <= pl[0][0]) return pl[0][1];
    for (let i = 1; i < pl.length; i++) if (pl[i][0] >= d) { const [a, u] = pl[i - 1], [b, v] = pl[i]; return u + (v - u) * (b > a ? (d - a) / (b - a) : 0); } return pl[pl.length - 1][1]; };
  // earned schedule: the date by which the plan expected the work done so far
  const esDate = m => { const pl = m.curve.plan, a = m.pct / 100; if (!pl.length) return null; if (a <= pl[0][1]) return pl[0][0];
    for (let i = 1; i < pl.length; i++) if (pl[i][1] >= a) { const [d0, v0] = pl[i - 1], [d1, v1] = pl[i]; return Math.round(d0 + (v1 > v0 ? (a - v0) / (v1 - v0) : 0) * (d1 - d0)); }
    return pl[pl.length - 1][0]; };
  const days = n => plural(Math.abs(n), "day");
  const msList = m => m.tasks.filter(t => t.ms && t.finish != null).sort((a, b) => a.finish - b.finish);
  const fcText = m => m.fc && m.fc.ok ? `P50 · 50/50 date ${F(m.fc.p50)}, P80 · 8 in 10 ${F(m.fc.p80)}` : "";
  // risks with a date and a workstream, for the radar and the report
  function riskList(m) {
    const R = [], rules = m.rules;
    for (const t of m.tasks) {
      if (t.done || t.finish == null) continue;
      if (t.od) R.push({ name: t.name, phase: t.phase, date: m.sd, sev: t.ms || t.late > 10 ? 3 : 2, kind: t.ms ? "Late milestone" : "Overdue", why: `${plural(t.late || 0, "working day")} overdue`, c: C.red });
      else if (t.var > 0 && (t.ms || t.var >= 5)) R.push({ name: t.name, phase: t.phase, date: t.finish, sev: t.ms ? (t.var > 10 ? 3 : 2) : 1, kind: "Slipping", why: `${plural(t.var, "working day")} later than baseline`, c: C.amber });
    }
    for (const p of m.phases) if (p.status === "At risk" && p.finish != null) R.push({ name: p.name, phase: p.name, date: Math.max(p.finish, m.sd), sev: 2, kind: "Watch", why: (p.var ?? 0) > rules.amberVar ? `Workstream finish ${plural(p.var, "working day")} late` : `${p.behind} pts behind plan`, c: C.violet });
    return R.sort((a, b) => b.sev - a.sev || a.date - b.date).slice(0, 40);
  }
  const when = (d, sd) => d <= sd ? "now" : `in ${days(d - sd)}`;

  // ---------- executive: the three questions, answered once ----------
  S.executive = m => {
    const gaps = m.trend.map(([d, a]) => Math.round(planAt(m, d) * 100 - a * 100)), g1 = gaps[gaps.length - 1], g0 = gaps.length > 1 ? gaps[gaps.length - 2] : null;
    const trend = g0 == null ? ["First version", "No trend yet: save next week's plan to see the direction."] : g1 - g0 >= 2 ? ["Getting worse", `The gap to plan grew from ${g0} to ${g1} pts since the last version.`]
      : g0 - g1 >= 2 ? ["Getting better", `The gap to plan closed from ${g0} to ${g1} pts since the last version.`] : ["Holding", `The gap to plan stayed near ${g1} pts.`];
    const p = m.prev, flap = (a, b) => `<span class="sx-flap" data-to="${esc(b)}">${esc(b)}</span>${a != null && a !== b ? `<small>was ${esc(a)}</small>` : ""}`;
    const fin = m.bfinish != null && m.var ? `${m.var > 0 ? "+" : "−"}${Math.abs(m.var)} days` : "On baseline";
    const say = `${esc(m.rag[0])}. ${m.pct}% done against ${Math.round(planAt(m, m.sd) * 100)}% planned${m.finish != null ? `, finishing ${esc(F(m.finish))}` : ""}.${m.fc?.ok ? ` ${esc(fcText(m))}.` : ""}`;
    return `<div class="sx sx-exec">${head(m, `${pill(m.rag[0], m.rag[1])} ${say}`)}
      <div class="sx-three">
        <section class="sx-panel"><h3>What changed${p ? ` since ${esc(F(fromIso(p.asOf)))}` : ""}</h3>
          <dl class="sx-board"><div><dt>Done</dt><dd>${flap(p ? p.pct + "%" : null, m.pct + "%")}</dd></div><div><dt>Overdue tasks</dt><dd>${flap(p ? String(p.overT) : null, String(m.overT))}</dd></div>
          <div><dt>Finish vs baseline</dt><dd>${flap(null, fin)}</dd></div><div><dt>Milestones done</dt><dd>${flap(null, `${m.nMsDone}/${m.nMs}`)}</dd></div></dl>
          ${p ? "" : `<p class="sx-fine">Save next week's plan to see what moved.</p>`}</section>
        <section class="sx-panel"><h3>Is it getting worse?</h3><p class="sx-ans">${trend[0]}</p>${field("seismof", 170)}<p class="sx-fine">${trend[1]}</p></section>
        <section class="sx-panel"><h3>When will it finish?</h3>${m.fc?.ok ? `<p class="sx-ans">${esc(F(m.fc.p80))}</p>${field("sandf", 170)}<p class="sx-fine">P50 · 50/50 date ${esc(F(m.fc.p50))} · P80 · 8 in 10 ${esc(F(m.fc.p80))} · plan ${esc(F(m.finish))}</p>`
          : `<p class="sx-ans">${esc(F(m.finish))}</p><p class="sx-fine">The plan's finish. ${esc(m.fc?.reason || "A forecast needs more finished tasks.")}</p>`}</section>
      </div>
      ${srT("Gap to plan by version", ["Version", "Points behind plan"], m.trend.map(([d], i) => [F(d), gaps[i]]))}</div>`;
  };

  // ---------- overview: the project clock ----------
  S.overview = m => {
    const es = esDate(m), gap = es == null ? 0 : m.sd - es, planned = Math.round(planAt(m, m.sd) * 100);
    const say = es == null ? "No plan dates to set the clock by." : gap > 3 ? `The work done so far is what the plan expected by ${esc(F(es))}. The project clock runs ${days(gap)} behind the calendar.`
      : gap < -3 ? `The work done so far is what the plan expected by ${esc(F(es))}: ${days(gap)} ahead of the calendar.` : "The work done so far is what the plan expected by today. The project clock keeps time.";
    return `<div class="sx">${head(m, say)}<div class="sx-split">${field("clockf", 620)}<div class="sx-side">
      ${stat("Done", m.pct + "%", `${planned}% planned by today`)}${stat("Project clock", es == null ? "—" : esc(F(es)), es == null ? "" : gap > 3 ? `${days(gap)} behind today` : gap < -3 ? `${days(gap)} ahead of today` : "Keeping time", gap > 3 ? "warn" : "")}
      ${stat("Plan finish", esc(F(m.finish)), m.bfinish != null && m.var ? `${m.var > 0 ? m.var + " working days after" : -m.var + " working days before"} baseline` : "As planned")}
      <ol class="sx-list">${m.phases.slice(0, 8).map((p, i) => `<li><i style="--c:${ST[p.status] || C.slate}"></i><span>${i + 1}. ${esc(p.name)}</span><b>${p.pct}%</b></li>`).join("")}</ol></div></div>
      ${keyL([["bar", TH.ink, "Today hand"], ["bar", TH.acc, "Project clock: where the work done sits on the plan"], ["ring", C.teal, "Workstream rings, filled as work is done"], ["dia", TH.ink, "Milestone"]])}</div>`;
  };

  // ---------- workstreams: a race against the plan's ghost ----------
  S.health = m => {
    const P = m.phases, gap = p => Math.round((p.exp || 0) * 100) - p.pct, behind = P.filter(p => p.status !== "Done" && gap(p) > 2), ahead = P.filter(p => p.status !== "Done" && gap(p) < -2);
    const worst = [...behind].sort((a, b) => gap(b) - gap(a))[0];
    const say = !P.length ? "No workstreams found. Add a Phase column, or outline levels, to group tasks."
      : `${behind.length ? `${plural(behind.length, "workstream")} ${behind.length === 1 ? "is" : "are"} behind ${behind.length === 1 ? "its" : "their"} ghost` : "No workstream is behind its ghost"}${ahead.length ? `, ${ahead.length} ahead` : ""}. Each ghost runs where the plan says that workstream should be today.${worst ? ` Furthest behind: ${esc(worst.name)}, ${gap(worst)} pts.` : ""}`;
    return `<div class="sx">${head(m, say)}${P.length ? field("racef", 70 + Math.min(8, P.length) * 64) : ""}
      ${keyL([["dot", C.teal, "Workstream, at its % done"], ["ring", TH.mute, "Ghost: where the plan expects it today"], ["bar", C.red, "Behind the ghost"], ["bar", C.green, "Ahead of the ghost"]])}
      ${srT("Workstreams against plan", ["Workstream", "Done", "Plan expects", "Status"], P.map(p => [p.name, p.pct + "%", Math.round((p.exp || 0) * 100) + "%", p.status]))}</div>`;
  };

  // ---------- timeline: a lens over time ----------
  S.timeline = m => `<div class="sx">${head(m, `Move across the timeline to magnify any stretch of it: weeks and single tasks open up under the lens. ${m.finish != null ? `Planned finish ${esc(F(m.finish))}.` : ""}`)}
    ${m.phases.length ? field("lensf", 90 + Math.min(8, m.phases.length) * 58, "sx-lens") : `<p class="st-empty">No workstreams found.</p>`}
    ${keyL([["bar", C.teal, "Workstream, filled as work is done"], ["bar", TH.mute, "Task, under the lens"], ["dia", TH.ink, "Milestone"], ["bar", TH.acc, "Today"]])}
    ${srT("Workstreams", ["Workstream", "Start", "Finish", "Done"], m.phases.map(p => [p.name, F(p.start), F(p.finish), p.pct + "%"]))}</div>`;

  // ---------- team: the next 4 weeks on each person's shelf ----------
  function shelves(m) {
    if (m._sh) return m._sh;
    const sd = m.sd, end = sd + HORIZON, cap = Math.max(1, nwd(sd, end));
    return (m._sh = { cap, people: m.team.slice(0, 8).map(r => {
      const blocks = m.tasks.filter(t => !t.ms && !t.done && owners(t.owner).includes(r.name)).map(t => {
        const s = Math.max(t.start ?? t.finish ?? sd, sd), f = Math.min(t.finish ?? s, end);
        return { t, d: t.finish != null && f >= s ? nwd(s, f) * (1 - (t.pct || 0)) : t.od ? 1 : 0 }; }).filter(b => b.d > 0.05)
        .sort((a, b) => (b.t.od - a.t.od) || a.t.finish - b.t.finish);
      return { ...r, blocks }; }) });
  }
  S.team = m => {
    if (!m.hasOwner) return `<div class="sx">${head(m, "Add an Owner column to the plan to see who carries the next 4 weeks.")}</div>`;
    const over = m.team.filter(r => r.load > 100);
    const say = over.length ? `${esc(over.slice(0, 3).map(r => r.name).join(", "))}${over.length > 3 ? ` and ${over.length - 3} more` : ""} ${over.length === 1 ? "has" : "have"} more work in the next 4 weeks than fits. What falls off the shelf won't get done in time without help.`
      : "Everyone's next 4 weeks fit on their shelf.";
    const n = Math.min(8, m.team.length);
    return `<div class="sx">${head(m, say)}${field("shelff", Math.ceil(n / Math.min(4, n)) * 220 + 10)}
      ${keyL([["sq", TH.acc, "A task: width is its working days in the next 4 weeks"], ["sq", C.red, "Overdue"], ["bar", TH.ink, "The shelf holds 4 weeks of one person's time"]])}
      ${srT("Load over the next 4 weeks", ["Person", "Load", "Open tasks", "Late"], m.team.map(r => [r.name, r.load + "%", r.open, r.late]))}</div>`;
  };

  // ---------- risks: the radar ----------
  S.risks = m => {
    const R = riskList(m), here = R.filter(r => r.date <= m.sd).length, soon = R.filter(r => r.date > m.sd && r.date <= m.sd + 30).length;
    const say = R.length ? `${here ? `${plural(here, "risk")} already here` : "Nothing has arrived yet"}, ${soon ? `${soon} arriving in the next 30 days` : "nothing arriving in the next 30 days"}. Distance from the centre is time until it hits; size is how much it matters.`
      : "The plan shows no overdue work, slipping milestones or workstreams at risk.";
    return `<div class="sx">${head(m, say)}<div class="sx-split">${field("radarf", 600)}<div class="sx-side"><ol class="sx-risks">${R.slice(0, 7).map(r =>
      `<li><i style="--c:${r.c}"></i><div><b>${esc(r.name)}</b><small>${esc(r.why)} · ${esc(r.phase || "")}</small></div><span>${esc(when(r.date, m.sd))}</span></li>`).join("") || `<li class="st-empty">No risks in the plan.</li>`}</ol></div></div>
      ${keyL([["dot", C.red, "Overdue or late milestone"], ["dot", C.amber, "Later than baseline"], ["dot", C.violet, "Workstream at risk"]])}
      ${srT("Risks", ["Risk", "Why", "Workstream", "When"], R.map(r => [r.name, r.why, r.phase || "", when(r.date, m.sd)]))}</div>`;
  };

  // ---------- delivery: the approach to the finish date ----------
  S.journey = m => {
    const pf = m.paceFinish, dv = pf != null && m.finish != null ? pf - m.finish : null;
    const say = pf == null ? "Not enough progress yet to project a landing." : `At the current pace the project touches down ${esc(F(pf))}, ${dv > 0 ? `${days(dv)} past the planned finish` : dv < 0 ? `${days(dv)} before the planned finish` : "right on the planned finish"}. The beam is the plan's approach with its tolerance.`;
    return `<div class="sx">${head(m, say)}${field("glidef", 540)}
      ${stats([stat("Height above the approach", m.behind > 0 ? `${m.behind} pts` : m.behind < 0 ? `${-m.behind} pts low` : "On the beam", m.behind > 0 ? "Behind plan: more work left than planned" : "Work left is as planned or less", m.behind > m.rules.amberBehind ? "warn" : ""),
        stat("Touchdown at current pace", esc(F(pf)), dv == null ? "" : dv > 0 ? `${days(dv)} long` : "On or before the plan", dv > 0 ? "warn" : ""), stat("Planned touchdown", esc(F(m.finish)), m.bfinish != null ? `Baseline ${esc(F(m.bfinish))}` : "No baseline"),
        m.fc?.ok ? stat("P80 · 8 in 10", esc(F(m.fc.p80)), `P50 · 50/50 date ${esc(F(m.fc.p50))}`) : stat("Work done", m.pct + "%", `${m.work.done} of ${m.work.total} working days`)])}</div>`;
  };

  // ---------- milestones: a domino run ----------
  S.milestones = m => {
    const L = msList(m), done = L.filter(t => t.done), next = L.find(t => !t.done), slip = L.filter(t => !t.done && t.var > 0);
    const say = !L.length ? "No milestones in the plan. Mark them with zero duration or a Milestone column." : `${done.length} of ${L.length} milestones have fallen.${next ? ` Next to fall: ${esc(next.name)}, ${esc(F(next.finish))}${next.finish >= m.sd ? `, ${when(next.finish, m.sd)}` : ", overdue"}.` : ""}${slip.length ? ` ${plural(slip.length, "milestone")} now later than baseline.` : ""}`;
    return `<div class="sx">${head(m, say)}${L.length ? field("dominof", 470) : ""}
      ${keyL([["sq", C.green, "Achieved: fallen"], ["sq", TH.acc, "Next to fall"], ["sq", C.amber, "Later than baseline, by the working days shown"], ["sq", C.red, "Overdue: should have fallen"]])}
      ${srT("Milestones", ["Milestone", "Date", "Status"], L.map(t => [t.name, F(t.finish), t.done ? "Done" : t.od ? "Overdue" : t.var > 0 ? `${t.var} days later than baseline` : "Planned"]))}</div>`;
  };

  // ---------- status report: a front page ----------
  S.report = m => {
    const R = riskList(m), L = msList(m), next = L.filter(t => !t.done).slice(0, 5), planned = Math.round(planAt(m, m.sd) * 100), p = m.prev;
    const worst = [...m.phases].filter(x => x.status === "Late" || x.status === "At risk").sort((a, b) => (b.late - a.late) || (b.behind - a.behind))[0];
    const hl = m.overT ? `${plural(m.overT, "task")} overdue${worst ? ` as ${worst.name} falls behind` : ""}` : m.bfinish != null && m.var > 0 ? `Finish now ${m.var} working days after baseline${worst ? `, ${worst.name} to watch` : ""}`
      : m.finish != null ? `On course for ${F(m.finish)}` : "Status this week";
    const ch = p ? [`Done moved from ${p.pct}% to ${m.pct}%.`, p.overT !== m.overT ? `Overdue tasks went from ${p.overT} to ${m.overT}.` : `Overdue tasks unchanged at ${m.overT}.`] : ["This is the first saved version, so there is nothing to compare yet."];
    return `<article class="sx-paper"><header class="sx-mast"><span>Status report</span><span>As of ${esc(F(m.sd))}</span>${pill(m.rag[0], m.rag[1])}</header>
      <h2 class="sx-hl">${esc(hl)}</h2><p class="sx-dek">${esc(m.title)}${m.client ? ` · ${esc(m.client)}` : ""}</p>
      <p class="sx-lede">${m.pct}% of the work is done against ${planned}% planned by today. ${m.finish != null ? `The plan finishes ${esc(F(m.finish))}${m.bfinish != null && m.var ? `, ${m.var > 0 ? m.var + " working days after" : -m.var + " working days before"} baseline` : ""}.` : ""}${m.fc?.ok ? ` ${esc(fcText(m))}.` : ""}</p>
      <div class="sx-cols"><section><h3>What changed</h3>${ch.map(x => `<p>${esc(x)}</p>`).join("")}</section>
        <section><h3>Coming up</h3><ul>${next.map(t => `<li><span>${esc(t.name)}</span><time>${esc(F(t.finish))}</time></li>`).join("") || "<li>No open milestones.</li>"}</ul></section>
        <section><h3>Watch list</h3><ul>${R.slice(0, 5).map(r => `<li><span>${esc(r.name)}<small>${esc(r.why)}</small></span></li>`).join("") || "<li>Nothing on the watch list.</li>"}</ul></section></div>
      <div class="sx-cols sx-cols2"><section><h3>Standings</h3><table class="sx-stand">${m.phases.map(x => `<tr><td>${esc(x.name)}</td><td><span class="sx-meter" style="--w:${x.pct}%;--c:${ST[x.status] || C.slate}"></span></td><td>${x.pct}%</td><td>${pill(ST_LABEL[x.status] || x.status, stCls(x.status))}</td></tr>`).join("")}</table></section>
        <section><h3>Finish forecast</h3>${m.fc?.ok ? `${field("sandf", 150)}<p class="sx-fine">${esc(fcText(m))}. Plan ${esc(F(m.finish))}.</p>` : `<p>${esc(m.fc?.reason || "A forecast needs more finished tasks.")}</p>`}</section></div></article>`;
  };
  // ---------- at a glance: a bento of cards built from whatever columns the file has ----------
  const PRI = [[/^(blocker|highest|critical|urgent|p0|p1|1|showstopper)$/i, 0], [/^(high|major|p2|2)$/i, 1], [/^(medium|normal|moderate|p3|3)$/i, 2], [/^(low|minor|p4|4)$/i, 3], [/^(lowest|trivial|p5|5|nice to have)$/i, 4]];
  const priRank = v => (PRI.find(([re]) => re.test(v.trim())) || [0, 9])[1];
  const stateCol = v => { const s = v.toLowerCase(); return /^(not started|to ?do|backlog|new|open|planned|not yet started|pending|queued|ready)$/.test(s) ? TH.mute : /done|complete|closed|resolved|finished|delivered|released|live|accepted|yes|true/.test(s) ? C.green
    : /cancel|won.?t|dropped|descoped|duplicate|removed/.test(s) ? C.slate : /block|hold|stuck|risk|late|overdue|escalat/.test(s) ? C.red
    : /progress|doing|active|review|test|qa|started|develop|build|wip|ongoing/.test(s) ? TH.acc : TH.mute; };
  const priCol = v => [C.red, C.orange, C.amber, C.slate, TH.mute, TH.mute][Math.min(5, priRank(v))];
  const SEM_ICON = { status: "check", priority: "flag", severity: "alert", sprint: "repeat", points: "hash", hours: "clock", cost: "coin", risk: "alert", label: "tag", component: "box",
    team: "users", place: "pin", client: "brief", owner: "users", phase: "layers", type: "shape", date: "cal", flag: "check", other: "grid" };
  const GI = { check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>', flag: '<path d="M5 21V4h11l-2 4 2 4H5"/>', alert: '<path d="M12 4 2.8 19.5h18.4z"/><path d="M12 10v4m0 3v.01"/>',
    repeat: '<path d="M4 12a8 8 0 0 1 14-5.3L20 9M20 4v5h-5M20 12a8 8 0 0 1-14 5.3L4 15M4 20v-5h5"/>', hash: '<path d="M5 9h14M5 15h14M10 4 8 20M16 4l-2 16"/>',
    clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>', coin: '<circle cx="12" cy="12" r="8.5"/><path d="M14.8 9.2c-.6-.8-1.6-1.2-2.8-1.2-1.7 0-3 .9-3 2.1 0 2.9 6 1.4 6 4.1 0 1.2-1.3 2.1-3 2.1-1.3 0-2.4-.5-3-1.4M12 6.5v11"/>',
    tag: '<path d="M4 4h7l9 9-7 7-9-9z"/><circle cx="8.5" cy="8.5" r="1.3"/>', box: '<path d="M12 3 4 7.5v9L12 21l8-4.5v-9z"/><path d="M4 7.5 12 12l8-4.5M12 12v9"/>',
    users: '<circle cx="9" cy="8.5" r="3.5"/><path d="M3 20c.8-3.3 3.2-5 6-5s5.2 1.7 6 5"/><path d="M16 4.6a3.5 3.5 0 0 1 0 7.8M18 15c1.6.6 2.6 2.4 3 5"/>',
    pin: '<path d="M12 21s-6.5-6.1-6.5-11a6.5 6.5 0 0 1 13 0c0 4.9-6.5 11-6.5 11z"/><circle cx="12" cy="10" r="2.4"/>', brief: '<rect x="3.5" y="7.5" width="17" height="12" rx="2"/><path d="M9 7.5V5.5h6v2M3.5 12.5h17"/>',
    layers: '<path d="M12 4 3 8.5l9 4.5 9-4.5z"/><path d="m3 13 9 4.5 9-4.5"/>', shape: '<rect x="4" y="4" width="7" height="7" rx="1.5"/><circle cx="16.5" cy="16.5" r="3.5"/><path d="m16.5 4 3.5 6h-7z"/>',
    cal: '<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4m6-4v4"/>', grid: '<rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/>',
    done: '<circle cx="12" cy="12" r="8.5"/><path d="m8.5 12.3 2.4 2.4 4.6-5"/>', list: '<path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01"/>', file: '<path d="M14 3.5H7a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8.5z"/><path d="M14 3.5v5h5"/>' };
  const gi = k => `<svg class="gl-ic" viewBox="0 0 24 24" aria-hidden="true">${GI[k] || GI.grid}</svg>`;
  const compact = (v, cur = "") => { const a = Math.abs(v), s = a >= 1e9 ? (v / 1e9).toFixed(1) + "B" : a >= 1e6 ? (v / 1e6).toFixed(1) + "M" : a >= 1e4 ? (v / 1e3).toFixed(0) + "k" : a >= 1e3 ? (v / 1e3).toFixed(1) + "k" : String(Math.round(v * 10) / 10); return cur + s.replace(/\.0(?=[kMB]|$)/, ""); };
  const pctTxt = (a, b) => b ? Math.round(a / b * 100) : 0;
  // what to chart: every column with a handful of repeating values, every number column, every extra date and yes/no column
  function glanceData(m) {
    if (m._gl) return m._gl;
    const T = m.tasks.filter(t => !t.ms || m.tasks.every(x => x.ms)), dims = [], meas = [], dates = [], flags = [], seen = new Set();
    const dim = (name, sem, vals, multi) => {
      const map = new Map(); T.forEach((t, i) => { const v = vals(t, i); (multi ? v : [v]).forEach(x => { x = String(x ?? "").trim(); if (!x) return; const r = map.get(x) || map.set(x, { v: x, n: 0, done: 0 }).get(x); r.n++; if (t.done) r.done++; }); });
      let rows = [...map.values()]; if (rows.length < 2 && sem !== "owner") return; if (rows.length > 40 && !multi) return;
      rows.sort(sem === "priority" || sem === "severity" ? (a, b) => priRank(a.v) - priRank(b.v) || b.n - a.n : sem === "sprint" ? (a, b) => a.v.localeCompare(b.v, undefined, { numeric: true }) : (a, b) => b.n - a.n);
      const key = rows.map(r => r.v).join("|"); if (seen.has(key)) return; seen.add(key);
      dims.push({ name, sem, rows, total: T.length, more: Math.max(0, rows.length - 7) });
    };
    if (m.cols.status >= 0) dim(m.heads[m.cols.status] || "Status", "status", t => t.st);
    if (m.cols.phase >= 0 || m.phases.length > 1) dim(m.cols.phase >= 0 ? m.heads[m.cols.phase] : "Workstream", "phase", t => t.phase);
    if (m.cols.owner >= 0) dim(m.heads[m.cols.owner] || "Owner", "owner", t => owners(t.owner), true);
    if (m.cols.ms >= 0 && !T.every(t => /^(yes|no|y|n|true|false|1|0|)$/i.test(t.type || ""))) dim(m.heads[m.cols.ms] || "Type", "type", t => t.type);
    (m.extra || []).forEach((c, i) => {
      if (c.type === "category" && c.sem === "label") dim(c.name, "label", t => (t.x[i] || "").split(/[;,|]/), true);
      else if (c.type === "category") dim(c.name, c.sem || "other", t => t.x[i]);
      else if (c.type === "text" && (c.sem === "label" || c.sem === "component" || c.sem === "sprint")) dim(c.name, c.sem, t => (t.x[i] || "").split(/[;,|]/), true);
      else if (c.type === "number") { let tot = 0, dn = 0, cnt = 0; const byP = new Map(); T.forEach(t => { const v = E.numOf(t.x[i]); if (v == null) return; tot += v; cnt++; if (t.done) dn += v; byP.set(t.phase || "—", (byP.get(t.phase || "—") || 0) + v); });
        const cur = (T.map(t => t.x[i]).find(v => /[$€£₹¥]/.test(v)) || "").match(/[$€£₹¥]/)?.[0] || "";
        meas.push({ name: c.name, sem: c.sem || "other", tot, dn, cnt, cur, unit: c.sem === "points" ? "pts" : c.sem === "hours" ? "h" : "", by: [...byP].sort((a, b) => b[1] - a[1]).slice(0, 5) }); }
      else if (c.type === "date") { const by = new Map(); T.forEach(t => { const d = E.parseDate(t.x[i], m.dayFirst); if (d == null) return; const [y, mo] = ymd(d), k = y * 12 + mo - 1; by.set(k, (by.get(k) || 0) + 1); });
        if (by.size) dates.push({ name: c.name, by: [...by].sort((a, b) => a[0] - b[0]).slice(-12) }); }
      else if (c.type === "flag") { const f = T.map(t => t.x[i]).filter(Boolean); flags.push({ name: c.name, yes: f.filter(v => /^(yes|y|true|✓|✔|x)$/i.test(v)).length, n: f.length }); }
    });
    const st = t => t.done ? "done" : t.od ? "late" : t.var > 0 ? "slip" : t.start != null && t.start <= m.sd ? "doing" : "todo";
    const order = [...T].filter(t => t.finish != null).sort((a, b) => a.finish - b.finish), B = 220, per = Math.max(1, Math.ceil(order.length / B)), rank = { late: 4, slip: 3, doing: 2, todo: 1, done: 0 };
    const track = []; for (let i = 0; i < order.length; i += per) { const g = order.slice(i, i + per), s = g.map(st).sort((a, b) => rank[b] - rank[a])[0]; track.push({ s, t: g[0], n: g.length }); }
    const share = ["done", "doing", "slip", "late", "todo"].map(k => [k, T.filter(t => st(t) === k).length]);
    return (m._gl = { T, dims, meas, dates, flags, track, per, share });
  }
  S.glance = m => {
    const G = glanceData(m), p = m.prev, planned = Math.round(planAt(m, m.sd) * 100), dP = p ? m.pct - p.pct : null;
    const SC = { done: C.green, doing: TH.acc, slip: C.amber, late: C.red, todo: TH.mute }, SL = { done: "Done", doing: "In progress", slip: "Later than baseline", late: "Overdue", todo: "Not started" };
    let i = 0; const card = (span, body, cls = "") => `<section class="gl-card gl-s${span} ${cls}" style="--i:${i++}">${body}</section>`;
    const num = (v, suf = "") => `<b class="gl-num" data-to="${v}" data-suf="${esc(suf)}">${esc(v)}${esc(suf)}</b>`;
    const chip = (v, good) => v == null || v === 0 ? "" : `<span class="gl-chip ${good ? "up" : "down"}">${v > 0 ? "▲" : "▼"} ${Math.abs(v)}</span>`;
    const ring = (parts, size = 120, sw = 12, center = "") => { const r = (size - sw) / 2, L = 2 * Math.PI * r, tot = parts.reduce((s, x) => s + x[1], 0) || 1; let off = 0;
      return `<svg class="gl-ring" viewBox="0 0 ${size} ${size}" width="${size}" height="${size}" aria-hidden="true"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" class="gl-ring-bg" stroke-width="${sw}"/>${parts.filter(x => x[1] > 0).map(([c, v]) => { const len = v / tot * L, s = `<circle cx="${size / 2}" cy="${size / 2}" r="${r}" stroke="${c}" stroke-width="${sw}" stroke-dasharray="${Math.max(0, len - (parts.length > 1 ? 2 : 0)).toFixed(2)} ${L.toFixed(2)}" stroke-dashoffset="${(-off).toFixed(2)}" stroke-linecap="${parts.length > 1 ? "butt" : "round"}"/>`; off += len; return s; }).join("")}</svg>${center}`; };
    const spark = pts => { if (pts.length < 2) return ""; const xs = pts.map(q => q[0]), x0 = Math.min(...xs), x1 = Math.max(...xs), vs = pts.map(q => q[1]), lo = Math.min(...vs), hi = Math.max(...vs), P = pts.map(([d, v]) => [(d - x0) / ((x1 - x0) || 1) * 200, 50 - (v - lo) / ((hi - lo) || 1) * 44]);
      const ln = P.map(q => q.map(n => n.toFixed(1)).join(",")).join(" "); return `<svg class="gl-spark" viewBox="0 0 200 56" preserveAspectRatio="none" aria-hidden="true"><defs><linearGradient id="gl-sg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="var(--acc)" stop-opacity=".35"/><stop offset="1" stop-color="var(--acc)" stop-opacity="0"/></linearGradient></defs><polygon points="0,56 ${ln} 200,56" fill="url(#gl-sg)"/><polyline points="${ln}" pathLength="1"/></svg>`; };
    const bars = (rows, total, colOf) => `<ul class="gl-bars">${rows.slice(0, 7).map(r => `<li><span class="gl-bl">${esc(r.v)}</span><span class="gl-bt" title="${r.done} of ${r.n} done"><i style="--w:${pctTxt(r.n, Math.max(...rows.map(x => x.n)))}%;--c:${colOf(r.v)}"></i><i class="gl-bd" style="--w:${pctTxt(r.done, Math.max(...rows.map(x => x.n)))}%"></i></span><b>${r.n}</b></li>`).join("")}</ul>`;
    const head2 = (ic, title, sub) => `<header class="gl-h">${gi(ic)}<div><h3>${esc(title)}</h3>${sub ? `<p>${sub}</p>` : ""}</div></header>`;
    const trendPts = m.trend.length > 1 ? m.trend : m.curve.plan.filter(q => q[0] <= m.sd);
    const out = [
      card(4, `${head2("done", "Done", `${planned}% planned by today`)}<div class="gl-hero">${ring([[TH.acc, m.pct], [TH.light ? "rgba(15,23,42,.06)" : "rgba(255,255,255,.06)", 100 - m.pct]], 132, 14, `<div class="gl-ringc">${num(m.pct, "%")}${chip(dP, true)}</div>`)}
        <div class="gl-hero-r">${spark(trendPts)}<p class="gl-fine">${m.trend.length > 1 ? `${m.trend.length} saved versions` : "Planned progress to today"}</p></div></div>`, "gl-glow"),
      card(2, `<span class="gl-stamp">${gi("list")}</span><span class="gl-k">Tasks done</span>${num(m.nTaskDone)}<span class="gl-of">of ${m.nTask}</span>`),
      card(2, `<span class="gl-stamp">${gi("alert")}</span><span class="gl-k">Overdue</span>${num(m.overT)}${chip(p ? m.overT - p.overT : null, false)}<span class="gl-of">${m.overT ? "past their finish date" : "nothing past due"}</span>`, m.overT ? "gl-bad" : ""),
      card(2, `<span class="gl-stamp">${gi("flag")}</span><span class="gl-k">Milestones</span>${num(m.nMsDone)}<span class="gl-of">of ${m.nMs} achieved</span>`),
      card(2, `<span class="gl-stamp">${gi("cal")}</span><span class="gl-k">Finish</span><b class="gl-date">${esc(F(m.finish))}</b><span class="gl-of">${m.bfinish != null && m.var ? `${m.var > 0 ? "+" : "−"}${Math.abs(m.var)} working days vs baseline` : m.fc?.ok ? `P80 · 8 in 10 ${esc(F(m.fc.p80))}` : "as planned"}</span>`, m.var > 0 ? "gl-warn" : ""),
      card(12, `${head2("grid", "Every task", `${G.track.length < G.T.length ? `Each bar is ${G.per} tasks, ` : "One bar per task, "}in finish order. Hover for the task.`)}
        <div class="gl-track">${G.track.map((b, j) => `<i style="--c:${SC[b.s]};--j:${j}" title="${esc(b.t.name)}${b.n > 1 ? ` and ${b.n - 1} more` : ""} · ${esc(F(b.t.finish))} · ${SL[b.s]}"></i>`).join("")}</div>
        <div class="gl-stack">${G.share.filter(s => s[1]).map(([k, v]) => `<i style="--w:${pctTxt(v, G.T.length)}%;--c:${SC[k]}"></i>`).join("")}</div>
        <ul class="gl-legend">${G.share.filter(s => s[1]).map(([k, v]) => `<li><i style="--c:${SC[k]}"></i>${SL[k]} <b>${v}</b><span>${pctTxt(v, G.T.length)}%</span></li>`).join("")}</ul>`),
    ];
    G.dims.forEach(d => { const colOf = d.sem === "status" ? stateCol : d.sem === "priority" || d.sem === "severity" ? priCol : () => TH.acc;
      const top = d.rows.slice(0, 7), dn = d.rows.reduce((s, r) => s + r.done, 0), nn = d.rows.reduce((s, r) => s + r.n, 0);
      if (d.sem === "status" || ((d.sem === "priority" || d.sem === "severity") && d.rows.length <= 6))
        out.push(card(4, `${head2(SEM_ICON[d.sem] || "grid", d.name, `${plural(d.rows.length, "value")} across ${plural(nn, "task")}`)}<div class="gl-donut">${ring(top.map(r => [colOf(r.v), r.n]), 120, 16, `<div class="gl-ringc">${num(top[0].n)}<small>${esc(top[0].v)}</small></div>`)}
          <ul class="gl-legend gl-col">${top.map(r => `<li><i style="--c:${colOf(r.v)}"></i>${esc(r.v)} <b>${r.n}</b><span>${pctTxt(r.n, nn)}%</span></li>`).join("")}</ul></div>`));
      else out.push(card(4, `${head2(SEM_ICON[d.sem] || "grid", d.name, `${plural(d.rows.length, "value")} · ${pctTxt(dn, nn)}% of their tasks done${d.more ? ` · top 7 shown` : ""}`)}${d.sem === "owner" ? `<div class="gl-avs">${top.slice(0, 6).map((r, j) => `<span class="gl-av" style="--h:${(j * 57 + 210) % 360}" title="${esc(r.v)}">${esc(r.v.split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase())}</span>`).join("")}${d.rows.length > 6 ? `<span class="gl-av gl-more">+${d.rows.length - 6}</span>` : ""}</div>` : ""}${bars(top, nn, colOf)}`)); });
    G.meas.forEach(x => out.push(card(4, `${head2(SEM_ICON[x.sem] || "hash", x.name, `${plural(x.cnt, "task")} with a value`)}<div class="gl-meas"><div><span class="gl-k">Total</span><b class="gl-big">${esc(compact(x.tot, x.cur))}${x.unit ? `<small> ${x.unit}</small>` : ""}</b>
      <span class="gl-of">${esc(compact(x.dn, x.cur))}${x.unit ? " " + x.unit : ""} on done tasks · ${pctTxt(x.dn, x.tot)}%</span></div>${ring([[C.green, x.dn], [TH.light ? "rgba(15,23,42,.06)" : "rgba(255,255,255,.07)", Math.max(0, x.tot - x.dn)]], 84, 10)}</div>
      <ul class="gl-bars gl-thin">${x.by.map(([k, v]) => `<li><span class="gl-bl">${esc(k)}</span><span class="gl-bt"><i style="--w:${pctTxt(v, x.by[0][1])}%;--c:${TH.acc}"></i></span><b>${esc(compact(v, x.cur))}</b></li>`).join("")}</ul>`)));
    G.dates.forEach(x => { const mx = Math.max(...x.by.map(b => b[1])); out.push(card(4, `${head2("cal", x.name, "Tasks per month")}<div class="gl-cols">${x.by.map(([k, v]) => `<div title="${MON[k % 12]} ${Math.floor(k / 12)}: ${v}"><i style="--h:${pctTxt(v, mx)}%"></i><span>${MON[k % 12]}</span></div>`).join("")}</div>`)); });
    G.flags.forEach(x => out.push(card(2, `<span class="gl-stamp">${gi("check")}</span><span class="gl-k">${esc(x.name)}</span>${num(pctTxt(x.yes, x.n), "%")}<span class="gl-of">${x.yes} of ${x.n} say yes</span>`)));
    // about this file: which header fed which field, so a wrong guess is easy to spot
    const used = E.FIELDS.filter(([k]) => m.cols[k] >= 0).map(([k, lab]) => [lab, m.heads[m.cols[k]], m.guessed.includes(k)]);
    const cnt = ty => (m.extra || []).filter(c => c.type === ty).length, notes = (m.extra || []).filter(c => c.type === "text" || c.type === "empty" || c.type === "id").map(c => c.name);
    out.push(card(12, `${head2("file", "Read from your file", `${plural(m.rowsRead, "row")} · ${plural(used.length, "column")} mapped · ${plural((m.extra || []).length, "other column")} (${cnt("category")} grouped, ${cnt("number")} totalled, ${cnt("date")} dated, ${cnt("flag")} yes/no)`)}
      <dl class="gl-grid">${used.map(([lab, h, g]) => `<div><dt>${esc(lab)}</dt><dd>${esc(h)}${g ? ` <span class="gl-chip warn">read from the values</span>` : ""}</dd></div>`).join("")}</dl>
      ${notes.length ? `<p class="gl-fine">Not charted (IDs, free text or empty): ${esc(notes.slice(0, 8).join(", "))}${notes.length > 8 ? ` and ${notes.length - 8} more` : ""}.</p>` : ""}`, "gl-about"));
    return `<div class="gl">${head(m, `${esc(m.rag[0])}: ${m.pct}% done against ${planned}% planned${m.finish != null ? `, finishing ${esc(F(m.finish))}` : ""}. Every card below comes from a column in your file.`)}<div class="gl-bento">${out.join("")}</div></div>`;
  };
  // ---------- now line: past | today | future on one field; late and slipping work bends the Today line ----------
  function nowData(m) {
    if (m._now) return m._now;
    const sd = m.sd, T = m.tasks.filter(t => t.finish != null);
    const names = m.phases.slice(0, 8).map(p => p.name), other = T.some(t => !names.includes(t.phase));
    const lanes = [...m.phases.slice(0, 8).map(p => ({ name: p.name, pct: p.pct, status: p.status })), ...(other || !names.length ? [{ name: names.length ? "Other" : "All tasks", pct: null, status: null }] : [])]
      .map(l => ({ ...l, tasks: T.filter(t => l.name === "Other" || l.name === "All tasks" ? !names.includes(t.phase) : t.phase === l.name) }));
    lanes.forEach(l => {
      l.late = l.tasks.filter(t => t.od); l.slip = l.tasks.filter(t => !t.done && !t.od && t.var > 0);
      l.raw = l.late.reduce((s, t) => s + (t.w || 1) * Math.max(1, t.late || 1), 0) + .4 * l.slip.reduce((s, t) => s + (t.w || 1) * t.var, 0);
    });
    const top = Math.max(...lanes.map(l => l.raw), 1);
    lanes.forEach(l => { l.press = l.raw ? .25 + .75 * Math.sqrt(l.raw / top) : 0; });
    const fc = m.fc && m.fc.ok && m.fc.dist ? m.fc : null, ends = [m.finish, sd, ...(fc ? [fc.dist[fc.dist.length - 1], fc.p80] : [])].filter(v => v != null);
    const t0 = Math.min(m.start ?? sd, ...T.map(t => t.finish)), t1 = Math.max(...ends);
    return (m._now = { lanes, fc, t0, t1: t1 + Math.max(7, (t1 - t0) * .04), late: lanes.flatMap(l => l.late), slip: lanes.flatMap(l => l.slip),
      worst: [...lanes].sort((a, b) => b.raw - a.raw)[0] });
  }
  S.now = m => {
    const D = nowData(m), fc = D.fc, w = D.worst && D.worst.raw ? D.worst.name : "";
    const say = D.late.length ? `${plural(D.late.length, "late task")} ${D.late.length === 1 ? "is" : "are"} pressing on today${w ? `, most of the pressure in ${esc(w)}` : ""}.`
      : D.slip.length ? `Nothing is overdue, but ${plural(D.slip.length, "open task")} now ${D.slip.length === 1 ? "finishes" : "finish"} later than baseline${w ? `, most in ${esc(w)}` : ""}. That is what bends the Today line.`
      : "Nothing is pressing on today: no task is overdue or later than baseline.";
    const card = (k, v, sub, cls = "") => `<div class="sx-stat ${cls}"><span>${k}</span><b>${v}</b><small>${sub}</small></div>`;
    return `<div class="nw">
      <div class="sx-head"><p class="sx-meta">As of ${esc(F(m.sd))}${m.client ? ` · ${esc(m.client)}` : ""}</p><h2>${esc(m.title)}</h2><p class="sx-say">${say}</p></div>
      <div class="sx-field" data-scene="nowfield" aria-hidden="true"></div>
      <ul class="sx-key"><li><i class="k-done"></i>Done</li><li><i class="k-open"></i>Still to do</li><li><i class="k-slip"></i>Later than baseline (tail shows the slip)</li><li><i class="k-late"></i>Late, held at today</li><li><i class="k-ms"></i>Milestone</li>${fc ? `<li><i class="k-grain"></i>One possible finish date</li>` : ""}</ul>
      <div class="sx-stats">${card("Pressing on today", D.late.length, D.late.length ? `${plural(D.late.length, "task")} past due` : "No task past due", D.late.length ? "bad" : "")}
        ${card("Later than baseline", D.slip.length, D.slip.length && w ? `Most in ${esc(w)}` : "Open tasks on baseline", D.slip.length ? "warn" : "")}
        ${card("Plan finish", esc(F(m.finish)), m.bfinish != null && m.var ? `${m.var > 0 ? m.var + " working days after" : -m.var + " working days before"} baseline` : "As planned")}
        ${fc ? card("P50 · 50/50 date", esc(F(fc.p50)), "Half the simulated finishes are earlier") + card("P80 · 8 in 10", esc(F(fc.p80)), "8 in 10 simulated finishes are earlier") : card("Finish forecast", "—", esc(m.fc?.reason || "Needs more finished tasks"))}</div>
      <table class="st-sr"><caption>Pressure on today by workstream</caption><tr><th>Workstream</th><th>Late tasks</th><th>Later than baseline</th></tr>
        ${D.lanes.map(l => `<tr><td>${esc(l.name)}</td><td>${l.late.length}</td><td>${l.slip.length}</td></tr>`).join("")}</table>
    </div>`;
  };
  // the look: mode, finish (flat, matte, gloss), accent colour and motion, for every screen; saved per browser
  const LOOK_KEY = "keelson-studio-look";
  const ACCENTS = [["#1f4fd8", "Keelson blue"], ["#6366f1", "Indigo"], ["#8b5cf6", "Violet"], ["#0d9488", "Teal"], ["#059669", "Emerald"], ["#d97706", "Amber"], ["#e11d48", "Rose"], ["#334155", "Slate"]];
  const HEX = /^#[0-9a-f]{6}$/i;
  function lookOf(saved, brand) {
    const L = { mode: "dark", finish: "gloss", motion: "full", accent: "#6366f1", ...saved };
    if (!["dark", "light"].includes(L.mode)) L.mode = "dark";
    if (!["flat", "matte", "gloss"].includes(L.finish)) L.finish = "gloss";
    if (!["off", "subtle", "full"].includes(L.motion)) L.motion = "full";
    L.brand = L.accent === "brand" && HEX.test(brand?.color || "");
    L.acc = L.brand ? brand.color : HEX.test(L.accent) ? L.accent : "#6366f1";
    return L;
  }
  const onAcc = hex => { const n = parseInt(hex.slice(1), 16); return (.299 * (n >> 16) + .587 * (n >> 8 & 255) + .114 * (n & 255)) / 255 > .62 ? "#0f172a" : "#fff"; };
  // ---------- motion (Pts) ----------
  let ptsP = null;
  const load = () => ptsP || (ptsP = new Promise((ok, no) => {
    if (root.Pts) return ok(root.Pts);
    const s = document.createElement("script");
    s.src = new URL("vendor/pts.min.js", SRC).href;
    s.onload = () => root.Pts ? ok(root.Pts) : no(new Error("Pts did not load"));
    s.onerror = () => { ptsP = null; no(new Error("Pts did not load")); };
    document.head.appendChild(s);
  }));
  const rnd = seed => () => (seed = seed * 16807 % 2147483647) / 2147483647;
  const ease3 = u => 1 - (1 - clamp(u, 0, 1)) ** 3;
  const FT = (ctx, wt, px) => { ctx.font = `${wt} ${px}px ${ctx._ff || (ctx._ff = getComputedStyle(ctx.canvas).fontFamily)}`; };
  const text = (ctx, s, x, y, col, align = "left") => { ctx.textAlign = align; ctx.fillStyle = col; ctx.fillText(s, x, y); ctx.textAlign = "left"; };
  const line = (ctx, a, b, c, d, col, wd) => { ctx.beginPath(); ctx.moveTo(a, b); ctx.lineTo(c, d); ctx.strokeStyle = col; ctx.lineWidth = wd; ctx.stroke(); };
  const diamond = (ctx, x, y, r) => { ctx.beginPath(); ctx.moveTo(x, y - r); ctx.lineTo(x + r, y); ctx.lineTo(x, y + r); ctx.lineTo(x - r, y); ctx.closePath(); };
  function tipBox(ctx, x, y, w, h, a, b) {
    const nm = a.length > 46 ? a.slice(0, 45) + "…" : a; FT(ctx, 700, 12); let tw = ctx.measureText(nm).width; FT(ctx, 500, 11.5); tw = Math.max(tw, ctx.measureText(b).width) + 22;
    const bx = x + tw + 16 > w ? x - tw - 12 : x + 12, by = clamp(y - 46, 4, h - 52);
    ctx.fillStyle = TH.card; ctx.strokeStyle = rgba(TH.acc, .5); ctx.lineWidth = 1; ctx.beginPath(); ctx.roundRect(bx, by, tw, 46, 9); ctx.fill(); ctx.stroke();
    FT(ctx, 700, 12); text(ctx, nm, bx + 11, by + 19, TH.ink); FT(ctx, 500, 11.5); text(ctx, b, bx + 11, by + 35, TH.mute);
  }
  function glow(ctx, x, y, r, col, a = 1) {
    a *= TH.glow; if (!a) return;
    const g = ctx.createRadialGradient(x, y, 0, x, y, r); g.addColorStop(0, rgba(col, a)); g.addColorStop(1, rgba(col, 0));
    ctx.fillStyle = g; ctx.beginPath(); ctx.arc(x, y, r, 0, 6.2832); ctx.fill();
  }

  const SCENES = {
    // the project clock: workstream rings, the Today hand, and the clock hand at the earned-schedule date
    clockf(m) {
      const es = esDate(m), t0 = m.start ?? m.sd, t1 = Math.max(m.finish ?? m.sd, m.bfinish ?? 0, m.sd), span = t1 - t0 || 1, P = m.phases.slice(0, 8), n = Math.max(1, P.length);
      const A = d => -Math.PI / 2 + clamp((d - t0) / span, 0, 1) * Math.PI * 2 * .94, MS = msList(m);
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, u = ease3((T - 200) / 1600), cx = w * .5, cy = h * .5, R = Math.min(w, h) * .4, rw = Math.min(20, R * .5 / n), r0 = R - 28 - n * rw;
        const hx = hover[0] >= 0 ? hover[0] * w : -999, hy = hover[1] * h, hr = Math.hypot(hx - cx, hy - cy);
        // dial: elapsed time, month ticks, baseline notch, milestones
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, A(t0), A(t0) + (A(m.sd) - A(t0)) * u); ctx.closePath(); ctx.fillStyle = rgba(TH.acc, TH.light ? .05 : .07); ctx.fill();
        ctx.beginPath(); ctx.arc(cx, cy, R, A(t0), A(t1)); ctx.strokeStyle = TH.line; ctx.lineWidth = 1; ctx.stroke();
        const [y0, m0] = ymd(t0); FT(ctx, 600, 11);
        for (let i = 0, y = y0, mo = m0; i < 60; i++, mo++) { if (mo > 12) { mo = 1; y++; } const d = E.fromYMD(y, mo, 1); if (d > t1) break; if (d < t0) continue; const a = A(d);
          ctx.beginPath(); ctx.moveTo(cx + Math.cos(a) * (R - 6), cy + Math.sin(a) * (R - 6)); ctx.lineTo(cx + Math.cos(a) * (R + 3), cy + Math.sin(a) * (R + 3)); ctx.strokeStyle = TH.mute; ctx.lineWidth = 1.2; ctx.stroke();
          text(ctx, MON[mo - 1] + (mo === 1 ? " " + String(y).slice(2) : ""), cx + Math.cos(a) * (R + 18), cy + Math.sin(a) * (R + 18) + 4, TH.mute, "center"); }
        MS.forEach(t2 => { const a = A(t2.finish), x = cx + Math.cos(a) * (R - 14), y = cy + Math.sin(a) * (R - 14); diamond(ctx, x, y, 4.5); ctx.fillStyle = t2.done ? TH.ink : t2.od ? C.red : t2.var > 0 ? C.amber : TH.mute; ctx.fill(); });
        if (m.bfinish != null && m.bfinish !== m.finish) { const a = A(m.bfinish); FT(ctx, 600, 10.5); text(ctx, "Baseline", cx + Math.cos(a) * (R + 34), cy + Math.sin(a) * (R + 34) + 4, TH.mute, "center"); }
        // rings: track from start to finish, filled by % done
        let hit = null;
        P.forEach((p, i) => { if (p.start == null || p.finish == null) return; const r = R - 28 - (i + .5) * rw, a0 = A(p.start), a1 = A(p.finish), col = ST[p.status] || C.slate;
          const on = Math.abs(hr - r) < rw / 2 && (hit = hit || p) === p;
          ctx.lineCap = "round"; ctx.beginPath(); ctx.arc(cx, cy, r, a0, Math.max(a0 + .01, a0 + (a1 - a0) * u)); ctx.strokeStyle = rgba(col, on ? .35 : .18); ctx.lineWidth = rw * .62; ctx.stroke();
          ctx.beginPath(); ctx.arc(cx, cy, r, a0, a0 + Math.max(.01, (a1 - a0) * p.pct / 100 * u)); ctx.strokeStyle = col; ctx.lineWidth = rw * (on ? .7 : .62); ctx.stroke(); ctx.lineCap = "butt";
          if (TH.glow && on) glow(ctx, cx + Math.cos(a1) * r, cy + Math.sin(a1) * r, rw * 1.6, col, .5); });
        // hands and the gap between them
        const aT = A(t0) + (A(m.sd) - A(t0)) * u, aE = es == null ? aT : A(t0) + (A(es) - A(t0)) * ease3((T - 600) / 1800), behind = es != null && es < m.sd - 3;
        if (es != null && Math.abs(aT - aE) > .005) { ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R * .8, Math.min(aT, aE), Math.max(aT, aE)); ctx.closePath(); ctx.fillStyle = rgba(behind ? C.red : C.green, .16 + .04 * Math.sin(T / 700)); ctx.fill(); }
        const hand = (a, len, wd, col) => { ctx.beginPath(); ctx.moveTo(cx - Math.cos(a) * 14, cy - Math.sin(a) * 14); ctx.lineTo(cx + Math.cos(a) * len, cy + Math.sin(a) * len); ctx.strokeStyle = col; ctx.lineWidth = wd; ctx.lineCap = "round"; ctx.stroke(); ctx.lineCap = "butt"; };
        hand(aT, R - 2, 2.2, TH.ink); if (es != null) { if (TH.glow) glow(ctx, cx + Math.cos(aE) * R * .8, cy + Math.sin(aE) * R * .8, 18, TH.acc, .6); hand(aE, R * .8, 4.5, TH.acc); }
        FT(ctx, 700, 11.5); text(ctx, `Today · ${F(m.sd)}`, cx + Math.cos(aT) * (R + 46), cy + Math.sin(aT) * (R + 46), TH.ink, "center");
        if (es != null && u > .9) { const la = Math.abs(aE - aT) < .25 ? aE - .25 * Math.sign(aT - aE || 1) : aE; text(ctx, `Work done · ${F(es)}`, cx + Math.cos(la) * R * .62, cy + Math.sin(la) * R * .62, TH.acc, "center"); }
        // hub
        ctx.beginPath(); ctx.arc(cx, cy, Math.max(30, r0 - 6), 0, 6.2832); ctx.fillStyle = TH.card; ctx.fill(); ctx.strokeStyle = TH.line; ctx.stroke();
        FT(ctx, 800, Math.min(44, r0 * .55)); text(ctx, Math.round(m.pct * u) + "%", cx, cy + 6, TH.ink, "center");
        FT(ctx, 600, 12); text(ctx, es == null ? "done" : behind ? `${days(m.sd - es)} behind` : es > m.sd + 3 ? `${days(es - m.sd)} ahead` : "on time", cx, cy + 26, behind ? C.red : TH.mute, "center");
        if (hit) tipBox(ctx, hx, hy, w, h, hit.name, `${hit.pct}% done · ${F(hit.start)} – ${F(hit.finish)} · ${ST_LABEL[hit.status] || hit.status}`);
      };
    },
    // every workstream races its ghost: the ghost is at the plan's expected % for today
    racef(m) {
      const P = m.phases.slice(0, 8), seeds = P.map((_, i) => i * 1.7);
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, x0 = 230, x1 = w - 150, lh = (h - 50) / Math.max(1, P.length), X = v => x0 + (x1 - x0) * clamp(v, 0, 100) / 100, hy = hover[1] >= 0 ? hover[1] * h : -999;
        FT(ctx, 600, 11); [0, 25, 50, 75, 100].forEach(v => { text(ctx, v + "%", X(v), 16, TH.mute, "center"); ctx.setLineDash([2, 5]); line(ctx, X(v), 24, X(v), h - 18, TH.line, 1); ctx.setLineDash([]); });
        ctx.fillStyle = rgba(TH.ink, .7); for (let r = 0; r < (h - 50) / 5; r++) for (let c = 0; c < 2; c++) if ((r + c) % 2) ctx.fillRect(x1 + 8 + c * 5, 28 + r * 5, 5, 5);
        let hit = null;
        P.forEach((p, i) => {
          const cy = 36 + lh * (i + .5), col = ST[p.status] || C.slate, expv = Math.round((p.exp || 0) * 100), g = expv - p.pct, on = Math.abs(hy - cy) < lh / 2; if (on) hit = p;
          ctx.fillStyle = on ? rgba(TH.acc, .07) : TH.light ? "rgba(15,23,42,.025)" : "rgba(148,163,255,.035)"; ctx.beginPath(); ctx.roundRect(x0 - 14, cy - lh * .36, x1 - x0 + 28, lh * .72, lh * .36); ctx.fill();
          FT(ctx, 700, 13); text(ctx, p.name.length > 24 ? p.name.slice(0, 23) + "…" : p.name, 16, cy - 2, TH.ink);
          FT(ctx, 500, 11.5); text(ctx, p.owners.length ? p.owners.join(", ").slice(0, 32) : `${plural(p.n, "task")}`, 16, cy + 14, TH.mute);
          const run = ease3((T - 300 - i * 90) / 1700), xr = X(p.pct * run), xg = X(expv * clamp((T - 300) / 1700, 0, 1)), bob = Math.sin(T / 140 + seeds[i]) * 1.6 * (p.status === "Done" ? 0 : 1);
          if (p.status !== "Done" && Math.abs(g) > 0 && run > .95) { line(ctx, Math.min(xr, xg), cy, Math.max(xr, xg), cy, rgba(g > 0 ? C.red : C.green, .75), 3); }
          const tg = ctx.createLinearGradient(xr - 70, 0, xr, 0); tg.addColorStop(0, rgba(col, 0)); tg.addColorStop(1, rgba(col, .5)); line(ctx, Math.max(x0, xr - 70), cy + bob, xr, cy + bob, tg, 6);
          ctx.setLineDash([3, 3]); ctx.beginPath(); ctx.arc(xg, cy, 10, 0, 6.2832); ctx.strokeStyle = rgba(TH.mute, .8); ctx.lineWidth = 1.5; ctx.stroke(); ctx.setLineDash([]);
          if (TH.glow) glow(ctx, xr, cy + bob, 22, col, .45);
          ctx.beginPath(); ctx.arc(xr, cy + bob, 11, 0, 6.2832); ctx.fillStyle = col; ctx.fill(); FT(ctx, 800, 10.5); text(ctx, String(i + 1), xr, cy + bob + 4, onAcc(col), "center");
          FT(ctx, 700, 12.5); const lab = p.status === "Done" ? `Done · ${F(p.finish)}` : p.pct === 0 && expv === 0 ? `Starts ${F(p.start)}` : g > 2 ? `${g} pts behind` : g < -2 ? `${-g} pts ahead` : "With its ghost";
          text(ctx, lab, x1 + 26, cy + 4, p.status === "Done" ? TH.mute : g > 2 ? C.red : g < -2 ? C.green : TH.ink);
        });
        if (hit) tipBox(ctx, Math.min(X(hit.pct), w - 260), hy, w, h, hit.name, `${hit.pct}% done · plan expects ${Math.round((hit.exp || 0) * 100)}% · ${plural(hit.n, "task")}${hit.late ? ` · ${hit.late} late` : ""}`);
      };
    },
    // a fisheye lens over the timeline: the stretch under the pointer opens up to weeks and tasks
    lensf(m) {
      const P = m.phases.slice(0, 8), t0 = Math.min(m.start ?? m.sd, m.sd), t1 = Math.max(m.finish ?? m.sd, m.bfinish ?? 0, m.sd) + 7, span = t1 - t0 || 1;
      const byP = P.map(p => m.tasks.filter(t => t.phase === p.name && t.finish != null && !t.ms)), msP = P.map(p => m.tasks.filter(t => t.phase === p.name && t.ms && t.finish != null));
      const fish = (u, f, D) => { if (u < f) { const d = (f - u) / (f || 1e-9); return f - f * ((D + 1) * d / (D * d + 1)); } const d = (u - f) / ((1 - f) || 1e-9); return f + (1 - f) * ((D + 1) * d / (D * d + 1)); };
      let F0 = (m.sd - t0) / span, D0 = 0;
      return ({ ctx, w, h, t, hover, still }) => {
        const x0 = 200, x1 = w - 20, top = 54, lh = (h - top - 20) / Math.max(1, P.length), inField = hover[0] >= 0 && hover[0] * w > x0;
        const fT = inField ? clamp((hover[0] * w - x0) / (x1 - x0), 0, 1) : (m.sd - t0) / span, dT = inField ? 5 : 2.2;
        if (still) { F0 = fT; D0 = dT; } else { F0 += (fT - F0) * .14; D0 += (dT - D0) * .08; }
        const X = d => x0 + fish(clamp((d - t0) / span, 0, 1), F0, D0) * (x1 - x0), wk = d => X(d + 7) - X(d);
        // the lens itself, then weeks and months
        const fx = x0 + fish(F0, F0, D0) * (x1 - x0), lg = ctx.createLinearGradient(fx - 160, 0, fx + 160, 0); lg.addColorStop(0, rgba(TH.acc, 0)); lg.addColorStop(.5, rgba(TH.acc, TH.light ? .07 : .1)); lg.addColorStop(1, rgba(TH.acc, 0));
        ctx.fillStyle = lg; ctx.fillRect(fx - 160, top - 30, 320, h - top + 20);
        let last = -99, lastW = -99; const [y0, m0] = ymd(t0);
        for (let d = t0 - ((t0 % 7) + 7) % 7; d <= t1; d += 7) { const x = X(d), s = wk(d); if (d < t0) continue; if (s > 7) line(ctx, x, top - 6, x, h - 16, TH.line, 1);
          if (s > 44 && x - lastW > 44) { FT(ctx, 500, 10.5); const [, mo, da] = ymd(d); text(ctx, `${da} ${MON[mo - 1]}`, x + 3, top - 10, TH.mute); lastW = x; } }
        for (let i = 0, y = y0, mo = m0; i < 80; i++, mo++) { if (mo > 12) { mo = 1; y++; } const d = E.fromYMD(y, mo, 1); if (d > t1) break; if (d < t0) continue; const x = X(d); line(ctx, x, top - 26, x, h - 16, rgba(TH.ink, .14), 1);
          if (x - last > 34) { FT(ctx, 700, 11.5); text(ctx, MON[mo - 1] + (mo === 1 ? " " + String(y).slice(2) : ""), x + 4, top - 26, TH.ink); last = x; } }
        // rows: the workstream bar, its tasks once there is room, its milestones
        let near = null, nd = 1e9; const hx = hover[0] * w, hy = hover[1] * h;
        P.forEach((p, i) => { const cy = top + lh * (i + .5), col = ST[p.status] || C.slate;
          FT(ctx, 700, 12.5); text(ctx, p.name.length > 24 ? p.name.slice(0, 23) + "…" : p.name, 14, cy - 3, TH.ink); FT(ctx, 500, 11.5); text(ctx, `${p.pct}% · ${F(p.finish)}`, 14, cy + 13, TH.mute);
          if (i) line(ctx, x0, top + lh * i, x1, top + lh * i, TH.line, 1);
          if (p.start == null || p.finish == null) return;
          const a = X(p.start), b = Math.max(a + 6, X(p.finish)), pf = X(p.start + (p.finish - p.start) * p.pct / 100);
          ctx.fillStyle = rgba(col, .22); ctx.beginPath(); ctx.roundRect(a, cy - 15, b - a, 12, 6); ctx.fill(); ctx.fillStyle = col; ctx.beginPath(); ctx.roundRect(a, cy - 15, Math.max(4, pf - a), 12, 6); ctx.fill();
          byP[i].forEach((tk, j) => { const s = X(tk.start ?? tk.finish), f = Math.max(s + 2, X(tk.finish)), room = wk(tk.finish); if (room < 16) return;
            const al = clamp((room - 16) / 30, 0, 1), y = cy + 4 + (j % 3) * 5; line(ctx, s, y, f, y, rgba(tk.od ? C.red : tk.done ? TH.mute : col, .9 * al), 3);
            if (room > 70 && f - s > 60) { FT(ctx, 500, 10); text(ctx, tk.name.slice(0, Math.floor((f - s) / 6)), s, y - 3, rgba(TH.ink, al)); }
            const dd = Math.abs(hy - y) < 5 && hx >= s - 2 && hx <= f + 2 ? Math.abs(hy - y) : 1e9; if (dd < nd) { nd = dd; near = tk; } });
          msP[i].forEach(tk => { const x = X(tk.finish); diamond(ctx, x, cy - 9, 6); ctx.fillStyle = tk.done ? TH.ink : tk.od ? C.red : tk.var > 0 ? C.amber : TH.card; ctx.fill(); ctx.strokeStyle = TH.ink; ctx.lineWidth = 1.2; ctx.stroke();
            if (wk(tk.finish) > 40) { FT(ctx, 600, 10.5); text(ctx, tk.name.slice(0, 26), x + 9, cy - 20, TH.ink); }
            if (Math.hypot(hx - x, hy - cy + 9) < 9) { nd = 0; near = tk; } });
        });
        const tx = X(m.sd); line(ctx, tx, top - 30, tx, h - 12, TH.acc, 2); FT(ctx, 700, 11); const lw = ctx.measureText("Today").width + 14;
        ctx.fillStyle = TH.acc; ctx.beginPath(); ctx.roundRect(tx - lw / 2, h - 16, lw, 16, 8); ctx.fill(); text(ctx, "Today", tx, h - 4, onAcc(TH.acc), "center");
        if (near) tipBox(ctx, hx, hy, w, h, near.name, `${F(near.start ?? near.finish)} – ${F(near.finish)} · ${near.done ? "Done" : near.od ? "Overdue" : (near.pct ? Math.round(near.pct * 100) + "% done" : "Not started")}`);
      };
    },
    // each person's next 4 weeks: tasks drop onto a shelf that holds 4 weeks; the shelf sags, what doesn't fit falls off
    shelff(m) {
      const S0 = shelves(m), n = S0.people.length, cols = Math.min(4, n);
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, cw = w / cols, ch = 220, hx = hover[0] * w, hy = hover[1] * h; let near = null;
        S0.people.forEach((r, i) => {
          const ox = (i % cols) * cw, oy = Math.floor(i / cols) * ch, L = cw * .7, sx = ox + 18, sy = oy + 118, over = r.load > 100;
          const sag = Math.min(18, (Math.min(r.load, 160) / 100) ** 2 * 7) * ease3((T - 600 - i * 80) / 900), Y = x => sy + sag * 4 * clamp((x - sx) / L, 0, 1) * (1 - clamp((x - sx) / L, 0, 1));
          FT(ctx, 700, 13.5); text(ctx, r.name.length > 22 ? r.name.slice(0, 21) + "…" : r.name, sx, oy + 26, TH.ink);
          FT(ctx, 600, 12); text(ctx, `${r.load}% of 4 weeks · ${plural(r.open, "open task")}${r.late ? ` · ${r.late} late` : ""}`, sx, oy + 44, over ? C.red : TH.mute);
          // wall brackets and the plank
          ctx.fillStyle = rgba(TH.ink, .25); [sx + 10, sx + L - 16].forEach(bx => { ctx.beginPath(); ctx.moveTo(bx, Y(bx) + 5); ctx.lineTo(bx + 6, Y(bx) + 5); ctx.lineTo(bx, Y(bx) + 22); ctx.closePath(); ctx.fill(); });
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.quadraticCurveTo(sx + L / 2, sy + sag * 2, sx + L, sy); ctx.lineTo(sx + L, sy + 6); ctx.quadraticCurveTo(sx + L / 2, sy + 6 + sag * 2, sx, sy + 6); ctx.closePath();
          ctx.fillStyle = over ? rgba(C.red, .8) : TH.ink; ctx.fill();
          ctx.setLineDash([3, 4]); line(ctx, sx + L, sy - 50, sx + L, sy + 4, TH.mute, 1); ctx.setLineDash([]); FT(ctx, 500, 10.5); text(ctx, "4 weeks", sx + L - 2, sy - 54, TH.mute, "right");
          const floor = oy + ch - 26; line(ctx, sx + L + 6, floor, ox + cw - 14, floor, TH.line, 1);
          let x = sx, fx = sx + L + 10, spilled = 0;
          r.blocks.forEach((b, j) => {
            const bw = Math.max(5, b.d / S0.cap * L) - 2, fits = x + bw <= sx + L + .5, col = b.t.od ? C.red : rgba(TH.acc, j % 2 ? .62 : .9), at = 500 + i * 80 + j * 110, u = clamp((T - at) / 520, 0, 1);
            let bx, by, rot = 0, ww = bw;
            if (fits) { bx = x; by = Y(x + bw / 2) - 24; x += bw + 2; }
            else { spilled += b.d; ww = Math.min(bw, ox + cw - 14 - fx); if (ww < 4) return; bx = fx; by = floor - 24; rot = .1 * ((j % 3) - 1); fx += ww + 3; }
            if (u <= 0) return; const yy = by - (1 - u * u) * (by - oy + 30);
            ctx.save(); ctx.translate(bx + ww / 2, yy + 12); ctx.rotate(rot * u); ctx.fillStyle = col; ctx.beginPath(); ctx.roundRect(-ww / 2, -12, ww, 24, 4); ctx.fill();
            if (ww > 54) { FT(ctx, 600, 10.5); ctx.save(); ctx.beginPath(); ctx.rect(-ww / 2 + 4, -12, ww - 8, 24); ctx.clip(); text(ctx, b.t.name, -ww / 2 + 6, 4, b.t.od ? "#fff" : onAcc(TH.acc)); ctx.restore(); }
            ctx.restore();
            if (hx >= bx && hx <= bx + ww && hy >= yy && hy <= yy + 24) near = { b, x: hx, y: hy };
          });
          if (spilled > 0) { FT(ctx, 700, 11.5); text(ctx, `Doesn't fit: ${Math.round(spilled)} working days`, ox + cw - 14, floor + 16, C.red, "right"); }
        });
        if (near) tipBox(ctx, near.x, near.y, w, h, near.b.t.name, `${Math.round(near.b.d * 10) / 10} working days in the next 4 weeks · due ${F(near.b.t.finish)}${near.b.t.od ? " · overdue" : ""}`);
      };
    },
    // the radar: distance is time until a risk hits, sectors are workstreams, a sweep leaves an afterglow
    radarf(m) {
      const R0 = riskList(m), names = [...new Set(m.phases.map(p => p.name))], ns = Math.max(1, names.length), H = 120, rr = rnd(31);
      const B = R0.map(r => { const s = Math.max(0, names.indexOf(r.phase)); return { ...r, a: -Math.PI / 2 + (s + .15 + rr() * .7) / ns * 6.2832, j: rr() }; });
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, cx = w * .5, cy = h * .5, R = Math.min(w * .44, h * .44), core = R * .13, u = ease3((T - 100) / 1200);
        const rad = d => d <= m.sd ? core * .4 : core + (R - core) * clamp((d - m.sd) / H, 0, 1);
        [[0, "Now"], [30, "30 days"], [60, "60 days"], [90, "90 days"], [H, "Later"]].forEach(([d, lab], k) => { const r = (k ? rad(m.sd + d) : core) * u;
          ctx.beginPath(); ctx.arc(cx, cy, r, 0, 6.2832); ctx.strokeStyle = k === 0 ? rgba(TH.acc, .5) : TH.line; ctx.lineWidth = 1; ctx.stroke(); FT(ctx, 600, 10.5); if (k) text(ctx, lab, cx + 4, cy - r - 4, TH.mute); });
        names.forEach((nm, s) => { const a = -Math.PI / 2 + s / ns * 6.2832; line(ctx, cx + Math.cos(a) * core, cy + Math.sin(a) * core, cx + Math.cos(a) * R * u, cy + Math.sin(a) * R * u, TH.line, 1);
          const am = a + .5 / ns * 6.2832, lx = cx + Math.cos(am) * (R + 18), ly = cy + Math.sin(am) * (R + 18); FT(ctx, 600, 11); text(ctx, nm.length > 18 ? nm.slice(0, 17) + "…" : nm, lx, ly + 4, TH.ink, Math.cos(am) > .2 ? "left" : Math.cos(am) < -.2 ? "right" : "center"); });
        const sw = still ? null : (T / 3800 * 6.2832) % 6.2832 - Math.PI / 2;
        if (sw != null && ctx.createConicGradient) { const g = ctx.createConicGradient(sw - 1.1, cx, cy); g.addColorStop(0, rgba(TH.acc, 0)); g.addColorStop(.175, rgba(TH.acc, TH.light ? .1 : .16)); g.addColorStop(.176, rgba(TH.acc, 0)); g.addColorStop(1, rgba(TH.acc, 0));
          ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, 0, 6.2832); ctx.fillStyle = g; ctx.fill(); line(ctx, cx, cy, cx + Math.cos(sw) * R, cy + Math.sin(sw) * R, rgba(TH.acc, .6), 1.5); }
        let near = null, nd = 14; const hx = hover[0] * w, hy = hover[1] * h;
        B.forEach((b, i) => {
          const r = b.date <= m.sd ? core * (.2 + .6 * b.j) : rad(b.date), x = cx + Math.cos(b.a) * r * u, y = cy + Math.sin(b.a) * r * u, sz = [0, 4, 6, 9][b.sev];
          const since = sw == null ? 0 : ((sw - b.a) % 6.2832 + 6.2832) % 6.2832, al = sw == null ? 1 : .35 + .65 * Math.exp(-since / 1.6);
          if (sw != null && since < .35) { ctx.beginPath(); ctx.arc(x, y, sz + since * 40, 0, 6.2832); ctx.strokeStyle = rgba(b.c, (1 - since / .35) * .7); ctx.lineWidth = 1.2; ctx.stroke(); }
          glow(ctx, x, y, sz * 3, b.c, .5 * al); ctx.beginPath(); ctx.arc(x, y, sz, 0, 6.2832); ctx.fillStyle = rgba(b.c, al); ctx.fill();
          const d = Math.hypot(hx - x, hy - y); if (d < nd) { nd = d; near = { b, x, y }; } });
        ctx.beginPath(); ctx.arc(cx, cy, core * .55, 0, 6.2832); ctx.fillStyle = TH.card; ctx.fill(); FT(ctx, 700, 11); text(ctx, "Today", cx, cy + 4, TH.ink, "center");
        if (near) tipBox(ctx, near.x, near.y, w, h, near.b.name, `${near.b.why} · ${near.b.date <= m.sd ? "already here" : when(near.b.date, m.sd)}`);
      };
    },
    // the approach: remaining work is altitude, the plan is the glide path, the finish date is the runway
    glidef(m) {
      const pf = m.paceFinish, t0 = m.curve.t0, tEnd = Math.max(m.finish ?? m.sd, m.bfinish ?? 0, Math.min(pf ?? 0, (m.finish ?? m.sd) + 400), m.sd) + 21, tol = m.rules.amberBehind / 100;
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, gx = 40, rx = w - 30, top = 40, gy = h - 92, X = d => gx + (d - t0) / (tEnd - t0) * (rx - gx), Y = a => gy - (gy - top) * clamp(a, 0, 1.05);
        // ground, runway from the planned finish, and its lights
        const gg = ctx.createLinearGradient(0, gy, 0, h); gg.addColorStop(0, rgba(TH.ink, TH.light ? .06 : .1)); gg.addColorStop(1, rgba(TH.ink, 0)); ctx.fillStyle = gg; ctx.fillRect(0, gy, w, h - gy);
        line(ctx, 0, gy, w, gy, TH.line, 1);
        if (m.finish != null) { const a = X(m.finish), b = Math.min(rx, a + Math.max(160, (rx - a) * .9));
          ctx.beginPath(); ctx.moveTo(a - 10, gy + 4); ctx.lineTo(b + 30, gy + 4); ctx.lineTo(b + 60, gy + 30); ctx.lineTo(a - 40, gy + 30); ctx.closePath(); ctx.fillStyle = rgba(TH.ink, TH.light ? .12 : .2); ctx.fill();
          ctx.setLineDash([14, 10]); line(ctx, a, gy + 17, b + 40, gy + 17, rgba(TH.light ? "#ffffff" : "#e2e8f0", .7), 2); ctx.setLineDash([]);
          ctx.fillStyle = rgba(TH.light ? "#ffffff" : "#e2e8f0", .8); for (let k = 0; k < 4; k++) ctx.fillRect(a - 8 + k * 6, gy + 8, 3, 18);
          const chase = (T / 90) % 12; for (let k = 0; k < 12; k++) { const x = a - 14 - k * 14, on = still ? .5 : Math.max(0, 1 - Math.abs(k - chase) / 1.5); ctx.fillStyle = rgba(C.amber, .25 + .75 * on); ctx.beginPath(); ctx.arc(x, gy + 17, 2.2, 0, 6.2832); ctx.fill(); }
          FT(ctx, 700, 11.5); text(ctx, `Planned touchdown · ${F(m.finish)}`, a, gy + 48, TH.ink, "center"); }
        if (m.bfinish != null && m.bfinish !== m.finish) { const x = X(m.bfinish); ctx.setLineDash([3, 4]); line(ctx, x, gy - 30, x, gy + 4, TH.mute, 1); ctx.setLineDash([]); FT(ctx, 600, 10.5); text(ctx, "Baseline", x, gy - 34, TH.mute, "center"); }
        // the beam: the planned approach with its tolerance
        const pl = m.curve.plan.filter(([d]) => d <= tEnd), up = pl.map(([d, v]) => [X(d), Y(1 - v + tol)]), dn = pl.map(([d, v]) => [X(d), Y(Math.max(0, 1 - v - tol))]).reverse();
        ctx.beginPath(); [...up, ...dn].forEach(([x, y], i) => i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)); ctx.closePath(); ctx.fillStyle = rgba(TH.acc, TH.light ? .08 : .1); ctx.fill();
        ctx.setLineDash([6, 6]); ctx.beginPath(); pl.forEach(([d, v], i) => i ? ctx.lineTo(X(d), Y(1 - v)) : ctx.moveTo(X(d), Y(1 - v))); ctx.strokeStyle = rgba(TH.acc, .8); ctx.lineWidth = 1.5; ctx.stroke(); ctx.setLineDash([]);
        FT(ctx, 600, 11); text(ctx, "Planned approach", X(t0) + 8, Y(1) - 10, TH.acc);
        // the flight so far, drawn in, then the line to touchdown at the current pace
        const act = m.curve.actual, u = ease3((T - 200) / 1800), k = u * (act.length - 1), seg = act.slice(0, Math.floor(k) + 1).map(([d, v]) => [X(d), Y(1 - v)]);
        if (act.length > 1 && k < act.length - 1) { const [a, b] = [act[Math.floor(k)], act[Math.ceil(k)]], f = k - Math.floor(k); seg.push([X(a[0] + (b[0] - a[0]) * f), Y(1 - (a[1] + (b[1] - a[1]) * f))]); }
        [-5, 5].forEach(o => { ctx.beginPath(); seg.forEach(([x, y], i) => i ? ctx.lineTo(x, y + o) : ctx.moveTo(x, y + o)); ctx.strokeStyle = rgba(TH.ink, .12); ctx.lineWidth = 1; ctx.stroke(); });
        ctx.beginPath(); seg.forEach(([x, y], i) => i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)); ctx.strokeStyle = TH.ink; ctx.lineWidth = 2.5; ctx.stroke();
        const [px, py] = seg[seg.length - 1] || [X(m.sd), Y(1 - m.pct / 100)];
        let ang = .25;
        if (pf != null && u >= 1) { const tx = X(Math.min(pf, tEnd)), late = pf > (m.finish ?? pf), p2 = clamp((T - 2000) / 900, 0, 1); ang = Math.atan2(gy - py, tx - px);
          ctx.setLineDash([7, 6]); line(ctx, px, py, px + (tx - px) * p2, py + (gy - py) * p2, late ? C.amber : C.green, 2); ctx.setLineDash([]);
          if (p2 >= 1) { ctx.beginPath(); ctx.arc(tx, gy, 6, 0, 6.2832); ctx.fillStyle = late ? C.amber : C.green; ctx.fill(); FT(ctx, 700, 11.5); text(ctx, `At current pace · ${F(pf)}`, tx, gy - 14, late ? C.amber : C.green, tx > rx - 90 ? "right" : "center");
            if (late && m.finish != null) { const a = X(m.finish); line(ctx, a, gy + 62, tx, gy + 62, C.amber, 2); line(ctx, a, gy + 56, a, gy + 68, C.amber, 2); line(ctx, tx, gy + 56, tx, gy + 68, C.amber, 2); FT(ctx, 700, 11); text(ctx, `${days(pf - m.finish)} long`, (a + tx) / 2, gy + 80, C.amber, "center"); } } }
        // the plane
        const bob = still ? 0 : Math.sin(T / 600) * 1.5; ctx.save(); ctx.translate(px, py + bob); ctx.rotate(clamp(ang, -.6, .6));
        ctx.fillStyle = TH.ink; ctx.beginPath(); ctx.moveTo(16, 0); ctx.quadraticCurveTo(10, -4, -12, -3); ctx.lineTo(-17, -11); ctx.lineTo(-20, -11); ctx.lineTo(-17, -2); ctx.lineTo(-18, 3); ctx.lineTo(-12, 3); ctx.quadraticCurveTo(10, 4, 16, 0); ctx.fill();
        ctx.beginPath(); ctx.moveTo(2, 1); ctx.lineTo(-8, 12); ctx.lineTo(-4, 12); ctx.lineTo(8, 1); ctx.closePath(); ctx.fill();
        const blink = still || (T % 1200) < 140; if (blink) { glow(ctx, 16, 0, 10, TH.acc, .9); ctx.fillStyle = TH.acc; ctx.beginPath(); ctx.arc(-19, -11, 2, 0, 6.2832); ctx.fill(); } ctx.restore();
        FT(ctx, 700, 11.5); text(ctx, `Today · ${m.pct}% done`, px, py - 22, TH.ink, "center");
        // months
        const [y0, m0] = ymd(t0); FT(ctx, 500, 11); let lx = -99;
        for (let i = 0, y = y0, mo = m0; i < 80; i++, mo++) { if (mo > 12) { mo = 1; y++; } const d = E.fromYMD(y, mo, 1); if (d > tEnd) break; if (d < t0) continue; const x = X(d); if (x - lx < 40) continue; lx = x; text(ctx, MON[mo - 1] + (mo === 1 ? " " + String(y).slice(2) : ""), x, h - 8, TH.mute, "center"); }
        // the pointer reads the approach at any date
        if (hover[0] >= 0) { const hx = hover[0] * w; if (hx > gx && hx < rx) { const d = t0 + (hx - gx) / (rx - gx) * (tEnd - t0); line(ctx, hx, top, hx, gy, rgba(TH.ink, .2), 1);
          let a = null; for (const q of act) if (q[0] <= d) a = q; tipBox(ctx, hx, hover[1] * h, w, h, F(Math.round(d)), `Plan ${Math.round(planAt(m, d) * 100)}% done${a && d <= m.sd ? ` · actual ${Math.round(a[1] * 100)}%` : ""}`); } }
      };
    },
    // milestones as dominoes: what's achieved has fallen in a chain, the next one wobbles
    dominof(m) {
      const L = msList(m).slice(0, 16), nextI = L.findIndex(t => !t.done);
      return ({ ctx, w, h, t, hover, still }) => {
        const T = still ? 1e9 : t, gx = 40, rx = w - 70, gy = h * .52, DH = Math.min(118, h * .3), DW = 20;
        const step = (rx - gx) / L.length, xs = L.map((_, i) => gx + (i + .5) * step);
        // today sits between the milestones it falls between, by date
        const k = L.findIndex(t2 => t2.finish > m.sd), X = () => k < 0 ? xs[xs.length - 1] + step * .5 : k === 0 ? xs[0] - step * .5
          : xs[k - 1] + (xs[k] - xs[k - 1]) * (m.sd - L[k - 1].finish) / ((L[k].finish - L[k - 1].finish) || 1);
        line(ctx, 0, gy, w, gy, TH.line, 1.5);
        const tx = X(); ctx.setLineDash([3, 4]); line(ctx, tx, gy - DH - 46, tx, gy + 6, TH.acc, 1.5); ctx.setLineDash([]); FT(ctx, 700, 11); text(ctx, `Today · ${F(m.sd)}`, tx, gy - DH - 52, TH.acc, "center");
        let near = null; const hx = hover[0] * w, hy = hover[1] * h, nDone = L.filter(x => x.done).length;
        L.forEach((t2, i) => {
          const x = xs[i], col = t2.done ? C.green : t2.od ? C.red : t2.var > 0 ? C.amber : i === nextI ? TH.acc : TH.mute, nextGap = i < L.length - 1 ? xs[i + 1] - x - DW : 999;
          const rest = Math.min(1.25, Math.asin(clamp(nextGap / DH, 0, 1)) || 1.25), at = 500 + i * 230, f = clamp((T - at) / 420, 0, 1);
          let ang = 0; if (t2.done) ang = rest * f * f; else if (t2.od) ang = .2 + Math.sin(T / 300) * .02 * (still ? 0 : 1); else if (i === nextI) { const go = still ? 1 : clamp((T - 500 - nDone * 230 - 300) / 400, 0, 1); ang = go * Math.sin(T / 260) * .05; }
          if (t2.var > 0 && !t2.done) { FT(ctx, 700, 11); const lab = `+${t2.var} d`, bw = ctx.measureText(lab).width + 12; ctx.fillStyle = rgba(C.amber, .18); ctx.beginPath(); ctx.roundRect(x - bw / 2, gy - DH - 26, bw, 18, 9); ctx.fill(); text(ctx, lab, x, gy - DH - 13, C.amber, "center"); }
          ctx.save(); ctx.translate(x + DW / 2, gy); ctx.rotate(ang);
          ctx.fillStyle = t2.done || t2.od || i === nextI || t2.var > 0 ? col : TH.card; ctx.strokeStyle = col; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.roundRect(-DW, -DH, DW, DH, 4); ctx.fill(); ctx.stroke();
          line(ctx, -DW + 4, -DH / 2, -4, -DH / 2, rgba(t2.done || t2.od || i === nextI || t2.var > 0 ? onAcc(col) : TH.ink, .5), 1.5);
          ctx.fillStyle = rgba(t2.done || t2.od || i === nextI || t2.var > 0 ? onAcc(col) : TH.ink, .6); [[-DW / 2, -DH * .75], [-DW / 2, -DH * .25]].forEach(([a, b]) => { ctx.beginPath(); ctx.arc(a, b, 2.4, 0, 6.2832); ctx.fill(); });
          ctx.restore();
          FT(ctx, 700, 11.5); const mx = Math.floor(step / 6.5), nm = t2.name.length > mx ? t2.name.slice(0, mx - 1) + "…" : t2.name, ly = gy + 26;
          text(ctx, nm, x, ly, TH.ink, "center"); FT(ctx, 500, 11); text(ctx, t2.done ? `Done · ${F(t2.afin ?? t2.finish)}` : t2.od ? `Overdue · ${F(t2.finish)}` : F(t2.finish), x, ly + 15, t2.od ? C.red : t2.var > 0 ? C.amber : TH.mute, "center");
          if (hx > x - DW && hx < x + DW && hy > gy - DH - 10 && hy < gy + 60) near = { t2, x, y: gy - DH / 2 };
        });
        if (near) { const t2 = near.t2; tipBox(ctx, near.x, near.y, w, h, t2.name, t2.done ? `Achieved · ${F(t2.afin ?? t2.finish)}` : t2.od ? `${plural(t2.late || 0, "working day")} overdue` : t2.var > 0 ? `${F(t2.finish)} · ${plural(t2.var, "working day")} later than baseline` : `Due ${F(t2.finish)}`); }
      };
    },
    // the gap to plan over the saved versions, drawn by a seismograph needle
    seismof(m) {
      const G = m.trend.map(([d, a]) => [d, planAt(m, d) * 100 - a * 100]), A = m.rules.amberBehind, Rd = m.rules.redBehind, top = Math.max(Rd * 1.4, ...G.map(g => g[1]), 5);
      return ({ ctx, w, h, t, still }) => {
        const T = still ? 1e9 : t, Y = v => h - 18 - (h - 30) * clamp(v, -top * .2, top) / top * .9 - (h - 30) * .1, n = G.length, X = i => n < 2 ? w * .82 : 14 + i / (n - 1) * (w * .82 - 14);
        [[Rd, top, C.red], [A, Rd, C.amber]].forEach(([a, b, c]) => { ctx.fillStyle = rgba(c, TH.light ? .07 : .1); ctx.fillRect(0, Y(b), w, Y(a) - Y(b)); });
        for (let x = (-(T / 40) % 16); x < w; x += 16) line(ctx, x, 0, x, h - 18, rgba(TH.ink, .05), 1);
        FT(ctx, 600, 10); text(ctx, "Behind plan, pts", 4, 11, TH.mute);
        const u = ease3((T - 300) / 1500), k = u * (n - 1), pts = G.slice(0, Math.floor(k) + 1).map(([, v], i) => [X(i), Y(v)]);
        if (n > 1 && k < n - 1) { const i = Math.floor(k), f = k - i; pts.push([X(i) + (X(i + 1) - X(i)) * f, Y(G[i][1] + (G[i + 1][1] - G[i][1]) * f)]); }
        const tremor = still ? 0 : Math.sin(T / 45) * 1.2 + Math.sin(T / 17) * .6, [ex, ey] = pts[pts.length - 1];
        ctx.beginPath(); pts.forEach(([x, y], i) => i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)); ctx.lineTo(ex, ey + tremor); ctx.strokeStyle = TH.ink; ctx.lineWidth = 2; ctx.stroke();
        line(ctx, ex, ey + tremor, w - 4, ey + tremor - 10, rgba(TH.ink, .55), 1.2); ctx.beginPath(); ctx.arc(ex, ey + tremor, 3.5, 0, 6.2832); ctx.fillStyle = TH.acc; ctx.fill();
        G.forEach(([d], i) => { if (i === 0 || i === n - 1) { FT(ctx, 500, 10); text(ctx, F(d), X(i), h - 4, TH.mute, i ? "right" : "left"); } });
      };
    },
    // the forecast's possible finish dates, falling into a pile
    sandf(m) {
      const fc = m.fc; if (!fc || !fc.ok) return () => {};
      const lo = Math.min(fc.dist[0], m.finish ?? fc.dist[0]) - 5, hi = Math.max(fc.dist[fc.dist.length - 1], m.finish ?? 0) + 5;
      return ({ ctx, w, h, t, still }) => {
        const T = still ? 1e9 : t, X = d => 10 + (d - lo) / (hi - lo) * (w - 20), base = h - 20, bw = 5.5, cnt = {};
        line(ctx, 0, base + 2, w, base + 2, TH.line, 1);
        fc.dist.forEach((d, i) => { const k = (i * 7919) % fc.dist.length, bin = Math.round(X(d) / bw); cnt[bin] = (cnt[bin] || 0) + 1; const y = base - (cnt[bin] - .5) * 5, u = (T - 300 - k * 10) / 500; if (u <= 0) return;
          ctx.fillStyle = d <= fc.p50 ? TH.acc : d <= fc.p80 ? rgba(TH.acc, .55) : rgba(TH.mute, .6); ctx.fillRect(bin * bw - 2.2, (u >= 1 ? y : 6 + (y - 6) * u * u) - 2.2, 4.4, 4.4); });
        const vl = (d, c, lab, y, dash) => { const x = X(d); ctx.setLineDash(dash); line(ctx, x, 14, x, base + 2, c, 1.2); ctx.setLineDash([]); FT(ctx, 700, 10.5); text(ctx, lab, clamp(x, 30, w - 30), y, c, "center"); };
        if (m.finish != null) vl(m.finish, TH.mute, "Plan", 10, [3, 3]);
        if (T > 1500) { vl(fc.p50, TH.acc, "P50", h - 6, []); vl(fc.p80, TH.light ? C.violet : "#c4b5fd", "P80", 10, []); }
      };
    },
    nowfield(m) {
      const D = nowData(m), fc = D.fc, sd = m.sd, n = D.lanes.length;
      const b = D.lanes.map(() => 0), v = D.lanes.map(() => 0);
      let L = null, lastT = 0;
      const layout = (w, h) => {
        const gx = 170, rx = w - 24, top = 34, laneB = fc ? h - 170 : h - 40, rainTop = laneB + 30, base = h - 30, lh = (laneB - top) / n;
        const X = d => gx + (d - D.t0) / (D.t1 - D.t0) * (rx - gx), M = X(sd), maxB = Math.min(70, (rx - M) * .25);
        const rr = rnd(97), P = [];
        D.lanes.forEach((ln, li) => {
          const cy = top + lh * (li + .5), late = [...ln.late].sort((a, c) => (c.late || 0) - (a.late || 0)), per = Math.max(1, Math.floor(lh * .78 / 9));
          ln.tasks.forEach(t => {
            const r = t.ms ? 5.5 : clamp(2 + Math.sqrt(t.w || 1) * .55, 2.4, 6.5), jy = cy + (rr() - .5) * lh * .62;
            const kind = t.od ? "late" : t.done ? "done" : t.var > 0 ? "slip" : "open";
            const p = { t, li, r, kind, ph: rr() * 6.28, x0: gx + rr() * (rx - gx), y0: top + rr() * (laneB - top), cy };
            if (kind === "late") { const k = late.indexOf(t), col = Math.floor(k / per), row = k % per; p.pile = col; p.ty = cy - lh * .39 + 4.5 + row * 9 + (col % 2) * 4.5; p.tx0 = X(t.finish); }
            else { p.ty = jy; p.tx = kind === "done" ? Math.min(X(t.finish), M - r - 3) : Math.max(X(t.finish), M + r + 3); if (kind === "slip" && t.bfin != null) p.bx = X(t.bfin); }
            p.delay = (kind === "done" ? 0 : kind === "late" ? 900 : 350) + rr() * 500;
            P.push(p);
          });
        });
        let G = [];
        if (fc) {
          const bw = 6.5, cnt = {}, order = fc.dist.map((d, i) => i).sort((a, c) => ((a * 7919) % 160) - ((c * 7919) % 160));
          G = order.map((i, k) => { const x = X(fc.dist[i]), bin = Math.round(x / bw); cnt[bin] = (cnt[bin] || 0) + 1;
            return { x: bin * bw, y: base - (cnt[bin] - .5) * 6, d: fc.dist[i], at: 1500 + k * 14 }; });
          const tall = Math.max(...Object.values(cnt)) * 6, room = base - rainTop - 52; if (tall > room) G.forEach(g => { g.y = base - (base - g.y) * room / tall; });
        }
        return { w, h, gx, rx, top, laneB, rainTop, base, lh, X, M, maxB, P, G };
      };
      // the Today line's x at height y: lane bulges eased between lane centres, pinned at both ends
      const bend = (y) => { const { top, lh, laneB } = L; if (y <= top || y >= laneB) return 0; const f = (y - top) / lh - .5;
        if (f <= 0) return b[0] * (.5 - .5 * Math.cos(Math.PI * clamp((y - top) / (lh * .5), 0, 1)));
        if (f >= n - 1) return b[n - 1] * (.5 - .5 * Math.cos(Math.PI * clamp((laneB - y) / (lh * .5), 0, 1)));
        const i = Math.floor(f), u = f - i, s = .5 - .5 * Math.cos(Math.PI * u); return b[i] * (1 - s) + b[i + 1] * s; };
      return ({ ctx, form, w, h, t, hover, still }) => {
        if (!L || L.w !== w || L.h !== h) L = layout(w, h);
        const { gx, rx, top, laneB, rainTop, base, lh, X, M, maxB, P, G } = L, font = getComputedStyle(ctx.canvas).fontFamily, T = still ? 1e9 : t;
        const hx = hover[0] >= 0 ? hover[0] * w : -999, hy = hover[1] >= 0 ? hover[1] * h : -999;
        // the membrane: springs towards each lane's pressure once the late work lands; the pointer can press it too
        const dt = still ? 0 : clamp((t - lastT) / 16.7, .3, 3); lastT = t;
        D.lanes.forEach((ln, i) => {
          const cy = top + lh * (i + .5), poke = Math.abs(hx - M - b[i]) < 46 && Math.abs(hy - cy) < lh * .6 ? (46 - Math.abs(hx - M - b[i])) * .5 : 0;
          const goal = ln.press * maxB * (T > 1500 ? 1 : 0) + poke;
          if (still) b[i] = goal; else { v[i] += (goal - b[i]) * .045 * dt; v[i] *= Math.pow(.88, dt); b[i] += v[i] * dt; }
        });
        // lanes, past tint and month ticks
        ctx.font = "600 12px " + font; ctx.textBaseline = "middle";
        ctx.fillStyle = TH.light ? "rgba(15,23,42,.025)" : "rgba(148,163,255,.035)"; ctx.fillRect(gx, top - 10, M - gx, laneB - top + 20);
        D.lanes.forEach((ln, i) => {
          const y = top + lh * i; if (i) form.strokeOnly(TH.line, 1).line([[gx, y], [rx, y]]);
          const nm = ln.name.length > 21 ? ln.name.slice(0, 20) + "…" : ln.name, cy = y + lh / 2;
          ctx.fillStyle = TH.ink; ctx.font = "600 12.5px " + font; ctx.fillText(nm, 14, cy - (ln.pct != null ? 7 : 0));
          if (ln.pct != null) { ctx.fillStyle = TH.mute; ctx.font = "500 11.5px " + font; ctx.fillText(`${ln.pct}% done${ln.late.length ? ` · ${ln.late.length} late` : ""}`, 14, cy + 9); }
        });
        ctx.textBaseline = "alphabetic";
        { const [y0, m0] = ymd(D.t0), months = Math.ceil((D.t1 - D.t0) / 30.4) + 1, step = Math.max(1, Math.ceil(months / 14)); ctx.fillStyle = TH.mute; ctx.font = "500 11px " + font;
          for (let i = 0, y = y0, mo = m0; i <= months; i++, mo++) { if (mo > 12) { mo = 1; y++; } const d = E.fromYMD(y, mo, 1), x = X(d); if (d < D.t0 || d > D.t1 || i % step) continue;
            form.strokeOnly(TH.line, 1).line([[x, laneB + 6], [x, laneB + 12]]); ctx.fillText(MON[mo - 1] + (mo === 1 ? " " + String(y).slice(2) : ""), x + 3, laneB + 20); } }
        // pressure: the area the Today line has been pushed into, then the line, echoed
        const ys = []; for (let y = top - 14; y <= laneB + 14; y += 3) ys.push(y);
        const line = ys.map(y => [M + bend(y), y]), hot = Math.max(...b) / (maxB || 1);
        if (hot > .02) { ctx.beginPath(); ctx.moveTo(M, top - 14); line.forEach(([x, y]) => ctx.lineTo(x, y)); ctx.lineTo(M, laneB + 14); ctx.closePath(); ctx.fillStyle = rgba(D.late.length ? C.red : C.amber, TH.light ? .09 : .13); ctx.fill(); }
        for (let k = 3; k >= 1; k--) form.strokeOnly(rgba(TH.acc, .12 / k + (TH.light ? .04 : .02)), 1).line(ys.map(y => [M + bend(y) * (1 + k * .22) + k * 7, y]));
        const g = ctx.createLinearGradient(0, top, 0, laneB);
        D.lanes.forEach((ln, i) => g.addColorStop((i + .5) / n, ln.late.length && ln.press > .2 ? C.red : ln.press > .2 ? C.amber : TH.acc));
        if (TH.glow) form.strokeOnly(rgba(TH.acc, .18 * TH.glow), 9).line(line);
        form.strokeOnly(g, 2.4).line(line);
        ctx.font = "700 11px " + font; const lab = `Today · ${F(sd)}`, lw = ctx.measureText(lab).width + 16;
        ctx.fillStyle = TH.acc; ctx.beginPath(); ctx.roundRect(M - lw / 2, top - 32, lw, 18, 9); ctx.fill(); ctx.fillStyle = onAcc(TH.acc); ctx.fillText(lab, M - lw / 2 + 8, top - 19);
        // plan finish and the forecast markers, through the whole field
        const vline = (x, col, dash, txt, y, left) => { ctx.setLineDash(dash); form.strokeOnly(rgba(col, .7), 1.2).line([[x, top - 8], [x, base]]); ctx.setLineDash([]);
          ctx.font = "700 11px " + font; ctx.fillStyle = col; const tw = ctx.measureText(txt).width; ctx.fillText(txt, left ? Math.max(gx, x - tw - 5) : Math.min(x + 5, rx - tw), y); };
        if (m.finish != null) vline(X(m.finish), TH.mute, [4, 4], `Plan finish · ${F(m.finish)}`, fc ? rainTop + 14 : laneB + 34, !!fc);
        // tasks: assemble from a scatter, done settles behind today, late is held against it, slipping trails back to baseline
        let near = null, nd = 14;
        P.forEach(p => {
          const u = ease3((T - p.delay) / 1400); let x, y;
          if (p.kind === "late") { const mx = M + b[p.li] - p.r - 3 - p.pile * 8; x = p.x0 + (mx - p.x0) * u + Math.sin(T / 70 + p.ph) * .7 * u; y = p.y0 + (p.ty - p.y0) * u; }
          else { x = p.x0 + (p.tx - p.x0) * u; y = p.y0 + (p.ty - p.y0) * u + (p.kind === "done" ? 0 : Math.sin(T / 1400 + p.ph) * 1.4 * u); }
          const dx = x - hx, dy = y - hy, dd = Math.hypot(dx, dy); if (dd < 40 && dd > .1) { x += dx / dd * (40 - dd) * .18; y += dy / dd * (40 - dd) * .18; }
          if (dd < nd) { nd = dd; near = { p, x, y }; }
          const col = p.kind === "late" ? C.red : p.kind === "slip" ? C.amber : p.kind === "done" ? TH.mute : TH.acc, a = .25 + .75 * u;
          if (p.kind === "slip" && p.bx != null && u > .6) { const tg = ctx.createLinearGradient(p.bx, 0, x, 0); tg.addColorStop(0, rgba(C.amber, 0)); tg.addColorStop(1, rgba(C.amber, .55 * u)); form.strokeOnly(tg, Math.max(1, p.r * .7)).line([[p.bx, y], [x, y]]); }
          if (p.kind === "late" && u > .6) { ctx.setLineDash([2, 4]); form.strokeOnly(rgba(C.red, .35), 1).line([[p.tx0, y], [x, y]]); ctx.setLineDash([]); }
          if (p.kind === "late") glow(ctx, x, y, p.r * 3.2, C.red, .35);
          if (p.t.ms) { diamond(ctx, x, y, p.r); if (p.kind === "open") { ctx.strokeStyle = rgba(col, a); ctx.lineWidth = 1.6; ctx.stroke(); } else { ctx.fillStyle = rgba(p.kind === "done" ? TH.ink : col, a); ctx.fill(); } return; }
          if (p.kind === "open") form.strokeOnly(rgba(col, a), 1.4).point([x, y], p.r, "circle"); else form.fillOnly(rgba(col, p.kind === "done" ? a * .55 : a)).point([x, y], p.r, "circle");
        });
        // the forecast: possible finish dates fall and pile up; the pile's shape is the uncertainty
        if (fc) {
          ctx.font = "700 12px " + font; ctx.fillStyle = TH.ink; ctx.fillText("When it finishes", gx - 156, rainTop + 6);
          ctx.font = "500 11.5px " + font; ctx.fillStyle = TH.mute; ctx.fillText(`${fc.dist.length} simulated finishes`, gx - 156, rainTop + 22);
          G.forEach(gr => { const u = (T - gr.at) / 650; if (u <= 0) return; const y = u >= 1 ? gr.y : rainTop + (gr.y - rainTop) * u * u;
            const col = gr.d <= fc.p50 ? TH.acc : gr.d <= fc.p80 ? rgba(TH.acc, .6) : rgba(TH.light ? "#64748b" : "#94a3b8", .7);
            ctx.globalAlpha = u >= 1 ? .78 + .22 * Math.sin(T / 650 - gr.x / 36) : 1; ctx.fillStyle = col; ctx.fillRect(gr.x - 2.6, y - 2.6, 5.2, 5.2); ctx.globalAlpha = 1; });
          form.strokeOnly(TH.line, 1).line([[gx, base + 2], [rx, base + 2]]);
          const shown = clamp((T - 1500) / (fc.dist.length * 14 + 650), 0, 1);
          if (shown > .5) { ctx.globalAlpha = clamp((shown - .5) * 3, 0, 1);
            vline(X(fc.p50), TH.acc, [], `P50 · 50/50 date · ${F(fc.p50)}`, rainTop + 14, X(fc.p50) > rx - 300 && X(fc.p50) - X(m.finish ?? 0) > 180);
            vline(X(fc.p80), TH.light ? C.violet : "#c4b5fd", [], `P80 · 8 in 10 · ${F(fc.p80)}`, rainTop + 30); ctx.globalAlpha = 1; }
        }
        // what the pointer is on
        if (near) { const { p, x, y } = near, t2 = p.t;
          const st = p.kind === "late" ? `Late by ${plural(t2.late || 0, "day")}, due ${F(t2.finish)}` : p.kind === "done" ? `Done · ${F(t2.finish)}` : p.kind === "slip" ? `Due ${F(t2.finish)} · ${plural(t2.var, "working day")} later than baseline` : `Due ${F(t2.finish)}`;
          const nm = t2.name.length > 44 ? t2.name.slice(0, 43) + "…" : t2.name; ctx.font = "700 12px " + font; const tw = Math.max(ctx.measureText(nm).width, (ctx.font = "500 11.5px " + font, ctx.measureText(st).width)) + 22;
          const bx = x + tw + 16 > w ? x - tw - 12 : x + 12, by = clamp(y - 44, 4, h - 52);
          ctx.fillStyle = TH.card; ctx.strokeStyle = rgba(TH.acc, .5); ctx.lineWidth = 1; ctx.beginPath(); ctx.roundRect(bx, by, tw, 46, 9); ctx.fill(); ctx.stroke();
          ctx.fillStyle = TH.ink; ctx.font = "700 12px " + font; ctx.fillText(nm, bx + 11, by + 19); ctx.fillStyle = TH.mute; ctx.font = "500 11.5px " + font; ctx.fillText(st, bx + 11, by + 35);
          form.strokeOnly(rgba(TH.acc, .9), 1.5).point([x, y], p.r + 4, "circle"); }
      };
    },
  };

  function scene(Pts, host, scr, draw, still) {
    const space = new Pts.CanvasSpace(host).setup({ bgcolor: "transparent", resize: true, retina: true });
    const form = space.getForm(), ptr = [-1, -1];
    const mv = e => { const r = host.getBoundingClientRect(); ptr[0] = (e.clientX - r.left) / r.width; ptr[1] = (e.clientY - r.top) / r.height; };
    const out = () => { ptr[0] = ptr[1] = -1; };
    scr.addEventListener("pointermove", mv); scr.addEventListener("pointerleave", out);
    const frame = t => { const w = space.width, h = space.height; if (w > 0 && h > 0) draw({ ctx: space.ctx, form, w, h, t: still ? 6000 : t, still, ptr: ptr[0] < 0 ? [.5, .5] : ptr, hover: ptr, scr, host }); };
    space.add({ animate: frame, resize: () => { if (still) space.playOnce(120); } });
    if (still) space.playOnce(300); else space.play();
    return () => { scr.removeEventListener("pointermove", mv); scr.removeEventListener("pointerleave", out); space.dispose(); };
  }
  // at a glance: numbers count up, and a soft highlight follows the pointer across the glass cards
  function glanceLive(scr, still) {
    if (!still) scr.querySelectorAll(".gl-num").forEach(el => { const to = +el.dataset.to, suf = el.dataset.suf || ""; if (!Number.isFinite(to) || !to) return;
      const t0 = performance.now(), tick = now => { const u = Math.min(1, (now - t0) / 1100); el.textContent = Math.round(to * (1 - (1 - u) ** 4)) + suf; if (u < 1) requestAnimationFrame(tick); }; el.textContent = "0" + suf; requestAnimationFrame(tick); });
    scr.querySelectorAll(".gl-card").forEach(c => c.addEventListener("pointermove", e => { const r = c.getBoundingClientRect(); c.style.setProperty("--mx", `${e.clientX - r.left}px`); c.style.setProperty("--my", `${e.clientY - r.top}px`); }));
  }
  // split-flap figures: each character flips through the board before it settles
  function flaps(scr, still, stops) {
    if (still) return; const CH = "0123456789%+−/ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    scr.querySelectorAll(".sx-flap").forEach((el, k) => { const to = el.dataset.to; let n = 0;
      const id = setInterval(() => { n++; el.textContent = [...to].map((c, i) => n > 6 + i * 2 + k * 3 || c === " " ? c : CH[(Math.random() * CH.length) | 0]).join(""); if (n > 8 + to.length * 2 + k * 3) { el.textContent = to; clearInterval(id); } }, 55);
      stops.push(() => clearInterval(id)); });
  }


  function mount(el, o) {
    const id = SCREENS.some(s => s[0] === o.screen) ? o.screen : "glance";
    const m = model(o.k, o), still = !!(root.matchMedia && root.matchMedia("(prefers-reduced-motion: reduce)").matches);
    const range = m.start != null ? `${F(m.start)} – ${F(m.finish)}` : F(m.finish);
    let saved = (() => { try { return JSON.parse(localStorage.getItem(LOOK_KEY)) || {}; } catch (e) { return {}; } })(), stops = [], gen = 0;
    const opt = (name, v, label, L, extra = "") => `<label class="st-opt"${extra}><input type="radio" name="${name}" value="${v}"${L[name] === v || (name === "accent" && L.accent === v) ? " checked" : ""}><span>${label}</span></label>`;
    const panel = L => `<details class="st-look"><summary class="st-btn">${ic('<circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 0 0 18z" fill="currentColor"/>')}Appearance</summary>
      <div class="st-look-pop" role="group" aria-label="Appearance">
        <fieldset><legend>Mode</legend>${opt("mode", "dark", "Dark", L)}${opt("mode", "light", "Light", L)}</fieldset>
        <fieldset><legend>Finish</legend>${opt("finish", "flat", "Flat", L)}${opt("finish", "matte", "Matte", L)}${opt("finish", "gloss", "Gloss", L)}</fieldset>
        <fieldset><legend>Motion</legend>${opt("motion", "off", "Off", L)}${opt("motion", "subtle", "Subtle", L)}${opt("motion", "full", "Full", L)}</fieldset>
        <fieldset class="st-swatches"><legend>Colour</legend>${ACCENTS.map(([c, n]) => opt("accent", c, `<i style="--c:${c}"></i><b class="st-sr">${n}</b>`, L, ` title="${n}"`)).join("")}
          ${HEX.test(o.brand?.color || "") ? opt("accent", "brand", `<i style="--c:${o.brand.color}"></i><b>Brand</b>`, L, ' title="Your brand colour, from Settings"') : ""}
          <label class="st-opt st-custom" title="Any colour"><input type="color" value="${L.acc}" aria-label="Custom colour"></label></fieldset>
      </div></details>`;
    function paint(reopen) {
      stops.forEach(f => f()); stops = [];
      const L = lookOf(saved, o.brand); theme(L); TH.acc = L.acc;
      const my = ++gen, ex = id === "executive", light = L.mode === "light", calm = still || L.motion === "off", once = calm || L.motion === "subtle";
      const logo = L.brand && /^data:image\//.test(o.brand?.logo || "") ? `<img class="st-logo" src="${esc(o.brand.logo)}" alt="Client logo">` : "";
      el.innerHTML = `<div class="st-scroll"><div class="st-screen ${light ? "st-light" : "st-dark"} st-${id}" data-screen="${id}" data-finish="${L.finish}" data-motion="${calm ? "off" : L.motion}" style="--acc:${L.acc};--on-acc:${onAcc(L.acc)}">
        <div class="st-bg" data-scene="${id}" aria-hidden="true"></div>
        <header class="st-bar"><span class="st-brand"><svg aria-hidden="true"><use href="#k-mark"/></svg>Keelson</span>${logo}
          <nav class="st-tabs" aria-label="Dashboard screens">${SCREENS.map(([k, n]) => `<a href="#dash/${k}"${k === id ? ' aria-current="page"' : ""}>${n}</a>`).join("")}</nav>
          ${ex ? "" : `<span class="st-range">${ic(I.cal)}${esc(range)}</span>`}${panel(L)}
          ${ex ? `<button class="st-btn" type="button" data-st-present>${ic('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>')}Present</button>` : ""}
          <button class="st-btn" type="button" data-st-export>${ic(I.dl)}Export</button></header>
        <div class="st-body">${S[id](m)}</div></div></div>`;
      const scr = el.querySelector(".st-screen");
      el.querySelector("[data-st-export]").addEventListener("click", () => { // the screen alone, landscape, colours kept
        const pg = document.createElement("style"); pg.textContent = "@page { size: landscape; margin: 0 }"; document.head.appendChild(pg);
        document.body.classList.add("st-print"); root.print(); document.body.classList.remove("st-print"); pg.remove(); });
      el.querySelector("[data-st-present]")?.addEventListener("click", () => document.fullscreenElement ? document.exitFullscreen() : scr.requestFullscreen?.().catch(() => {}));
      const look = el.querySelector(".st-look");
      if (reopen) { look.open = true; look.querySelector(`input[name="${reopen}"]:checked, input[type="color"]`)?.focus(); }
      look.addEventListener("change", e => {
        const t = e.target, k = t.type === "color" ? "accent" : t.name; saved = { ...saved, [k]: t.value };
        try { localStorage.setItem(LOOK_KEY, JSON.stringify(saved)); } catch (err) {}
        paint(t.type === "color" ? "color" : k);
      });
      document.addEventListener("click", function off(e) { if (!el.contains(look)) return document.removeEventListener("click", off); if (!look.contains(e.target)) look.open = false; });
      flaps(scr, calm, stops); glanceLive(scr, calm);
      load().then(P => { if (my !== gen) return; scr.classList.add("st-live");
        scr.querySelectorAll("[data-scene]").forEach(h => { const make = SCENES[h.dataset.scene]; if (make) stops.push(scene(P, h, scr, make(m), once)); }); }, () => {}); // no motion is fine; the figures are all in the HTML
    }
    paint();
    return { destroy() { gen++; stops.forEach(f => f()); el.innerHTML = ""; }, model: m };
  }

  root.KeelsonStudio = { load, mount, model, SCREENS };
  if (typeof module !== "undefined") module.exports = root.KeelsonStudio;
})(typeof window !== "undefined" ? window : globalThis);
