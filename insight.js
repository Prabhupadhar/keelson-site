/* Keelson plan checks: four questions a PM answers by hand, from the engine's result (k).
     quality   14 checks on how the plan is built, after the DCMA 14-point assessment (limits are a starting point, not a verdict)
     drivers   the critical and near-critical tasks: what is driving the finish date
     workload  open tasks per person per week for the coming weeks
     slips     tasks whose finish moved later since the last version, and a log of the reason for each
   Links, float, the critical flag and constraints come from extra columns of the plan: MS Project .xml and Primavera .xer
   carry them (plan-read.js), an Excel export has them when the columns are included. A check without its column is "not checked".
   Pure: no DOM, no clock, no network. Tests: tests/web/insight.test.js. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const HIGH = 44, NEAR = 5, NONE = "No owner";
  const COLS = { pred: ["Predecessors", "Predecessor", "Depends On", "Dependencies", "Dependency", "Blocked By"], succ: ["Successors", "Successor"],
    float: ["Total Slack", "Total Float", "Float", "Slack"], crit: ["Critical"], cstr: ["Constraint Type", "Primary Constraint", "Constraint"] };
  const HARD = ["must start on", "must finish on", "start no later than", "finish no later than", "start on", "finish on", "mandatory start", "mandatory finish", "start on or before", "finish on or before"];
  const ok = k => !!(k && k.tasks && k.cols.name >= 0 && k.cols.finish >= 0);
  const who = o => E.noOwner(o) ? NONE : String(o).replace(/ /g, " ").replace(/\s+/g, " ").trim();
  const D = v => E.fmt(v, true);
  // extra column by any of its usual names -> a reader of that cell for a task, or null when the plan has no such column
  function reader(k, names) { const want = names.map(E.norm), i = (k.extra || []).findIndex(c => want.includes(E.norm(c.name))); return i < 0 ? null : t => (t.x && t.x[i]) || ""; }
  // "5 days", "-3d", "2 wks", "16 hrs", "1.5" -> working days
  function daysOf(s) { const m = /^\s*(-?[\d.,]+)\s*([a-z]*)/i.exec(String(s || "")); if (!m) return null; const v = parseFloat(m[1].replace(/,/g, "")); if (!isFinite(v)) return null;
    const u = m[2].toLowerCase(); return u.startsWith("h") ? v / 8 : u.startsWith("w") ? v * 5 : u.startsWith("mo") ? v * 20 : v; }
  // "3, 5FS+2 days, A1010SS-1d" -> [{type, lag}]
  function links(s) { return String(s || "").split(/[,;]/).map(x => x.trim()).filter(Boolean).map(x => { const m = /^(.+?)(FS|SS|FF|SF)?\s*(?:([+-])\s*([\d.]+)\s*[a-z%]*)?$/i.exec(x);
    return { type: (m && m[2] ? m[2] : "FS").toUpperCase(), lag: m && m[3] ? (m[3] === "-" ? -1 : 1) * (parseFloat(m[4]) || 0) : 0 }; }); }
  const pct = (a, b) => b ? Math.round(a / b * 1000) / 10 : 0, share = (a, b) => `${a} of ${b} (${pct(a, b)}%)`;

  // ---- quality: 14 checks ----
  function quality(k) {
    if (!ok(k)) return { checks: [], pass: 0, of: 0 };
    const sd = k.statusDate, T = k.tasks, inc = T.filter(t => !t.done), R = Object.fromEntries(Object.entries(COLS).map(([key, names]) => [key, reader(k, names)]));
    const row = (t, detail) => ({ name: t.name, owner: who(t.owner), detail: detail || "" }), fl = t => daysOf(R.float(t));
    const allLinks = R.pred ? inc.flatMap(t => links(R.pred(t)).map(l => ({ t, ...l }))) : [];
    const out = [], add = (id, name, need, why, run) => { const c = { id, n: out.length + 1, name, state: "na", value: "", limit: "", rows: [], why: "" };
      if (need) Object.assign(c, run()); else c.why = why; out.push(c); };
    const most = (rows, base, lim = 5) => ({ rows, value: share(rows.length, base), limit: `at most ${lim}%`, state: pct(rows.length, base) <= lim ? "pass" : "fail" });
    const none = (rows, value) => ({ rows, value: value ?? String(rows.length), limit: "none", state: rows.length ? "fail" : "pass" });
    const noCol = c => `This plan has no ${c} column.`;
    add("logic", R.succ ? "Missing links" : "Missing predecessors", R.pred && inc.length, noCol("Predecessors"), () => most(inc.map(t => { const p = !R.pred(t), s = R.succ ? !R.succ(t) : false;
      return p || s ? row(t, p && s ? "no predecessor, no successor" : p ? "no predecessor" : "no successor") : null; }).filter(Boolean), inc.length));
    add("leads", "Leads (negative lag)", R.pred && allLinks.length, R.pred ? "No links between open tasks." : noCol("Predecessors"), () => { const b = allLinks.filter(l => l.lag < 0);
      return none([...new Set(b.map(l => l.t))].map(t => row(t, "starts before its predecessor allows")), `${b.length} of ${allLinks.length} links`); });
    add("lags", "Lags (waiting time on a link)", R.pred && allLinks.length, R.pred ? "No links between open tasks." : noCol("Predecessors"), () => { const b = allLinks.filter(l => l.lag > 0);
      return { ...most(b, allLinks.length), rows: [...new Set(b.map(l => l.t))].map(t => row(t, "has a lag on a link")) }; });
    add("fs", "Finish-to-start links", R.pred && allLinks.length, R.pred ? "No links between open tasks." : noCol("Predecessors"), () => { const fs = allLinks.filter(l => l.type === "FS").length, b = allLinks.filter(l => l.type !== "FS");
      return { rows: [...new Set(b.map(l => l.t))].map(t => row(t, "linked start-to-start, finish-to-finish or start-to-finish")), value: share(fs, allLinks.length), limit: "at least 90%", state: pct(fs, allLinks.length) >= 90 ? "pass" : "fail" }; });
    add("hard", "Hard constraints", R.cstr && inc.length, noCol("Constraint Type"), () => most(inc.filter(t => HARD.includes(E.norm(R.cstr(t)))).map(t => row(t, R.cstr(t))), inc.length));
    add("highfloat", `High float (over ${HIGH} working days)`, R.float && inc.length, noCol("Total Slack or Total Float"), () => most(inc.filter(t => fl(t) > HIGH).map(t => row(t, `${Math.round(fl(t))} working days of float`)), inc.length));
    add("negfloat", "Negative float", R.float && inc.length, noCol("Total Slack or Total Float"), () => none(inc.filter(t => fl(t) < 0).map(t => row(t, `${Math.round(fl(t) * 10) / 10} working days`))));
    const work = inc.filter(t => !t.ms && t.start != null && t.finish != null);
    add("long", `Long tasks (over ${HIGH} working days)`, k.cols.start >= 0 && work.length, "This plan has no Start column.", () => most(work.filter(t => E.wdiff(t.start, t.finish) + 1 > HIGH).map(t => row(t, `${E.wdiff(t.start, t.finish) + 1} working days`)), work.length));
    add("dates", "Dates on the wrong side of the status date", T.length, "", () => none([...inc.filter(t => t.finish != null && t.finish < sd).map(t => row(t, `not finished, but its finish date (${D(t.finish)}) has passed`)),
      ...inc.filter(t => !(t.finish != null && t.finish < sd) && t.pct === 0 && t.start != null && t.start < sd).map(t => row(t, `not started, but its start date (${D(t.start)}) has passed`)),
      ...T.filter(t => t.done && t.afin != null && t.afin > sd).map(t => row(t, `finished on ${D(t.afin)}, after the status date`))]));
    add("owner", "Tasks with nobody assigned", k.cols.owner >= 0 && inc.length, "This plan has no owner or resource column.", () => none(inc.filter(t => E.noOwner(t.owner)).map(t => row(t))));
    const due = T.filter(t => t.bfin != null && t.bfin <= sd);
    add("missed", "Missed tasks (late against the baseline)", k.cols.bfin >= 0 && due.length, k.cols.bfin >= 0 ? "No task has a baseline finish on or before the status date." : "This plan has no baseline finish column.",
      () => most(due.filter(t => !t.done || (t.afin != null && t.afin > t.bfin)).map(t => row(t, t.done ? `finished ${D(t.afin)}, baseline ${D(t.bfin)}` : `not finished, baseline ${D(t.bfin)}`)), due.length));
    add("cptest", "Critical path test", false, "Not checkable from a file: in your scheduling tool, add 600 days to a critical task and see that the finish date moves by the same.", null);
    const cpl = k.finish != null ? E.wdiff(sd, k.finish) : 0, floats = R.float ? inc.map(fl).filter(v => v != null) : [];
    add("cpli", "Critical path length index", floats.length && cpl > 0, R.float ? "The finish date has passed or no open task has float." : noCol("Total Slack or Total Float"), () => { const v = (cpl + Math.min(...floats)) / cpl;
      return { value: v.toFixed(2), limit: "at least 0.95", state: v >= 0.95 ? "pass" : "fail" }; });
    add("bei", "Baseline execution index", k.cols.bfin >= 0 && due.length, k.cols.bfin >= 0 ? "No task has a baseline finish on or before the status date." : "This plan has no baseline finish column.", () => { const v = T.filter(t => t.done).length / due.length;
      return { value: v.toFixed(2), limit: "at least 0.95", state: v >= 0.95 ? "pass" : "fail", rows: due.filter(t => !t.done).map(t => row(t, `baseline ${D(t.bfin)}, not finished`)) }; });
    const checked = out.filter(c => c.state !== "na");
    return { checks: out, pass: checked.filter(c => c.state === "pass").length, of: checked.length };
  }

  // ---- drivers: critical and near-critical tasks ----
  function drivers(k) {
    if (!ok(k)) return { has: false, critical: [], near: [] };
    const F = reader(k, COLS.float), C = reader(k, COLS.crit); if (!F && !C) return { has: false, critical: [], near: [] };
    const yes = v => /^(yes|y|true|1)$/i.test(String(v).trim()), item = t => ({ name: t.name, owner: who(t.owner), start: t.start, finish: t.finish, float: F ? daysOf(F(t)) : null, ms: !!t.ms });
    const inc = k.tasks.filter(t => !t.done && t.finish != null), byFin = (a, b) => a.finish - b.finish, isCrit = t => (C && yes(C(t))) || (F && daysOf(F(t)) != null && daysOf(F(t)) <= 0);
    return { has: true, critical: inc.filter(isCrit).map(item).sort(byFin), near: F ? inc.filter(t => !isCrit(t) && daysOf(F(t)) != null && daysOf(F(t)) <= NEAR).map(item).sort(byFin) : [] };
  }

  // ---- workload: open tasks per person per week ----
  function workload(k, weeks = 8) {
    if (!ok(k)) return { weeks: [], rows: [] };
    const sd = k.statusDate, w0 = sd - E.wd(sd), starts = Array.from({ length: weeks }, (_, i) => w0 + i * 7), by = {};
    for (const t of k.tasks) {
      if (t.done || t.ms || t.finish == null) continue;
      // overdue work still has to be done: it sits in this week until it is finished or re-dated
      const a = t.finish < w0 ? w0 : (t.start ?? t.finish), b = t.finish < w0 ? w0 : t.finish, r = by[who(t.owner)] ||= { owner: who(t.owner), counts: starts.map(() => 0), tasks: 0 };
      let hit = false; starts.forEach((s, i) => { if (a <= s + 6 && b >= s) { r.counts[i]++; hit = true; } }); if (hit) r.tasks++;
    }
    const rows = Object.values(by).filter(r => r.tasks).map(r => ({ ...r, peak: Math.max(...r.counts) })).sort((a, b) => (a.owner === NONE) - (b.owner === NONE) || b.peak - a.peak || a.owner.localeCompare(b.owner));
    return { weeks: starts, rows };
  }
  const workloadSheet = w => [{ name: "Workload", rows: [["Owner", ...w.weeks.map(s => `Week of ${E.fmt(s)}`), "Busiest week"], ...w.rows.map(r => [r.owner, ...r.counts, r.peak])] }];

  // ---- slips: what moved later, and why ----
  const REASONS = ["Scope changed", "Waiting on another team or supplier", "Waiting on the client", "People not available", "Took longer than estimated", "Rework or defects", "Re-planned", "Other"];
  // changes: diff.js changes between the last version and this one. One entry per move: a task that slips again gets a new one
  const slips = changes => (changes || []).filter(c => c.type === "finishLater" || (c.type === "milestoneMoved" && c.wd > 0)).map(c => ({ id: `${c.key}>${c.to}`, key: c.key, name: c.name, owner: who(c.owner),
    from: c.from, to: c.to, wd: c.wd, milestone: !!c.milestone })).sort((a, b) => b.wd - a.wd);
  // log: [{id, name, owner, from, to, wd, milestone, reason, note, asOf}] kept by the app; setting a reason adds or replaces that move's entry, clearing it removes it
  function setReason(log, slip, reason, note, asOf) { const rest = (log || []).filter(x => x.id !== slip.id); const r = REASONS.includes(reason) ? reason : "", n = String(note || "").trim().slice(0, 300);
    return r || n ? [...rest, { ...slip, reason: r, note: n, asOf }].sort((a, b) => a.asOf - b.asOf || b.wd - a.wd) : rest; }
  const slipSheet = log => [{ name: "Slip log", rows: [["Recorded as of", "Task", "Owner", "Was due", "Now due", "Working days later", "Reason", "Note"],
    ...(log || []).map(x => [{ d: x.asOf }, x.name, x.owner, { d: x.from }, { d: x.to }, x.wd, x.reason, x.note])] }];
  const slipSummary = log => { const by = {}; for (const x of log || []) { const r = by[x.reason || "No reason given"] ||= { reason: x.reason || "No reason given", n: 0, wd: 0 }; r.n++; r.wd += x.wd; } return Object.values(by).sort((a, b) => b.wd - a.wd); };

  const api = { quality, drivers, workload, workloadSheet, slips, setReason, slipSheet, slipSummary, links, daysOf, REASONS, HIGH, NEAR, NONE };
  root.KeelsonInsight = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
