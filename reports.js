/* Keelson Pro report builders: k (from KeelsonEngine.compute) + opts -> ReportDoc (plain JSON, see docs/PRO-ARCHITECTURE.md §2).
   Pure and deterministic: no DOM, no clock (the status date is k.statusDate). Every figure comes from k.
   Tested by tests/web/reports.test.js. */
(function (root) {
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const C = { navy: "#1f3a5f", teal: "#1a7f86", amber: "#f79009", red: "#d92d20", green: "#067647", grey: "#98a2b3", ink: "#1d2939", muted: "#667085",
    planned: "#9ccfd3", line: "#e4e7ec", lane: "#f7f9fb", bar: "#d0dae6" };
  const BAR = { "Done": C.grey, "On track": C.teal, "Planned": C.planned, "At risk": C.amber, "Late": C.red };
  const FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
  const PERIOD = 7, CAP = 20, SLIP_CAP = 15, HIST = 6;
  // milestone keywords the app passes as msKeys by default (a task whose name contains one is a milestone)
  const DEFAULT_MS_KEYS = ["go-live", "go live", "sign-off", "signed off", "sign off", "approval", "approved", "decision", "complete", "closed", "gate"];
  const NOBASE = "No baseline: finish slippage not assessed";

  const list = [
    { id: "onepager", name: "Executive one-pager", orientation: "landscape",
      description: "One page for the sponsor: overall status and why, progress, forecast finish, milestones, timeline, decisions needed." },
    { id: "weekly", name: "Weekly status report", orientation: "portrait",
      description: "The one-pager headline plus changes since the last report, full commentary with risks, changes and decisions, and this week's completed, due and overdue work." },
    { id: "roadmap", name: "Milestone roadmap", orientation: "landscape",
      description: "A clean milestone timeline by phase for a steering committee or client slide, with forecast against baseline." },
    { id: "variance", name: "Schedule variance and delay report", orientation: "portrait",
      description: "Where the schedule has slipped against baseline: open slipped tasks worst first, by owner and by phase." },
  ];

  // ---------- small helpers ----------
  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const cut = (s, n) => { s = String(s ?? ""); return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s; };
  const pctN = v => Math.round((v || 0) * 100);
  const pctT = v => pctN(v) + "%";
  const plural = (n, w, ws) => `${n} ${n === 1 || n === -1 ? w : (ws || w + "s")}`;
  const WORDS = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"];
  const lead = n => n >= 1 && n <= 9 ? WORDS[n] : String(n);                       // a number that starts a sentence
  const Plural = (n, w, ws) => `${lead(n)} ${n === 1 ? w : (ws || w + "s")}`;
  const and = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a[a.length - 1];
  // signed working days from a to b: the engine's rule (weekend-safe); local copy until the engine exports it
  const wdiff = (a, b) => E.wdiff ? E.wdiff(a, b) : a === b ? 0 : b > a ? E.nwd(a + 1, b) : -E.nwd(b + 1, a);
  const wdT = n => n === 0 ? "on baseline" : n > 0 ? `+${plural(n, "working day")}` : `${plural(-n, "working day")} early`;
  const wd = n => n === 0 ? "0" : `${n > 0 ? "+" : ""}${n}\u00a0wd`;               // narrow cells and tiles; never wraps
  const pts = n => `${n > 0 ? "+" : ""}${n} ${Math.abs(n) === 1 ? "pt" : "pts"}`;
  const fromIso = s => { const p = String(s || "").split("-").map(Number); return p.length === 3 && p.every(Number.isFinite) ? E.fromYMD(...p) : null; };
  const yearOf = v => +E.iso(v).slice(0, 4);
  const F = v => v == null ? "—" : E.fmt(v, true);
  const mkD = k => v => v == null ? "—" : E.fmt(v, k.statusDate == null || yearOf(v) !== yearOf(k.statusDate)); // year only when it differs
  const mkS = k => v => v == null ? "—" : E.fmt(v) + (k.statusDate == null || yearOf(v) !== yearOf(k.statusDate) ? " " + String(yearOf(v)).slice(2) : ""); // "24 Nov 26"
  const TONE = { "Done": "grey", "Met": "grey", "On track": "teal", "Planned": null, "At risk": "amber", "Late": "red", "Overdue": "red" };
  const chip = s => TONE[s] ? { text: s, tone: TONE[s], chip: true } : { text: s, chip: true };
  const L = (label, width) => ({ label, align: "left", ...(width ? { width } : {}) });
  const R = (label, width) => ({ label, align: "right", ...(width ? { width } : {}) });
  const M = (label, width) => ({ label, align: "center", ...(width ? { width } : {}) });
  const owner = o => o.owner || "—";
  const ragTone = r => ({ Red: "red", Amber: "amber", Green: "green" }[r] || null);
  const more = (n, w, ws) => n > 0 ? { type: "note", text: `+${n} more ${n === 1 ? w : (ws || w + "s")} not shown.` } : null;
  const compact = a => a.filter(Boolean);
  const range = (a, b) => yearOf(a) === yearOf(b) ? `${E.fmt(a)} – ${E.fmt(b, true)}` : `${E.fmt(a, true)} – ${E.fmt(b, true)}`;

  // engine fields that newer engines provide, with the same rule computed here when an older engine lacks them
  const openLeaves = k => (k.tasks || []).filter(t => !t.done);
  const stale = k => k.stale ?? { count: (k.tasks || []).filter(t => t.od).length, oldest: null };
  const finishPassed = k => k.finishPassed ?? (k.finish != null && k.finish < k.statusDate && openLeaves(k).length > 0);
  const bCover = k => k.bCover ?? (a => ({ with: a.filter(t => t.bfin != null).length, of: a.length }))((k.tasks || []).filter(t => !t.ms));

  // every phase in plan order (k.phases holds at most 8). The engine's phase (status, var, exp, behind) wins where it
  // exists; phases outside k.phases get the same rule: Done, Late, At risk by tolerance, Planned, On track.
  function allPhases(k) {
    const order = [], sd = k.statusDate, r = k.rules || {};
    for (const t of k.tasks || []) if (t.phase && !order.includes(t.phase)) order.push(t.phase);
    const mx = a => a.length ? Math.max(...a) : null, mn = a => a.length ? Math.min(...a) : null;
    const ex = t => t.start != null && t.finish != null ? (sd < t.start ? 0 : sd >= t.finish ? 1 : Math.min(1, E.nwd(t.start, sd) / t.w)) : (t.finish != null && sd >= t.finish ? 1 : 0);
    return order.map(name => {
      const pl = k.tasks.filter(t => t.phase === name), pt = pl.filter(t => !t.ms), pw = pt.reduce((s, t) => s + t.w, 0);
      const pe = t => t.done ? 1 : (t.pct ?? 0);
      const pct = pw ? pt.reduce((s, t) => s + t.w * pe(t), 0) / pw : pl.filter(t => t.done).length / Math.max(1, pl.length);
      const exp = pw ? pt.reduce((s, t) => s + t.w * ex(t), 0) / pw : 0;
      const finish = mx(pl.map(t => t.finish).filter(v => v != null)), bfin = mx(pl.map(t => t.bfin).filter(v => v != null));
      const start = mn(pl.map(t => t.start).filter(v => v != null));
      const v = finish != null && bfin != null ? wdiff(bfin, finish) : null, behind = pctN(exp) - pctN(pct);
      const status = pl.every(t => t.done) ? "Done" : pl.some(t => t.od) ? "Late" : (v ?? 0) > (r.amberVar ?? 0) || behind > (r.amberBehind ?? 100) ? "At risk"
        : pct === 0 && start != null && start > sd ? "Planned" : "On track";
      const own = { name, pct, exp, start, finish, bfin, var: v, behind, status, open: pt.filter(t => !t.done).length };
      const kp = (k.phases || []).find(p => p.name === name);
      return kp ? { ...own, ...kp, bfin: kp.bfin ?? bfin, var: kp.var ?? v } : own;
    });
  }
  const msStatus = (t, k) => t.done ? "Met" : t.finish < k.statusDate ? "Overdue" : (t.var ?? 0) > ((k.rules || {}).amberVar ?? 0) ? "At risk" : "On track";
  const allMs = k => (k.tasks || []).filter(t => t.ms && t.finish != null).sort((a, b) => a.finish - b.finish || a.row - b.row)
    .map(t => ({ ...t, status: msStatus(t, k) }));
  const findMs = (k, m) => m && (k.tasks || []).find(t => t.ms && t.name === m.name && t.finish === m.finish);
  const byLate = (a, b) => b.late - a.late || a.row - b.row;
  const overdueTasks = k => (k.tasks || []).filter(t => t.od && !t.ms).sort(byLate);
  const dueSoon = k => (k.tasks || []).filter(t => !t.ms && !t.done && t.finish != null && t.finish >= k.statusDate && t.finish <= k.statusDate + k.upcomingDays)
    .sort((a, b) => a.finish - b.finish || a.row - b.row);
  const completedRecently = (k, days = PERIOD) => (k.tasks || []).filter(t => t.done && t.finish != null && t.finish > k.statusDate - days && t.finish <= k.statusDate)
    .sort((a, b) => a.finish - b.finish || a.row - b.row);
  const slipped = k => (k.tasks || []).filter(t => !t.ms && (t.var ?? 0) > 0).sort((a, b) => b.var - a.var || a.row - b.row);

  // ---------- snapshot + deltas ----------
  // task identity across reports: the engine's key (name|phase, "#n" on repeats); a key that is not unique is skipped
  const baseKey = t => String(t.key ?? `${t.name}|${t.phase || ""}`).replace(/\s*\(?#\d+\)?$/, "");
  function taskKeys(k) {
    const n = {}, nn = {};
    for (const t of k.tasks || []) { n[baseKey(t)] = (n[baseKey(t)] || 0) + 1; nn[t.name] = (nn[t.name] || 0) + 1; }
    return (k.tasks || []).filter(t => t.name).map(t => ({ t, key: n[baseKey(t)] === 1 ? baseKey(t) : null, byName: nn[t.name] === 1 }));
  }
  const point = s => s && typeof s.asOf === "string" ? { asOf: s.asOf, rag: s.rag ?? "—", pct: typeof s.pct === "number" ? s.pct : null, var: typeof s.var === "number" ? s.var : null } : null;
  // the last HIST report points (oldest first), ending with `cur`; earlier points come from `hist` and `prev`
  function trail(cur, hist, prev) {
    const a = [...(Array.isArray(hist) ? hist : []), ...(prev && Array.isArray(prev.history) ? prev.history : []), prev].map(point)
      .filter(p => p && p.asOf < cur.asOf);
    const by = {}; for (const p of a) by[p.asOf] = p;
    return [...Object.values(by).sort((x, y) => x.asOf < y.asOf ? -1 : 1), cur].slice(-HIST);
  }
  function snapshot(k, prev) {
    const tasks = {};
    for (const { t, key } of taskKeys(k)) if (key) tasks[key] = E.iso(t.finish);
    const phases = {};
    for (const p of allPhases(k)) phases[p.name] = p.pct;
    const s = { v: 2, asOf: E.iso(k.statusDate), rag: k.rag, pct: k.pct, exp: k.exp, finish: E.iso(k.finish), bfinish: E.iso(k.bfinish), var: k.var,
      overT: k.overT, overMs: k.overMs, nMs: k.nMs, nMsDone: k.nMsDone, nTask: k.nTask, nTaskDone: k.nTaskDone, phases, tasks };
    s.history = trail(point(s), null, prev);
    return s;
  }

  // the "since last report" figures; null when there is no usable previous snapshot
  function deltas(k, prev) {
    if (!prev || typeof prev !== "object") return null;
    const num = v => typeof v === "number" && Number.isFinite(v);
    const pf = fromIso(prev.finish);
    const moved = [];
    if (prev.tasks) for (const { t, key, byName } of taskKeys(k)) { // v1 snapshots were keyed by name
      const was = fromIso(key && key in prev.tasks ? prev.tasks[key] : prev.v !== 2 && byName ? prev.tasks[t.name] : null);
      if (was != null && t.finish != null && was !== t.finish)
        moved.push({ name: t.name, owner: t.owner, phase: t.phase, was, finish: t.finish, by: wdiff(was, t.finish), row: t.row });
    }
    moved.sort((a, b) => Math.abs(b.by) - Math.abs(a.by) || b.by - a.by || a.row - b.row);
    return {
      asOf: fromIso(prev.asOf), rag: prev.rag ?? null,
      pct: num(prev.pct) ? pctN(k.pct) - pctN(prev.pct) : null,
      finish: pf != null && k.finish != null ? wdiff(pf, k.finish) : null, prevFinish: pf,
      overT: num(prev.overT) ? k.overT - prev.overT : null,
      overMs: num(prev.overMs) ? k.overMs - prev.overMs : null,
      nMsDone: num(prev.nMsDone) ? k.nMsDone - prev.nMsDone : null,
      phases: Object.fromEntries(allPhases(k).filter(p => prev.phases && num(prev.phases[p.name])).map(p => [p.name, pctN(p.pct) - pctN(prev.phases[p.name])])),
      moved, prev };
  }
  const sgn = n => n > 0 ? "+" + n : String(n);
  const tone = (n, goodUp) => !n ? null : (n > 0) === goodUp ? "green" : "red";

  // ---------- commentary ----------
  function commentary(k, opts = {}) {
    const sd = k.statusDate, D = mkD(k), edits = (opts && opts.edits) || {}, short = !!(opts && opts.compact); // compact: one-pager columns
    const ph = allPhases(k), ms = allMs(k), open = openLeaves(k), inP = v => v != null && v > sd - PERIOD && v <= sd;
    const hi = [], co = [], nx = [];
    if (!k.rows) hi.push("The plan has no rows to report on yet.");
    else if (!open.length) { // everything is done: say so plainly, no "next 14 days" filler
      const v = k.bfinish == null ? "" : k.var > 0 ? `, ${plural(k.var, "working day")} after baseline` : k.var < 0 ? `, ${plural(-k.var, "working day")} ahead of baseline` : ", on baseline";
      hi.push(`All work is complete; the project finished on ${F(k.finish)}${v}.`);
      co.push(k.bfinish != null && k.var > 0 ? `The project finished ${plural(k.var, "working day")} after its baseline of ${F(k.bfinish)}.` : "No open concerns: all work is complete.");
      nx.push("No open work remains.");
    } else {
      // highlights: what was delivered this period, then phases (or milestones), then progress
      const got = [...ms.filter(m => m.done && inP(m.finish)), ...ph.filter(p => p.status === "Done" && inP(p.finish)).map(p => ({ name: p.name + " phase", finish: p.finish }))]
        .sort((a, b) => a.finish - b.finish);
      const items = got.length ? got : completedRecently(k).filter(t => !t.ms), nShow = short ? 2 : 3;
      if (items.length) hi.push(`Completed this period: ${and(items.slice(0, nShow).map(m => `${m.name} (${D(m.finish)})`).concat(items.length > nShow ? [`${items.length - nShow} more`] : []))}.`);
      const doneP = ph.filter(p => p.status === "Done").length;
      if (k.nTask && ph.length && doneP) hi.push(`${lead(doneP)} of ${ph.length} phases ${doneP === 1 ? "is" : "are"} complete.`);
      else if (!k.nTask && k.nMs) hi.push(`${lead(k.nMsDone)} of ${plural(k.nMs, "milestone")} met.`);
      if (k.nTask) {
        const d = pctN(k.pct) - pctN(k.exp);
        if (!pctN(k.exp) && !pctN(k.pct)) {
          const first = Math.min(...open.filter(t => !t.ms).map(t => t.start ?? t.finish).filter(v => v != null && v >= sd));
          hi.push(`Work has not started${Number.isFinite(first) ? `; the first tasks begin on ${F(first)}` : ""}.`);
        } else hi.push(`${pctT(k.pct)} complete against ${pctT(k.exp)} planned, ${Math.abs(d) <= 2 ? "broadly on plan" : d < 0 ? `${plural(-d, "point")} behind plan` : `${plural(d, "point")} ahead of plan`}; ${k.nTaskDone} of ${plural(k.nTask, "task")} done.`);
      }

      // concerns: finish slip, a forecast that has gone stale, overdue work, progress, missing baseline
      const slipMs = ms.filter(m => !m.done && (m.var ?? 0) > 0);
      if (k.bfinish != null && k.var > 0) co.push(`Forecast finish ${F(k.finish)} is ${plural(k.var, "working day")} after the baseline of ${F(k.bfinish)}${k.var <= k.rules.amberVar ? ", within tolerance" : ""}` +
        (slipMs.length && !short ? `; ${slipMs.length > 1 ? `${slipMs.length} milestones are forecast after baseline, the next being ${slipMs[0].name} (+${plural(slipMs[0].var, "working day")})` : `milestone ${slipMs[0].name} is forecast ${plural(slipMs[0].var, "working day")} after its baseline`}.` : "."));
      else if (slipMs.length) co.push(`${Plural(slipMs.length, "milestone")} ${slipMs.length > 1 ? "are" : "is"} forecast after baseline, the next being ${slipMs[0].name} (+${plural(slipMs[0].var, "working day")}), though the overall finish holds.`);
      const passed = finishPassed(k);
      if (passed) co.push(`The forecast finish of ${F(k.finish)} has passed with work still open; the plan needs re-forecasting.`);
      const od = overdueTasks(k), odMs = ms.filter(m => m.status === "Overdue");
      if (od.length || odMs.length) {
        const subj = od.length && odMs.length ? `${Plural(od.length, "task")} and ${plural(odMs.length, "milestone")} are overdue`
          : od.length ? `${Plural(od.length, "task")} ${od.length > 1 ? "are" : "is"} overdue` : `${Plural(odMs.length, "milestone")} ${odMs.length > 1 ? "are" : "is"} overdue`;
        const names = od.length ? od.slice(0, short ? 1 : 2).map((t, i) => `${t.name} (${t.owner || "no owner"}, ${i ? t.late : plural(t.late, "working day")})`)
          : odMs.slice(0, short ? 1 : 2).map(m => `${m.name} (due ${D(m.finish)})`);
        const n = od.length || odMs.length;
        co.push(`${subj}${n > 1 ? ", the oldest being " : ": "}${and(names)}` + (!passed && !short ? `; ${n > 1 ? "their" : "its"} dates need re-forecasting.` : "."));
      }
      if (k.behind > k.rules.amberBehind) co.push(`Progress is ${plural(k.behind, "point")} behind plan.`);
      if (k.bfinish == null) co.push("The plan has no baseline dates, so slippage against baseline cannot be measured.");
      if (!co.length) co.push("Nothing is overdue and the forecast finish is on or ahead of baseline.");
      if (short && co.length > 1 && k.bfinish != null && k.var > 0 && k.var <= k.rules.amberVar) co.push(co.shift()); // a slip within tolerance is not the lead concern

      const up = dueSoon(k), nUp = short ? 1 : 3;
      nx.push(!up.length ? `No tasks are due in the next ${plural(k.upcomingDays, "day")}.`
        : `${Plural(up.length, "task")} ${up.length > 1 ? "are" : "is"} due in the next ${plural(k.upcomingDays, "day")}${up.length > nUp ? ", including" : ":"} ` +
          and(up.slice(0, nUp).map(t => `${t.name} (${t.owner ? t.owner + ", " : ""}${D(t.finish)})`)) + ".");
      const nm = findMs(k, k.nextMs);
      if (k.nextMs) nx.push(`Next milestone: ${k.nextMs.name} on ${F(k.nextMs.finish)}` +
        (nm && nm.bfin != null ? (nm.var > 0 ? `, ${plural(nm.var, "working day")} after its baseline of ${F(nm.bfin)}.` : nm.var < 0 ? `, ahead of its baseline of ${F(nm.bfin)}.` : ", on baseline.") : "."));
      else if (k.nMs && k.nMsDone === k.nMs) nx.push("All milestones have been met.");
      else if (k.nMs) nx.push("No open milestone is dated on or after the status date.");
    }
    const block = (id, label, text, placeholder) => {
      const ed = typeof edits[id] === "string" ? edits[id] : null;
      return { type: "text", id, label, text: ed ?? text, editable: true, ...(placeholder && ed == null ? { placeholder: true, exportText: placeholder } : {}) };
    };
    // compact: at most 2 sentences and about 120 characters (the second sentence is dropped if it would not fit)
    const join = a => !short ? a.slice(0, 3).join(" ") : a.slice(0, 2).reduce((t, x, i) => !i ? x : (t + " " + x).length <= 120 ? t + " " + x : t, "");
    return {
      highlights: block("highlights", "Highlights", join(hi)),
      concerns: block("concerns", "Concerns", join(co)),
      next: block("next", `Next ${plural(k.upcomingDays ?? 14, "day")}`, join(nx)),
      decisions: block("decisions", "Decisions and support needed", "Add decisions or support needed from the sponsor.", "No decisions requested this period."),
    };
  }
  // an editable list the user fills in (risks, changes); unedited it is a placeholder that exports as "None reported"
  function userList(id, prompt, edits) {
    const e = edits && edits[id], items = (Array.isArray(e) ? e : typeof e === "string" ? e.split("\n") : []).map(s => String(s).trim()).filter(Boolean);
    return e == null ? { type: "list", id, items: [prompt], editable: true, placeholder: true, exportText: "None reported this period." }
      : { type: "list", id, items: items.length ? items : ["None reported this period."], editable: true };
  }

  // ---------- svg ----------
  function months(t0, t1) {
    const out = [], [y, m] = E.iso(t0).split("-").map(Number);
    for (let i = 1; i < 120; i++) { const v = E.fromYMD(y, m + i, 1); if (v > t1) break; out.push(v); }
    return out;
  }
  function axis(t0, t1, x0, x1, y, h, fs = 11) {
    const ms = months(t0, t1), every = Math.ceil(ms.length / 12), span = t1 - t0 + 1, pos = v => +(x0 + (v - t0) / span * (x1 - x0)).toFixed(1);
    const out = [];
    ms.forEach((v, i) => {
      out.push(`<line x1="${pos(v)}" y1="${y + 6}" x2="${pos(v)}" y2="${y + h}" stroke="${C.line}" stroke-width="1"/>`);
      if (i % every === 0) {
        const [yy, mm] = E.iso(v).split("-").map(Number), lab = E.fmt(v).split(" ")[1] + (mm === 1 || i === 0 ? " " + yy : "");
        out.push(`<text x="${pos(v) + 3}" y="${y}" font-size="${fs}" fill="${C.muted}">${esc(lab)}</text>`);
      }
    });
    return { svg: out.join(""), pos, span };
  }
  const diamond = (x, y, r, fill, stroke) => { const n = v => +v.toFixed(1);
    return `<path d="M${n(x)} ${n(y - r)}L${n(x + r)} ${n(y)}L${n(x)} ${n(y + r)}L${n(x - r)} ${n(y)}Z" fill="${fill}"${stroke ? ` stroke="${stroke}" stroke-width="1.5"` : ""}/>`; };
  const wrap = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="${FONT}" fill="${C.ink}"><rect width="${w}" height="${h}" fill="#ffffff"/>${body}</svg>`;
  const nowLine = (k, ax, t0, t1, y0, y1) => k.statusDate >= t0 && k.statusDate <= t1
    ? `<line x1="${ax.pos(k.statusDate + .5)}" y1="${y0}" x2="${ax.pos(k.statusDate + .5)}" y2="${y1}" stroke="${C.red}" stroke-width="1.5" stroke-dasharray="4 3"/>` : "";
  function legend(items, x, y, fs = 11) { // items: [kind, colour, label]
    let cx = x; const out = [];
    for (const [kind, col, lab] of items) {
      if (kind === "bar") out.push(`<rect x="${cx}" y="${y - 8}" width="14" height="9" rx="2" fill="${col}"/>`);
      else if (kind === "ms") out.push(diamond(cx + 6, y - 4, 6, col));
      else if (kind === "label") { out.push(`<text x="${cx}" y="${y}" font-size="${fs}" font-weight="600" fill="${C.ink}">${esc(lab)}</text>`); cx += Math.max(lab.length * fs * .62 + 10, 78); continue; }
      else if (kind === "hollow") out.push(diamond(cx + 6, y - 4, 5, "#ffffff", col));
      else if (kind === "ring") out.push(diamond(cx + 6, y - 4, 6, C.teal, col));
      else out.push(`<line x1="${cx + 6}" y1="${y - 10}" x2="${cx + 6}" y2="${y + 2}" stroke="${col}" stroke-width="1.5" stroke-dasharray="4 3"/>`);
      out.push(`<text x="${cx + 19}" y="${y}" font-size="${fs}" fill="${C.muted}">${esc(lab)}</text>`);
      cx += 19 + lab.length * fs * .56 + 18;
    }
    return out.join("");
  }
  const MS_FILL = { "Met": C.navy, "Overdue": C.red, "At risk": C.amber, "On track": C.teal };
  const MS_ORDER = ["Met", "On track", "At risk", "Overdue"], PH_ORDER = ["Done", "On track", "Planned", "At risk", "Late"];
  // only the statuses the chart shows (v6)
  const barLegend = sts => PH_ORDER.filter(s => sts.includes(s)).map(s => ["bar", BAR[s], s]);
  const msLegend = (sts, base, ring) => [...MS_ORDER.filter(s => sts.includes(s)).map(s => ["ms", MS_FILL[s], s]),
    ...(base ? [["hollow", C.muted, "Baseline"]] : []), ...(ring ? [["ring", C.amber, "Just after baseline"]] : [])];
  // baseline marker + link line, or (when the two markers would overlap: always under 4 px) an amber ring on the forecast diamond (v2)
  const near = (bx, x, gap) => Math.abs(bx - x) < Math.max(4, gap);

  // one-pager chart: a row per phase (% done, finish, status, bar coloured by status) and a row per milestone
  // (due, status, diamond at forecast with hollow baseline marker), status-date line and month axis. One svg keeps
  // the page on one A4 landscape sheet; table rows would need twice the height.
  const INK = { "Done": "#475467", "Met": "#475467", "On track": "#0f5e63", "Planned": "#3538cd", "At risk": "#93370d", "Late": "#b42318", "Overdue": "#b42318" };
  function timelineSvg(k, msRows) {
    if (k.start == null) return null;
    const S = mkS(k), mt = msRows.map(m => findMs(k, m) || { ...m, var: null, bfin: null }).map(t => ({ ...t, status: msStatus(t, k) }));
    const t0 = Math.min(k.start, ...mt.map(m => m.bfin ?? m.finish)), t1 = Math.max(k.finish ?? t0, k.bfinish ?? t0, t0, ...mt.map(m => m.bfin ?? m.finish));
    const rows = k.phases.length ? k.phases : [{ name: "Whole plan", pct: k.pct, start: k.start, finish: k.finish, status: k.overT ? "Late" : "On track" }];
    const W = 1000, X0 = 420, RH = 14, f = 12, top = 20, msTop = top + rows.length * RH + (mt.length ? 18 : 0), bot = msTop + mt.length * RH, H = bot + 22;
    const XF = 290, XS = 362; // FINISH/DUE and STATUS columns: short dates ("24 Nov 26") end before STATUS starts
    const ax = axis(t0, t1, X0, W - 10, 13, bot - 13, f), out = [];
    const head = (y, a) => a.map(([x, t, anchor]) => `<text x="${x}" y="${y}" font-size="${f - 1}" font-weight="600" fill="${C.muted}"${anchor ? ` text-anchor="${anchor}"` : ""}>${esc(t)}</text>`).join("");
    const cells = (y, i, a) => { const ty = y + 11; return (i % 2 === 0 ? `<rect x="0" y="${y}" width="${W}" height="${RH}" fill="${C.lane}"/>` : "") +
      a.map(([x, t, extra]) => `<text x="${x}" y="${ty}" font-size="${f}"${extra || ""}>${esc(t)}</text>`).join(""); };
    out.push(head(13, [[8, "PHASE"], [278, "DONE", "end"], [XF, "FINISH"], [XS, "STATUS"]]));
    rows.forEach((p, i) => {
      const y = top + i * RH, a = p.start ?? p.finish, b = p.finish ?? a;
      out.push(cells(y, i, [[8, cut(p.name, 34)], [278, pctT(p.pct), ` text-anchor="end"`], [XF, S(p.finish)], [XS, p.status, ` font-weight="600" fill="${INK[p.status] || C.ink}"`]]));
      if (a != null) out.push(`<rect x="${ax.pos(a)}" y="${y + 2}" width="${Math.max(4, +((b - a + 1) / ax.span * (W - 10 - X0)).toFixed(1))}" height="10" rx="3" fill="${BAR[p.status] || C.teal}"/>`);
    });
    let base = false, ring = false;
    if (mt.length) {
      const onlyOpen = mt.every(m => !m.done);
      out.push(`<line x1="0" y1="${msTop - 17}" x2="${W}" y2="${msTop - 17}" stroke="${C.line}"/>`,
        head(msTop - 5, [[8, `${onlyOpen ? "OPEN " : ""}MILESTONES (${k.ms} MET)`.toUpperCase()], [XF, "DUE"], [XS, "STATUS"]]));
      mt.forEach((m, i) => {
        const y = msTop + i * RH, x = ax.pos(m.finish + .5), bx = m.bfin != null && m.bfin !== m.finish ? ax.pos(m.bfin + .5) : null, close = bx != null && near(bx, x, 10.5);
        out.push(cells(y, i, [[8, cut(m.name, 34)], [XF, S(m.finish)], [XS, m.status, ` font-weight="600" fill="${INK[m.status] || C.ink}"`]]));
        if (bx != null && !close) { base = true;
          out.push(`<line x1="${bx}" y1="${y + 7}" x2="${x}" y2="${y + 7}" stroke="${C.muted}" stroke-dasharray="2 2"/>`, diamond(bx, y + 7, 4.5, "#ffffff", C.muted)); }
        if (close && m.finish > m.bfin) ring = true;
        out.push(diamond(x, y + 7, 6, MS_FILL[m.status], close && m.finish > m.bfin ? C.amber : "#ffffff"));
      });
    }
    out.unshift(ax.svg);
    out.push(nowLine(k, ax, t0, t1, top - 4, bot));
    out.push(legend([...barLegend(rows.map(p => p.status)), ...(mt.length ? msLegend(mt.map(m => m.status), base, ring) : []),
      ["line", C.red, "Status date " + E.fmt(k.statusDate)]], 8, H - 5, f));
    const D = mkD(k);
    return { type: "svg", svg: wrap(W, H, out.join("")), width: W, height: H,
      alt: `Progress by phase and milestones, ${F(t0)} to ${F(t1)}, status date ${F(k.statusDate)}. Phases: ` +
        rows.map(p => `${p.name} ${pctT(p.pct)} done, finish ${D(p.finish)}, ${p.status}`).join("; ") +
        (mt.length ? `. Milestones: ` + mt.map(m => `${m.name} due ${D(m.finish)}, ${m.status}`).join("; ") : "") + "." };
  }

  // roadmap: phases as lanes (bar coloured by phase status), labelled milestone diamonds at forecast (hollow marker at
  // baseline when it differs), labels staggered above/below so they never overlap, status-date line labelled.
  const ROAD_LANES = 12, ROAD_H = 560, LANE_MIN = 32;
  function roadLanes(k) {
    const ms = allMs(k), ph = allPhases(k);
    const lead0 = ph.findIndex(p => p.status !== "Done"), off = Math.max(0, Math.min(lead0 < 0 ? 0 : lead0, ph.length - ROAD_LANES));
    const shown = ph.slice(off, off + ROAD_LANES);
    let lanes = shown.map(p => ({ name: p.name, start: p.start ?? p.finish, finish: p.finish, status: p.status, ms: ms.filter(m => m.phase === p.name) }));
    const loose = ms.filter(m => !ph.some(p => p.name === m.phase));
    if (!lanes.length) lanes = [{ name: "Project", start: k.start, finish: k.finish, status: null, sub: `${pctT(k.pct)} complete`, ms }];
    else if (loose.length) { if (lanes.length === ROAD_LANES) lanes.pop(); lanes.push({ name: "Other milestones", start: null, finish: null, status: null, ms: loose }); }
    const hidden = ph.length - lanes.filter(l => shown.includes(ph.find(p => p.name === l.name))).length;
    return { lanes, ms, off, after: Math.max(0, hidden - off) };
  }
  function roadmapSvg(k) {
    if (k.start == null) return null;
    const sd = k.statusDate, D = mkD(k), { lanes } = roadLanes(k), f = 12;
    const vis = lanes.flatMap(l => l.ms), sp = lanes.filter(l => l.start != null);
    const t0 = Math.min(...sp.map(l => l.start), ...vis.map(m => Math.min(m.finish, m.bfin ?? m.finish)), sd);
    const t1 = Math.max(...sp.map(l => l.finish), ...vis.map(m => Math.max(m.finish, m.bfin ?? m.finish)), t0 + 1);
    const W = 1000, LW = 228, X1 = W - 12, top = 38, foot = 48;
    const span = t1 - t0 + 1, pos = v => +(LW + (v - t0) / span * (X1 - LW)).toFixed(1);
    const sx = sd >= t0 && sd <= t1 ? pos(sd + .5) : null;
    // place labels: slots 0,2,4.. above the bar, 1,3,5.. below; first slot whose last label ends before this one starts.
    // A label never runs within 4 px of the status-date line: it moves to the side of the line its diamond is on.
    const layout = n => lanes.map(ln => {
      const ends = [], lab = ln.ms.map(m => {
        const name = n ? cut(m.name, n) : "", date = D(m.finish) + (!m.done && (m.var ?? 0) > 0 ? ` (+${m.var} wd)` : "");
        const w = name.length * 7 + (name ? 6 : 0) + date.length * 6.5, x = pos(m.finish + .5);
        let l = Math.min(X1 - w, Math.max(LW + 4, x - w / 2));
        if (sx != null && l < sx + 4 && l + w > sx - 4) { // the diamond's side first, else the other side when that one is too narrow
          const left = sx - 4 - w >= LW + 4, right = sx + 4 + w <= X1;
          l = (x <= sx ? left || !right : !right && left) ? Math.max(LW + 4, sx - 4 - w) : Math.min(X1 - w, sx + 4);
        }
        let s = 0; while (ends[s] != null && ends[s] + 10 > l) s++;
        ends[s] = l + w;
        return { m, name, date, x, l, s };
      });
      const na = Math.ceil(Math.max(0, ...lab.map(o => o.s + 1)) / 2), nb = Math.floor(Math.max(0, ...lab.map(o => o.s + 1)) / 2);
      const up = na ? 12 + 14 * na : 16, down = Math.max(nb ? 12 + 14 * nb : 16, LANE_MIN - up);
      return { ln, lab, up, h: up + down };
    });
    let L;
    for (const n of [30, 20, 12, 0]) { // ponytail: shorter labels until the chart fits; a plan with dozens of clustered milestones still overflows
      L = layout(n);
      if (top + L.reduce((t, o) => t + o.h, 0) + foot <= ROAD_H) break;
    }
    const lanesH = L.reduce((t, o) => t + o.h, 0), H = top + lanesH + foot, out = [];
    const mo = months(t0, t1), every = Math.ceil(mo.length / 10);
    mo.forEach((v, i) => {
      out.push(`<line x1="${pos(v)}" y1="${top - 6}" x2="${pos(v)}" y2="${top + lanesH}" stroke="${C.line}"/>`);
      if (i % every === 0) { const [yy, mm] = E.iso(v).split("-").map(Number);
        out.push(`<text x="${pos(v) + 4}" y="${top - 9}" font-size="${f}" fill="${C.muted}">${esc(E.fmt(v).split(" ")[1] + (mm === 1 || i === 0 ? " " + yy : ""))}</text>`); }
    });
    let y = top, base = false, ring = false;
    L.forEach(({ ln, lab, up, h }, i) => {
      const mid = y + up;
      if (i % 2 === 0) out.unshift(`<rect x="0" y="${y}" width="${W}" height="${h}" fill="${C.lane}"/>`);
      out.push(`<text x="10" y="${mid - 2}" font-size="13" font-weight="600" fill="${C.navy}">${esc(cut(ln.name, 28))}</text>`);
      const sub = ln.status || ln.sub || "";
      if (sub) out.push(`<text x="10" y="${mid + 13}" font-size="${f}" font-weight="${ln.status ? 600 : 400}" fill="${ln.status ? INK[ln.status] : C.muted}">${esc(sub)}</text>`);
      if (ln.start != null && ln.finish != null)
        out.push(`<rect x="${pos(ln.start)}" y="${mid - 4}" width="${Math.max(4, +((ln.finish - ln.start + 1) / span * (X1 - LW)).toFixed(1))}" height="8" rx="4" fill="${BAR[ln.status] || C.bar}"/>`);
      for (const { m, x } of lab) if (m.bfin != null && m.bfin !== m.finish && !near(pos(m.bfin + .5), x, 14)) {
        const bx = pos(m.bfin + .5); base = true;
        out.push(`<line x1="${bx}" y1="${mid}" x2="${x}" y2="${mid}" stroke="${C.ink}" stroke-width="1.2" stroke-dasharray="2 2"/>`, diamond(bx, mid, 6, "#ffffff", C.ink));
      }
      for (const { m, name, date, x, l, s } of lab) {
        const r = m.bfin != null && m.finish > m.bfin && near(pos(m.bfin + .5), x, 14); ring = ring || r;
        out.push(diamond(x, mid, 8, MS_FILL[m.status], r ? C.amber : "#ffffff"));
        const ty = s % 2 === 0 ? mid - 13 - 14 * (s / 2) : mid + 21 + 14 * ((s - 1) / 2);
        out.push(`<text x="${+l.toFixed(1)}" y="${ty}" font-size="${f}">${name ? `<tspan font-weight="600">${esc(name)}</tspan> ` : ""}<tspan fill="${C.muted}">${esc(date)}</tspan></text>`);
      }
      y += h;
    });
    if (sx != null) out.push(`<line x1="${sx}" y1="2" x2="${sx}" y2="${top + lanesH}" stroke="${C.red}" stroke-width="1.5" stroke-dasharray="4 3"/>`,
      `<text x="${Math.min(X1, sx + 4)}" y="12" font-size="${f}" font-weight="600" fill="${C.red}" text-anchor="${sx + 130 > X1 ? "end" : "start"}">Status date ${esc(E.fmt(sd))}</text>`);
    const sts = lanes.map(l => l.status).filter(Boolean), mss = vis.map(m => m.status);
    out.push(!sts.length ? "" : legend([["label", null, "Phase"], ...barLegend(sts)], 10, H - 26, f),
      !mss.length ? "" : legend([["label", null, "Milestone"], ...msLegend(mss, false, ring), ...(base ? [["hollow", C.ink, "Baseline date"]] : [])], 10, H - 6, f));
    return { type: "svg", svg: wrap(W, H, out.join("")), width: W, height: H,
      alt: `Milestone roadmap, status date ${F(sd)}: ` + lanes.map(l => `${l.name}${l.status ? ` (${l.status})` : ""}` +
        (l.ms.length ? ": " + l.ms.map(m => `${m.name} ${D(m.finish)} ${m.status}`).join(", ") : "")).join("; ") + "." };
  }

  // weekly trend: one point per report (oldest first), dot coloured by overall status, line = % complete
  function trendSvg(pts) {
    const W = 680, H = 118, x0 = 70, x1 = W - 40, n = pts.length, X = i => +(x0 + i * (x1 - x0) / Math.max(1, n - 1)).toFixed(1), Y = p => +(84 - (p || 0) * 60).toFixed(1);
    const col = r => ({ Red: C.red, Amber: C.amber, Green: C.green }[r] || C.grey), out = [];
    for (const g of [0, .5, 1]) out.push(`<line x1="${x0 - 14}" y1="${Y(g)}" x2="${x1 + 14}" y2="${Y(g)}" stroke="${C.line}"/>`,
      `<text x="${x0 - 20}" y="${Y(g) + 4}" font-size="11" fill="${C.muted}" text-anchor="end">${g * 100}%</text>`);
    const line = pts.map((p, i) => p.pct == null ? null : `${X(i)},${Y(p.pct)}`).filter(Boolean);
    if (line.length > 1) out.push(`<polyline points="${line.join(" ")}" fill="none" stroke="${C.teal}" stroke-width="2"/>`);
    pts.forEach((p, i) => {
      const y = Y(p.pct), d = fromIso(p.asOf);
      out.push(`<circle cx="${X(i)}" cy="${y}" r="6" fill="${col(p.rag)}" stroke="#ffffff" stroke-width="1.5"/>`,
        p.pct == null ? "" : `<text x="${X(i)}" y="${y - 10}" font-size="11" font-weight="600" text-anchor="middle">${pctT(p.pct)}</text>`,
        `<text x="${X(i)}" y="${H - 14}" font-size="11" fill="${C.muted}" text-anchor="middle">${esc(d == null ? p.asOf : E.fmt(d))}</text>`);
    });
    out.push(`<text x="8" y="${H - 1}" font-size="11" fill="${C.muted}">Dot colour = overall status (red, amber, green); line = % complete at each report.</text>`);
    return { type: "svg", svg: wrap(W, H, out.join("")), width: W, height: H,
      alt: `Trend over the last ${n} reports: ` + pts.map(p => `${p.asOf} ${p.rag}${p.pct == null ? "" : ", " + pctT(p.pct) + " complete"}`).join("; ") + "." };
  }

  // ---------- shared blocks ----------
  function tiles(k) {
    const D = mkD(k), r = k.rules, st = stale(k), cov = bCover(k), passed = finishPassed(k), noBase = k.bfinish == null && k.rows > 0;
    const ragSub = k.ragCalc && k.ragCalc !== k.rag ? [`Calculated: ${k.ragCalc}`, k.ragReason].filter(Boolean).join(" · ")
      : [...(k.why || []), ...(noBase && !(k.why || []).includes(NOBASE) ? [NOBASE] : [])].join(" · ");
    const old = st.count || passed;
    const finSub = k.finish == null ? "" : [
      old && (st.count ? `Forecast not updated: ${plural(st.count, "open task")} ${st.count === 1 ? "is" : "are"} past ${st.count === 1 ? "its" : "their"} finish date` : "Forecast not updated: the finish date has passed with work open"),
      k.bfinish == null ? !old && "No baseline in plan" : old ? `Baseline ${D(k.bfinish)} (${k.var ? wd(k.var) : "on baseline"})` // short beside the stale note
        : `Baseline ${D(k.bfinish)} · ${wdT(k.var)}${k.var > 0 && k.var <= r.amberVar ? ", within tolerance" : ""}`,
      k.bfinish != null && cov.with < cov.of && `baseline covers ${cov.with} of ${plural(cov.of, "task")}`].filter(Boolean).join(" · ");
    const msSub = !k.nMs ? "No milestones in plan" : k.nextMs ? `Next: ${cut(k.nextMs.name, 40)} · ${D(k.nextMs.finish)}` : k.nMsDone === k.nMs ? "All met" : "None upcoming";
    const nOpen = k.nTask - k.nTaskDone;
    return { type: "tiles", items: [
      { label: "Overall status", value: k.rag, sub: ragSub, tone: ragTone(k.rag), solid: true },
      { label: "Complete", value: pctT(k.pct), sub: k.nTask ? `Expected ${pctT(k.exp)} by now · ${k.nTaskDone} of ${plural(k.nTask, "task")} done` : "",
        tone: k.behind > r.redBehind ? "red" : k.behind > r.amberBehind ? "amber" : null },
      { label: "Forecast finish", value: F(k.finish), sub: finSub, tone: k.finish != null && old ? "red" : k.bfinish == null ? null : k.var > r.redVar ? "red" : k.var > r.amberVar ? "amber" : "green" },
      { label: "Milestones met", value: k.ms, sub: msSub, tone: k.overMs >= r.redMs ? "red" : k.overMs ? "amber" : null },
      { label: "Overdue tasks", value: String(k.overT), sub: [k.overMs && plural(k.overMs, "overdue milestone"), nOpen ? `of ${plural(nOpen, "open task")}` : "No open tasks"].filter(Boolean).join(" · "),
        tone: k.overT > r.amberOver ? "amber" : null },
    ] };
  }
  function dataCheck(k) {
    const c = k.checks || {}, sd = k.statusDate, open = openLeaves(k).length;
    const issues = [c.noFinish && `${plural(c.noFinish, "task")} without a finish date`, c.badDates && `${plural(c.badDates, "row")} with unreadable dates`,
      c.order && `${plural(c.order, "task")} finishing before ${c.order === 1 ? "its" : "their"} start`, c.mismatch && plural(c.mismatch, "status/% mismatch", "status/% mismatches"),
      c.capped && "only the first 1,000 rows read",
      k.rows && k.start != null && sd < k.start && `the status date ${F(sd)} is before the plan starts (${F(k.start)})`,
      k.rows && k.finish != null && sd > k.finish && open && `the status date ${F(sd)} is after the forecast finish (${F(k.finish)}) with ${plural(open, "task")} still open`].filter(Boolean);
    return { type: "note", text: (issues.length ? `Data check: ${plural(k.rows, "row")} read · ${issues.join(" · ")}.`
      : `Data check: ${plural(k.rows, "row")} read cleanly${k.nCancel ? ` · ${k.nCancel} cancelled left out` : ""}.`) + " Days late and variances are working days (Mon–Fri)." };
  }
  // the tolerances behind the RAG, from k.rules, so a reader can check the status (M1)
  function rulesText(k) {
    const r = k.rules || {};
    return `Status rules: Amber if the forecast finish is more than ${plural(r.amberVar, "working day")} after baseline, any milestone is overdue, more than ${plural(r.amberOver, "task")} ${r.amberOver === 1 ? "is" : "are"} overdue ` +
      `or progress is more than ${plural(r.amberBehind, "point")} behind plan; Red above ${r.redVar} days, ${plural(r.redMs, "overdue milestone")} or ${r.redBehind} points. Schedule only; cost and scope not assessed.`;
  }
  const rulesNote = k => ({ type: "note", text: rulesText(k) });
  const taskRows = (k, a, { phase, late } = {}) => a.map(t => compact([t.name, owner(t), phase ? t.phase || "—" : null, mkD(k)(t.finish),
    late ? { text: String(t.late), tone: "red", bold: true } : null]));
  const taskCols = ({ phase, late, due = "Due" } = {}) => compact([L("Task", phase ? "38%" : "50%"), L("Owner"), phase ? L("Phase") : null, M(due, "13%"), late ? R("Days late", "11%") : null]);
  function phaseNote(k) {
    if (!k.nPhases) return { type: "note", text: "No phases found: add a Phase column, WBS numbers or outline levels." };
    const n = k.nPhases - k.phOff - 8, t = [k.phOff ? `${plural(k.phOff, "earlier completed phase")} not shown` : "", n > 0 ? `+${plural(n, "more phase")} not shown` : ""].filter(Boolean).join(" · ");
    return t ? { type: "note", text: t + "." } : null;
  }
  function msNote(k) {
    return !k.nMs ? { type: "note", text: "No milestones found: add a 0-day task, a Milestone column, or a milestone keyword." }
      : k.msOpen > 8 ? { type: "note", text: `+${plural(k.msOpen - 8, "more open milestone")} not shown.` } : null;
  }
  const msTable = (k, ms, empty) => ({ type: "table", columns: [L("Milestone", "34%"), L("Phase", "22%"), M("Baseline", "12%"), M("Forecast", "12%"), R("Var", "8%"), M("Status", "12%")],
    rows: ms.map(m => [m.name, m.phase || "—", mkD(k)(m.bfin), mkD(k)(m.finish), m.var == null ? "—" : { text: wd(m.var), tone: m.var > 0 && !m.done ? "amber" : null }, chip(m.status)]), empty });
  function header(id, k, opts) {
    const meta = list.find(r => r.id === id), sd = k.statusDate;
    return { id, name: meta.name, title: opts.title || "Project status",
      subtitle: [opts.client, meta.name, id === "weekly" && sd != null ? `Period ${range(sd - (PERIOD - 1), sd)}` : "", opts.preparedBy ? `Prepared by ${opts.preparedBy}` : ""].filter(Boolean).join(" · "),
      asOf: E.iso(sd), rag: k.rag || "—", orientation: meta.orientation };
  }

  // ---------- reports ----------
  function onepager(k, opts) {
    // the chart holds at most 11 rows (8 phases + 3 milestones, or up to 6 milestones with fewer phases) to stay on one page
    const cm = commentary(k, { ...opts, compact: true }), N = Math.min(6, Math.max(3, 11 - Math.max(1, k.phases.length)));
    const open = k.msList.filter(m => m.status !== "Done"), met = k.msList.filter(m => m.status === "Done");
    const msTop = [...met.slice(Math.max(0, met.length - Math.max(0, N - open.length))), ...open.slice(0, N)], nOpen = Math.min(N, open.length);
    const tl = timelineSvg(k, msTop), pn = phaseNote(k), dc = dataCheck(k), n = k.msOpen - nOpen;
    const notes = [pn && pn.text, !k.nMs ? msNote(k).text : n > 0 ? `+${plural(n, "open milestone")} in the roadmap.` : ""].filter(Boolean);
    return compact([
      tiles(k),
      { type: "columns", columns: [[cm.highlights], [cm.concerns], [cm.next]] },
      cm.decisions,
      tl || { type: "note", text: "No dates found to draw a timeline." },
      { type: "note", text: [...notes, dc.text.replace(" Days late and variances are working days (Mon–Fri).", "").replace(/ read cleanly(.*)\.$/, ", no issues$1."), rulesText(k)].join(" ") }, // one footer note line: notes, data check, status rules
    ]);
  }

  function sinceLast(k, dl) {
    const D = mkD(k), p = dl.prev, cell = (n, goodUp, f = String) => n == null ? "—" : { text: n === 0 ? "No change" : n > 0 && f === String ? sgn(n) : f(n), tone: tone(n, goodUp) };
    const RAGS = ["Green", "Amber", "Red"];
    const rows = [
      ["Overall status", p.rag || "—", k.rag, p.rag && p.rag !== k.rag ? { text: `${p.rag} to ${k.rag}`, tone: RAGS.indexOf(k.rag) > RAGS.indexOf(p.rag) ? "red" : "green" } : "No change"],
      ["% complete", typeof p.pct === "number" ? pctT(p.pct) : "—", pctT(k.pct), cell(dl.pct, true, pts)],
      ["Forecast finish", dl.prevFinish != null ? D(dl.prevFinish) : "—", D(k.finish),
        dl.finish == null ? "—" : { text: dl.finish === 0 ? "No change" : dl.finish > 0 ? `${plural(dl.finish, "working day")} later` : `${plural(-dl.finish, "working day")} earlier`, tone: tone(dl.finish, false) }],
      ["Overdue tasks", p.overT ?? "—", k.overT, cell(dl.overT, false)],
      ["Overdue milestones", p.overMs ?? "—", k.overMs, cell(dl.overMs, false)],
      ["Milestones met", p.nMsDone ?? "—", k.nMsDone, cell(dl.nMsDone, true)],
    ].map(r => r.map((c, i) => i && typeof c === "number" ? String(c) : c));
    const later = dl.moved.filter(m => m.by > 0).length, earlier = dl.moved.filter(m => m.by < 0).length;
    const out = [{ type: "heading", text: `Since last report (${dl.asOf != null ? F(dl.asOf) : "previous"})` },
      { type: "table", columns: [L("Measure", "34%"), R("Last report", "22%"), R("This report", "22%"), R("Change", "22%")], rows }];
    if (dl.moved.length) {
      out.push({ type: "note", text: `${Plural(dl.moved.length, "task")} changed finish date since the last report: ${later} later, ${earlier} earlier.` },
        { type: "table", columns: [L("Task", "40%"), L("Owner"), M("Was", "14%"), M("Now", "14%"), R("Moved", "14%")],
          rows: dl.moved.slice(0, 10).map(m => [m.name, owner(m), D(m.was), D(m.finish), { text: wd(m.by), tone: m.by > 0 ? "red" : "green" }]) },
        more(dl.moved.length - 10, "date change"));
    }
    return compact(out);
  }

  function weekly(k, opts) {
    const D = mkD(k), cm = commentary(k, opts), dl = deltas(k, opts.prev), ph = allPhases(k), ms = allMs(k);
    const done7 = completedRecently(k), up = dueSoon(k), od = overdueTasks(k);
    const hist = trail(point({ asOf: E.iso(k.statusDate), rag: k.rag, pct: k.pct, var: k.var }), opts.history, opts.prev);
    const phCols = [L("Phase", "40%"), R("Progress", "14%"), M("Finish", "16%"), M("Status", "16%")];
    if (dl) phCols.push(R("Change", "14%"));
    return compact([
      tiles(k),
      ...(dl ? sinceLast(k, dl) : []),
      ...(hist.length > 1 && k.rows ? [{ type: "heading", text: `Trend over the last ${hist.length} reports` }, trendSvg(hist)] : []),
      { type: "heading", text: "Commentary" }, cm.highlights, cm.concerns, cm.next, cm.decisions,
      { type: "heading", text: "Key risks and issues" }, userList("risks", "Add the key risks and issues for this period, one per line.", opts.edits),
      { type: "heading", text: "Changes requested" }, userList("changes", "Add any change requests raised this period, one per line.", opts.edits),
      { type: "heading", text: `Completed in the last ${PERIOD} days` },
      { type: "table", columns: [L("Task", "44%"), L("Owner"), L("Phase"), M("Finished", "14%")],
        rows: done7.slice(0, CAP).map(t => [t.ms ? { text: t.name, bold: true } : t.name, owner(t), t.phase || "—", D(t.finish)]),
        empty: `Nothing was completed in the last ${PERIOD} days.` }, more(done7.length - CAP, "completed item"),
      { type: "heading", text: `Due in the next ${plural(k.upcomingDays, "day")}` },
      { type: "table", columns: taskCols({ phase: true }), rows: taskRows(k, up.slice(0, CAP), { phase: true }), empty: "Nothing is due in this window." },
      more(up.length - CAP, "task due soon", "tasks due soon"),
      { type: "heading", text: od.length ? `Overdue · ${od.length}` : "Overdue" },
      { type: "table", columns: taskCols({ phase: true, late: true }), rows: taskRows(k, od.slice(0, CAP), { phase: true, late: true }), empty: "Nothing is overdue." },
      more(od.length - CAP, "overdue task"),
      { type: "heading", text: k.nMs ? `Milestones · ${k.ms} met` : "Milestones" },
      msTable(k, ms, "No milestones found in the plan."),
      { type: "heading", text: "Phases" },
      { type: "table", columns: phCols, rows: ph.map(p => {
        const r = [p.name, pctT(p.pct), D(p.finish), chip(p.status)];
        if (dl) { const c = dl.phases[p.name]; r.push(c == null ? "New" : c === 0 ? "No change" : { text: pts(c), tone: c > 0 ? "green" : "red" }); }
        return r; }), empty: "No phases found: add a Phase column, WBS numbers or outline levels." },
      dataCheck(k), rulesNote(k),
    ]);
  }

  function roadmap(k) {
    const ms = allMs(k), svg = roadmapSvg(k), nm = findMs(k, k.nextMs);
    const sum = !k.rows ? "The plan has no rows to report on yet." : [
      k.nMs ? `${lead(k.nMsDone)} of ${plural(k.nMs, "milestone")} met.` : "No milestones found in the plan.",
      k.nextMs ? `Next: ${k.nextMs.name} on ${F(k.nextMs.finish)}` + (nm && nm.bfin != null && nm.var != null ? (nm.var > 0 ? `, ${plural(nm.var, "working day")} after its baseline.`
        : nm.var < 0 ? `, ${plural(-nm.var, "working day")} ahead of its baseline.` : ", on baseline.") : ".") : "",
      k.finish != null ? `Forecast finish ${F(k.finish)}${k.bfinish != null ? ` against a baseline of ${F(k.bfinish)} (${wdT(k.var)})` : ""}.` : ""].filter(Boolean).join(" ");
    const r = roadLanes(k), t = [r.off ? `${plural(r.off, "earlier completed phase")} not shown` : "", r.after ? `+${plural(r.after, "more phase")} not shown` : ""].filter(Boolean).join(" · ");
    return compact([
      { type: "text", id: "summary", text: sum, editable: false },
      svg,
      t ? { type: "note", text: t + "; all milestones are in the table below." } : null,
      { type: "heading", text: "Milestones" },
      msTable(k, ms, "No milestones found: add a 0-day task, a Milestone column, or a milestone keyword."),
      k.bfinish == null && ms.length ? { type: "note", text: "The plan has no baseline dates, so milestones are shown at their forecast dates only." } : null,
      rulesNote(k),
    ]);
  }

  function variance(k) {
    const D = mkD(k), tasks = (k.tasks || []).filter(t => !t.ms), ph = allPhases(k);
    const hasBase = (k.tasks || []).some(t => t.bfin != null);
    const owners = [...new Set(tasks.map(t => t.owner || ""))];
    const ownerRows = f => owners.map(o => ({ o, a: tasks.filter(t => (t.owner || "") === o) })).map(({ o, a }) => ({ o, ...f(a), open: a.filter(t => !t.done).length, over: a.filter(t => t.od).length }));
    const tstat = t => t.done ? chip("Done") : t.od ? chip("Overdue") : { text: "Open", chip: true };
    const nOpen = k.nTask - k.nTaskDone, ofOpen = nOpen ? `of ${plural(nOpen, "open task")}` : "No open tasks";
    if (!hasBase) {
      const od = overdueTasks(k), worst = od[0];
      const by = ownerRows(a => ({ worst: Math.max(0, ...a.filter(t => t.od).map(t => t.late)) })).filter(r => r.over)
        .sort((a, b) => b.worst - a.worst || b.over - a.over || a.o.localeCompare(b.o));
      return compact([
        { type: "tiles", items: [
          { label: "Finish variance", value: "—", sub: "No baseline in plan", tone: null },
          { label: "Overdue tasks", value: String(k.overT), sub: ofOpen, tone: k.overT > k.rules.amberOver ? "amber" : null },
          { label: "Worst delay", value: worst ? `${worst.late} wd` : "—", sub: worst ? `Working days late · ${cut(worst.name, 40)}` : "Nothing overdue", tone: worst ? "red" : null },
          { label: "Overdue milestones", value: String(k.overMs), sub: k.nMs ? `of ${plural(k.nMs, "milestone")}` : "No milestones in plan", tone: k.overMs ? "red" : null }] },
        { type: "note", text: "This plan has no baseline dates, so slippage against baseline cannot be measured. Delay is shown as working days late on overdue work instead. Add a Baseline Finish column to track slippage." },
        { type: "heading", text: "Overdue tasks" },
        { type: "table", columns: taskCols({ phase: true, late: true }), rows: taskRows(k, od.slice(0, SLIP_CAP), { phase: true, late: true }), empty: "Nothing is overdue." },
        more(od.length - SLIP_CAP, "overdue task"),
        { type: "heading", text: "By owner" },
        { type: "table", columns: [L("Owner", "40%"), R("Open tasks", "20%"), R("Overdue", "20%"), R("Worst delay", "20%")],
          rows: by.map(r => [r.o || "(no owner)", String(r.open), { text: String(r.over), tone: "red" }, `${r.worst} wd`]), empty: "No owner has overdue work." },
        { type: "heading", text: "By phase" },
        { type: "table", columns: [L("Phase", "40%"), M("Forecast finish", "18%"), R("Overdue", "14%"), R("Worst delay", "14%"), M("Status", "14%")],
          rows: ph.map(p => { const o = od.filter(t => t.phase === p.name); return [p.name, D(p.finish), String(o.length), o.length ? `${o[0].late} wd` : "—", chip(p.status)]; }),
          empty: "No phases found: add a Phase column, WBS numbers or outline levels." },
        dataCheck(k), rulesNote(k),
      ]);
    }
    const sl = slipped(k), openSl = sl.filter(t => !t.done), doneSl = sl.filter(t => t.done), worst = openSl[0], based = tasks.filter(t => t.bfin != null).length;
    const slMs = allMs(k).filter(m => !m.done && (m.var ?? 0) > 0);
    const by = ownerRows(a => { const s = a.filter(t => !t.done && (t.var ?? 0) > 0); return { slip: s.length, worst: Math.max(0, ...s.map(t => t.var)) }; })
      .filter(r => r.slip || r.over).sort((a, b) => b.worst - a.worst || b.over - a.over || b.slip - a.slip || a.o.localeCompare(b.o));
    return compact([
      { type: "tiles", items: [
        { label: "Finish variance", value: k.bfinish == null ? "—" : k.var === 0 ? "On baseline" : wd(k.var),
          sub: `Working days · forecast ${D(k.finish)}, baseline ${D(k.bfinish)}`, tone: k.var > k.rules.redVar ? "red" : k.var > k.rules.amberVar ? "amber" : "green" },
        { label: "Tasks slipped", value: String(sl.length), sub: `of ${plural(based, "task")} with a baseline · ${openSl.length} still open`, tone: openSl.length ? "amber" : null },
        { label: "Worst open slip", value: worst ? wd(worst.var) : "—", sub: worst ? cut(worst.name, 40) : "No open task behind baseline", tone: worst ? "amber" : null },
        { label: "Overdue tasks", value: String(k.overT), sub: [k.overMs && plural(k.overMs, "overdue milestone"), ofOpen].filter(Boolean).join(" · "),
          tone: k.overT > k.rules.amberOver ? "amber" : null }] },
      based < tasks.length ? { type: "note", text: `${lead(based)} of ${plural(tasks.length, "task")} ${based === 1 ? "has" : "have"} a baseline finish; the rest cannot be measured for slippage.` } : null,
      { type: "heading", text: "Open tasks behind baseline" },
      { type: "table", columns: [L("Task", "28%"), L("Owner", "15%"), L("Phase", "20%"), M("Baseline", "10%"), M("Forecast", "10%"), R("Slip", "8%"), M("Status", "9%")],
        rows: openSl.slice(0, SLIP_CAP).map(t => [t.name, owner(t), t.phase || "—", D(t.bfin), D(t.finish), { text: wd(t.var), tone: "amber", bold: true }, tstat(t)]),
        empty: "No open task is forecast to finish after its baseline." },
      more(openSl.length - SLIP_CAP, "open slipped task"),
      doneSl.length ? { type: "note", text: `${Plural(doneSl.length, "completed task")} finished after baseline (worst +${plural(doneSl[0].var, "working day")}).` } : null,
      slMs.length ? { type: "note", text: `${Plural(slMs.length, "open milestone")} also forecast after baseline: ${and(slMs.slice(0, 4).map(m => `${m.name} (+${plural(m.var, "working day")})`).concat(slMs.length > 4 ? [`${slMs.length - 4} more`] : []))}.` } : null,
      { type: "heading", text: "By owner" },
      { type: "table", columns: [L("Owner", "36%"), R("Open tasks", "16%"), R("Overdue", "16%"), R("Open slipped", "16%"), R("Worst slip", "16%")],
        rows: by.map(r => [r.o || "(no owner)", String(r.open), r.over ? { text: String(r.over), tone: "red" } : "0", String(r.slip), r.worst ? wd(r.worst) : "—"]),
        empty: "No owner has open slipped or overdue work." },
      { type: "heading", text: "By phase" },
      { type: "table", columns: [L("Phase", "36%"), M("Baseline finish", "17%"), M("Forecast finish", "17%"), R("Variance", "14%"), M("Status", "16%")],
        rows: ph.map(p => { const v = p.var ?? (p.finish != null && p.bfin != null ? wdiff(p.bfin, p.finish) : null);
          return [p.name, D(p.bfin), D(p.finish), v == null ? "—" : { text: wd(v), tone: v > 0 ? "amber" : null }, chip(p.status)]; }),
        empty: "No phases found: add a Phase column, WBS numbers or outline levels." },
      dataCheck(k), rulesNote(k),
    ]);
  }

  const BUILD = { onepager, weekly, roadmap, variance };
  function build(id, k, opts) {
    if (!BUILD[id]) throw new Error("Unknown report: " + id);
    opts = opts || {};
    return { ...header(id, k, opts), blocks: BUILD[id](k, opts) };
  }

  root.KeelsonReports = { list, build, snapshot, commentary, deltas, allPhases, completedRecently, esc, DEFAULT_MS_KEYS };
  if (typeof module !== "undefined") module.exports = root.KeelsonReports;
})(typeof window !== "undefined" ? window : globalThis);
