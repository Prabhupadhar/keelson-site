/* Keelson dashboard engine for the browser: the same rules as the Excel workbook (src/dashboard.py), run on pasted
   text. Pure functions, no network. tests/web/engine.test.js checks it against the Python reference on five plan shapes. */
(function (root) {
  const FIELDS = [ // key, label, required, header names (first match wins; a column claimed earlier is not reused)
    ["name", "Task name", true, ["Task Name", "Name", "Task", "Activity", "Activity Name", "Task Title", "Title", "Deliverable", "Work Item", "Item", "Milestone / Task", "Description", "Summary"]],
    ["finish", "Finish / due date", true, ["Finish", "Finish Date", "Finish_Date", "End", "End Date", "Due", "Due Date", "Due date", "Forecast Finish", "Scheduled Finish", "Planned Finish", "Planned End", "Target Date", "Target End", "Deadline", "Completion Date", "Date"]],
    ["start", "Start date", false, ["Start", "Start Date", "Start_Date", "Begin", "Begin Date", "Forecast Start", "Scheduled Start", "Planned Start"]],
    ["bfin", "Baseline finish", false, ["Baseline Finish", "Baseline_Finish", "Baseline Finish Date", "Baseline End", "Baseline Due", "Baseline", "Original Finish", "Original End", "Original Due Date", "Original Due"]],
    ["pct", "% complete", false, ["% Complete", "Percent Complete", "Percent_Complete", "% Done", "Pct Complete", "Complete %", "Completion %", "Progress", "Completion", "% Work Complete"]],
    ["status", "Status", false, ["Status", "Task Status", "State", "Progress Status", "Completed At", "Completed"]],
    ["owner", "Owner", false, ["Owner", "Resource Names", "Resource_Names", "Resources", "Assigned To", "Assignee", "Responsible", "Task Owner", "Lead", "Person", "Accountable"]],
    ["phase", "Phase / workstream", false, ["Phase", "Workstream", "Work Stream", "Stage", "Stream", "Epic Link", "Epic Name", "Epic", "Section", "Section/Column", "Swimlane", "Category"]],
    ["wbs", "WBS / outline number", false, ["WBS", "WBS Code", "Outline Number", "Outline_Number", "WBS ID", "Ref", "No.", "#"]],
    ["lvl", "Outline level", false, ["Outline Level", "Outline_Level", "Level", "Indent", "Hierarchy Level"]],
    ["dur", "Duration", false, ["Duration", "Duration (days)", "Dur", "Days", "Work Days", "Planned Duration"]],
    ["ms", "Milestone / type", false, ["Milestone", "Is Milestone", "Milestone?", "Milestone Flag", "Type", "Task Type", "Issue Type", "Item Type"]],
    ["sumf", "Summary flag", false, ["Summary", "Is Summary", "Summary Task"]],
  ];
  const DONE = ["Complete", "Completed", "Done", "Closed", "Finished", "Resolved", "Delivered", "Achieved", "Accepted", "Signed off", "Released", "Live"].map(s => s.toLowerCase());
  const CANCEL = ["Cancelled", "Canceled", "Won't Do", "Wont Do", "Won't Fix", "Dropped", "Descoped", "Removed", "Not Required", "Duplicate", "N/A"].map(s => s.toLowerCase());
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const RULES = { amberVar: 5, redVar: 15, amberOver: 3, redMs: 2, amberBehind: 10, redBehind: 20 };
  const N_PHASE = 8, N_MS = 8, N_LIST = 6, HDR_SCAN = 10, CAP = 1000, DAY = 864e5;

  // dates are whole days since 1970-01-01 (UTC), so arithmetic never meets daylight saving
  const fromYMD = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / DAY);
  const ymd = v => { const d = new Date(v * DAY); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; };
  const iso = v => v == null ? null : ymd(v).map((x, i) => String(x).padStart(i ? 2 : 4, "0")).join("-");
  const fmt = (v, year) => { if (v == null) return ""; const [y, m, d] = ymd(v); return `${d} ${MON[m - 1]}` + (year ? ` ${y}` : ""); };
  const wd = v => (new Date(v * DAY).getUTCDay() + 6) % 7; // 0 = Monday
  const nwd = (a, b) => { // Excel NETWORKDAYS, a <= b, Monday to Friday
    if (b < a) return -nwd(b, a);
    const n = b - a + 1, full = Math.floor(n / 7);
    let c = full * 5;
    for (let i = 0, s = wd(a + full * 7); i < n % 7; i++) if ((s + i) % 7 < 5) c++;
    return c;
  };
  // signed working days from a to b, counting b but not a (0 when equal, also when both fall on a weekend)
  const wdiff = (a, b) => a === b ? 0 : b > a ? nwd(a + 1, b) : -nwd(b + 1, a);
  const isNum = s => /^-?\d+(\.\d+)?$/.test(s);
  const low = s => String(s ?? "").trim().toLowerCase();
  const truthy = s => ["yes", "y", "true", "1", "x"].includes(low(s));

  function parseTable(text) {
    // control characters become spaces (tab, LF and CR are kept: they are the separators)
    const lines = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, " ").replace(/\r\n?/g, "\n").replace(/\n+$/, "").split("\n");
    const head = lines.slice(0, 12).join("\n");
    const sep = head.includes("\t") ? "\t" : (head.split(";").length > head.split(",").length ? ";" : ",");
    if (sep === "\t") return lines.map(l => l.split("\t").map(c => c.trim()));
    const rows = []; let row = [], cell = "", q = false; // RFC 4180 CSV, quotes may span lines
    const src = lines.join("\n");
    for (let i = 0; i < src.length; i++) {
      const ch = src[i];
      if (q) { if (ch === '"') { if (src[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === sep) { row.push(cell.trim()); cell = ""; }
      else if (ch === "\n") { row.push(cell.trim()); rows.push(row); row = []; cell = ""; }
      else cell += ch;
    }
    row.push(cell.trim()); rows.push(row);
    return rows;
  }

  // header row = the row (1-10) with most recognised header names; columns by synonym order
  function detect(grid) {
    const syn = new Set(FIELDS.flatMap(f => f[3].map(low)));
    let best = 0, hdr = 0;
    for (let i = 0; i < Math.min(HDR_SCAN, grid.length); i++) {
      const s = grid[i].filter(c => syn.has(low(c))).length;
      if (s > best) { best = s; hdr = i; }
    }
    const heads = grid[hdr] || [], cols = {}, used = [];
    for (const [k, , , names] of FIELDS) {
      let det = -1;
      for (const n of names) { const j = heads.findIndex(h => low(h) === low(n)); if (j >= 0) { det = j; break; } }
      cols[k] = det >= 0 && !used.includes(det) ? det : -1;
      used.push(cols[k]);
    }
    return { hdr, heads, cols };
  }

  // one text date -> [y, m, d] parts and day-first evidence, like the workbook's _t/_s/_a/_b/_c columns
  function dateParts(raw) {
    let t = String(raw ?? "").trim();
    if (!t) return null;
    if (isNum(t) && +t > 20000 && +t < 80000) return { serial: Math.floor(+t) - 25569 }; // Excel serial
    if (WEEKDAYS.includes(t.slice(0, 3).toLowerCase())) t = t.slice((t + " ").indexOf(" ") + 1).trim();
    let s = /[\/\-.]/.test(t) ? t.split(" ")[0] : t;
    s = s.replace(/[ \-.]/g, "/").replace(/,/g, "");
    const p = s.split("/");
    if (p.length < 2) return { bad: true };
    const [a, b] = p, c = (p.slice(2).join("/") || "").slice(0, 4);
    const mon = x => { const i = MON.findIndex(m => m.toLowerCase() === x.slice(0, 3).toLowerCase()); return i < 0 ? NaN : i + 1; };
    const ev = a.length !== 4 && isNum(a) && isNum(b) ? (+a > 12 && +b <= 12 ? 1 : +b > 12 && +a <= 12 ? -1 : 0) : 0;
    return { a, b, c, mon, ev };
  }
  function dateValue(p, dayFirst) {
    if (!p || p.bad) return null;
    if (p.serial != null) return p.serial;
    const { a, b, c, mon } = p;
    let y, m, d;
    if (a.length === 4) { y = +a; m = isNum(b) ? +b : mon(b); d = parseInt(c.slice(0, 2), 10); }
    else {
      const yy = parseInt(c.slice(0, 4), 10); y = yy < 100 ? 2000 + yy : yy;
      if (!isNum(a)) { m = mon(a); d = +b; } else if (!isNum(b)) { m = mon(b); d = +a; }
      else if (dayFirst) { m = +b; d = +a; } else { m = +a; d = +b; }
    }
    if (![y, m, d].every(Number.isFinite) || m < 1 || m > 12 || d < 1 || d > 31 || y < 1950 || y > 2200) return null;
    return fromYMD(y, m, d);
  }

  function compute(text, opt = {}) {
    const grid = parseTable(text || "");
    const det = detect(grid);
    const cols = { ...det.cols, ...(opt.cols || {}) };
    const sd = opt.statusDate ?? Math.floor(Date.now() / DAY);
    const rules = { ...RULES, ...(opt.rules || {}) };
    const upcoming = opt.upcomingDays ?? 14;
    const keys = (opt.msKeys || []).map(low).filter(Boolean);
    const body = grid.slice(det.hdr + 1, det.hdr + 1 + CAP);
    const capped = grid.length - det.hdr - 1 > CAP;
    const get = (r, k) => cols[k] >= 0 ? (r[cols[k]] ?? "").trim() : "";

    const R = body.map(r => {
      const o = {}; FIELDS.forEach(([k]) => o["r_" + k] = get(r, k)); return o;
    });
    let evDay = 0, evMon = 0;
    for (const o of R) for (const k of ["start", "finish", "bfin"]) {
      o["p_" + k] = dateParts(o["r_" + k]);
      const e = o["p_" + k]?.ev || 0; if (e > 0) evDay++; if (e < 0) evMon++;
    }
    const dayFirst = opt.dateOrder === "day" ? true : opt.dateOrder === "month" ? false : evDay > evMon;
    let maxPa = 0;
    for (const o of R) {
      const st = o.r_status;
      o.st = !st ? "" : dateValue(dateParts(st), dayFirst) != null ? "Done" : DONE.includes(low(st)) ? "Done" : CANCEL.includes(low(st)) ? "Cancelled" : "Open";
      o.nm = o.st === "Cancelled" ? "" : o.r_name.trim();
      o.act = o.nm !== "";
      for (const k of ["start", "finish", "bfin"]) o[k] = dateValue(o["p_" + k], dayFirst);
      const pa = o.r_pct === "" ? null : parseFloat(o.r_pct.replace("%", ""));
      o.pa = Number.isFinite(pa) ? pa : null;
      if (o.pa != null) maxPa = Math.max(maxPa, o.pa);
      o.lv = !o.act ? null : cols.lvl >= 0 ? (isNum(o.r_lvl) ? +o.r_lvl : null) : o.r_wbs ? o.r_wbs.split(".").length : null;
      const dm = o.r_dur.replace(/d/g, " ").split(" ")[0];
      o.dn = o.r_dur === "" ? null : isNum(dm) ? +dm : null;
    }
    const scale = maxPa > 1.0001 ? 100 : 1;
    R.forEach((o, i) => {
      const nx = R[i + 1];
      o.pct = o.pa == null ? null : Math.min(1, Math.max(0, o.pa / scale));
      o.done = o.act && (o.st === "Done" || (o.pct ?? 0) >= 1);
      o.pe = o.done ? 1 : (o.pct ?? 0);
      const msc = low(o.r_ms);
      o.sum = !o.act ? false : truthy(o.r_sumf) || msc.includes("summary") || msc === "epic" ? true
        : !nx ? false
        : cols.lvl >= 0 ? (o.lv != null && nx.lv != null && nx.lv > o.lv)
        : !!(o.r_wbs && nx.r_wbs && nx.r_wbs.startsWith(o.r_wbs + "."));
      o.ms = o.act && !o.sum && (truthy(o.r_ms) || msc.includes("milestone") || o.dn === 0 || keys.some(k => low(o.nm).includes(k)));
      o.leaf = o.act && !o.sum; o.isms = o.leaf && o.ms; o.istask = o.leaf && !o.ms;
      o.row = i;
    });
    const act = R.filter(o => o.act);
    const lvs = act.map(o => o.lv).filter(v => v != null);
    const minL = lvs.length ? Math.min(...lvs) : null;
    const topL = minL == null ? null : lvs.filter(v => v === minL).length === 1 ? minL + 1 : minL;
    let ph = "";
    for (const o of R) {
      if (o.act) {
        if (cols.phase >= 0) { if (o.r_phase.trim()) ph = o.r_phase.trim(); }
        else if (o.lv != null && o.lv === topL) ph = o.nm;
      }
      o.ph = ph;
      const s = o.start, f = o.finish;
      o.w = s != null && f != null ? (f >= s ? Math.max(1, nwd(s, f)) : 1) : 1;
      o.e = s != null && f != null ? (sd < s ? 0 : sd >= f ? 1 : Math.min(1, nwd(s, sd) / o.w)) : (f != null && sd >= f ? 1 : 0);
      o.od = o.leaf && !o.done && f != null && f < sd;
      o.late = o.od ? Math.max(1, wdiff(f, sd)) : null;
      o.var = f != null && o.bfin != null ? wdiff(o.bfin, f) : null;
    }
    const leaf = R.filter(o => o.leaf), tasks = leaf.filter(o => o.istask), ms = leaf.filter(o => o.isms);
    const sum = (a, f) => a.reduce((t, o) => t + f(o), 0);
    const ww = sum(tasks, o => o.w);
    const k = { rows: act.length, cols, heads: det.heads, hdrRow: det.hdr + 1, dayFirst, evidence: evDay + evMon, statusDate: sd };
    k.pct = ww ? sum(tasks, o => o.w * o.pe) / ww : leaf.length ? leaf.filter(o => o.done).length / leaf.length : 0;
    k.exp = ww ? sum(tasks, o => o.w * o.e) / ww : 0;
    const fins = leaf.map(o => o.finish).filter(v => v != null), bfs = leaf.map(o => o.bfin).filter(v => v != null);
    const starts = leaf.map(o => o.start).filter(v => v != null);
    k.finish = fins.length ? Math.max(...fins) : null;
    k.bfinish = bfs.length ? Math.max(...bfs) : null;
    k.start = starts.length ? Math.min(...starts) : fins.length ? Math.min(...fins) : null;
    k.var = k.finish == null || k.bfinish == null ? 0 : wdiff(k.bfinish, k.finish);
    k.behind = ww ? Math.round(k.exp * 100) - Math.round(k.pct * 100) : 0; // the printed % figures, subtracted
    k.overT = tasks.filter(o => o.od).length; k.overMs = ms.filter(o => o.od).length;
    k.nTask = tasks.length; k.nTaskDone = tasks.filter(o => o.done).length;
    k.nMs = ms.length; k.nMsDone = ms.filter(o => o.done).length;
    k.ms = ms.length ? `${k.nMsDone} of ${ms.length}` : "—";
    k.nCancel = R.filter(o => o.st === "Cancelled").length;
    const red = k.var > rules.redVar || k.overMs >= rules.redMs || k.behind > rules.redBehind;
    const amber = k.var > rules.amberVar || k.overMs >= 1 || k.overT > rules.amberOver || k.behind > rules.amberBehind;
    const noFin = leaf.length - fins.length; // integer tests: > 20% is noFin * 5 > leaves, > 50% is noFin * 2 > leaves
    const rule = red ? "Red" : amber ? "Amber" : "Green";
    if (!act.length) { k.ragCalc = "—"; k.why = []; }
    else if (!fins.length || noFin * 2 > leaf.length) { k.ragCalc = "—"; k.why = ["Not enough dated tasks to assess status"]; }
    else {
      k.ragCalc = rule === "Green" && noFin * 5 > leaf.length ? "Amber" : rule;
      k.why = rule === "Green" ? [k.ragCalc === "Amber" ? `${noFin} tasks have no finish date` : "All rules within tolerance"] : [
        k.var > rules.amberVar && `Finish ${k.var} days late`,
        k.overMs > 0 && `${k.overMs} milestone${k.overMs > 1 ? "s" : ""} overdue`,
        k.overT > rules.amberOver && `${k.overT} tasks overdue`,
        k.behind > rules.amberBehind && `${k.behind} pts behind plan`].filter(Boolean);
      if (!bfs.length) k.why.push("No baseline: finish slippage not assessed");
    }
    const ov = opt.ragOverride;
    const useOv = act.length && ov && ["Red", "Amber", "Green"].includes(ov.rag);
    k.rag = useOv ? ov.rag : k.ragCalc;
    k.ragReason = useOv ? String(ov.reason ?? "") : null;
    const openLate = leaf.filter(o => o.od); // open leaves due before the status date: their forecast is stale
    k.stale = { count: openLate.length, oldest: openLate.length ? Math.min(...openLate.map(o => o.finish)) : null };
    k.finishPassed = k.finish != null && k.finish < sd && leaf.some(o => !o.done);
    k.bCover = { with: bfs.length, of: leaf.length };

    // phases in order of first leaf; leading completed phases skipped when there are more than 8
    const order = [];
    for (const o of leaf) if (o.ph && !order.includes(o.ph)) order.push(o.ph);
    const open = order.map(p => leaf.some(o => o.ph === p && !o.done));
    const lead = open.indexOf(true) >= 0 ? open.indexOf(true) : 0;
    k.phOff = Math.max(0, Math.min(lead, order.length - N_PHASE));
    k.nPhases = order.length;
    const mx = a => a.length ? Math.max(...a) : null, mn = a => a.length ? Math.min(...a) : null;
    k.phases = order.slice(k.phOff, k.phOff + N_PHASE).map(p => {
      const pl = leaf.filter(o => o.ph === p), pt = pl.filter(o => o.istask), pw = sum(pt, o => o.w);
      const pct = pw ? sum(pt, o => o.w * o.pe) / pw : pl.filter(o => o.done).length / Math.max(1, pl.length);
      const fin = mx(pl.map(o => o.finish).filter(v => v != null)), bfin = mx(pl.map(o => o.bfin).filter(v => v != null));
      const st = mn(pl.map(o => o.start).filter(v => v != null));
      const exp = pw ? sum(pt, o => o.w * o.e) / pw : 0;
      const pvar = fin != null && bfin != null ? wdiff(bfin, fin) : null;
      const behind = pw ? Math.round(exp * 100) - Math.round(pct * 100) : 0;
      const status = pl.every(o => o.done) ? "Done" : pl.some(o => o.od) ? "Late"
        : (pvar ?? 0) > rules.amberVar || behind > rules.amberBehind ? "At risk"
        : pct === 0 && st != null && st > sd ? "Planned" : "On track";
      return { name: p, pct, exp, var: pvar, behind, start: st, finish: fin, status, ms: pl.filter(o => o.isms && o.finish != null).map(o => o.finish) };
    });
    const by = (a, f) => [...a].sort((x, y) => f(x) - f(y) || x.row - y.row);
    const openMs = by(ms.filter(o => !o.done && o.finish != null), o => o.finish);
    const doneMs = by(ms.filter(o => o.done && o.finish != null), o => o.finish);
    const dn = Math.min(doneMs.length, Math.max(0, N_MS - openMs.length));
    k.msOpen = openMs.length;
    k.msList = [...doneMs.slice(doneMs.length - dn), ...openMs].slice(0, N_MS).map(o => ({
      name: o.nm, finish: o.finish, status: o.done ? "Done" : o.finish < sd ? "Late" : (o.var ?? 0) > 0 ? `+${o.var} days` : "On track" }));
    const next = openMs.find(o => o.finish >= sd);
    k.nextMs = next ? { name: next.nm, finish: next.finish } : null;
    const item = o => ({ name: o.nm, owner: o.r_owner, finish: o.finish, late: o.late });
    const overdue = by(tasks.filter(o => o.od), o => -o.late);
    k.overdue = overdue.slice(0, N_LIST).map(item); k.overdueMore = Math.max(0, overdue.length - N_LIST);
    const up = by(tasks.filter(o => !o.done && o.finish != null && o.finish >= sd && o.finish <= sd + upcoming), o => o.finish);
    k.upcoming = up.slice(0, N_LIST).map(item); k.upcomingMore = Math.max(0, up.length - N_LIST); k.upcomingDays = upcoming;
    // every leaf row, for report builders (dates are day numbers; see fromYMD/iso)
    const seen = {};
    k.tasks = leaf.map(o => ({ key: (b => (seen[b] = (seen[b] || 0) + 1) > 1 ? `${b}#${seen[b]}` : b)(o.nm + "|" + o.ph), row: o.row, name: o.nm, owner: o.r_owner, phase: o.ph, start: o.start, finish: o.finish, bfin: o.bfin,
      pct: o.pct, done: o.done, ms: o.isms, od: o.od, late: o.late, var: o.var, w: o.w }));
    k.rules = rules;
    k.checks = {
      noFinish: leaf.filter(o => o.finish == null).length,
      badDates: act.filter(o => ["start", "finish", "bfin"].some(f => o["r_" + f] !== "" && o[f] == null)).length,
      order: leaf.filter(o => o.start != null && o.finish != null && o.finish < o.start).length,
      mismatch: cols.status >= 0 && cols.pct >= 0 ? leaf.filter(o => o.pct != null && ((o.st === "Done" && o.pct < 1) || (o.st === "Open" && o.pct >= 1))).length : 0,
      capped };
    return k;
  }

  const parseDate = (t, dayFirst) => dateValue(dateParts(t), dayFirst);
  root.KeelsonEngine = { compute, parseTable, detect, parseDate, FIELDS, RULES, fromYMD, iso, fmt, nwd, wdiff, wd, DAY };
  if (typeof module !== "undefined") module.exports = root.KeelsonEngine;
})(typeof window !== "undefined" ? window : globalThis);
