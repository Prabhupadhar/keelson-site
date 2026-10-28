/* Keelson dashboard engine for the browser: the same rules as the Excel workbook (src/dashboard.py), run on pasted
   text. Pure functions, no network. tests/web/engine.test.js checks it against the Python reference on five plan shapes. */
(function (root) {
  const FIELDS = [ // key, label, required, header names (first match wins; a column claimed earlier is not reused)
    ["name", "Task name", true, ["Task Name", "Name", "Task", "Activity", "Activity Name", "Task Title", "Title", "Deliverable", "Work Item", "Item", "Milestone / Task", "Description", "Summary", "Work Item Title", "Item Name", "Subject", "Issue", "Story", "Ticket", "Card Name", "Action", "Step", "Requirement", "Feature", "Line Item", "Activity Description", "Task Description", "Deliverable Name", "Milestone Name"]],
    ["finish", "Finish / due date", true, ["Finish", "Finish Date", "Finish_Date", "End", "End Date", "Due", "Due Date", "Due date", "Forecast Finish", "Scheduled Finish", "Planned Finish", "Planned End", "Target Date", "Target End", "Deadline", "Completion Date", "Date", "Due On", "Due By", "End Time", "Finish Time", "Target Finish", "Target Completion", "Planned Completion", "Planned End Date", "Planned Finish Date", "Expected Finish", "Expected End", "Estimated Finish", "Est Finish", "Forecast End", "Forecast Date", "Forecast Finish Date", "Committed Date", "Delivery Date", "Timeline End", "End On", "Finish By", "Closing Date"]],
    ["start", "Start date", false, ["Start", "Start Date", "Start_Date", "Begin", "Begin Date", "Forecast Start", "Scheduled Start", "Planned Start", "Planned Start Date", "Start Time", "Start On", "Timeline Start", "Expected Start", "Estimated Start", "Target Start", "Kick-off", "Kickoff", "Kick off Date"]],
    ["bfin", "Baseline finish", false, ["Baseline Finish", "Baseline_Finish", "Baseline Finish Date", "Baseline End", "Baseline Due", "Baseline", "Original Finish", "Original End", "Original Due Date", "Original Due", "Baseline1 Finish", "BL Finish", "Original Target Date", "Initial Due Date", "Baseline End Date", "Baseline Due Date", "BL Project Finish"]],
    ["pct", "% complete", false, ["% Complete", "Percent Complete", "Percent_Complete", "% Done", "Pct Complete", "Complete %", "Completion %", "Progress", "Completion", "% Work Complete", "Percent Done", "% Progress", "Progress %", "Percentage Complete", "Percent", "Done %", "PctComplete", "Complete", "Physical % Complete"]],
    ["status", "Status", false, ["Status", "Task Status", "State", "Progress Status", "Completed At", "Completed", "Workflow State", "Item Status", "Issue Status", "Current Status", "Status Category", "Is Done", "Done", "Completed?"]],
    ["owner", "Owner", false, ["Owner", "Resource Names", "Resource_Names", "Resources", "Assigned To", "Assignee", "Responsible", "Task Owner", "Lead", "Person", "Accountable", "Assignees", "Assigned", "Owner Name", "Resource", "Resource Name", "Assigned Resource", "Responsible Person", "Team Member", "Action Owner", "Task Assignee", "People", "Who", "Assignee Name", "Owners", "Members"]],
    ["phase", "Phase / workstream", false, ["Phase", "Workstream", "Work Stream", "Stage", "Stream", "Epic Link", "Epic Name", "Epic", "Section", "Section/Column", "Swimlane", "Category", "Workstream Name", "Work Package", "Project Phase", "Phase Name", "Stage Name", "Parent", "Parent Name", "Parent Summary", "Epic Summary", "Group", "List", "List Name", "Area Path", "Track", "Stream Name"]],
    ["wbs", "WBS / outline number", false, ["WBS", "WBS Code", "Outline Number", "Outline_Number", "WBS ID", "Ref", "No.", "#", "WBS Number", "WBS No", "S.No", "Sl No", "Item No", "Outline"]],
    ["lvl", "Outline level", false, ["Outline Level", "Outline_Level", "Level", "Indent", "Hierarchy Level", "Indent Level", "Depth"]],
    ["dur", "Duration", false, ["Duration", "Duration (days)", "Dur", "Days", "Work Days", "Planned Duration", "Original Duration", "Estimated Duration", "Duration Days", "Planned Days"]],
    ["ms", "Milestone / type", false, ["Milestone", "Is Milestone", "Milestone?", "Milestone Flag", "Type", "Task Type", "Issue Type", "Item Type", "Is a Milestone", "Key Milestone", "Work Item Type", "Activity Type", "Record Type"]],
    ["sumf", "Summary flag", false, ["Summary", "Is Summary", "Summary Task", "Summary Row", "Is Parent"]],
    // v2 (plan diff, history, forecast): last, so they never take a column from the fields above
    ["uid", "Task ID", false, ["Unique ID", "Unique_ID", "UID", "Issue key", "Issue Key", "Key", "ID", "Task ID", "Issue id", "Work Item ID", "Activity ID", "Row ID", "Item ID", "Ticket ID", "Task Number"]],
    ["afin", "Actual finish", false, ["Actual Finish", "Actual_Finish", "Actual End", "Actual Finish Date", "Resolved", "Resolution Date", "Completed Date", "Done Date", "Actual Completion", "Completed On", "Closed Date", "Date Completed", "Actual End Date", "Date Closed", "Finished At", "Resolved Date", "Done On", "Closed"]],
    ["bstart", "Baseline start", false, ["Baseline Start", "Baseline_Start", "Baseline Start Date", "Original Start", "Baseline1 Start", "BL Start", "Original Start Date", "BL Project Start"]],
  ];
  const DONE = ["Complete", "Completed", "Done", "Closed", "Finished", "Resolved", "Delivered", "Achieved", "Accepted", "Signed off", "Released", "Live", "Yes", "Y", "True", "✓", "✔", "☑"].map(s => s.toLowerCase());
  const CANCEL = ["Cancelled", "Canceled", "Won't Do", "Wont Do", "Won't Fix", "Dropped", "Descoped", "Removed", "Not Required", "Duplicate", "N/A"].map(s => s.toLowerCase());
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  // milestone keywords the app passes by default (KeelsonReports.DEFAULT_MS_KEYS re-exports this list)
  const DEFAULT_MS_KEYS = ["go-live", "go live", "sign-off", "signed off", "sign off", "approval", "approved", "decision", "closed", "stage gate", "phase complete"];
  const RULES = { amberVar: 5, redVar: 15, amberOver: 3, redMs: 2, amberBehind: 10, redBehind: 20 };
  const N_PHASE = 8, N_MS = 8, N_LIST = 6, HDR_SCAN = 10, CAP = 1000, DAY = 864e5;

  // dates are whole days since 1970-01-01 (UTC), so arithmetic never meets daylight saving
  const fromYMD = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / DAY);
  const ymd = v => { const d = new Date(v * DAY); return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()]; };
  const iso = v => v == null ? null : ymd(v).map((x, i) => String(x).padStart(i ? 2 : 4, "0")).join("-");
  const fmt = (v, year) => { if (v == null) return ""; const [y, m, d] = ymd(v); return `${d} ${MON[m - 1]}` + (year ? ` ${y}` : ""); };
  const wd = v => (new Date(v * DAY).getUTCDay() + 6) % 7; // 0 = Monday (a weekday label, not a working-day test)

  // work calendar: one module-level calendar that every working-day helper consults (default Sat+Sun weekend, no holidays)
  const WEEKENDS = { "sat-sun": [6, 0], "fri-sat": [5, 6], "thu-fri": [4, 5], "sun": [0], "fri": [5] }; // JS getDay numbers, 0 = Sun
  const dow = v => ((v + 4) % 7 + 7) % 7; // getUTCDay without a Date: 1970-01-01 was a Thursday
  let CAL, OFF, HOL, HOLS; // OFF[getDay] = 1 on weekend days; HOL = holiday day numbers; HOLS = sorted holidays on working days
  function setCalendar(c = {}) {
    const we = c.weekend == null ? [6, 0] : c.weekend;
    if (!Array.isArray(we) || we.length > 3 || new Set(we).size !== we.length || !we.every(d => Number.isInteger(d) && d >= 0 && d <= 6))
      throw new Error("Weekend must be 0 to 3 different days, each 0 (Sunday) to 6 (Saturday).");
    const hs = c.holidays == null ? [] : c.holidays;
    if (!Array.isArray(hs)) throw new Error("Holidays must be a list of YYYY-MM-DD dates.");
    const lo = fromYMD(1950, 1, 1), hi = fromYMD(2200, 12, 31), set = new Set();
    for (const h of hs) {
      if (set.size >= 500 || typeof h !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(h)) continue;
      const v = fromYMD(...h.split("-").map(Number));
      if (iso(v) === h && v >= lo && v <= hi) set.add(v);
    }
    const off = new Uint8Array(7); for (const d of we) off[d] = 1;
    OFF = off; HOL = set; HOLS = [...set].filter(v => !off[dow(v)]).sort((a, b) => a - b);
    CAL = { weekend: we.slice(), holidays: [...set].sort((a, b) => a - b).map(iso), ignored: hs.length - set.size };
    return calendar();
  }
  const calendar = () => ({ weekend: CAL.weekend.slice(), holidays: CAL.holidays.slice(), ignored: CAL.ignored });
  const isWd = v => !OFF[dow(v)] && !(HOLS.length && HOL.has(v));
  // the working day n working days after d (d itself counts as day 0, whatever it is)
  const addWd = (d, n) => { for (; n > 0;) { d++; if (isWd(d)) n--; } return d; };
  const lowerB = (x) => { let l = 0, h = HOLS.length; while (l < h) { const m = (l + h) >> 1; if (HOLS[m] < x) l = m + 1; else h = m; } return l; };
  // the working week for report text: "Mon–Fri" (long: "Monday to Friday"), "Sun–Thu, 12 holidays", or a list when not one run
  const DN = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], DL = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  function calText(long) {
    const N = long ? DL : DN, on = [0, 1, 2, 3, 4, 5, 6].filter(d => !OFF[d]);
    let t;
    if (on.length === 7) t = long ? "every day of the week" : "all 7 days";
    else {
      const first = on.find(d => OFF[(d + 6) % 7]); // a working day after a weekend day
      const run = on.map((_, i) => (first + i) % 7), week = DN.map((_, i) => (first + i) % 7).filter(d => !OFF[d]);
      t = run.every(d => !OFF[d]) ? `${N[run[0]]}${long ? " to " : "–"}${N[run[run.length - 1]]}` : week.map(d => N[d]).join(", ");
    }
    const n = CAL.holidays.length;
    return n ? `${t}, ${n} holiday${n === 1 ? "" : "s"}` : t;
  }
  const nwd = (a, b) => { // Excel NETWORKDAYS(a, b, holidays) on the active calendar, a <= b
    if (b < a) return -nwd(b, a);
    const n = b - a + 1, full = Math.floor(n / 7);
    let c = full * (7 - CAL.weekend.length);
    for (let i = 0, s = dow(a + full * 7); i < n % 7; i++) if (!OFF[(s + i) % 7]) c++;
    return HOLS.length ? c - (lowerB(b + 1) - lowerB(a)) : c;
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

  // headers compare loosely: case, spaces, underscores, dots, dashes, colons, a trailing "?" and bracketed units like "(days)" don't matter
  const norm = s => String(s ?? "").replace(/^\uFEFF/, "").toLowerCase().replace(/^\s*custom ?field\s*\(([^)]*)\)/, "$1").replace(/\([^)]*\)|\[[^\]]*\]/g, " ").replace(/[_./:\-–—?*]+/g, " ").replace(/%/g, " % ").replace(/\s+/g, " ").trim();
  // a loose match (the header contains a synonym's words) must not take a column that belongs to a sibling field
  const NOT = { finish: /\b(baseline|bl|original|actual|early|late|latest|earliest|float|slack|variance|duration|start|begin)\b/, start: /\b(baseline|bl|original|actual|early|late|latest|earliest|float|slack|variance|finish|end|due)\b/,
    bfin: /\bstart\b/, bstart: /\b(finish|end)\b/, afin: /\bstart\b/, name: /\b(id|key|number|no|date|owner|status|type|code)\b/, pct: /\b(work|cost)\b/ };
  const LOOSE = ["name", "finish", "start", "bfin", "pct", "status", "owner", "phase", "dur", "uid", "afin", "bstart"];
  // header row = the row (1-10) with most recognised header names; columns by synonym order: exact, then loose, then by content
  function detect(grid) {
    const syn = new Set(FIELDS.flatMap(f => f[3].map(norm)));
    let best = 0, hdr = 0;
    for (let i = 0; i < Math.min(HDR_SCAN, grid.length); i++) {
      const s = grid[i].filter(c => syn.has(norm(c))).length;
      if (s > best) { best = s; hdr = i; }
    }
    const heads = grid[hdr] || [], H = heads.map(norm), cols = {}, guessed = [];
    FIELDS.forEach(([k]) => { cols[k] = -1; });
    const free = j => j >= 0 && !Object.values(cols).includes(j);
    const pass = test => { for (const [k, , , names] of FIELDS) if (cols[k] < 0) for (const n of names) { const j = heads.findIndex((h, i) => free(i) && test(h, H[i], n, k)); if (j >= 0) { cols[k] = j; break; } } };
    pass((h, _, n) => low(h) === low(n));
    pass((_, h, n) => h !== "" && h === norm(n));
    pass((_, h, n, k) => { const w = norm(n); if (!LOOSE.includes(k) || w.length < 5 || (NOT[k] && NOT[k].test(h))) return false; return ` ${h} `.includes(` ${w} `); });
    // no name or finish header: read the columns themselves
    const body = grid.slice(hdr + 1, hdr + 201).filter(r => r.some(c => String(c ?? "").trim()));
    const vals = j => body.map(r => String(r[j] ?? "").trim()).filter(Boolean);
    const dateShare = j => { const v = vals(j); return v.length < 3 ? 0 : v.filter(x => { const p = dateParts(x); return dateValue(p, false) != null || dateValue(p, true) != null; }).length / v.length; };
    const dateMean = j => { const v = vals(j).map(x => { const p = dateParts(x); return dateValue(p, false) ?? dateValue(p, true); }).filter(x => x != null); return v.reduce((a, b) => a + b, 0) / (v.length || 1); };
    if (cols.finish < 0) {
      const ds = heads.map((_, j) => j).filter(j => free(j) && dateShare(j) >= .6).sort((a, b) => dateMean(b) - dateMean(a));
      if (ds.length) { cols.finish = ds[0]; guessed.push("finish"); if (cols.start < 0 && ds.length > 1) { cols.start = ds[ds.length - 1]; guessed.push("start"); } }
    }
    if (cols.name < 0) {
      const textish = j => { const v = vals(j); if (v.length < 3) return 0; const t = v.filter(x => /\p{L}{2}/u.test(x) && !isNum(x.replace(/[%,]/g, ""))).length / v.length; return t < .6 || dateShare(j) > .3 ? 0 : v.reduce((a, x) => a + x.length, 0) / v.length * (new Set(v).size / v.length); };
      const j = heads.map((_, i) => i).filter(free).sort((a, b) => textish(b) - textish(a))[0];
      if (j != null && textish(j) > 3) { cols.name = j; guessed.push("name"); }
    }
    return { hdr, heads, cols, guessed };
  }

  // the columns no field claimed: what each one holds, so a dashboard can chart it
  const SEM = [["priority", /\b(priority|prio|urgency|importance)\b/], ["severity", /\b(severity|impact)\b/], ["sprint", /\b(sprint|iteration|cycle|release|version|fix version)\b/],
    ["points", /\b(story points?|points|sp|estimate|estimation|effort|size|complexity)\b/], ["hours", /\b(hours|hrs|work|remaining work|time spent|logged|timesheet)\b/],
    ["cost", /\b(cost|budget|spend|amount|price|value|actuals?|forecast cost|eac|capex|opex)\b/], ["risk", /\b(risk|rag|health)\b/],
    ["label", /\b(labels?|tags?|keywords?)\b/], ["component", /\b(components?|module|area|feature area|system|application|app)\b/],
    ["team", /\b(team|department|squad|function|vendor|supplier|org)\b/], ["place", /\b(region|location|site|country|city|office)\b/], ["client", /\b(client|customer|account)\b/]];
  const money = v => v.replace(/[$€£₹¥,\s]|usd|eur|gbp|inr|hrs?|h$|d$|pts?$/gi, "");
  function profile(heads, rows, cols, dayFirst) {
    const taken = new Set(Object.values(cols).filter(j => j >= 0));
    return heads.map((h, col) => ({ h: String(h ?? "").trim(), col })).filter(c => c.h && !taken.has(c.col)).map(({ h, col }) => {
      const v = rows.map(r => String(r[col] ?? "").trim()), f = v.filter(Boolean), n = f.length, nh = norm(h);
      const sem = (SEM.find(([, re]) => re.test(nh)) || [null])[0];
      if (n === 0) return { name: h, col, type: "empty", sem };
      const num = f.filter(x => isNum(money(x).replace(/%$/, ""))).length / n, dt = f.filter(x => dateValue(dateParts(x), dayFirst) != null).length / n;
      const distinct = new Set(f.map(low)).size, yes = f.filter(x => /^(yes|no|y|n|true|false|✓|✔|x)$/i.test(x)).length / n;
      const idLike = /\b(id|key|no|number|ref|code|s no|sl no|row)\b/.test(nh) || (num >= .8 && !sem && distinct === n && n > 5);
      const type = num >= .8 && idLike ? "id" : num >= .8 && distinct > 2 ? "number" : dt >= .8 ? "date" : yes >= .9 ? "flag" : distinct <= Math.max(12, n * .3) && distinct < n ? "category" : "text";
      return { name: h, col, type, sem, filled: n / rows.length, distinct };
    });
  }
  const numOf = v => { const x = parseFloat(money(String(v ?? "")).replace(/%$/, "")); return Number.isFinite(x) ? x : null; };

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
    // milestone keywords match whole words, any case; the default ones only on rows of at most 1 working day or with no start
    // (a keyword in a 6-day "Prepare approval pack" is not a milestone), keywords the user adds on any row
    const keys = [...new Set((opt.msKeys || []).map(low).filter(Boolean))].map(k => ({ def: DEFAULT_MS_KEYS.includes(k),
      re: new RegExp(`(?:^|[^\\p{L}\\p{N}])${k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(?:$|[^\\p{L}\\p{N}])`, "iu") }));
    const body = grid.slice(det.hdr + 1, det.hdr + 1 + CAP);
    const capped = grid.length - det.hdr - 1 > CAP;
    const get = (r, k) => cols[k] >= 0 ? (r[cols[k]] ?? "").trim() : "";

    const R = body.map(r => {
      const o = { raw: r }; FIELDS.forEach(([k]) => o["r_" + k] = get(r, k)); return o;
    });
    let evDay = 0, evMon = 0;
    for (const o of R) for (const k of ["start", "finish", "bfin", "afin", "bstart"]) {
      o["p_" + k] = dateParts(o["r_" + k]);
      const e = o["p_" + k]?.ev || 0; if (e > 0) evDay++; if (e < 0) evMon++;
    }
    const dayFirst = opt.dateOrder === "day" ? true : opt.dateOrder === "month" ? false : evDay > evMon;
    let maxPa = 0;
    for (const o of R) {
      // no status, % or actual-finish column: a Kanban list called Done (Trello, ClickUp lists) still means done
      const st = o.r_status || (cols.status < 0 && cols.pct < 0 && cols.afin < 0 && DONE.includes(low(o.r_phase)) ? o.r_phase : "");
      o.st = !st ? "" : dateValue(dateParts(st), dayFirst) != null ? "Done" : DONE.includes(low(st)) ? "Done" : CANCEL.includes(low(st)) ? "Cancelled" : "Open";
      o.nm = o.st === "Cancelled" ? "" : o.r_name.trim();
      o.act = o.nm !== "";
      for (const k of ["start", "finish", "bfin", "afin", "bstart"]) o[k] = dateValue(o["p_" + k], dayFirst);
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
      o.ms = o.act && !o.sum && (truthy(o.r_ms) || msc.includes("milestone") || o.dn === 0
        || keys.some(k => k.re.test(o.nm) && (!k.def || o.start == null || o.finish == null || nwd(o.start, o.finish) <= 1)));
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
    const k = { rows: act.length, cols, heads: det.heads, guessed: det.guessed.filter(g => !(opt.cols && g in opt.cols)), hdrRow: det.hdr + 1, dayFirst, evidence: evDay + evMon, statusDate: sd };
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
    k.extra = profile(det.heads, leaf.map(o => o.raw), cols, dayFirst);
    // every leaf row, for report builders (dates are day numbers; see fromYMD/iso)
    const seen = {};
    k.tasks = leaf.map(o => ({ key: (b => (seen[b] = (seen[b] || 0) + 1) > 1 ? `${b}#${seen[b]}` : b)(o.nm + "|" + o.ph), row: o.row, name: o.nm, owner: o.r_owner, phase: o.ph, start: o.start, finish: o.finish, bfin: o.bfin,
      pct: o.pct, done: o.done, ms: o.isms, od: o.od, late: o.late, var: o.var, w: o.w,
      uid: o.r_uid, afin: o.afin, bstart: o.bstart, st: o.r_status, type: o.r_ms, x: k.extra.map(c => String(o.raw[c.col] ?? "").trim()) }));
    // rows the status column marks cancelled (not leaves, so not in k.tasks): the plan diff tells cancelled from removed
    k.cancelled = R.filter(o => o.st === "Cancelled" && o.r_name).map(o => ({ uid: o.r_uid, name: o.r_name, phase: (cols.phase >= 0 && o.r_phase) || o.ph, owner: o.r_owner, row: o.row }));
    k.rules = rules;
    k.checks = {
      noFinish: leaf.filter(o => o.finish == null).length,
      badDates: act.filter(o => ["start", "finish", "bfin"].some(f => o["r_" + f] !== "" && o[f] == null)).length,
      order: leaf.filter(o => o.start != null && o.finish != null && o.finish < o.start).length,
      mismatch: cols.status >= 0 && cols.pct >= 0 ? leaf.filter(o => o.pct != null && ((o.st === "Done" && o.pct < 1) || (o.st === "Open" && o.pct >= 1))).length : 0,
      capped };
    return k;
  }

  setCalendar();
  const parseDate = (t, dayFirst) => dateValue(dateParts(t), dayFirst);
  root.KeelsonEngine = { compute, parseTable, detect, norm, numOf, parseDate, FIELDS, RULES, DEFAULT_MS_KEYS, fromYMD, iso, fmt, nwd, wdiff, wd, DAY,
    setCalendar, calendar, calText, isWd, addWd, WEEKENDS };
  if (typeof module !== "undefined") module.exports = root.KeelsonEngine;
})(typeof window !== "undefined" ? window : globalThis);
