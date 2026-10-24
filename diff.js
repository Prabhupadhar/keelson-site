/* Keelson plan diff: what changed between two Versions (site/history.js; format docs/PRO-ARCHITECTURE.md section 5).
   match -> diff -> report (a "Changes since <date>" ReportDoc) / block (Blocks for other reports). Pure: no DOM, no clock,
   no network. Plan text is untrusted: it goes into ReportDoc as plain text, and the renderers escape it. Tests: tests/web/diff.test.js. */
(function (root) {
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const FUZZY = 0.8, CAP = 30, BLOCK_ROWS = 8;
  const ORDER = ["milestoneMoved", "baselineMoved", "finishLater", "finishEarlier", "added", "removed", "cancelled", "startMoved",
    "ownerChanged", "reopened", "pctBackwards", "renamed", "phaseMoved"];

  // ---------- text helpers (same rules as reports.js) ----------
  const norm = s => String(s ?? "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim(); // lowercase, punctuation out, spaces collapsed
  const clean = s => String(s ?? "").replace(/[\u0000-\u001F\u007F]/g, " ");
  const cut = (s, n) => { s = clean(s).trim(); return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s; };
  const plural = (n, w, ws) => `${n} ${n === 1 || n === -1 ? w : (ws || w + "s")}`;
  const and = a => a.length < 2 ? a.join("") : a.slice(0, -1).join(", ") + " and " + a[a.length - 1];
  const wdCell = n => n == null ? "—" : n === 0 ? "0 wd" : `${n > 0 ? "+" : ""}${n} wd`;
  const fromIso = s => { const p = String(s || "").split("-").map(Number); return p.length === 3 && p.every(Number.isFinite) ? E.fromYMD(...p) : null; };
  const year = v => +E.iso(v).slice(0, 4);
  const dater = asOf => { const y = asOf ? +asOf.slice(0, 4) : null; return v => v == null ? "—" : E.fmt(v, y == null || year(v) !== y); };
  const shortD = asOf => { const y = asOf ? +asOf.slice(0, 4) : null; return v => v == null ? "—" : E.fmt(v) + (y == null || year(v) !== y ? " " + String(year(v)).slice(2) : ""); }; // "18 Nov 26" in tables
  const fmtIso = s => { const v = fromIso(s); return v == null ? String(s || "") : E.fmt(v, true); };
  const pctT = v => v == null ? "—" : Math.round(v * 100) + "%";
  const L = (label, width) => ({ label, align: "left", ...(width ? { width } : {}) });
  const R = (label, width) => ({ label, align: "right", ...(width ? { width } : {}) });

  // ---------- matching ----------
  const words = s => new Set(norm(s).split(" ").filter(Boolean));
  const jaccard = (a, b) => { let i = 0; for (const w of a) if (b.has(w)) i++; const u = a.size + b.size - i; return u ? i / u : 0; };
  const hasIds = v => v.tasks.length > 0 && v.tasks.filter(t => t.uid).length >= 0.8 * v.tasks.length;

  // pairs prev/cur tasks: phase + normalised name (repeats by order); then Task ID, kept only when neither row has an
  // exact-name partner left in the other version (an MS Project "ID" renumbers on insert) and the names are the same or
  // close (word overlap >= 0.8, a rename to confirm); then unique name across phases; then one-to-one fuzzy names within
  // a phase = probable renames. Rows cancelled in cur are matched last.
  function match(prev, cur) {
    const A = prev.tasks || [], B = cur.tasks || [], ua = new Array(A.length).fill(false), ub = new Array(B.length).fill(false);
    const pairs = [];
    const pair = (i, j, how) => { ua[i] = ub[j] = true; pairs.push({ a: A[i], b: B[j], how, ia: i, ib: j }); };
    const group = (idx, used, T, keyOf) => { const m = new Map(); idx.forEach(i => { if (used[i]) return; const k = keyOf(T[i]); if (!m.has(k)) m.set(k, []); m.get(k).push(i); }); return m; };
    const all = n => Array.from({ length: n }, (_, i) => i);
    const pk = t => t.phase + "\u0001" + norm(t.name);
    const gb = group(all(B.length), ub, B, pk);
    for (const [k, is] of group(all(A.length), ua, A, pk)) { const js = gb.get(k) || []; for (let x = 0; x < Math.min(is.length, js.length); x++) pair(is[x], js[x], "name"); }
    // IDs are trusted when tasks matched by name kept their ID (an MS Project "ID" is a row number, renumbered on insert)
    const same = pairs.filter(p => p.a.uid && p.b.uid), kept = same.filter(p => p.a.uid === p.b.uid).length;
    const by = hasIds(prev) && hasIds(cur) && kept >= 0.8 * same.length ? "id" : "name";
    if (by === "id") {
      const left = (T, used) => { const s = new Set(); T.forEach((t, i) => { if (!used[i]) s.add(norm(t.name)); }); return s; };
      const na = left(A, ua), nb = left(B, ub), at = new Map();
      B.forEach((t, j) => { if (!ub[j] && t.uid && !at.has(t.uid)) at.set(t.uid, j); });
      A.forEach((t, i) => {
        const j = !ua[i] && t.uid ? at.get(t.uid) : undefined; if (j === undefined) return;
        const x = norm(t.name), y = norm(B[j].name);
        if (x !== y && (nb.has(x) || na.has(y) || jaccard(words(x), words(y)) < FUZZY)) return;
        at.delete(t.uid); pair(i, j, x === y ? "id" : "fuzzy");
      });
    }
    const nb = group(all(B.length), ub, B, t => norm(t.name));
    for (const [k, is] of group(all(A.length), ua, A, t => norm(t.name))) { const js = nb.get(k) || []; if (k && is.length === 1 && js.length === 1) pair(is[0], js[0], "name"); }
    // fuzzy: per phase, best scores first, each task used once
    const fa = group(all(A.length), ua, A, t => t.phase), fb = group(all(B.length), ub, B, t => t.phase);
    for (const [ph, is] of fa) {
      const js = fb.get(ph); if (!js) continue;
      const wb = js.map(j => words(B[j].name)), c = [];
      for (const i of is) { const wa = words(A[i].name); js.forEach((j, x) => { const s = jaccard(wa, wb[x]); if (s >= FUZZY) c.push([s, i, j]); }); }
      c.sort((p, q) => q[0] - p[0] || p[1] - q[1] || p[2] - q[2]);
      for (const [, i, j] of c) if (!ua[i] && !ub[j]) pair(i, j, "fuzzy");
    }
    // prev tasks that are cancelled rows in cur
    const C = cur.cancelled || [], uc = new Array(C.length).fill(false), cancelled = [];
    const tryC = (i, f) => { const x = C.findIndex((c, n) => !uc[n] && f(c)); if (x >= 0) { uc[x] = ua[i] = true; cancelled.push({ a: A[i], c: C[x], ia: i }); } };
    A.forEach((t, i) => { if (!ua[i] && by === "id" && t.uid) tryC(i, c => c.uid === t.uid); });
    A.forEach((t, i) => { if (!ua[i]) tryC(i, c => c.phase === t.phase && norm(c.name) === norm(t.name)); });
    A.forEach((t, i) => { if (!ua[i]) tryC(i, c => norm(c.name) === norm(t.name)); });
    return { by, pairs, cancelled, renames: pairs.filter(p => p.how === "fuzzy"),
      added: B.map((t, j) => ({ t, j })).filter(x => !ub[x.j]), removed: A.map((t, i) => ({ t, i })).filter(x => !ua[x.i]) };
  }

  // ---------- diff ----------
  // opts: {renames: {accepted: [cur key], split: [cur key]}, minDays: 1, exclude: [change id], starts: false}
  // exclude leaves changes out of a client version, except baseline changes (they always show). starts: true lists
  // start-only moves with the changes instead of in the appendix.
  function diff(prev, cur, opts = {}) {
    const minDays = opts.minDays ?? 1, rn = opts.renames || {};
    const split = new Set(rn.split || []), accepted = new Set(rn.accepted || []);
    const excl = new Set((opts.exclude || []).filter(id => !String(id).startsWith("baselineMoved:")));
    const m = match(prev, cur);
    const changes = [], appendix = [];
    const ch = (type, t, row, extra) => ({ id: type + ":" + t.key, type, key: t.key, name: t.name, phase: t.phase, owner: t.owner,
      milestone: !!t.ms, from: null, to: null, wd: null, row, ...extra });
    const push = (c, small) => (small ? appendix : changes).push(c);
    const added = m.added.map(x => [x.t, x.j]), removed = m.removed.map(x => [x.t, x.i]);
    const pairs = m.pairs.filter(p => !(p.how === "fuzzy" && split.has(p.b.key)));
    for (const p of m.pairs) if (p.how === "fuzzy" && split.has(p.b.key)) { added.push([p.b, p.ib]); removed.push([p.a, p.ia]); }
    for (const [t, j] of added) push(ch("added", t, j, { to: t.finish }));
    for (const [t, i] of removed) push(ch("removed", t, 1e6 + i, { from: t.finish }));
    for (const { a, ia } of m.cancelled) push(ch("cancelled", a, 1e6 + ia, { from: a.finish }));
    // a phase renamed: every matched task of A now sits in B, A is gone and B is new -> one change, not one per task
    const phA = new Set((prev.tasks || []).map(t => t.phase)), phB = new Set((cur.tasks || []).map(t => t.phase)), into = new Map();
    for (const p of pairs) { if (!into.has(p.a.phase)) into.set(p.a.phase, new Set()); into.get(p.a.phase).add(p.b.phase); }
    const phRen = new Map();
    for (const [x, s] of into) { const y = [...s][0]; if (x && y && s.size === 1 && x !== y && !phB.has(x) && !phA.has(y)) phRen.set(x, y); }
    for (const [x, y] of phRen) if ([...phRen.values()].filter(v => v === y).length > 1) phRen.delete(x); // two phases merged: not a rename
    for (const [x, y] of phRen) {
      const ps = pairs.filter(p => p.a.phase === x);
      changes.push({ id: "phaseMoved:phase:" + x, type: "phaseMoved", key: "phase:" + x, name: `Phase '${x}' renamed '${y}'`, phase: y, owner: "",
        milestone: false, from: x, to: y, wd: null, row: Math.min(...ps.map(p => p.ib)), phaseRenamed: true, tasks: ps.length });
    }
    const mv = (x, y) => x == null || y == null || x === y ? null : E.wdiff(x, y);
    for (const { a, b, how, ib } of pairs) {
      const fw = mv(a.finish, b.finish), sw = mv(a.start, b.start);
      const st = sw == null ? {} : { startFrom: a.start, startTo: b.start, startWd: sw };
      const bw = a.bfin != null && a.bfin !== b.bfin ? (b.bfin == null ? null : E.wdiff(a.bfin, b.bfin)) : undefined;
      const rebased = bw != null && Math.abs(bw) >= minDays && fw != null; // finish and baseline moved: one baseline change
      if (bw !== undefined) {
        const x = { from: a.bfin, to: b.bfin, wd: bw };
        if (rebased) {
          const vw = E.wdiff(b.bfin, b.finish) - E.wdiff(a.bfin, a.finish);
          Object.assign(x, { finishFrom: a.finish, finishTo: b.finish, finishWd: fw, varianceWd: vw, hidden: vw === 0 && fw !== 0, ...st,
            detail: `re-baselined: finish ${wdCell(fw)}, baseline ${wdCell(bw)}, variance ${vw === 0 ? "unchanged" : wdCell(vw)}` });
        }
        push(ch("baselineMoved", b, ib, x), bw != null && Math.abs(bw) < minDays);
      }
      if (fw != null && !rebased) push(ch(a.ms || b.ms ? "milestoneMoved" : fw > 0 || (fw === 0 && b.finish > a.finish) ? "finishLater" : "finishEarlier",
        b, ib, { from: a.finish, to: b.finish, wd: fw, ...st }), Math.abs(fw) < minDays);
      else if (fw == null && sw != null) push(ch("startMoved", b, ib, { from: a.start, to: b.start, wd: sw }), !opts.starts || Math.abs(sw) < minDays);
      if (norm(a.owner) !== norm(b.owner)) push(ch("ownerChanged", b, ib, { from: a.owner, to: b.owner }));
      if (a.done && !b.done) push(ch("reopened", b, ib, { from: 1, to: b.pct }));
      else if (a.pct != null && b.pct != null && Math.round(b.pct * 100) < Math.round(a.pct * 100)) push(ch("pctBackwards", b, ib, { from: a.pct, to: b.pct }));
      if (norm(a.name) !== norm(b.name)) push(ch("renamed", b, ib, { from: a.name, to: b.name, probable: how === "fuzzy" && !accepted.has(b.key) }));
      if (a.phase !== b.phase && phRen.get(a.phase) !== b.phase) push(ch("phaseMoved", b, ib, { from: a.phase, to: b.phase }));
    }
    // milestones: go-live (else the latest milestone) first, then later moves before earlier, then by size
    const msl = changes.filter(c => c.type === "milestoneMoved");
    const lead = msl.find(c => /\bgo[\s-]?live\b/i.test(c.name)) || msl.reduce((x, c) => !x || c.to > x.to ? c : x, null);
    const rank = c => c.type !== "milestoneMoved" ? [0, 0] : [c === lead ? 0 : 1, c.wd > 0 ? 0 : 1];
    const sort = list => list.filter(c => !excl.has(c.id)).sort((x, y) => ORDER.indexOf(x.type) - ORDER.indexOf(y.type) ||
      rank(x)[0] - rank(y)[0] || rank(x)[1] - rank(y)[1] || Math.abs(y.wd ?? 0) - Math.abs(x.wd ?? 0) || x.row - y.row);
    const cs = sort(changes), ap = sort(appendix);
    const counts = Object.fromEntries(ORDER.map(t => [t, cs.filter(c => c.type === t).length]));
    counts.appendix = ap.length; counts.total = cs.length;
    // phase end dates (last non-milestone finish) for "driven by"
    const ends = v => { const e = new Map(); for (const t of v.tasks || []) if (!t.ms && t.finish != null) e.set(t.phase, Math.max(e.get(t.phase) ?? -Infinity, t.finish)); return e; };
    const eb = ends(cur), phases = [];
    for (const [ph, x] of ends(prev)) { const p = phRen.get(ph) ?? ph, y = eb.get(p); if (p && y != null && y !== x) phases.push({ phase: p, from: x, to: y, wd: E.wdiff(x, y) }); }
    const facts = { finishFrom: prev.finish ?? null, finishTo: cur.finish ?? null, phases };
    return { by: m.by, counts, changes: cs, appendix: ap, hidden: changes.length + appendix.length - cs.length - ap.length, facts,
      renames: m.renames.map(p => ({ key: p.b.key, from: p.a.name, to: p.b.name, phase: p.b.phase })),
      headline: headline(cs, prev.asOf, cur.asOf, ap.length, facts), prevAsOf: prev.asOf, asOf: cur.asOf };
  }

  // the sentence a sponsor reads first, e.g. "Since 16 Feb: finish 28 May → 1 Jun (+2 wd); Go-live 14 → 18 May (+2 wd),
  // driven by Testing. 1 task added (Extra regression cycle), 1 removed (Archive legacy reports). 3 tasks finish later: the
  // largest move was +2 wd (User acceptance testing); 2 tasks followed." Later/earlier counts are tasks, milestones included.
  function headline(cs, prevAsOf, asOf, nSmall = 0, f = {}) {
    const D = dater(asOf), of = t => cs.filter(c => c.type === t);
    const rng = (a, b) => { const x = D(a), y = D(b), i = x.indexOf(" "), j = y.indexOf(" "); return i > 0 && x.slice(i) === y.slice(j) ? `${x.slice(0, i)} → ${y}` : `${x} → ${y}`; };
    const move = (a, b, w) => `${rng(a, b)} (${wdCell(w)})`;
    const small = nSmall ? ` ${plural(nSmall, "smaller date change")} (start-only, or under a working day) in the appendix.` : "";
    if (!cs.length) return `No changes since ${D(fromIso(prevAsOf))}.${small}`;
    const out = [];
    // 1. the dates a sponsor cares about: project finish, go-live (or the latest milestone), the phase that moved most
    const lead = [], fw = f.finishFrom != null && f.finishTo != null ? E.wdiff(f.finishFrom, f.finishTo) : null;
    const ms = of("milestoneMoved"), same = ms.length && ms[0].from === f.finishFrom && ms[0].to === f.finishTo;
    if (fw != null) lead.push(f.finishFrom === f.finishTo ? `finish unchanged (${D(f.finishTo)})` : `finish${same ? ` and ${cut(ms[0].name, 50)}` : ""} ${move(f.finishFrom, f.finishTo, fw)}`);
    if (ms.length && !same) lead.push(`${cut(ms[0].name, 50)} ${move(ms[0].from, ms[0].to, ms[0].wd)}`);
    const sign = Math.sign(ms.length ? ms[0].wd : fw || 0);
    const ph = (f.phases || []).filter(p => p.wd && (!sign || Math.sign(p.wd) === sign)).sort((x, y) => Math.abs(y.wd) - Math.abs(x.wd) || x.from - y.from)[0];
    if (ph && sign && lead.length) lead[lead.length - 1] += `, driven by ${cut(ph.phase, 40)}`;
    else if (ph) lead.push(`${cut(ph.phase, 40)} phase end ${move(ph.from, ph.to, ph.wd)}`);
    if (ms.length > 1) lead.push(`${plural(ms.length - 1, "other milestone")} moved`);
    // 2. scope
    const add = of("added"), rem = of("removed"), can = of("cancelled"), c1 = [];
    const top = list => { const m = new Map(); for (const c of list) if (c.phase) m.set(c.phase, (m.get(c.phase) || 0) + 1); return [...m].sort((x, y) => y[1] - x[1])[0] || [null, 0]; };
    const named = (list, verb) => `${c1.length ? list.length : plural(list.length, "task")} ${verb}` + (list.length === 1 ? ` (${cut(list[0].name, 50)})` : "");
    if (add.length) {
      const [p, n] = top(add);
      c1.push(add.length === 1 ? named(add, "added") : plural(add.length, "task") + " added" + (!p ? "" : n === add.length ? ` (all in ${cut(p, 40)})` : n * 2 > add.length ? ` (${n} in ${cut(p, 40)})` : ""));
    }
    if (rem.length) c1.push(named(rem, "removed"));
    if (can.length) c1.push(named(can, "cancelled"));
    // 3. task dates: tasks, milestones included
    const later = [...of("finishLater"), ...ms.filter(c => c.wd > 0)].sort((x, y) => y.wd - x.wd), earlier = [...of("finishEarlier"), ...ms.filter(c => c.wd < 0)];
    const nMs = l => { const k = l.filter(c => c.milestone).length; return k ? ` (incl. ${plural(k, "milestone")})` : ""; };
    const c3 = [];
    if (later.length === 1) c3.push(`1 task finishes later${nMs(later)}: ${wdCell(later[0].wd)} (${cut(later[0].name, 50)})`);
    else if (later.length) c3.push(`${later.length} tasks finish later${nMs(later)}: the largest move was ${wdCell(later[0].wd)} (${cut(later[0].name, 50)}); ${plural(later.length - 1, "task")} followed`);
    if (earlier.length) c3.push(`${plural(earlier.length, "task")} ${earlier.length === 1 ? "finishes" : "finish"} earlier${nMs(earlier)}, by up to ${wdCell(Math.min(...earlier.map(c => c.wd)))}`);
    const bl = of("baselineMoved"), hid = bl.filter(c => c.hidden).length;
    const also = [], rnm = of("renamed"), probable = rnm.filter(c => c.probable).length;
    const n = (t, fn) => { const k = of(t).length; if (k) also.push(fn(k)); };
    n("startMoved", k => `${plural(k, "start date")} moved`);
    n("ownerChanged", k => `${plural(k, "owner change")}`);
    n("reopened", k => `${plural(k, "task")} reopened`);
    n("pctBackwards", k => `% complete went down on ${plural(k, "task")}`);
    if (rnm.length) also.push(`${plural(rnm.length, "task")} renamed${probable ? ` (${probable} to confirm)` : ""}`);
    for (const c of of("phaseMoved").filter(c => c.phaseRenamed)) also.push(`phase '${cut(c.from, 40)}' renamed '${cut(c.to, 40)}'`);
    const pm = of("phaseMoved").filter(c => !c.phaseRenamed).length; if (pm) also.push(`${plural(pm, "task")} moved to another phase`);
    const cap = s => s[0].toUpperCase() + s.slice(1);
    if (lead.length) out.push(`Since ${D(fromIso(prevAsOf))}: ${lead.join("; ")}.`);
    if (c1.length) out.push(cap(c1.join(", ")) + ".");
    if (c3.length) out.push(cap(c3.join("; ")) + ".");
    if (bl.length) out.push(`Baseline changed on ${plural(bl.length, "task")}${hid ? `, of which ${hid} hidden by re-baseline (the finish moved with its baseline)` : ""}.`);
    if (also.length) out.push((out.length ? "Also " : `Since ${D(fromIso(prevAsOf))}: `) + and(also) + ".");
    return out.join(" ") + small;
  }

  // ---------- ReportDoc ----------
  const ownerT = s => clean(s).trim() || "—";
  const phaseT = s => clean(s).trim() || "—";
  function tables(d, asOf, cap = CAP) {
    const D = shortD(asOf), out = [], of = t => d.changes.filter(c => c.type === t);
    const arrow = (a, b) => `${a} → ${b}`;
    const more = (n, w) => n > 0 ? { type: "note", text: `+${n} more ${w} not shown.` } : null;
    const section = (title, list, columns, row, w = "changes") => {
      if (!list.length) return;
      out.push({ type: "heading", text: `${title} (${list.length})` }, { type: "table", columns, rows: list.slice(0, cap).map(row) });
      const m = more(list.length - cap, w); if (m) out.push(m);
    };
    const dateCols = first => [L(first, "32%"), L("Phase", "17%"), L("Owner", "15%"), L("From → to", "25%"), R("Moved", "11%")];
    const late = c => ({ text: wdCell(c.wd), tone: c.wd > 0 ? "red" : c.wd < 0 ? "green" : null });
    const nm = c => { const t = cut(c.name, 90) + (c.detail ? ` (${c.detail})` : ""); return c.milestone ? { text: t, bold: true } : t; };
    const dateRow = (c, tone) => [nm(c), phaseT(c.phase), ownerT(c.owner),
      tone ? { text: arrow(D(c.from), D(c.to)), tone } : arrow(D(c.from), D(c.to)), tone ? { text: wdCell(c.wd), tone } : late(c)];
    section("Milestone dates moved", of("milestoneMoved"), dateCols("Milestone"), c => dateRow(c));
    section("Baseline changed", of("baselineMoved"), dateCols("Task"), c => dateRow(c, "red"));
    section("Finish moved later", of("finishLater"), dateCols("Task"), c => dateRow(c));
    section("Finish moved earlier", of("finishEarlier"), dateCols("Task"), c => dateRow(c));
    section("Added", of("added"), [L("Task", "44%"), L("Phase", "22%"), L("Owner", "18%"), R("Finish", "16%")],
      c => [c.milestone ? { text: cut(c.name, 90), bold: true } : cut(c.name, 90), phaseT(c.phase), ownerT(c.owner), D(c.to)]);
    section("Removed or cancelled", [...of("removed"), ...of("cancelled")], [L("Task", "40%"), L("Phase", "20%"), L("Owner", "16%"), R("Was due", "12%"), R("", "12%")],
      c => [cut(c.name, 90), phaseT(c.phase), ownerT(c.owner), D(c.from), c.type === "cancelled" ? { text: "Cancelled", tone: "grey", chip: true } : { text: "Removed", chip: true }]);
    section("Start moved", of("startMoved"), dateCols("Task"), c => [...dateRow(c).slice(0, 4), wdCell(c.wd)]);
    section("Owner changed", of("ownerChanged"), [L("Task", "44%"), L("Phase", "22%"), L("From → to", "34%")],
      c => [cut(c.name, 90), phaseT(c.phase), arrow(ownerT(c.from), ownerT(c.to))]);
    section("Progress went backwards", [...of("reopened"), ...of("pctBackwards")], [L("Task", "40%"), L("Phase", "20%"), L("Owner", "16%"), R("From → to", "24%")],
      c => [cut(c.name, 90), phaseT(c.phase), ownerT(c.owner), { text: c.type === "reopened" ? `Done → ${pctT(c.to)} (reopened)` : arrow(pctT(c.from), pctT(c.to)), tone: "amber" }]);
    section("Renamed", of("renamed"), [L("Was", "36%"), L("Now", "36%"), L("Phase", "16%"), L("", "12%")],
      c => [cut(c.from, 90), cut(c.to, 90), phaseT(c.phase), c.probable ? { text: "To confirm", tone: "amber", chip: true } : ""]);
    section("Moved to another phase", of("phaseMoved"), [L("Task", "40%"), L("From → to", "44%"), L("Owner", "16%")],
      c => [cut(c.name, 90), arrow(phaseT(c.from), phaseT(c.to)), ownerT(c.owner)]);
    return out;
  }
  const KIND = { finishLater: "Finish", finishEarlier: "Finish", startMoved: "Start", baselineMoved: "Baseline", milestoneMoved: "Milestone" };

  // opts: diff opts + {title, client, preparedBy, edits: {[textBlockId]: string}}
  function report(prev, cur, opts = {}) {
    const d = diff(prev, cur, opts), c = d.counts, D = shortD(cur.asOf), since = fmtIso(prev.asOf);
    const of = t => d.changes.filter(x => x.type === t);
    const later = [...of("finishLater"), ...of("milestoneMoved").filter(x => x.wd > 0)], nMs = later.filter(x => x.milestone).length; // same as the headline
    const gone = c.removed + c.cancelled;
    const edits = opts.edits || {}, txt = (id, t) => typeof edits[id] === "string" ? edits[id] : t;
    const blocks = [
      { type: "tiles", items: [
        { label: "Added", value: String(c.added), sub: c.added ? `${plural(c.added, "new task")}` : "No new tasks", tone: null },
        { label: "Removed", value: String(gone), sub: c.cancelled ? `${c.cancelled} cancelled` : gone ? `${plural(gone, "task")} no longer in the plan` : "Nothing removed", tone: gone ? "amber" : null },
        { label: "Moved later", value: String(later.length), sub: later.length ? `Tasks incl. ${plural(nMs, "milestone")}; largest ${wdCell(Math.max(...later.map(x => x.wd)))}` : "No finish dates slipped", tone: later.length ? "amber" : "green" },
        { label: "Baseline changed", value: String(c.baselineMoved), sub: c.baselineMoved ? `${plural(c.baselineMoved, "task")} re-baselined` + ((h => h ? `, ${h} hiding a slip` : "")(of("baselineMoved").filter(x => x.hidden).length)) : "Baseline unchanged", tone: c.baselineMoved ? "red" : "green" },
      ] },
      { type: "text", id: "changes-summary", label: "What changed", text: txt("changes-summary", d.headline), editable: true },
      ...(d.hidden ? [{ type: "note", text: `${plural(d.hidden, "change")} hidden from this version.` }] : []),
      ...tables(d, cur.asOf, opts.cap ?? CAP),
    ];
    if (d.appendix.length) {
      const cap = opts.cap ?? CAP;
      blocks.push({ type: "heading", text: `Appendix: start-only moves and dates moved by less than ${plural(opts.minDays ?? 1, "working day")} (${d.appendix.length})` },
        { type: "table", columns: [L("Task", "40%"), L("Phase", "20%"), L("Date", "14%"), R("From → to", "26%")],
          rows: d.appendix.slice(0, cap).map(x => [cut(x.name, 90), phaseT(x.phase), KIND[x.type] || "Date", `${D(x.from)} → ${D(x.to)}`]) });
      if (d.appendix.length > cap) blocks.push({ type: "note", text: `+${d.appendix.length - cap} more not shown.` });
    }
    const probable = d.changes.filter(x => x.probable).length;
    blocks.push({ type: "note", text: `Compares the plan as of ${since} with the plan as of ${fmtIso(cur.asOf)}. ` +
      (d.by === "id" ? "Tasks are matched on phase and task name, then on task ID." : "Tasks are matched on phase and task name.") +
      (probable ? ` ${plural(probable, "rename")} matched on similar names ${probable === 1 ? "is" : "are"} marked "To confirm".` : "") +
      " Working days are Monday to Friday." });
    return { id: "changes", name: `Changes since ${since}`, title: clean(opts.title || "Plan changes"),
      subtitle: clean(["What changed in the plan", opts.client && `for ${opts.client}`].filter(Boolean).join(" ") + (opts.preparedBy ? ` · Prepared by ${opts.preparedBy}` : "")),
      asOf: cur.asOf, rag: cur.rag || "—", orientation: "portrait", blocks };
  }

  // a short section for another report: heading, the sentence, the biggest changes (milestones and baseline first)
  function block(prev, cur, opts = {}) {
    const d = diff(prev, cur, opts), D = shortD(cur.asOf), edits = opts.edits || {};
    const WHAT = { milestoneMoved: "Milestone moved", baselineMoved: "Baseline changed", finishLater: "Finish later", finishEarlier: "Finish earlier",
      added: "Added", removed: "Removed", cancelled: "Cancelled", startMoved: "Start moved", ownerChanged: "Owner changed", reopened: "Reopened",
      pctBackwards: "% complete down", renamed: "Renamed", phaseMoved: "Moved phase" };
    const span = x => ["ownerChanged", "renamed", "phaseMoved"].includes(x.type) ? `${cut(x.from, 30) || "—"} → ${cut(x.to, 30) || "—"}`
      : ["reopened", "pctBackwards"].includes(x.type) ? `${x.type === "reopened" ? "Done" : pctT(x.from)} → ${pctT(x.to)}` : `${D(x.from)} → ${D(x.to)}`;
    const rows = d.changes.slice(0, BLOCK_ROWS).map(x => [WHAT[x.type], cut(x.name, 70), span(x),
      x.wd == null ? "" : { text: wdCell(x.wd), tone: x.type === "baselineMoved" ? "red" : x.wd > 0 ? "red" : x.wd < 0 ? "green" : null }]);
    const out = [{ type: "heading", text: `Changes since ${fmtIso(prev.asOf)}` },
      { type: "text", id: "changes-block", text: typeof edits["changes-block"] === "string" ? edits["changes-block"] : d.headline, editable: true }];
    if (rows.length) out.push({ type: "table", columns: [L("Change", "20%"), L("Task", "42%"), L("From → to", "26%"), R("Moved", "12%")], rows });
    if (d.changes.length > BLOCK_ROWS) out.push({ type: "note", text: `+${d.changes.length - BLOCK_ROWS} more changes in the Changes report.` });
    return out;
  }

  const api = { match, diff, report, block, headline, norm, jaccard: (a, b) => jaccard(words(a), words(b)), ORDER };
  root.KeelsonDiff = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
