/* Keelson slip trend and re-baseline audit trail: docs/V2-FEATURES.md section B.
   trend(versions) -> analysis; report(versions, {notes, edits, title, client, preparedBy}) -> ReportDoc "Slip trend" (landscape).
   versions = stored history (docs/PRO-ARCHITECTURE.md section 5): {asOf, rag, pct, finish, bfinish, var, tasks:[{key, uid, name, phase, owner,
   start, finish, bfin, bstart, afin, pct, done, ms, st}]}; dates as "YYYY-MM-DD" or day numbers. Order received does not matter: sorted by asOf.
   Pure and deterministic. Tested by tests/web/trend.test.js. */
(function (root) {
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const C = { navy: "#1f3a5f", teal: "#1a7f86", amber: "#f79009", red: "#d92d20", grey: "#98a2b3", ink: "#1d2939", muted: "#667085", line: "#e4e7ec", lane: "#f7f9fb" };
  // milestone series: brand navy/teal/amber first, then hues checked for colour-blind separation from their neighbours; every line is
  // also direct-labelled and has its own marker shape (and dashes from slot 5), so colour never carries identity alone
  const SERIES = ["#1f3a5f", "#1a7f86", "#f79009", "#7a5af8", "#c11574", "#2a78d6", "#6b8e23", "#8f5a2e"];
  const SHAPES = ["circle", "square", "triangle", "diamond", "circle", "square", "triangle", "diamond"];
  const FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
  const MIN_V = 3, MAX_MS = 8, DRIFT_N = 6, TIME_PACE = 0.9;
  // a landing date is shown only from a fit worth extrapolating: 4+ reports, R² >= 0.7, landing within 2x the remaining span
  const VERDICT_N = 4, LAND_R2 = 0.7, LAND_SPAN = 2;

  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const cut = (s, n) => { s = String(s ?? ""); return s.length > n ? s.slice(0, n - 1).trimEnd() + "…" : s; };
  const plural = (n, w, ws) => `${n} ${n === 1 ? w : (ws || w + "s")}`;
  const day = v => { if (v == null || v === "") return null; if (typeof v === "number") return Number.isFinite(v) ? v : null;
    const p = String(v).slice(0, 10).split("-").map(Number); return p.length === 3 && p.every(Number.isFinite) ? E.fromYMD(...p) : null; };
  const yearOf = v => +E.iso(v).slice(0, 4);
  const F = v => v == null ? "—" : E.fmt(v, true);
  // "3 → 17 May", "26 Apr → 10 May", "20 Dec 2026 → 8 Jan 2027"
  const span = (a, b) => { const [ya, ma] = E.iso(a).split("-"), [yb, mb] = E.iso(b).split("-");
    return `${ya !== yb ? E.fmt(a, true) : ma !== mb ? E.fmt(a) : E.fmt(a).split(" ")[0]} → ${E.fmt(b, ya !== yb)}`; };
  const wdS = n => n === 0 ? "0" : `${n > 0 ? "+" : ""}${n} wd`;
  const hash = s => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); };
  // task identity across versions: the plan's ID (uid) when both sides have one, else the engine key (phase + name)
  const idOf = t => t.uid != null && t.uid !== "" ? "u:" + t.uid : "k:" + (t.key ?? `${t.name}|${t.phase || ""}`);

  function prep(versions) {
    const by = {};
    for (const v of Array.isArray(versions) ? versions : []) { const d = day(v && v.asOf); if (d != null) by[d] = v; } // one per report date; the later one given wins
    return Object.keys(by).map(Number).sort((a, b) => a - b).map(d => {
      const v = by[d], tasks = (v.tasks || []).map(t => ({ ...t, id: idOf(t), finish: day(t.finish), bfin: day(t.bfin), bstart: day(t.bstart) }));
      const fins = tasks.map(t => t.finish).filter(x => x != null), bfs = tasks.map(t => t.bfin).filter(x => x != null);
      return { asOf: d, rag: v.rag, tasks, map: new Map(tasks.map(t => [t.id, t])),
        finish: day(v.finish) ?? (fins.length ? Math.max(...fins) : null), bfinish: day(v.bfinish) ?? (bfs.length ? Math.max(...bfs) : null) };
    });
  }

  function rebaselines(V) {
    const out = [];
    for (let i = 1; i < V.length; i++) {
      const a = V[i - 1], b = V[i], changed = [];
      for (const t of b.tasks) { const o = a.map.get(t.id); if (!o) continue;
        if ((o.bfin != null && t.bfin != null && o.bfin !== t.bfin) || (o.bstart != null && t.bstart != null && o.bstart !== t.bstart)) changed.push({ o, t }); }
      const mk = (key, name, from, to, detail) => ({ id: `rb-${E.iso(b.asOf)}-${hash(key)}`, detected: E.iso(b.asOf), key, name, from: E.iso(from), to: E.iso(to),
        moved: from != null && to != null ? E.wdiff(from, to) : null, detail, note: null });
      // a project re-baseline is logged once, with the tasks it touched as detail; otherwise one row per task
      if (a.bfinish != null && b.bfinish != null && a.bfinish !== b.bfinish) {
        // name the milestones whose baseline moved ("Go-live 3 → 17 May"): that is what a sponsor agreed to
        const msm = changed.filter(c => c.t.ms && c.o.bfin != null && c.t.bfin != null && c.o.bfin !== c.t.bfin).map(c => `${c.t.name} ${span(c.o.bfin, c.t.bfin)}`);
        const ms = msm.length ? `; milestones: ${msm.slice(0, 4).join(", ")}${msm.length > 4 ? ` and ${msm.length - 4} more` : ""}` : "";
        out.push(mk("project", "Whole project (baseline finish)", a.bfinish, b.bfinish, changed.length ? `${plural(changed.length, "task baseline")} changed${changed.length <= 3 ? ": " + changed.map(c => c.t.name).join(", ") : ""}${ms}` : ""));
      }
      else for (const { o, t } of changed) out.push(mk(t.uid != null && t.uid !== "" ? String(t.uid) : t.key ?? t.name, t.name, o.bfin ?? o.bstart, t.bfin ?? t.bstart,
        o.bfin === t.bfin ? "baseline start changed" : t.phase ? t.phase : ""));
    }
    return out;
  }

  function trend(versions) {
    const V = prep(versions), n = V.length;
    const rb = rebaselines(V);
    if (n < MIN_V) return { ok: false, reason: `Needs ${MIN_V} reports; you have ${n}`, n, milestones: [], phases: [], rebaselines: rb, drift: null, versions: V.map(v => E.iso(v.asOf)) };
    const last = V[n - 1];

    // milestones: every one seen, a point per version where it has a date; up to 8 shown (biggest movers, always the last one to finish)
    const seen = new Map();
    for (const v of V) for (const t of v.tasks) if (t.ms && t.finish != null) {
      const m = seen.get(t.id) || { key: t.id, name: t.name, points: [] };
      m.name = t.name; m.points.push({ asOf: E.iso(v.asOf), finish: E.iso(t.finish), d: v.asOf, f: t.finish, done: !!t.done }); seen.set(t.id, m);
    }
    // history, not trend: a milestone already done before the first report, or dated before it and never moved, is left off
    for (const [id, m] of seen) { const p0 = m.points[0];
      if (p0.f < V[0].asOf && (p0.done || m.points.every(p => p.f === p0.f))) seen.delete(id); }
    let ms = [...seen.values()].map(m => ({ ...m, last: m.points[m.points.length - 1].f, moved: E.wdiff(m.points[0].f, m.points[m.points.length - 1].f) }));
    if (ms.length > MAX_MS) {
      const end = ms.reduce((a, b) => b.last > a.last ? b : a);
      const pick = ms.filter(m => m !== end).sort((a, b) => Math.abs(b.moved) - Math.abs(a.moved) || a.last - b.last).slice(0, MAX_MS - 1);
      ms = [...pick, end];
    }
    ms.sort((a, b) => a.last - b.last || (a.name < b.name ? -1 : 1));
    const nMs = seen.size;

    // phases: finish = latest task finish in that phase at each report
    const order = [];
    for (const v of [...V].reverse()) for (const t of v.tasks) if (t.phase && !order.includes(t.phase)) order.push(t.phase);
    const phases = order.map(name => {
      const f = V.map(v => { const a = v.tasks.filter(t => t.phase === name && t.finish != null).map(t => t.finish); return a.length ? Math.max(...a) : null; });
      let slipped = 0, cum = 0, latest = null;
      for (let i = 1; i < n; i++) if (f[i] != null && f[i - 1] != null) { const d = E.wdiff(f[i - 1], f[i]); if (d > 0) { slipped++; cum += d; } latest = d; }
      return { name, versionsSlipped: slipped, of: n - 1, cumulativeWd: cum, latestMove: latest, finish: E.iso(f[n - 1] ?? f.filter(x => x != null).pop()) };
    }).sort((a, b) => b.versionsSlipped - a.versionsSlipped || b.cumulativeWd - a.cumulativeWd || order.indexOf(a.name) - order.indexOf(b.name));

    // drift: least-squares slope of the forecast finish (working days) against report date (working days) over the last 6 reports
    const pts = V.filter(v => v.finish != null).slice(-DRIFT_N);
    let drift = null;
    if (pts.length >= MIN_V) {
      const x0 = pts[0].asOf, f0 = pts[0].finish, xs = pts.map(v => E.wdiff(x0, v.asOf)), ys = pts.map(v => E.wdiff(f0, v.finish));
      const mx = xs.reduce((a, b) => a + b, 0) / xs.length, my = ys.reduce((a, b) => a + b, 0) / ys.length;
      const sxx = xs.reduce((a, x) => a + (x - mx) ** 2, 0), sxy = xs.reduce((a, x, i) => a + (x - mx) * (ys[i] - my), 0), syy = ys.reduce((a, y) => a + (y - my) ** 2, 0);
      const s = sxx ? sxy / sxx : 0, r2 = sxx && syy ? sxy * sxy / (sxx * syy) : 1;
      // where the fitted line meets "finish = report date": the finish it is heading for (only while slower than time passes)
      const c = my - s * mx, gap = E.wdiff(x0, f0); // finish(t) = f0 + c + s*t (wd from x0); report date(t) = x0 + t
      const lp = pts[pts.length - 1], left = E.wdiff(lp.asOf, lp.finish), tl = s > 0.01 && s < TIME_PACE ? (gap + c) / (1 - s) : null;
      const fits = tl != null && tl >= 0 && pts.length >= VERDICT_N && r2 >= LAND_R2, near = fits && tl - xs[xs.length - 1] <= LAND_SPAN * Math.max(1, left);
      const landing = near ? wdAdd(x0, Math.round(tl)) : null, noLanding = near ? null : fits ? "far" : "irregular";
      // the latest run, apart from the fit: how many reports in a row have carried the same finish, else the last move
      const all = V.filter(v => v.finish != null);
      let k = 1; while (k < all.length && all[all.length - 1 - k].finish === lp.finish) k++;
      const latest = k >= 2 ? { held: k, since: E.iso(all[all.length - k].asOf), moved: 0 } : { held: 1, since: null, moved: all.length > 1 ? E.wdiff(all[all.length - 2].finish, lp.finish) : 0 };
      drift = { slope: s, slopeWdPerWeek: s * (7 - E.calendar().weekend.length), n: pts.length, r2, timePace: s >= TIME_PACE, verdict: pts.length >= VERDICT_N, from: E.iso(pts[0].asOf), to: E.iso(lp.asOf), landing, noLanding, latest };
    }
    return { ok: true, n, nMs, milestones: ms.map(({ key, name, points }) => ({ key, name, points: points.map(({ asOf, finish }) => ({ asOf, finish })) })),
      phases, rebaselines: rb, drift, versions: V.map(v => E.iso(v.asOf)), latest: { asOf: E.iso(last.asOf), finish: E.iso(last.finish), rag: last.rag } };
  }
  const wdAdd = (d, n) => { while (!E.isWd(d)) d--; return E.addWd(d, n); };

  function driftText(dr) {
    if (!dr) return null;
    const a = day(dr.from), b = day(dr.to), per = Math.abs(dr.slopeWdPerWeek).toFixed(1);
    const span = `over the last ${dr.n} reports (${E.fmt(a, yearOf(a) !== yearOf(b))} to ${F(b)})`;
    const L = dr.latest, latest = !L ? "" : L.held >= 2 ? ` Latest: held since ${E.fmt(day(L.since), yearOf(day(L.since)) !== yearOf(b))} (${L.held} reports).`
      : L.moved ? ` Latest: ${Math.abs(L.moved)} wd ${L.moved > 0 ? "later" : "earlier"} than the report before.` : "";
    let main;
    if (Math.abs(dr.slopeWdPerWeek) < 0.05) main = `The forecast finish has held steady ${span}.`;
    else if (dr.slopeWdPerWeek < 0) main = `The forecast finish moved ${per} wd earlier per week ${span}.`;
    else if (dr.timePace) main = `The forecast finish is slipping ${per} wd per week ${span}: as fast as time passes, so the trend points to no finish date.`;
    else main = `The forecast finish is slipping ${per} wd per week ${span} (slower than time passes); ` +
      (dr.landing != null ? `if that continues it lands around ${F(dr.landing)}.` : dr.noLanding === "far" ? "if that continues it lands too far out to project a date." : "too irregular to project a finish date.");
    return main + latest + (dr.verdict ? "" : ` A verdict needs ${VERDICT_N} reports.`);
  }

  // ---------- chart ----------
  const wrap = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="${FONT}" fill="${C.ink}"><rect width="${w}" height="${h}" fill="#ffffff"/>${body}</svg>`;
  function marker(shape, x, y, col, r = 4.5) {
    const n = v => +v.toFixed(1), st = `fill="${col}" stroke="#ffffff" stroke-width="1.5"`;
    if (shape === "square") return `<rect x="${n(x - r)}" y="${n(y - r)}" width="${n(2 * r)}" height="${n(2 * r)}" ${st}/>`;
    if (shape === "triangle") return `<path d="M${n(x)} ${n(y - r - 1)}L${n(x + r + 1)} ${n(y + r)}L${n(x - r - 1)} ${n(y + r)}Z" ${st}/>`;
    if (shape === "diamond") return `<path d="M${n(x)} ${n(y - r - 1)}L${n(x + r + 1)} ${n(y)}L${n(x)} ${n(y + r + 1)}L${n(x - r - 1)} ${n(y)}Z" ${st}/>`;
    return `<circle cx="${n(x)}" cy="${n(y)}" r="${r}" ${st}/>`;
  }
  function chart(t) {
    const W = 1000, H = 390, X0 = 78, X1 = W - 230, Y0 = 20, Y1 = H - 92, f = 12;
    const ds = t.versions.map(day), lines = t.milestones.map(m => m.points.map(p => [day(p.asOf), day(p.finish)]));
    const allF = lines.flat().map(p => p[1]);
    const xa = ds[0] - 3, xb = ds[ds.length - 1] + 3, ya0 = Math.min(...allF, ds[0]), yb0 = Math.max(...allF, ds[ds.length - 1]);
    const pad = Math.max(3, Math.round((yb0 - ya0) * .06)), ya = ya0 - pad, yb = yb0 + pad;
    const X = v => +(X0 + (v - xa) / (xb - xa) * (X1 - X0)).toFixed(1), Y = v => +(Y1 - (v - ya) / (yb - ya) * (Y1 - Y0)).toFixed(1);
    const D = v => E.fmt(v, yearOf(v) !== yearOf(ds[ds.length - 1]));
    const o = [];
    // y grid: month starts, or Mondays when the range is short
    const span = yb - ya, ticks = [];
    if (span > 100) { const [yy, mm] = E.iso(ya).split("-").map(Number); for (let i = 1; i < 60; i++) { const v = E.fromYMD(yy, mm + i, 1); if (v > yb) break; ticks.push(v); } }
    else { const step = span > 50 ? 14 : 7; let v = ya + ((7 - E.wd(ya)) % 7); for (; v <= yb; v += step) ticks.push(v); }
    const every = Math.ceil(ticks.length / 8);
    ticks.forEach((v, i) => { if (i % every) return;
      o.push(`<line x1="${X0}" y1="${Y(v)}" x2="${X1}" y2="${Y(v)}" stroke="${C.line}"/>`, `<text x="${X0 - 8}" y="${Y(v) + 4}" font-size="${f}" fill="${C.muted}" text-anchor="end">${esc(D(v))}</text>`); });
    // x: one tick per report (thinned past 10)
    const ev = Math.ceil(ds.length / 10);
    ds.forEach((v, i) => { if (i % ev && i !== ds.length - 1) return;
      o.push(`<line x1="${X(v)}" y1="${Y1}" x2="${X(v)}" y2="${Y1 + 5}" stroke="${C.muted}"/>`, `<text x="${X(v)}" y="${Y1 + 19}" font-size="${f}" fill="${C.muted}" text-anchor="middle">${esc(D(v))}</text>`); });
    o.push(`<line x1="${X0}" y1="${Y1}" x2="${X1}" y2="${Y1}" stroke="${C.muted}"/>`,
      `<text x="${(X0 + X1) / 2}" y="${Y1 + 38}" font-size="${f}" fill="${C.muted}" text-anchor="middle">Report date</text>`,
      `<text x="14" y="${(Y0 + Y1) / 2}" font-size="${f}" fill="${C.muted}" text-anchor="middle" transform="rotate(-90 14 ${(Y0 + Y1) / 2})">Forecast date</text>`);
    // diagonal: forecast date = report date, clipped to the plot
    const a = Math.max(xa, ya), b = Math.min(xb, yb);
    if (b > a) o.push(`<line x1="${X(a)}" y1="${Y(a)}" x2="${X(b)}" y2="${Y(b)}" stroke="${C.grey}" stroke-width="1.5" stroke-dasharray="6 4"/>`,
      `<text x="${X(b) - 4}" y="${Y(b) - 8}" font-size="${f}" fill="${C.muted}" text-anchor="end">Forecast = report date</text>`);
    // lines, markers, then end labels nudged apart
    const labs = [];
    lines.forEach((pts, i) => {
      const col = SERIES[i], dash = i >= 4 ? ` stroke-dasharray="7 4"` : "";
      if (pts.length > 1) o.push(`<polyline points="${pts.map(p => `${X(p[0])},${Y(p[1])}`).join(" ")}" fill="none" stroke="${col}" stroke-width="2"${dash}/>`);
      pts.forEach(p => o.push(marker(SHAPES[i], X(p[0]), Y(p[1]), col)));
      const lp = pts[pts.length - 1]; labs.push({ i, x: X(lp[0]), y: Y(lp[1]), name: cut(t.milestones[i].name, 24), date: D(lp[1]) });
    });
    labs.sort((p, q) => p.y - q.y);
    for (let j = 0; j < labs.length; j++) labs[j].ly = Math.max(labs[j].y, j ? labs[j - 1].ly + 28 : Y0 + 10);
    for (let j = labs.length - 1; j >= 0; j--) labs[j].ly = Math.min(labs[j].ly, j === labs.length - 1 ? Y1 - 6 : labs[j + 1].ly - 28);
    for (const L of labs) {
      const lx = X1 + 14;
      o.push(`<line x1="${L.x + 6}" y1="${L.y}" x2="${lx - 3}" y2="${L.ly}" stroke="${C.line}"/>`, marker(SHAPES[L.i], lx + 4, L.ly - 4, SERIES[L.i], 4),
        `<text x="${lx + 13}" y="${L.ly}" font-size="${f}" font-weight="600">${esc(L.name)}</text>`, `<text x="${lx + 13}" y="${L.ly + 14}" font-size="${f}" fill="${C.muted}">${esc(L.date)}</text>`);
    }
    // legend: series, then how to read
    let lx = X0, ly = H - 22;
    t.milestones.forEach((m, i) => {
      const nm = cut(m.name, 22), w = 26 + nm.length * f * .56 + 14;
      if (lx + w > W - 10) { lx = X0; ly += 18; }
      o.push(`<line x1="${lx}" y1="${ly - 4}" x2="${lx + 18}" y2="${ly - 4}" stroke="${SERIES[i]}" stroke-width="2"${i >= 4 ? ` stroke-dasharray="7 4"` : ""}/>`, marker(SHAPES[i], lx + 9, ly - 4, SERIES[i], 4),
        `<text x="${lx + 24}" y="${ly}" font-size="${f}" fill="${C.ink}">${esc(nm)}</text>`);
      lx += w;
    });
    o.push(`<text x="${X0}" y="${H - 40}" font-size="${f}" fill="${C.muted}">Flat line = holding its date · rising = slipping · a line parallel to the diagonal is moving as fast as time passes.</text>`);
    return { type: "svg", svg: wrap(W, H, o.join("")), width: W, height: H,
      alt: `Milestone trend over ${t.n} reports (${F(ds[0])} to ${F(ds[ds.length - 1])}): ` +
        t.milestones.map(m => `${m.name} ${m.points.map(p => E.fmt(day(p.finish))).join(", ")}`).join("; ") + "." };
  }

  // ---------- report ----------
  function report(versions, opts = {}) {
    const t = trend(versions), V = prep(versions), last = V[V.length - 1];
    const notes = { ...(opts.notes || {}), ...(opts.edits || {}) };
    const head = { id: "trend", name: "Slip trend", title: opts.title || "Project status",
      subtitle: [opts.client, "Slip trend and re-baseline audit trail", opts.preparedBy ? `Prepared by ${opts.preparedBy}` : ""].filter(Boolean).join(" · "),
      asOf: last ? E.iso(last.asOf) : null, rag: (last && last.rag) || "—", orientation: "landscape" };
    const rb = t.rebaselines.map(r => ({ ...r, note: typeof notes[r.id] === "string" && notes[r.id].trim() ? notes[r.id] : null }));
    const unNoted = rb.filter(r => !r.note).length;
    const audit = V.length < 2 ? [] : [
      { type: "heading", text: rb.length ? `Re-baseline log · ${rb.length}` : "Re-baseline log" },
      { type: "table", columns: [{ label: "Detected", align: "center", width: "11%" }, { label: "Task or project", align: "left", width: "27%" }, { label: "Old baseline", align: "center", width: "11%" },
        { label: "New baseline", align: "center", width: "11%" }, { label: "Moved", align: "right", width: "8%" }, { label: "PM note", align: "left", width: "32%" }],
        rows: rb.map(r => [F(day(r.detected)), r.detail ? `${r.name} (${r.detail})` : r.name, F(day(r.from)), F(day(r.to)),
          r.moved == null ? "—" : { text: wdS(r.moved), tone: r.moved > 0 ? "red" : null },
          r.note ? r.note : { text: "No note: explain this re-baseline", tone: "red", chip: true }]),
        empty: `No baseline changes across the ${plural(V.length, "report")}.` },
      unNoted ? { type: "note", text: `${plural(unNoted, "re-baseline")} ${unNoted === 1 ? "has" : "have"} no note. Each stays flagged in this log until the PM records why it happened and who agreed it.` } : null,
      ...rb.map(r => ({ type: "text", id: r.id, label: `Note: ${cut(r.name, 50)}, re-baselined ${F(day(r.detected))}`, editable: true,
        text: r.note ?? "Why was this re-baselined, and who agreed it?", ...(r.note ? {} : { placeholder: true, exportText: "No note recorded." }) })),
    ];
    if (!t.ok) return { ...head, blocks: [{ type: "note", text: `${t.reason}. The milestone trend, slipping phases and drift rate need ${MIN_V} saved reports; ` +
      (V.length >= 2 ? "the re-baseline audit trail below already covers the two you have." : "save this report each week and they build up.") }, ...audit].filter(Boolean) };

    const dr = t.drift, worst = t.phases.find(p => p.versionsSlipped > 0);
    const phRows = t.phases.filter(p => p.versionsSlipped > 0).slice(0, 12);
    return { ...head, blocks: [
      { type: "tiles", items: [
        { label: "Reports compared", value: String(t.n), sub: `${F(day(t.versions[0]))} to ${F(day(t.versions[t.n - 1]))}`, tone: null },
        { label: "Forecast drift", value: dr ? `${dr.slopeWdPerWeek >= 0 ? "+" : ""}${dr.slopeWdPerWeek.toFixed(1)} wd/wk` : "—",
          sub: !dr ? "No forecast finish in the reports" : !dr.verdict ? `Verdict needs ${VERDICT_N} reports` : dr.timePace ? "As fast as time passes" : dr.slopeWdPerWeek > 0.05 ? "Slipping, slower than time passes" : "Holding or improving",
          tone: !dr || !dr.verdict ? null : dr.timePace ? "red" : dr.slopeWdPerWeek > 0.05 ? "amber" : "green" },
        { label: "Re-baselines", value: String(rb.length), sub: rb.length ? `${unNoted} without a note` : "Baseline unchanged", tone: unNoted ? "red" : null },
        { label: "Phase slipping most", value: worst ? cut(worst.name, 22) : "None", sub: worst ? `Later in ${worst.versionsSlipped} of ${worst.of} reports · ${wdS(worst.cumulativeWd)} in total` : "No phase finish moved later", tone: worst ? "amber" : null }] },
      dr ? { type: "text", id: "trend-drift", text: driftText(dr), editable: false } : null,
      { type: "heading", text: "Milestone trend" },
      t.milestones.length ? chart(t) : { type: "note", text: "No milestones found in the saved reports: add a 0-day task, a Milestone column, or a milestone keyword." },
      t.nMs > t.milestones.length ? { type: "note", text: `Showing ${t.milestones.length} of ${t.nMs} milestones: the ones that moved most, and the last to finish.` } : null,
      { type: "heading", text: "Phases that keep slipping" },
      { type: "table", columns: [{ label: "Phase", align: "left", width: "40%" }, { label: "Reports with a later finish", align: "right", width: "18%" }, { label: "Total slipped", align: "right", width: "14%" },
        { label: "Latest move", align: "right", width: "14%" }, { label: "Forecast finish", align: "center", width: "14%" }],
        rows: phRows.map(p => [p.name, { text: `${p.versionsSlipped} of ${p.of}`, bold: p.versionsSlipped * 2 > p.of, tone: p.versionsSlipped * 2 > p.of ? "red" : null },
          wdS(p.cumulativeWd), p.latestMove == null ? "—" : { text: wdS(p.latestMove), tone: p.latestMove > 0 ? "amber" : p.latestMove < 0 ? "green" : null }, F(day(p.finish))]),
        empty: "No phase finish moved later between reports." },
      ...audit,
      { type: "note", text: `Moves are working days (${E.calText()}) between consecutive saved reports, matched on the plan's ID column where it has one, else on phase and task name. Drift is a least-squares slope over up to the last 6 reports.` },
    ].filter(Boolean) };
  }

  root.KeelsonTrend = { trend, report, driftText };
  if (typeof module !== "undefined") module.exports = root.KeelsonTrend;
})(typeof window !== "undefined" ? window : globalThis);
