/* Keelson finish-date confidence (P50 / P80) from the plan's own history: docs/V2-FEATURES.md section C.
   forecast(k, opts) -> result; report(k, opts) -> ReportDoc "Finish-date confidence"; block(k, opts) -> Block[] for other reports.
   Pure and deterministic: no DOM, no clock; the PRNG is seeded from a hash of the tasks, so the same file gives the same dates.
   Tested by tests/web/forecast.test.js (acceptance tests + the back-test in tests/web/backtest.js). */
(function (root) {
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const C = { navy: "#1f3a5f", teal: "#1a7f86", amber: "#f79009", red: "#d92d20", grey: "#98a2b3", ink: "#1d2939", muted: "#667085", line: "#e4e7ec", lane: "#f7f9fb" };
  const FONT = "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif";
  const MIN_N = 10, MIN_DONE = 0.2, RESET = 0.9, PHASE_N = 5, DISAGREE = 10;
  // tests/web/backtest.js at 1,000 projects (forecast.test.js checks these still match, rounded)
  const BACKTEST = { projects: 1000, p80: 81, p50: 53 };
  // Calibration (docs/V2-FEATURES.md C, "Calibration"; tests/web/backtest.js): the remaining work is replayed as a network rebuilt
  // from the plan's dates, each task overrunning like a completed (or still-running overdue) task; the pace of each replay is a
  // bootstrap of the history times exp(z * PROJ_SD * cv), a project-level term for "the remaining work as a whole can run differently".
  const PROJ_SD = 0.2;

  const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const plural = (n, w, ws) => `${n} ${n === 1 ? w : (ws || w + "s")}`;
  const yearOf = v => +E.iso(v).slice(0, 4);
  const mkD = sd => v => v == null ? "—" : E.fmt(v, sd == null || yearOf(v) !== yearOf(sd)); // year only when it differs
  const F = v => v == null ? "—" : E.fmt(v, true);
  const x2 = f => f.toFixed(2) + "x";
  const addWd = (d, n) => E.addWd(d, Math.max(0, Math.round(n))); // on the engine's work calendar

  // same file -> same dates: the seed is a hash of the reference class (completed tasks the bootstrap resamples) in a canonical
  // order, plus the status date; column and row order do not matter, and open work does not change the random stream, so adding
  // an overdue task can only move P80 later (see the censoring note below)
  const row = t => JSON.stringify([t.uid ?? "", t.key ?? t.name, t.phase ?? "", t.start ?? null, t.finish ?? null, t.bstart ?? null, t.bfin ?? null, t.afin ?? null]);
  function hash(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
  function mulberry32(a) {
    return () => { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
  }
  // counter-based uniform draw for (seed, task, iteration): independent of every other task
  const mix = x => { x ^= x >>> 16; x = Math.imul(x, 0x7feb352d); x ^= x >>> 15; x = Math.imul(x, 0x846ca68b); x ^= x >>> 16; return x >>> 0; };
  const u01 = (seed, id, i) => mix(mix(seed ^ id) + Math.imul(i + 1, 0x9E3779B9)) / 4294967296;
  const gauss = rnd => Math.sqrt(-2 * Math.log(1 - rnd())) * Math.cos(2 * Math.PI * rnd());
  const pctl = (sorted, p) => sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(p * sorted.length) - 1))]; // nearest rank

  function refuse(reason, extra) { return { ok: false, reason, n: 0, censored: 0, factor: null, p50: null, p80: null, planFinish: null, remaining: null, sentence: "", assumptions: [], phases: [], notes: [], ...extra }; }

  function forecast(k, opts = {}) {
    const sd = k.statusDate, D = mkD(sd), iters = opts.iterations || 5000;
    const all = (k.tasks || []).filter(t => !t.ms);
    const base = { planFinish: k.finish ?? null };
    if (!all.some(t => t.bfin != null)) return refuse("The plan has no baseline dates, so there is no track record to forecast from. Add Baseline Start and Baseline Finish columns to the export.", base);
    if (!all.some(t => t.bstart != null)) return refuse("The plan has baseline finish dates but no baseline start, so the baselined duration of each task cannot be known. Add a Baseline Start column to the export.", base);
    if (!all.some(t => !t.done)) return refuse("All work is complete, so there is no remaining work to forecast.", base);
    if (k.finish == null) return refuse("The plan has no forecast finish date.", base);
    if (sd > k.finish) return refuse(`The status date ${F(sd)} is after the forecast finish ${F(k.finish)} with work still open. Re-forecast the plan, then try again.`, base);

    const bd = t => Math.max(1, E.nwd(t.bstart, t.bfin));
    const based = all.filter(t => t.bstart != null && t.bfin != null && t.bfin >= t.bstart);
    const ref = [], cens = [];
    for (const t of based) {
      const af = t.afin ?? (t.done ? t.finish : null);
      if (t.done && af != null && t.start != null && af >= t.start) ref.push({ t, a: Math.max(1, E.nwd(t.start, af)), b: bd(t), af });
      else if (!t.done && t.start != null && t.start <= sd && sd > t.bfin) cens.push({ t, b: bd(t), lb: E.nwd(t.start, sd) / bd(t) });
    }
    const n = ref.length, R = E.wdiff(sd, k.finish);
    const out = { ...base, n, censored: cens.length, remaining: R };
    if (n < MIN_N) return refuse(`Only ${plural(n, "completed task")} ${n === 1 ? "has" : "have"} a baseline start, baseline finish and actual finish; at least ${MIN_N} are needed before the plan's own history says anything.`, out);
    const tot = based.reduce((s, t) => s + bd(t), 0), doneW = based.reduce((s, t) => s + bd(t) * (t.done ? 1 : Math.min(1, t.pct || 0)), 0);
    if (doneW < MIN_DONE * tot) return refuse(`Only ${Math.round(doneW / tot * 100)}% of the baselined working days are complete; at least ${MIN_DONE * 100}% are needed before the track record is worth projecting.`, out);
    const onBase = ref.filter(o => o.af === o.t.bfin).length;
    if (onBase >= RESET * n) return refuse(`${onBase} of ${n} completed tasks finished exactly on their baseline date: the baseline looks reset to actuals, so it cannot show how long work really took. Use the export from before the baseline was reset.`, out);

    ref.forEach(o => o.s = row(o.t)); ref.sort((x, y) => x.s < y.s ? -1 : x.s > y.s ? 1 : 0);
    const seed = opts.seed ?? hash(sd + "|" + ref.map(o => o.s).join("\n"));
    const ratios = ref.map(o => o.a / o.b);
    const factor = ref.reduce((s, o) => s + o.a, 0) / ref.reduce((s, o) => s + o.b, 0);
    const mean = ratios.reduce((s, r) => s + r, 0) / n, cv = Math.sqrt(ratios.reduce((s, r) => s + (r - mean) ** 2, 0) / Math.max(1, n - 1)) / mean;
    // the remaining work as a network rebuilt from the plan's own dates: a task that starts the working day after another
    // finishes is taken to wait for it (finish-to-start); a task with no such predecessor keeps its planned start.
    // Offsets are working days after the status date; in-progress tasks carry on from tomorrow for their remaining days.
    const net = all.map((t, ri) => ({ t, ri })).filter(({ t }) => !t.done && t.finish != null).map(({ t, ri }) => {
      const inProg = t.start != null && t.start <= sd, fin = Math.max(1, E.wdiff(sd, t.finish));
      const dur = inProg ? fin : t.start != null ? Math.max(1, E.nwd(t.start, t.finish)) : 1, st = inProg ? 1 : Math.max(1, fin - dur + 1);
      // an in-progress task with a baseline: its total length is drawn like any task (baseline x overrun), less the days already run
      const el = inProg && t.bstart != null && t.bfin != null && t.bfin >= t.bstart && t.start != null ? E.nwd(t.start, sd) : null;
      return { st, dur, fin: st + dur - 1, id: hash(row(t)), inProg, ph: t.phase || "", ri, pred: [], el, b: el != null ? bd(t) : 0 };
    }).sort((x, y) => x.st - y.st || x.fin - y.fin || x.id - y.id);
    // a task's shape is drawn in proportion to baseline duration (the same weighting as the factor), so a 2-day task's rounding
    // noise counts for less than a 20-day task's record
    // overdue open tasks are in the pool too (the big overruns are the ones still running): one drawn takes a completed ratio at or
    // above its ratio so far (or that ratio, when none is). Draws read the pool's weighted quantile function, so a pool that shifts
    // later moves every draw later or leaves it alone.
    const srt = [...ratios].sort((a, b) => a - b), atoms = ref.map((o, j) => [ratios[j], o.b]);
    if (opts.censShape !== false) for (const c of cens) { const k0 = srt.findIndex(r => r >= c.lb);
      if (k0 < 0) atoms.push([c.lb, c.b]); else for (let j = k0; j < n; j++) atoms.push([srt[j], c.b / (n - k0)]); }
    atoms.sort((x, y) => x[0] - y[0]);
    const cum = []; atoms.reduce((c, a, j) => (cum[j] = c + a[1]), 0);
    const pick = u => { const x = u * cum[cum.length - 1]; let lo = 0, hi = cum.length - 1; while (lo < hi) { const m = (lo + hi) >> 1; if (cum[m] <= x) lo = m + 1; else hi = m; } return atoms[lo][0]; };
    const byFin = new Map();
    net.forEach((o, j) => { const a = byFin.get(o.fin); if (a) a.push(j); else byFin.set(o.fin, [j]); });
    // a hand-off between phases: a task that starts right after a task of another phase finishes waits for that whole phase
    // (its tasks replayed so far, i.e. those planned to start earlier: a running maximum per phase keeps this linear)
    const phs = [...new Set(net.map(o => o.ph))];
    net.forEach(o => { o.pi = phs.indexOf(o.ph); o.pred = -1; o.hand = []; if (o.inProg) return;
      const cand = byFin.get(o.st - 1) || [], same = cand.filter(q => net[q].ph === o.ph);
      // within a phase, parallel chains often touch by chance: wait only for the nearest such task listed above it (plan row order)
      const above = same.filter(q => net[q].ri < o.ri);
      if (same.length) o.pred = (above.length ? above : same).reduce((a, q) => Math.abs(net[q].ri - o.ri) < Math.abs(net[a].ri - o.ri) ? q : a);
      o.hand = [...new Set(cand.filter(q => net[q].ph !== o.ph).map(q => net[q].pi))]; });
    const phMax = new Float64Array(phs.length);
    const rnd = mulberry32(seed), Ts = new Float64Array(iters), ps = (opts.projSd ?? PROJ_SD) * cv, fj = new Float64Array(net.length);
    for (let i = 0; i < iters; i++) {
      // 1. the history's overall pace: a bootstrap resample of the completed tasks' duration-weighted factor
      let A = 0, B = 0;
      for (let j = 0; j < n; j++) { const o = ref[Math.floor(rnd() * n)]; A += o.a; B += o.b; }
      // censored (overdue open) tasks join every resample: each takes at least lb x its baseline and no less than the resample's
      // own factor F, i.e. the fixed point F = (A + sum max(lb, F) b) / (B + sum b); g(F) rises with slope < 1, so iteration converges
      let f = A / B;
      if (cens.length) for (let it = 0; it < 50; it++) {
        let a2 = A, b2 = B; for (const c of cens) { a2 += Math.max(c.lb, f) * c.b; b2 += c.b; }
        const nf = a2 / b2; if (Math.abs(nf - f) < 1e-12) break; f = nf;
      }
      // 2. a project-level term: the remaining work as a whole may run faster or slower than the past
      f *= Math.exp(gauss(rnd) * ps);
      // 3. every remaining task gets its own overrun, shaped like a completed task's (ratio / mean ratio), and the network is replayed;
      //    parallel paths now merge on the slowest one. Draws hash (seed, task, iteration), so adding a task never reshuffles the others.
      let T = 0; phMax.fill(-Infinity);
      for (let j = 0; j < net.length; j++) {
        const o = net[j], u = u01(seed, o.id, i), r = f * pick(u) / factor;
        let st = o.st;
        if (o.pred >= 0 || o.hand.length) { st = o.pred >= 0 ? fj[o.pred] : -Infinity; for (const q of o.hand) if (phMax[q] > st) st = phMax[q]; st += 1; }
        fj[j] = st + (o.el != null ? Math.max(o.dur * r, o.b * r - o.el) : o.dur * r) - 1;
        if (fj[j] > T) T = fj[j];
        if (fj[j] > phMax[o.pi]) phMax[o.pi] = fj[j];
      }
      Ts[i] = T;
    }
    Ts.sort();
    const t50 = pctl(Ts, .5), t80 = pctl(Ts, .8), f50 = R ? t50 / R : 1, f80 = R ? t80 / R : 1;
    const p50 = addWd(sd, t50), p80 = Math.max(p50, addWd(sd, t80));

    // per-phase factors (plan order) where a phase has enough completed samples
    const order = [];
    for (const t of all) if (!order.includes(t.phase || "")) order.push(t.phase || "");
    const phases = order.map(name => { const a = ref.filter(o => (o.t.phase || "") === name);
      return { name: name || "(no phase)", n: a.length, factor: a.length >= PHASE_N ? a.reduce((s, o) => s + o.a, 0) / a.reduce((s, o) => s + o.b, 0) : null }; })
      .filter(p => p.n > 0);
    const notes = [], q = phases.filter(p => p.factor != null);
    if (q.length >= 2 && q[q.length - 1].factor > q.slice(0, -1).reduce((s, p) => s + p.factor, 0) / (q.length - 1) + 0.1)
      notes.push(`P80 likely optimistic: the latest phase with enough history (${q[q.length - 1].name}, ${x2(q[q.length - 1].factor)}) is running slower than the earlier ones, so remaining work may be riskier than the average.`);
    if (cv < 0.02) notes.push("Very uniform history: completed tasks all took almost exactly the same multiple of their baseline, so P50 and P80 sit together. Real projects rarely behave this evenly; check the dates were not copied from the baseline.");
    // the PM's own re-forecasts as a second method: shown only when they point well past P80 (or never settle)
    let callout = null;
    if (opts.versions && opts.versions.length >= 4) {
      const T = root.KeelsonTrend || (typeof require !== "undefined" ? require("./trend.js") : null), dr = T && T.trend(opts.versions).drift;
      if (dr && dr.verdict && (dr.timePace || (dr.landing != null && E.wdiff(p80, dr.landing) > DISAGREE)))
        callout = `Your own re-forecasts have slipped ${dr.slopeWdPerWeek.toFixed(1)} wd a week; if that continues the finish lands after P80. Treat P80 as optimistic.`;
    }

    const pf = k.finish, cmp = pf < p50 ? `The current plan date of ${D(pf)} is earlier than either.`
      : pf < p80 ? `The current plan date of ${D(pf)} is at or after P50 but earlier than P80.` : `The current plan date of ${D(pf)} is at or after both.`;
    const sentence = `${n} completed tasks took on average ${x2(factor)} their baselined time. If the remaining work behaves the same way, and parallel work waits for its slowest path, ` +
      `the 50/50 date is ${D(p50)} (P50); 8 in 10 replays finish by ${D(p80)} (P80). ${cmp} This ignores new scope and risks not yet in the plan.`;
    const assumptions = [
      "Remaining work grows like completed work did: the same spread of overruns and underruns.",
      "Task order is read from the plan's dates; lags and start-to-start links are missed.",
      "New scope and named risks are not modelled; add them to the plan to see their effect.",
      "The spread is understated if the remaining work is riskier than the work done so far.",
    ];
    const dist = Array.from({ length: 160 }, (_, i) => addWd(sd, Ts[Math.floor((i + .5) * Ts.length / 160)])); // 160 evenly spaced finish dates, for drawing
    return { ok: true, ...out, seed, factor, f50, f80, p50, p80, dist, cv, sentence, assumptions, phases, notes, callout };
  }

  // ---------- report ----------
  const wrap = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="${FONT}" fill="${C.ink}"><rect width="${w}" height="${h}" fill="#ffffff"/>${body}</svg>`;
  // date strip: status date, plan date, P50, P80 on one line; the P50-P80 range shaded; labels alternate above/below so they never collide
  function strip(k, f) {
    const sd = k.statusDate, D = mkD(sd), W = 700, H = 122, X0 = 34, X1 = W - 34, y = 66;
    const pts = [[sd, "Status date", C.muted, "line"], [f.planFinish, "Plan date", C.navy, "dot"], [f.p50, "P50 · 50/50 date", C.teal, "dot"], [f.p80, "P80 · 8 in 10", C.amber, "dot"]];
    const t0 = sd, t1 = Math.max(...pts.map(p => p[0])), span = Math.max(1, t1 - t0), x = v => +(X0 + (v - t0) / span * (X1 - X0)).toFixed(1);
    const o = [`<line x1="${X0}" y1="${y}" x2="${X1}" y2="${y}" stroke="${C.line}" stroke-width="6" stroke-linecap="round"/>`,
      `<rect x="${x(f.p50)}" y="${y - 5}" width="${Math.max(2, x(f.p80) - x(f.p50))}" height="10" rx="5" fill="#d5f0f1"/>`];
    // order by date; a label goes above unless the previous above-label would overlap it
    const lab = pts.map((p, i) => ({ v: p[0], t: p[1], c: p[2], kind: p[3], i })).sort((a, b) => a.v - b.v || a.i - b.i);
    let lastUp = -1e9, lastDn = -1e9;
    for (const L of lab) {
      const cx = x(L.v), w = Math.max(L.t.length * 7.2, 64), anchor = cx - w / 2 < 4 ? "start" : cx + w / 2 > W - 4 ? "end" : "middle";
      const left = anchor === "start" ? cx : anchor === "end" ? cx - w : cx - w / 2, up = left > lastUp + 8 || left <= lastDn + 8 ? left > lastUp + 8 : false;
      if (up) lastUp = left + w; else lastDn = left + w;
      const ty = up ? y - 36 : y + 34;
      if (L.kind === "line") o.push(`<line x1="${cx}" y1="${y - 14}" x2="${cx}" y2="${y + 14}" stroke="${C.muted}" stroke-width="1.5" stroke-dasharray="4 3"/>`);
      else o.push(`<circle cx="${cx}" cy="${y}" r="8" fill="${L.c}" stroke="#ffffff" stroke-width="2"/>`);
      o.push(`<line x1="${cx}" y1="${up ? y - 12 : y + 12}" x2="${cx}" y2="${up ? y - 30 : y + 28}" stroke="${C.line}"/>`,
        `<text x="${cx}" y="${ty}" font-size="13" font-weight="600" text-anchor="${anchor}">${esc(L.t)}</text>`,
        `<text x="${cx}" y="${ty + (up ? -16 : 16)}" font-size="13" fill="${C.muted}" text-anchor="${anchor}">${esc(D(L.v))}</text>`);
    }
    return { type: "svg", svg: wrap(W, H, o.join("")), width: W, height: H,
      alt: `Date strip: status date ${F(sd)}, current plan date ${F(f.planFinish)}, P50 ${F(f.p50)}, P80 ${F(f.p80)}.` };
  }

  const HOW = f => `Keelson measured how long each of the ${plural(f.n, "completed task")} with a baseline start, baseline finish and actual finish really took against its baselined working days, ` +
    `then replayed the remaining ${plural(f.remaining, "working day")} of work 5,000 times. The order of that work comes from the plan's dates: a task that starts the working day after another finishes waits for it; ` +
    `one that starts straight after another phase ends waits for that whole phase. In each replay every remaining task overruns like a completed task picked at random` +
    `${f.censored ? ` (overdue open tasks included, at no less than they have run so far)` : ""}, at a pace resampled from the history with an allowance for the rest of the project running faster or slower. ` +
    `Parallel work finishes when its slowest path does. P50 is the date half the replays finish by; P80 the date 8 in 10 finish by. Working days are ${E.calText()}. ` +
    `Checked on ${BACKTEST.projects.toLocaleString("en-GB")} simulated projects (P80 held in ${BACKTEST.p80}%, P50 in ${BACKTEST.p50}%); not yet validated on real project histories. ` +
    `Dependencies are inferred from dates; a Predecessors column, if present, is not used yet.`;

  // a visible text block, not a grey note: the re-forecast trend says P80 is optimistic
  const callout = f => f.callout ? { type: "text", id: "forecast-drift", label: "Check against your re-forecasts", text: f.callout, editable: false } : null;

  function block(k, opts = {}) {
    const f = opts.result || forecast(k, opts);
    if (!f.ok) return [{ type: "heading", text: "Finish-date confidence" }, { type: "note", text: `No P50 / P80 forecast: ${f.reason}` }];
    return [{ type: "heading", text: "Finish-date confidence" }, strip(k, f), { type: "text", id: "forecast-sentence", text: f.sentence, editable: false }, callout(f),
      ...f.notes.filter(t => /^(P80 likely|Very uniform)/.test(t)).map(text => ({ type: "note", text }))].filter(Boolean);
  }

  function report(k, opts = {}) {
    const f = opts.result || forecast(k, opts), sd = k.statusDate, D = mkD(sd);
    const head = { id: "forecast", name: "Finish-date confidence", title: opts.title || "Project status",
      subtitle: [opts.client, "Finish-date confidence", opts.preparedBy ? `Prepared by ${opts.preparedBy}` : ""].filter(Boolean).join(" · "),
      asOf: E.iso(sd), rag: k.rag || "—", orientation: "portrait" };
    if (!f.ok) return { ...head, blocks: [
      { type: "note", text: `Keelson will not estimate P50 / P80 for this plan: ${f.reason}` },
      { type: "heading", text: "What the forecast needs" },
      { type: "list", items: [`At least ${MIN_N} completed tasks with a baseline start, baseline finish and actual (or finished) date.`,
        `At least ${MIN_DONE * 100}% of the baselined working days complete.`, "A baseline that was set before the work was done, not reset to actuals.",
        "A forecast finish on or after the status date."] }] };
    return { ...head, blocks: [
      { type: "tiles", items: [
        { label: "Plan date", value: D(f.planFinish), sub: `Forecast finish in the plan · ${plural(f.remaining, "working day")} from the status date`, tone: null },
        { label: "P50 · 50/50 date", value: D(f.p50), sub: `Half the replays finish by then · ${E.wdiff(f.planFinish, f.p50) >= 0 ? "+" : ""}${E.wdiff(f.planFinish, f.p50)} wd on plan`, tone: E.wdiff(f.planFinish, f.p50) > 0 ? "amber" : "teal" },
        { label: "P80 · 8 in 10", value: D(f.p80), sub: `8 in 10 replays finish by then · ${E.wdiff(f.planFinish, f.p80) >= 0 ? "+" : ""}${E.wdiff(f.planFinish, f.p80)} wd on plan`, tone: E.wdiff(f.planFinish, f.p80) > 10 ? "red" : E.wdiff(f.planFinish, f.p80) > 0 ? "amber" : "teal" },
        { label: "Track record", value: x2(f.factor), sub: `Actual time over baselined time, ${plural(f.n, "completed task")}`, tone: null }] },
      strip(k, f),
      { type: "text", id: "forecast-sentence", label: "In one sentence", text: f.sentence, editable: false },
      callout(f),
      ...f.notes.map(text => ({ type: "note", text })),
      { type: "heading", text: "How this is calculated" },
      { type: "text", id: "forecast-how", text: HOW(f), editable: false },
      { type: "columns", columns: [[
        { type: "heading", text: "Assumptions" },
        { type: "list", id: "forecast-assumptions", items: f.assumptions }], [
        { type: "heading", text: "Track record by phase" },
        { type: "table", columns: [{ label: "Phase", align: "left", width: "50%" }, { label: "Tasks", align: "right", width: "18%" }, { label: "Actual / baseline", align: "right", width: "32%" }],
          rows: f.phases.map(p => [p.name, String(p.n), p.factor == null ? { text: `Under ${PHASE_N} tasks`, tone: "grey" } : { text: x2(p.factor), tone: p.factor > 1.1 ? "amber" : null, bold: true }]),
          empty: "No completed tasks by phase." }]] },
      { type: "note", text: `A range, not a promise: P50 and P80 describe how this project has behaved so far, not what it will do. +n wd = working days (${E.calText()}) later than the plan date.` },
    ].filter(Boolean) };
  }

  root.KeelsonForecast = { forecast, report, block, mulberry32, addWd, PROJ_SD, BACKTEST };
  if (typeof module !== "undefined") module.exports = root.KeelsonForecast;
})(typeof window !== "undefined" ? window : globalThis);
