/* Keelson plan health check: counts of rows that fail simple, checkable schedule checks, with the rows named.
   Input is the engine's result (k.tasks, k.cols, k.statusDate) and, optionally, the previous saved version. No score:
   a count per check can be verified against the plan, a single number can't. Free in the Quick check and the app. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const LONG = 44; // working days; the DCMA 14-point "high duration" threshold
  const SHOW = 8; // rows listed per check; the rest are counted

  function run(k, prev) {
    if (!k || !k.tasks || k.cols.name < 0) return [];
    const sd = k.statusDate, T = k.tasks, has = f => k.cols[f] >= 0;
    const work = T.filter(t => !t.ms), dur = t => E.wdiff(t.start, t.finish) + 1;
    const out = [];
    const add = (id, title, why, rows, detail) => out.push({ id, title, why, n: rows.length, rows: rows.slice(0, SHOW).map(t => ({ name: t.name, row: k.hdrRow + 1 + t.row, owner: E.noOwner(t.owner) ? "" : t.owner, detail: detail ? detail(t) : "" })) });
    if (has("owner")) add("owner", "Tasks with no owner", "Nobody is asked for an update, so nobody gives one.", work.filter(t => E.noOwner(t.owner)));
    add("finish", "Tasks with no finish date", "They drop off timelines, due lists and the forecast.", T.filter(t => t.finish == null));
    if (has("bfin")) add("baseline", "Tasks with no baseline finish", "Slippage can't be measured for them.", T.filter(t => t.bfin == null));
    add("order", "Finish before start", "One of the two dates is wrong.", T.filter(t => t.start != null && t.finish != null && t.finish < t.start), t => `${E.fmt(t.start)} → ${E.fmt(t.finish)}`);
    add("long", `Tasks longer than ${LONG} working days`, "Long tasks hide slippage; split them into pieces you can track. Based on the DCMA 14-point check.",
      work.filter(t => t.start != null && t.finish != null && dur(t) > LONG).sort((a, b) => dur(b) - dur(a)), t => `${dur(t)} wd`);
    if (has("pct") || has("status")) {
      add("late-start", "Overdue and not started", "Past their finish date with no progress recorded.", work.filter(t => !t.done && t.finish != null && t.finish < sd && !(t.pct > 0)).sort((a, b) => a.finish - b.finish), t => `due ${E.fmt(t.finish)}`);
      add("early-pct", "Progress before the start date", "Work is reported on tasks that haven't started yet. Check the start date or the % done.", work.filter(t => t.pct > 0 && t.start != null && t.start > sd), t => `${Math.round(t.pct * 100)}%, starts ${E.fmt(t.start)}`);
    }
    add("ms-dur", "Milestones with a duration", "A milestone is a point in time. Give it zero duration or make it a task.", T.filter(t => t.ms && t.start != null && t.finish != null && t.finish > t.start), t => `${dur(t)} wd`);
    if (prev && prev.tasks && (has("pct") || has("status"))) {
      const was = {}; for (const t of prev.tasks) was[t.key] = t;
      add("stale", `No progress since ${E.fmt(E.parseDate(prev.asOf))}`, "Under way and due, but the % done hasn't moved since the last saved version.",
        work.filter(t => { const p = was[t.key]; return p && !t.done && t.pct != null && p.pct != null && t.pct === p.pct && t.start != null && t.start <= sd; }), t => `${Math.round(t.pct * 100)}%`);
    }
    return out;
  }

  // one shared view for the Quick check and the app's Check step
  function html(res, esc) {
    if (!res.length) return "";
    const bad = res.filter(c => c.n), ok = res.filter(c => !c.n);
    return `<ul class="hc">${bad.map(c => `<li class="hc-i"><details><summary><b>${c.n}</b> ${esc(c.title)}</summary><p>${esc(c.why)}</p>
      <table class="hc-tb"><thead><tr><th>Task</th><th class="hc-row">Row</th><th>Owner</th><th></th></tr></thead><tbody>${c.rows.map(r => `<tr><td>${esc(r.name)}</td><td class="hc-row">${r.row}</td><td>${esc(r.owner)}</td><td>${esc(r.detail)}</td></tr>`).join("")}</tbody></table>
      ${c.n > c.rows.length ? `<p class="hc-more">+${c.n - c.rows.length} more</p>` : ""}</details></li>`).join("")}
      ${ok.length ? `<li class="hc-ok">✓ None found: ${ok.map(c => esc(c.title.charAt(0).toLowerCase() + c.title.slice(1))).join(" · ")}</li>` : ""}</ul>`;
  }
  const summary = res => { const n = res.filter(c => c.n).length; return !res.length ? "" : n ? `${n} to look at` : "all clear"; };

  const api = { run, html, summary, LONG };
  root.KeelsonHealth = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
