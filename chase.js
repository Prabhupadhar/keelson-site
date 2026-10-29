/* Keelson chase list: who to ask for an update this week, grouped by owner, ready to paste into an email or chat.
   From the engine's result (k.tasks, k.statusDate) and, optionally, the previous saved version (for stale progress).
   Pure: no DOM, no clock, no network. Keelson never sends anything; the PM copies the text. Pro in the app.
   Tests: tests/web/chase.test.js. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const SOON = 5; // working days ahead for "due soon"
  const NONE = "No owner";
  const ASK = { overdue: "please update % done or confirm a new date", soon: "please confirm it will finish on time", stale: "please update % done" };

  function build(k, prev) {
    if (!k || !k.tasks || k.cols.name < 0 || k.cols.finish < 0) return [];
    const sd = k.statusDate, until = E.addWd(sd, SOON), was = {};
    if (prev && prev.tasks) for (const t of prev.tasks) was[t.key] = t;
    const items = [];
    for (const t of k.tasks) {
      if (t.done || t.finish == null) continue;
      const p = was[t.key];
      if (t.od) items.push({ t, kind: "overdue" }); // the same overdue rule as the reports
      else if (t.ms) continue;
      else if (t.finish <= until) items.push({ t, kind: "soon" });
      else if (p && t.pct != null && p.pct != null && t.pct === p.pct && t.start != null && t.start <= sd) items.push({ t, kind: "stale" });
    }
    const by = {};
    for (const { t, kind } of items) {
      const o = E.noOwner(t.owner) ? NONE : String(t.owner).replace(/\u00a0/g, " ").trim();
      (by[o] = by[o] || []).push({ name: t.name, finish: t.finish, late: kind === "overdue" ? t.late : 0, pct: t.pct, ms: !!t.ms, kind, ask: ASK[kind] });
    }
    const order = { overdue: 0, soon: 1, stale: 2 };
    return Object.entries(by).map(([owner, list]) => ({ owner, items: list.sort((a, b) => order[a.kind] - order[b.kind] || b.late - a.late || a.finish - b.finish) }))
      .sort((a, b) => (a.owner === NONE) - (b.owner === NONE) || b.items.filter(i => i.kind === "overdue").length - a.items.filter(i => i.kind === "overdue").length || a.owner.localeCompare(b.owner));
  }

  const line = i => `- ${i.name}: ${i.kind === "overdue" ? `due ${E.fmt(i.finish)}, ${i.late} working day${i.late === 1 ? "" : "s"} late` : i.kind === "soon" ? `due ${E.fmt(i.finish)}, ${i.pct > 0 ? `${Math.round(i.pct * 100)}% done` : "not started"}` : `due ${E.fmt(i.finish)}, still ${Math.round(i.pct * 100)}% since the last update`}; ${i.ask}`;
  // one message per owner; plain text so it pastes cleanly into email, Teams or Slack
  function text(g, project) {
    const first = g.owner === NONE ? "" : g.owner.split(/[\s,;]+/)[0];
    return [g.owner === NONE ? "These tasks have no owner in the plan. Who is picking them up?" : `Hi ${first},`, "",
      `Could you send me an update on ${g.items.length === 1 ? "this" : "these"}${project ? ` for the ${project} status report` : ""}?`,
      ...g.items.map(line), "", "Thanks"].join("\n");
  }
  const all = (groups, project) => groups.map(g => `To: ${g.owner}\n\n${text(g, project)}`).join("\n\n---\n\n");

  const api = { build, text, all, SOON, NONE };
  root.KeelsonChase = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
