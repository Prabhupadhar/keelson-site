/* Keelson paperwork: the four things a PM writes by hand every day or week, as plain text ready to paste.
     today      what is overdue, due today, due in the next 5 working days and starting, by person (stand-up lines)
     message    this week's status as an email or chat message for the sponsor, the client or the team
     lookahead  every open task in the next 2, 4 or 6 weeks, by person; also as rows for an Excel sheet
     agenda     the next progress meeting's agenda from the plan and the actions log
   From the engine's result (k) and, optionally, the actions log and the counts of what changed. Every number is one
   the reports already show. Pure: no DOM, no clock, no network. Keelson sends nothing. Tests: tests/web/brief.test.js. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const SOON = 5, NONE = "No owner", TOP = 8;
  const D = v => E.fmt(v, true), d = v => E.fmt(v);
  const who = o => E.noOwner(o) ? NONE : String(o).replace(/ /g, " ").replace(/\s+/g, " ").trim();
  const n = (c, w) => `${c} ${w}${c === 1 ? "" : "s"}`;
  const pc = x => Math.round(x * 100) + "%";
  const ok = k => !!(k && k.tasks && k.cols.name >= 0 && k.cols.finish >= 0);
  const openTasks = k => k.tasks.filter(t => !t.done && t.finish != null);
  const byOwner = (items, sort) => { const by = {}; for (const it of items) (by[who(it.t.owner)] = by[who(it.t.owner)] || []).push(it);
    return Object.entries(by).map(([owner, list]) => ({ owner, items: list.sort(sort) })).sort((a, b) => (a.owner === NONE) - (b.owner === NONE) || a.owner.localeCompare(b.owner)); };
  const more = (list, line) => [...list.slice(0, TOP).map(line), ...(list.length > TOP ? [`- and ${list.length - TOP} more`] : [])];

  // the facts every message shares, in the reports' own words
  const STATE = { Green: "on track", Amber: "at risk", Red: "off track" };
  const status = k => STATE[k.rag] ? `Status: ${k.rag}, ${STATE[k.rag]}. ${(k.ragReason || (k.why || []).join("; "))}.`.replace(/\.\.$/, ".") : "Status: not enough dated tasks to judge.";
  const finish = k => k.finish == null ? "" : `Forecast finish: ${D(k.finish)}` + (k.bfinish == null ? "." : k.var > 0 ? `, ${n(k.var, "working day")} after the baseline of ${D(k.bfinish)}.`
    : k.var < 0 ? `, ${n(-k.var, "working day")} before the baseline of ${D(k.bfinish)}.` : ", on the baseline.");
  const progress = k => k.pct == null ? "" : `Progress: ${pc(k.pct)} complete${k.exp != null ? ` against ${pc(k.exp)} expected by now` : ""}; ${k.nTaskDone} of ${n(k.nTask, "task")} done.`;
  const nextMs = k => k.nextMs ? `Next milestone: ${k.nextMs.name}, ${D(k.nextMs.finish)}.` : "";
  // changes: { since: day number, counts } from the comparison the reports use (diff.js counts)
  function changed(c) {
    if (!c || !c.counts) return "";
    const x = c.counts, later = (x.finishLater || 0), parts = [x.added && `${n(x.added, "task")} added`, x.removed && `${x.removed} removed`, later && `${later} finishing later`,
      x.finishEarlier && `${x.finishEarlier} finishing earlier`, x.milestoneMoved && `${n(x.milestoneMoved, "milestone")} moved`].filter(Boolean);
    return `Since ${D(c.since)}: ${parts.length ? parts.join(", ") : "no dates moved and no tasks added or removed"}.`;
  }

  // ---- today: stand-up lines ----
  function today(k) {
    if (!ok(k)) return [];
    const sd = k.statusDate, until = E.addWd(sd, SOON), order = { overdue: 0, today: 1, soon: 2, start: 3 }, items = [];
    for (const t of openTasks(k)) {
      if (t.od) items.push({ t, kind: "overdue" });
      else if (t.finish === sd) items.push({ t, kind: "today" });
      else if (t.finish > sd && t.finish <= until) items.push({ t, kind: "soon" });
      else if (t.start != null && t.start >= sd && t.start <= until) items.push({ t, kind: "start" });
    }
    return byOwner(items, (a, b) => order[a.kind] - order[b.kind] || a.t.finish - b.t.finish).map(g => ({ owner: g.owner,
      items: g.items.map(({ t, kind }) => ({ name: t.name, kind, start: t.start, finish: t.finish, late: t.late || 0, pct: t.pct, ms: !!t.ms })) }));
  }
  const todayLine = i => i.kind === "overdue" ? `- Overdue: ${i.name} (due ${d(i.finish)}, ${n(i.late, "working day")} late)` : i.kind === "today" ? `- Due today: ${i.name}`
    : i.kind === "soon" ? `- Due ${d(i.finish)}: ${i.name}${i.pct != null ? ` (${pc(i.pct)} done)` : ""}` : `- Starts ${d(i.start)}: ${i.name}`;
  function todayText(groups, project, sd) {
    const head = `Stand-up${project ? ` · ${project}` : ""} · ${D(sd)}`;
    if (!groups.length) return `${head}\n\nNothing overdue, nothing due and nothing starting in the next ${SOON} working days.`;
    return [head, ...groups.map(g => `\n${g.owner}\n${g.items.map(todayLine).join("\n")}`)].join("\n");
  }

  // ---- message: sponsor, client or team ----
  const AUDIENCE = { sponsor: "Sponsor", client: "Client", team: "Team" };
  function message(k, o = {}) {
    if (!ok(k)) return "";
    const a = AUDIENCE[o.audience] ? o.audience : "sponsor", sd = k.statusDate, proj = o.project || "the project";
    const over = openTasks(k).filter(t => t.od).sort((x, y) => y.late - x.late), until = E.addWd(sd, a === "team" ? SOON : 10);
    const due = openTasks(k).filter(t => !t.od && t.finish >= sd && t.finish <= until).sort((x, y) => x.finish - y.finish);
    const owner = t => a === "client" || who(t.owner) === NONE ? "" : `, ${who(t.owner)}`;
    const decisions = (o.actions || []).filter(x => x.type === "Decision" && x.status === "Open");
    const out = [`Subject: ${proj}: status as of ${D(sd)}`, "", a === "team" ? "Hi all," : "Hi,", "", `Here is where ${proj} stands as of ${D(sd)}.`, "",
      ...[status(k), finish(k), progress(k), nextMs(k), changed(o.changes)].filter(Boolean)];
    if (a !== "client" || over.length) out.push("", over.length ? `Overdue (${over.length}):` : "Overdue: none.", ...more(over, t => `- ${t.name} (due ${d(t.finish)}, ${n(t.late, "working day")} late${owner(t)})`));
    if (due.length) out.push("", `Due in the next ${a === "team" ? SOON : 10} working days (${due.length}):`, ...more(due, t => `- ${t.name} (${d(t.finish)}${owner(t)})`));
    if (a === "sponsor") out.push("", decisions.length ? `Decisions needed (${decisions.length}):` : "Decisions needed: none this period.", ...more(decisions, x => `- ${x.text}${x.due != null ? ` (by ${d(x.due)})` : ""}`));
    if (a === "team") out.push("", "Please update your % done and dates in the plan before the next report.");
    return out.join("\n") + "\n\nThanks";
  }

  // ---- look-ahead: by person ----
  function lookahead(k, weeks) {
    if (!ok(k)) return [];
    const sd = k.statusDate, end = sd + (weeks || 2) * 7, items = [];
    for (const t of openTasks(k)) {
      const st = t.start ?? t.finish;
      if (!t.od && (st > end || t.finish < sd)) continue;
      items.push({ t, note: t.od ? "Overdue" : t.ms ? "Milestone" : st >= sd && t.finish <= end ? "Starts and finishes" : st >= sd ? "Starts" : t.finish <= end ? "Finishes" : "Continues" });
    }
    return byOwner(items, (a, b) => (b.t.od ? 1 : 0) - (a.t.od ? 1 : 0) || a.t.finish - b.t.finish).map(g => ({ owner: g.owner,
      items: g.items.map(({ t, note }) => ({ name: t.name, phase: t.phase || "", start: t.start, finish: t.finish, pct: t.pct, late: t.late || 0, note })) }));
  }
  function lookaheadText(groups, project, sd, weeks) {
    const head = `${weeks}-week look-ahead${project ? ` · ${project}` : ""} · ${D(sd)} to ${D(sd + weeks * 7)}`;
    if (!groups.length) return `${head}\n\nNo open tasks in this window.`;
    return [head, ...groups.map(g => `\n${g.owner} (${g.items.length})\n${g.items.map(i => `- ${i.name}: ${i.start != null && i.start !== i.finish ? `${d(i.start)} to ` : ""}${d(i.finish)}` +
      `${i.pct != null ? `, ${pc(i.pct)} done` : ""}${i.note === "Overdue" ? `, overdue by ${n(i.late, "working day")}` : i.note === "Milestone" ? ", milestone" : ""}`).join("\n")}`)].join("\n");
  }
  // one sheet for xlsx-write.js: dates as {d}, % as a number
  function lookaheadSheet(groups, weeks) {
    const rows = [["Owner", "Task", "Phase", "Start", "Finish", "% done", "This window"]];
    for (const g of groups) for (const i of g.items) rows.push([g.owner, i.name, i.phase, i.start != null ? { d: i.start } : null, { d: i.finish }, i.pct != null ? Math.round(i.pct * 100) : null, i.note]);
    return [{ name: `${weeks}-week look-ahead`, rows }];
  }

  // ---- agenda: the next progress meeting ----
  function agenda(k, o = {}) {
    if (!ok(k)) return "";
    const sd = k.statusDate, acts = (o.actions || []).filter(x => x.status === "Open"), of = t => acts.filter(x => x.type === t);
    const over = openTasks(k).filter(t => t.od).sort((x, y) => y.late - x.late), until = E.addWd(sd, 10);
    const due = openTasks(k).filter(t => !t.od && t.finish >= sd && t.finish <= until).sort((x, y) => x.finish - y.finish);
    const own = t => who(t.owner) === NONE ? "no owner" : who(t.owner);
    const act = x => `   - ${x.text}${E.noOwner(x.owner) ? "" : ` (${who(x.owner)}${x.due != null ? `, due ${d(x.due)}${x.due < sd ? ", overdue" : ""}` : ""})`}`;
    const actions = [...of("Action"), ...of("Issue")].sort((x, y) => (x.due ?? 1e9) - (y.due ?? 1e9));
    const secs = [
      ["Where we stand", [status(k), finish(k), progress(k), changed(o.changes)].filter(Boolean).map(x => `   ${x}`)],
      over.length && [`Overdue tasks (${over.length})`, more(over, t => `   - ${t.name} (${own(t)}, due ${d(t.finish)}, ${n(t.late, "working day")} late)`)],
      due.length && [`Due in the next 10 working days (${due.length})`, more(due, t => `   - ${t.name} (${own(t)}, ${d(t.finish)})`)],
      k.nextMs && ["Milestones", [`   ${nextMs(k)}`]],
      actions.length && [`Open actions (${actions.length})`, more(actions, act)],
      of("Decision").length && [`Decisions needed (${of("Decision").length})`, more(of("Decision"), act)],
      of("Risk").length && [`Open risks (${of("Risk").length})`, more(of("Risk"), act)],
      ["Any other business", []]].filter(Boolean);
    return [`Agenda · ${o.project || "Project"} progress meeting · plan as of ${D(sd)}`, "", ...secs.flatMap(([t, body], i) => [`${i + 1}. ${t}`, ...body, ""])].join("\n").trimEnd();
  }

  const api = { today, todayText, message, lookahead, lookaheadText, lookaheadSheet, agenda, AUDIENCE, SOON, NONE };
  root.KeelsonBrief = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
