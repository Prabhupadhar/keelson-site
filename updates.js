/* Keelson update round trip: one small sheet per owner out, the filled sheets back in, one change list to apply.
   round(k) picks each owner's open tasks; sheet() lays one owner's out for xlsx-write.js; read() takes a returned file
   (xlsx-read.js sheets, or CSV/pasted rows); review() checks the replies against the plan as it is now.
   Pure: no DOM, no clock, no network. Keelson never sends the sheets and never writes into the plan file.
   Pro in the app. Spec: docs/PM-TOOLKIT.md section 5. Tests: tests/web/updates.test.js. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const AHEAD = 10; // working days: tasks active or due in the next two weeks
  const NONE = "No owner", MARK = "keelson-updates", VER = "1";
  const HEAD = ["Ref", "ID", "Task", "Phase", "Start", "Finish now", "% now", "New %", "New finish", "Done?", "Blocked by", "Comment"];
  const C = Object.fromEntries(HEAD.map((h, i) => [h, i]));
  const LIM = { tasks: 200, text: 500, files: 50 };
  const clean = s => String(s ?? "").replace(/[\u0000-\u001f\u007f]/g, " ").replace(/ /g, " ").replace(/\s+/g, " ").trim();
  const norm = s => clean(s).toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
  const ownerOf = t => E.noOwner(t.owner) ? NONE : clean(t.owner);
  const pctOf = t => t.pct == null ? null : Math.round(t.pct * 100); // the engine keeps % complete as 0-1

  // each owner's open tasks that are late, under way or due within AHEAD working days; owners A-Z, "No owner" last
  function round(k, ahead = AHEAD) {
    if (!k || !k.tasks || k.cols.name < 0 || k.cols.finish < 0) return { asOf: null, groups: [] };
    const sd = k.statusDate, until = E.addWd(sd, ahead), by = {};
    for (const t of k.tasks) {
      if (t.done) continue;
      if (!(t.od || (t.finish != null && t.finish <= until) || (t.start != null && t.start <= until))) continue;
      (by[ownerOf(t)] = by[ownerOf(t)] || []).push(t);
    }
    const groups = Object.entries(by).map(([owner, ts]) => ({ owner, tasks: ts.sort((a, b) => (b.od ? 1 : 0) - (a.od ? 1 : 0) || (a.finish ?? 1e9) - (b.finish ?? 1e9)).slice(0, LIM.tasks) }))
      .sort((a, b) => (a.owner === NONE) - (b.owner === NONE) || a.owner.localeCompare(b.owner));
    return { asOf: sd, groups };
  }

  // one owner's workbook for KeelsonXlsxWrite.book: their tasks, blue columns to fill, a hidden tab that ties it to the round
  function sheet(g, asOf, project) {
    const d = v => v == null ? null : { d: v };
    const rows = [HEAD, ...g.tasks.map((t, i) => [`U${i + 1}`, t.uid || "", t.name, t.phase || "", d(t.start), d(t.finish), pctOf(t) ?? "", null, null, null, null, null])];
    const meta = [["kind", MARK], ["version", VER], ["round", E.iso(asOf)], ["owner", g.owner], ["project", clean(project)], ["ref", "key", "id", "phase", "name"],
      ...g.tasks.map((t, i) => [`U${i + 1}`, t.key, t.uid || "", t.phase || "", t.name])];
    const n = C["New %"];
    return [{ name: "Your tasks", rows, widths: [6, 10, 44, 18, 12, 12, 8, 9, 13, 9, 22, 36], protect: true,
      input: { [n]: "number", [n + 1]: "date", [n + 2]: "text", [n + 3]: "text", [n + 4]: "text" },
      checks: [{ col: n, min: 0, max: 100 }, { col: n + 1, date: true }, { col: n + 2, list: ["Yes", "No"] }] },
      { name: "keelson", rows: meta, hidden: true }];
  }
  const fileName = (g, asOf) => `Update - ${clean(g.owner).replace(/[\\/:*?"<>|]+/g, " ").slice(0, 60)} - ${E.fmt(asOf, true)}.xlsx`;

  // the note that goes with the sheet; plain text for email or chat
  function message(g, asOf, project, by) {
    const first = g.owner === NONE ? "" : g.owner.split(/[\s,;]+/)[0], n = g.tasks.length, late = g.tasks.filter(t => t.od).length;
    return [g.owner === NONE ? "These tasks have no owner in the plan. Who can update them?" : `Hi ${first},`, "",
      `Attached ${n === 1 ? "is the task" : `are the ${n} tasks`} I have for you${project ? ` on ${clean(project)}` : ""}${late ? ` (${late} past ${late === 1 ? "its" : "their"} finish date)` : ""}.`,
      `Could you fill in the blue columns${by != null ? ` by ${E.fmt(by)}` : ""}: new % done, a new finish date if it has moved, Done? and anything blocking you.`,
      "Leave a cell empty if nothing changed, and send the file back as it is.", "", "Thanks"].join("\n");
  }

  // ---------- in ----------
  const num = v => { const s = clean(v).replace(/,/g, "."), m = /^(-?\d+(?:\.\d+)?)\s*%?$/.exec(s); if (!m) return NaN; const x = +m[1]; return !s.includes("%") && s.includes(".") && x > 0 && x <= 1 ? Math.round(x * 100) : x; };
  const yes = v => /^(y|yes|done|complete|completed|true|x|✓|✔)$/i.test(clean(v)), no = v => /^(n|no|false|not yet)$/i.test(clean(v));
  const findHead = rows => { for (let i = 0; i < Math.min(rows.length, 10); i++) { const h = (rows[i] || []).map(norm); if (h.includes("task") && (h.includes("new") || h.includes("new finish"))) return i; } return -1; };

  // sheets: [{name, rows, hidden}] from xlsx-read.js, or {text} for CSV / pasted rows
  // -> {owner, round, meta (bool), rows: [{ref, key, id, name, phase, pct, pctRaw, finish, finishRaw, done, blocked, comment}]} or {error}
  function read(input, dayFirst = true) {
    const sheets = input && input.text != null ? [{ name: "", rows: E.parseTable(String(input.text)), hidden: false }] : (input || []);
    const metaS = sheets.find(s => s.rows && s.rows[0] && clean(s.rows[0][0]) === "kind" && clean(s.rows[0][1]) === MARK);
    const meta = {}, refs = {};
    if (metaS) { let inRefs = false; for (const r of metaS.rows) { if (!r) continue; if (clean(r[0]) === "ref") { inRefs = true; continue; } if (inRefs) refs[clean(r[0])] = { key: String(r[1] ?? ""), id: clean(r[2]), phase: clean(r[3]), name: clean(r[4]) }; else meta[clean(r[0])] = clean(r[1]); } }
    const data = sheets.filter(s => s !== metaS && s.rows).map(s => ({ s, h: findHead(s.rows) })).find(x => x.h >= 0);
    if (!data) return { error: "This doesn't look like an update sheet: no Task and New % or New finish columns." };
    const hr = data.s.rows[data.h].map(norm), col = name => hr.indexOf(norm(name));
    const at = Object.fromEntries(HEAD.map(h => [h, col(h)])), cell = (r, h) => at[h] >= 0 ? r[at[h]] : "";
    const rows = [];
    for (const r of data.s.rows.slice(data.h + 1)) {
      if (!r || !r.some(c => clean(c))) continue;
      const ref = clean(cell(r, "Ref")), m = refs[ref], pctRaw = clean(cell(r, "New %")), finishRaw = clean(cell(r, "New finish")), doneRaw = clean(cell(r, "Done?"));
      const fin = finishRaw ? E.parseDate(finishRaw, dayFirst) : null;
      rows.push({ ref, key: m ? m.key : null, id: clean(cell(r, "ID")) || (m ? m.id : ""), name: clean(cell(r, "Task")).slice(0, LIM.text) || (m ? m.name : ""), phase: clean(cell(r, "Phase")) || (m ? m.phase : ""),
        pct: pctRaw ? num(pctRaw) : null, pctRaw, finish: finishRaw ? (fin ?? NaN) : null, finishRaw, done: doneRaw ? (yes(doneRaw) ? true : no(doneRaw) ? false : null) : null, doneRaw,
        blocked: clean(cell(r, "Blocked by")).slice(0, LIM.text), comment: clean(cell(r, "Comment")).slice(0, LIM.text) });
    }
    return { owner: meta.owner || "", round: meta.round ? E.parseDate(meta.round, false) : null, meta: !!metaS, rows };
  }

  // replies: [{file, owner, round, rows}] from read(); k: the plan loaded now; sent: the open round's owners (optional)
  // -> {changes: [{owner, file, t, field, old, new, flags: [..]}], notes: [{owner, t, blocked, comment}], unknown: [{owner, row}],
  //     conflicts: [key], missing: [owner]}
  function review(k, replies, sent) {
    const tasks = (k && k.tasks) || [], byKey = new Map(tasks.map(t => [t.key, t])), sd = k && k.statusDate;
    const byId = new Map(), byPN = new Map(), byN = new Map();
    const add = (m, x, t) => { if (!x) return; m.set(x, m.has(x) ? null : t); }; // null: more than one task, so don't guess
    for (const t of tasks) { add(byId, clean(t.uid), t); add(byPN, (t.phase || "") + "\u0001" + norm(t.name), t); add(byN, norm(t.name), t); }
    const find = r => (r.key && byKey.get(r.key)) || (r.id && byId.get(r.id)) || byPN.get((r.phase || "") + "\u0001" + norm(r.name)) || byN.get(norm(r.name)) || null;
    const changes = [], notes = [], unknown = [], seen = new Map();
    for (const rep of replies || []) {
      const owner = rep.owner || NONE;
      for (const r of rep.rows) {
        const touched = r.pct != null || r.finish != null || r.done != null || r.blocked || r.comment;
        if (!touched) continue;
        const t = find(r);
        if (!t) { unknown.push({ owner, file: rep.file, row: r }); continue; }
        const push = (field, old, nu, flags = []) => { changes.push({ owner, file: rep.file, t, field, old, new: nu, flags }); const s = seen.get(t.key) || new Set(); s.add(owner); seen.set(t.key, s); };
        const done = r.done === true || (r.pct === 100 && r.done !== false);
        const was = pctOf(t);
        if (r.pct != null) {
          const f = []; if (!Number.isFinite(r.pct) || r.pct < 0 || r.pct > 100) f.push(`“${r.pctRaw}” isn't a % from 0 to 100`); else if (was != null && r.pct < was) f.push("% went down");
          if (!done && t.od && r.finish == null) f.push("late, and no new date given");
          if (!Number.isFinite(r.pct) || r.pct !== was) push("% done", was, Number.isFinite(r.pct) ? r.pct : r.pctRaw, f);
        }
        if (r.done === true && !t.done && r.pct !== 100) push("% done", was, 100);
        if (r.done == null && r.doneRaw) push("Done?", "", r.doneRaw, [`“${r.doneRaw}” isn't Yes or No`]);
        if (r.finish != null) {
          const f = [];
          if (!Number.isFinite(r.finish)) f.push(`“${r.finishRaw}” isn't a date`);
          else { if (t.start != null && r.finish < t.start) f.push("finishes before it starts"); if (!done && sd != null && r.finish < sd) f.push("new date is already past"); }
          if (!Number.isFinite(r.finish) || r.finish !== t.finish) push("Finish", t.finish, Number.isFinite(r.finish) ? r.finish : r.finishRaw, f);
        }
        if (r.blocked || r.comment) notes.push({ owner, t, blocked: r.blocked, comment: r.comment });
      }
    }
    const conflicts = [...seen].filter(([, s]) => s.size > 1).map(([key]) => key);
    for (const c of changes) if (conflicts.includes(c.t.key)) c.flags.push("also updated by " + [...seen.get(c.t.key)].filter(o => o !== c.owner).join(", "));
    const got = new Set((replies || []).map(r => r.owner).filter(Boolean));
    return { changes, notes, unknown, conflicts, missing: (sent || []).filter(o => !got.has(o)) };
  }

  // the change list to apply in the plan tool, for KeelsonXlsxWrite.book
  const show = (field, v) => v == null || v === "" ? "" : field === "Finish" && Number.isFinite(v) ? { d: v } : v;
  function changeSheet(rv, asOf) {
    const rows = [["ID", "Task", "Phase", "Field", "Now in plan", "New", "Owner", "Check", "Blocked by", "Comment"]];
    const note = new Map(rv.notes.map(n => [n.t.key + "\u0001" + n.owner, n]));
    for (const c of rv.changes) { const n = note.get(c.t.key + "\u0001" + c.owner); note.delete(c.t.key + "\u0001" + c.owner);
      rows.push([c.t.uid || "", c.t.name, c.t.phase || "", c.field, show(c.field, c.old), show(c.field, c.new), c.owner, c.flags.join("; "), n ? n.blocked : "", n ? n.comment : ""]); }
    for (const n of note.values()) rows.push([n.t.uid || "", n.t.name, n.t.phase || "", "Note", "", "", n.owner, "", n.blocked, n.comment]);
    for (const u of rv.unknown) rows.push([u.row.id, u.row.name, u.row.phase, "Not in the plan", "", "", u.owner, "task not found in the plan you have loaded", u.row.blocked, u.row.comment]);
    return [{ name: `Changes ${E.fmt(asOf, true)}`, rows, widths: [10, 44, 18, 10, 13, 13, 16, 30, 22, 36] }];
  }

  const api = { round, sheet, fileName, message, read, review, changeSheet, HEAD, NONE, MARK, AHEAD, LIM };
  root.KeelsonUpdates = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
