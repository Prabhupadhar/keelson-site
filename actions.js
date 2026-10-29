/* Keelson actions, decisions and risks log: read meeting notes (marker lines) or a table (Keelson's template, a RAID log,
   a table pasted back from the PM's own AI tool), merge into a per-project log carried forward week to week, and
   summarise it for the reports and the chase list. Pure: no DOM, no storage, no clock, no network. No AI runs here.
   Dates are day numbers (KeelsonEngine). Tests: tests/web/actions.test.js. Pro in the app. */
(function (root) {
  "use strict";
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const TYPES = ["Action", "Decision", "Risk", "Issue"];
  const LIM = { items: 2000, text: 400, field: 120 };
  const clip = (s, n) => String(s ?? "").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").replace(/ /g, " ").replace(/\s+/g, " ").trim().slice(0, n);
  const key = s => E.norm(String(s ?? "")).replace(/[^\p{L}\p{N} ]+/gu, "").replace(/\s+/g, " ").trim();

  // ---------- marker words, English first, then es, fr, de, pt ----------
  const MARK = {
    Action: ["action", "actions", "action item", "ai", "ap", "todo", "to do", "task", "acción", "accion", "tarea", "maßnahme", "massnahme", "aufgabe", "ação", "acao", "tarefa"],
    Decision: ["decision", "decided", "agreed", "decisión", "decisión", "acuerdo", "décision", "entscheidung", "beschluss", "decisão", "decisao"],
    Risk: ["risk", "riesgo", "risque", "risiko", "risco"],
    Issue: ["issue", "problem", "blocker", "problema", "problème", "probleme", "incidencia"],
  };
  const MARKS = Object.entries(MARK).flatMap(([t, ws]) => ws.map(w => [w, t])).sort((a, b) => b[0].length - a[0].length);
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const ALT = MARKS.map(m => esc(m[0])).join("|");
  const MARK_RE = new RegExp(`^\\s*(?:[-*•·▪◦]|\\d+[.)])?\\s*(?:\\[\\s*([xX✓ ]?)\\s*\\]\\s*)?(?:[\\[(](${ALT})[\\])]\\s*[:：]?|(${ALT})\\s*[:：])\\s*(.+)$`, "iu");

  // ---------- dates in free text ----------
  const MON = "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december|ene|abr|ago|dic|janv|févr|fevr|avr|mai|juin|juil|août|aout|déc|mär|mrz|okt|dez|fev|out|set";
  const DATE_RES = [
    /\b\d{4}-\d{1,2}-\d{1,2}\b/,
    /\b\d{1,2}[/.]\d{1,2}(?:[/.]\d{2,4})?\b/,
    new RegExp(`\\b(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\\.?,?\\s+)?\\d{1,2}(?:st|nd|rd|th)?\\s+(?:${MON})\\.?(?:\\s+\\d{4})?\\b`, "i"),
    new RegExp(`\\b(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\\.?,?\\s+)?(?:${MON})\\.?\\s+\\d{1,2}(?:st|nd|rd|th)?(?:,?\\s+\\d{4})?\\b`, "i"),
  ];
  // a date string, with the year taken from the meeting when it is left out (and rolled on if that lands well before it)
  function dateIn(s, ref, dayFirst) {
    s = String(s).replace(/\b(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/i, "").replace(/(\d)(st|nd|rd|th)\b/i, "$1").replace(/,/g, " ").replace(/\s+/g, " ").trim();
    let d = E.parseDate(s, dayFirst);
    if (d == null && ref != null && !/\d{4}/.test(s)) {
      const y = +E.iso(ref).slice(0, 4);
      d = E.parseDate(/^\d{1,2}[/.]\d{1,2}$/.test(s) ? `${s}/${y}`.replace(/\./g, "/") : `${s} ${y}`, dayFirst);
      if (d != null && d < ref - 60) d = E.parseDate(/^\d{1,2}[/.]\d{1,2}$/.test(s) ? `${s}/${y + 1}`.replace(/\./g, "/") : `${s} ${y + 1}`, dayFirst);
    }
    const n = /^(\d{1,2})[/.](\d{1,2})/.exec(s); // reject 31/02 and the like instead of rolling them into March
    if (d != null && n && +E.iso(d).slice(8) !== +(dayFirst ? n[1] : n[2])) return null;
    return d;
  }
  // a weekday in front belongs to the date only when it is that date's weekday ("since Monday Dec 14": Dec 14 is a Tuesday)
  const WD = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"], wdOf = d => WD[new Date(Date.UTC(...E.iso(d).split("-").map((v, i) => i === 1 ? v - 1 : +v))).getUTCDay()];
  const findDate = (s, ref, dayFirst) => {
    for (const re of DATE_RES) {
      const m = re.exec(s); if (!m) continue;
      const d = dateIn(m[0], ref, dayFirst); if (d == null) continue;
      const w = /^(mon|tue|wed|thu|fri|sat|sun)[a-z]*\.?,?\s+/i.exec(m[0]);
      return w && w[1].toLowerCase() !== wdOf(d) ? { d, at: m.index + w[0].length, len: m[0].length - w[0].length } : { d, at: m.index, len: m[0].length };
    }
    return null;
  };

  // ---------- meeting notes: one item per marker line; unmarked lines are reported as skipped ----------
  const NOT_NAME = `(?!(?:${MON}|mon|tue|wed|thu|fri|sat|sun|by|due|on)\\b)`; // "@Tomás Jun 23": Jun is the date, not a surname
  const AT_RE = new RegExp(`@([\\p{L}][\\p{L}\\p{N}_'’.-]*(?:\\s+${NOT_NAME}\\p{Lu}[\\p{L}'’.-]*)?)`, "iu");
  const NAME = /^\p{Lu}[\p{L}'’.-]*(?:\s+\p{Lu}[\p{L}'’.-]*){0,2}$/u;
  function parseNotes(text, opt = {}) {
    const ref = opt.meetingDate ?? null, dayFirst = opt.dayFirst !== false, items = [], skipped = [];
    const lines = String(text ?? "").split(/\r\n|\r|\n/).slice(0, 5000);
    lines.forEach((raw, i) => {
      const line = raw.replace(/ /g, " ");
      if (!line.trim()) return;
      const m = MARK_RE.exec(line);
      if (!m) { if (/\p{L}/u.test(line)) skipped.push({ line: i + 1, text: clip(line, LIM.text) }); return; }
      const type = MARKS.find(x => x[0] === (m[2] || m[3]).toLowerCase())?.[1] || "Action";
      let rest = m[4], owner = "", due = null, status = /[xX✓]/.test(m[1] || "") ? "Done" : undefined; // only an explicit mark changes a saved status
      if (/\((?:done|closed|complete[d]?)\)\s*$/i.test(rest)) { status = "Done"; rest = rest.replace(/\((?:done|closed|complete[d]?)\)\s*$/i, ""); }
      const at = AT_RE.exec(rest);
      if (at) { owner = at[1].replace(/[.’']+$/, ""); rest = rest.replace(at[0], " \u2063 "); } // a mark where the name was, so the words around it don't join
      if (type === "Action" || type === "Risk" || type === "Issue") { // decisions keep their dates as text
        const parts = rest.split(/\s+[|–—-]\s+|\s*\|\s*/).map(p => p.trim()).filter(Boolean);
        if (parts.length > 1) {
          const keep = [];
          for (const p of parts) {
            const pd = !due && dateIn(p.replace(/^(?:by|due|bis|avant|antes de|até|para el|le)\s+/i, ""), ref, dayFirst);
            if (pd != null && pd !== false) { due = pd; continue; }
            if (!owner && NAME.test(p) && parts.length > 2) { owner = p; continue; }
            keep.push(p);
          }
          if (!owner && keep.length > 1 && NAME.test(keep[keep.length - 1])) owner = keep.pop();
          else if (!owner && keep.length > 1 && NAME.test(keep[0])) owner = keep.shift();
          rest = keep.join(" – ");
        }
        if (due == null) { // risks and issues: the date is the review date
          const f = findDate(rest, ref, dayFirst);
          if (f) { due = f.d; rest = (rest.slice(0, f.at).replace(/\s*\b(?:by|due|on|bis|avant le|avant|antes del|antes de|até|para el|para|le)\s*$/i, "") + " " + rest.slice(f.at + f.len)).trim(); }
        }
        if (!owner && type === "Action") { // "Ravi to send the contract"
          const to = /^(\p{Lu}[\p{L}'’.-]*(?:\s+(?:and|&|y|et|und|e)\s+\p{Lu}[\p{L}'’.-]*|\s+\p{Lu}[\p{L}'’.-]*){0,2})\s+(?:to|will|va a|va|doit|soll|wird|vai)\s+(.+)$/u.exec(rest);
          if (to) { owner = to[1]; rest = to[2].charAt(0).toUpperCase() + to[2].slice(1); }
        }
      }
      rest = rest.replace(/\s*\u2063\s*/g, " ").replace(/\s+[–—-]\s*$|^\s*[–—-]\s+/g, "").replace(/\s{2,}/g, " ").trim();
      if (!/[\p{L}\p{N}]/u.test(rest)) return;
      items.push({ type, text: clip(rest, LIM.text), owner: clip(owner, LIM.field), due, status, raised: ref, ref: "", line: i + 1 });
    });
    return { items, skipped };
  }

  // ---------- tables: the template, RAID logs, exports, an AI tool's answer ----------
  const COLS = {
    ref: ["id", "ref", "reference", "no", "number", "#", "key", "item id", "action id", "risk id", "issue id", "raid id", "log id"],
    type: ["type", "item type", "category", "kind", "raid", "raid type", "log type", "tipo", "typ", "art"],
    text: ["item", "description", "action", "actions", "action item", "title", "summary", "details", "decision", "risk", "issue", "subject", "task", "risk description", "issue description", "action description", "descripción", "descripcion", "beschreibung", "descrição", "descricao", "élément", "intitulé"],
    owner: ["owner", "action owner", "risk owner", "issue owner", "assigned to", "assignee", "responsible", "who", "by whom", "actionee", "responsable", "verantwortlich", "responsável", "propietario"],
    due: ["due", "due date", "target date", "target", "by when", "deadline", "date due", "review date", "next review", "target resolution date", "fecha límite", "fälligkeit", "fällig", "échéance", "prazo"],
    status: ["status", "state", "rag", "estado", "statut", "état", "zustand", "situação"],
    raised: ["meeting date", "raised", "date raised", "raised on", "date", "logged", "date logged", "opened", "created", "date identified", "identified", "meeting"],
    firstDue: ["first due", "original due", "original due date", "baseline due"],
    closed: ["closed", "date closed", "closed on", "completed on", "resolved on"],
    score: ["score", "rating", "risk score", "severity", "priority", "exposure"],
    mitigation: ["mitigation", "response", "mitigation plan", "action plan", "contingency", "resolution"],
  };
  const COL_KEYS = Object.fromEntries(Object.entries(COLS).map(([f, ws]) => [f, ws.map(key)]));
  function headerMap(row) {
    const cols = {}, used = new Set();
    for (const [f, ws] of Object.entries(COL_KEYS)) {
      const j = row.findIndex((h, j) => !used.has(j) && ws.includes(key(h)));
      if (j >= 0) { cols[f] = j; used.add(j); }
    }
    return cols;
  }
  const typeOf = (s, fallback) => {
    const w = key(s); if (!w) return fallback;
    for (const t of TYPES) if (MARK[t].some(m => w === key(m) || w.startsWith(key(m) + " ") || w === key(m) + "s")) return t;
    return /^r$/.test(w) ? "Risk" : /^a$/.test(w) ? "Action" : /^i$/.test(w) ? "Issue" : /^d$/.test(w) ? "Decision" : fallback;
  };
  function parseTableText(text, opt = {}) {
    const grid = E.parseTable(String(text ?? "")).slice(0, LIM.items + 50), dayFirst = opt.dayFirst !== false;
    let hdr = -1, cols = {};
    for (let i = 0; i < Math.min(grid.length, 20); i++) { const c = headerMap(grid[i]); if (c.text != null && Object.keys(c).length >= 2) { hdr = i; cols = c; break; } }
    if (hdr < 0) return null;
    const textHead = key(grid[hdr][cols.text]), fallback = opt.type || typeOf(textHead.replace(/ description$/, ""), "Action");
    const items = [], skipped = [];
    const day = v => { const s = String(v ?? "").trim(); if (!s || /^(-+|—|n\/?a|tbd|tbc|#n\/a|\?)$/i.test(s)) return null; return dateIn(s, opt.meetingDate ?? null, dayFirst) ?? null; };
    for (let i = hdr + 1; i < grid.length; i++) {
      const r = grid[i], t = clip(r[cols.text], LIM.text);
      if (!t) continue;
      if (key(t) === textHead || /^(total|totals|sum)\b/i.test(t)) { skipped.push({ line: i + 1, text: t }); continue; } // repeated header, totals
      const st = cols.status != null ? E.statusWord(r[cols.status]) : "";
      items.push({ type: cols.type != null ? typeOf(r[cols.type], fallback) : fallback, text: t, owner: E.noOwner(r[cols.owner]) ? "" : clip(r[cols.owner], LIM.field),
        due: cols.due != null ? day(r[cols.due]) : null, status: st || undefined,
        raised: cols.raised != null ? day(r[cols.raised]) : opt.meetingDate ?? null, ref: cols.ref != null ? clip(r[cols.ref], 40) : "",
        firstDue: cols.firstDue != null ? day(r[cols.firstDue]) : null, closed: cols.closed != null ? day(r[cols.closed]) : null,
        score: cols.score != null ? clip(r[cols.score], 40) : "", mitigation: cols.mitigation != null ? clip(r[cols.mitigation], LIM.text) : "", line: i + 1 });
    }
    return { items: items.slice(0, LIM.items), skipped, headers: Object.fromEntries(Object.entries(cols).map(([f, j]) => [f, grid[hdr][j]])) };
  }
  // pasted text: a table when it has a recognisable header row, otherwise meeting notes
  const read = (text, opt = {}) => { const t = parseTableText(text, opt); return t && t.items.length ? { kind: "table", ...t } : { kind: "notes", ...parseNotes(text, opt) }; };

  // ---------- the log: merge new items week to week ----------
  const idOf = it => it.ref ? "ref:" + key(it.ref) : it.type.toLowerCase() + ":" + key(it.text);
  function merge(log, incoming, asOf) {
    const out = (log || []).map(x => ({ ...x })), at = {}, res = { added: 0, updated: 0, same: 0 };
    out.forEach((x, i) => { at[x.id] = i; if (x.ref) at[x.type.toLowerCase() + ":" + key(x.text)] ??= i; });
    for (const it of incoming) {
      const id = idOf(it), i = at[id] ?? at[it.type.toLowerCase() + ":" + key(it.text)];
      if (i == null) {
        if (out.length >= LIM.items) break;
        const st = it.status || "Open";
        out.push({ id, ref: it.ref || "", type: it.type, text: it.text, owner: it.owner || "", due: it.due ?? null, firstDue: it.firstDue ?? it.due ?? null,
          status: st, raised: it.raised ?? asOf, added: asOf, updated: asOf, closed: st !== "Open" ? it.closed ?? asOf : null, score: it.score || "", mitigation: it.mitigation || "" });
        at[id] = out.length - 1; res.added++; continue;
      }
      const x = out[i], before = JSON.stringify(x);
      if (it.owner) x.owner = it.owner;
      if (it.due != null) { x.firstDue ??= x.due ?? it.due; x.due = it.due; }
      if (it.status && it.status !== x.status) { x.status = it.status; x.closed = it.status === "Open" ? null : it.closed ?? asOf; }
      if (it.score) x.score = it.score;
      if (it.mitigation) x.mitigation = it.mitigation;
      if (JSON.stringify(x) !== before) { x.updated = asOf; res.updated++; } else res.same++;
    }
    return { log: out, ...res };
  }
  const setStatus = (log, id, status, asOf) => log.map(x => x.id !== id ? x : { ...x, status, closed: status === "Open" ? null : asOf, updated: asOf });
  const remove = (log, id) => log.filter(x => x.id !== id);

  // ---------- what the reports and the chase list need ----------
  function summary(log, sd, days = 7) { // "this week": from the last weekly status date (sd - 7) to sd, both included
    const open = (log || []).filter(x => x.status === "Open"), since = sd - days;
    const acts = open.filter(x => x.type === "Action" || x.type === "Issue");
    const late = x => x.due != null && x.due < sd;
    return {
      openActions: acts.length,
      overdue: acts.filter(late).sort((a, b) => a.due - b.due),
      dueSoon: acts.filter(x => x.due != null && x.due >= sd && x.due <= E.addWd(sd, 5)).sort((a, b) => a.due - b.due),
      slipped: acts.filter(x => x.due != null && x.firstDue != null && x.due > x.firstDue).map(x => ({ ...x, slip: E.wdiff(x.firstDue, x.due) })).sort((a, b) => b.slip - a.slip),
      noOwner: acts.filter(x => !x.owner),
      decisions: (log || []).filter(x => x.type === "Decision" && (x.raised ?? x.added) >= since && (x.raised ?? x.added) <= sd),
      risks: open.filter(x => x.type === "Risk"),
      closedRecently: (log || []).filter(x => x.status === "Done" && x.closed != null && x.closed >= since && x.closed <= sd),
      added: (log || []).filter(x => x.added >= since && x.added <= sd),
    };
  }

  // ---------- out: CSV in the template's columns (Excel opens it; Keelson reads it back) ----------
  const CSV_HEAD = ["ID", "Type", "Item", "Owner", "Due", "Status", "Meeting date", "First due", "Closed", "Score", "Mitigation"];
  const cell = v => { let s = String(v ?? ""); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }; // no formulas from plan text
  const d = v => v == null ? "" : E.iso(v);
  const toCSV = log => "﻿" + [CSV_HEAD, ...(log || []).map(x => [x.ref || "", x.type, x.text, x.owner, d(x.due), x.status, d(x.raised), d(x.firstDue), d(x.closed), x.score, x.mitigation])]
    .map(r => r.map(cell).join(",")).join("\r\n") + "\r\n";
  const TEMPLATE = "﻿" + [CSV_HEAD.slice(0, 7).join(","), "A-1,Action,Send the vendor contract for signature,Ravi Nair,2027-02-12,Open,2027-02-05", "D-1,Decision,Go-live stays on 3 Jun,,,Open,2027-02-05",
    "R-1,Risk,Data migration may slip if the legacy export is late,Grace Okafor,2027-02-19,Open,2027-02-05", "I-1,Issue,Test environment down since Monday,Sam Ortiz,2027-02-09,Open,2027-02-05"].join("\r\n") + "\r\n";
  const PROMPT = `From the meeting notes or transcript below, list every action, decision, risk and issue as a table with exactly these columns, tab-separated, one row each, header row first:
Type	Item	Owner	Due	Status	Meeting date
Type is Action, Decision, Risk or Issue. Item is one short sentence. Owner is the person's name as written, or blank. Due is YYYY-MM-DD, or blank if no date was given. Status is Open unless the notes say it is done. Meeting date is YYYY-MM-DD. Do not invent owners or dates. Output only the table.

Notes:
`;

  const api = { TYPES, read, parseNotes, parseTable: parseTableText, merge, setStatus, remove, summary, toCSV, TEMPLATE, PROMPT, CSV_HEAD, idOf, LIM };
  root.KeelsonActions = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
