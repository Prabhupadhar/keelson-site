/* Keelson schedule-file reader: MS Project XML (Save As > XML Format) and Primavera P6 XER, read in the browser.
   read(text) -> { kind: "xml" | "xer", title, sd: "YYYY-MM-DD" | "", rows, skipped } or null when the text is neither.
   rows is a sheet (header row first) with the column names the engine already knows, so everything downstream is unchanged.
   Pure functions, no DOM: tests/web/plan-read.test.js runs it under Node. */
(function (root) {
  const HEAD = ["Unique ID", "WBS", "Outline Level", "Task Name", "Phase", "Start", "Finish", "Baseline Start", "Baseline Finish", "Actual Finish",
    "% Complete", "Duration", "Milestone", "Summary", "Resource Names", "Predecessors",
    "Successors", "Total Slack", "Critical", "Constraint Type"]; // the last five feed the plan quality check and the critical-task list (insight.js)
  const day = s => (String(s || "").match(/^\d{4}-\d{2}-\d{2}/) || [""])[0];
  const num = v => { const n = parseFloat(v); return isFinite(n) ? n : 0; };
  const days = n => String(Math.round(n * 100) / 100);
  const yes = b => b ? "Yes" : "No";
  const lag = dd => !dd ? "" : (dd > 0 ? "+" : "") + days(dd) + "d"; // "5FS+2d", "5SS-1d": the way MS Project writes a link
  // successors are the links read the other way: ids in the order the tasks appear
  function withSuccessors(rows, links) { const by = {}; for (const [to, from] of links) (by[from] ||= []).push(to); const iu = HEAD.indexOf("Unique ID"), is = HEAD.indexOf("Successors");
    for (const r of rows.slice(1)) r[is] = (by[r[iu]] || []).join(", "); return rows; }

  // ---- MS Project XML ----
  const ENT = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };
  const ent = s => s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (m, e) => {
    if (e[0] !== "#") return ENT[e.toLowerCase()];
    const c = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return c > 0 && c <= 0x10ffff ? String.fromCodePoint(c) : " ";
  });
  const tag = (s, n) => { const m = s.match(new RegExp(`<${n}>([\\s\\S]*?)</${n}>`)); return m ? ent(m[1].replace(/^<!\[CDATA\[|\]\]>$/g, "")).trim() : ""; };
  const blocks = (s, n) => s.match(new RegExp(`<${n}>[\\s\\S]*?</${n}>`, "g")) || [];
  const LINK = ["FF", "", "SF", "SS"]; // PredecessorLink Type: 0 FF, 1 FS (written bare), 2 SF, 3 SS
  const CSTR = ["", "As Late As Possible", "Must Start On", "Must Finish On", "Start No Earlier Than", "Start No Later Than", "Finish No Earlier Than", "Finish No Later Than"]; // ConstraintType 0-7

  function fromXml(text) {
    const at = text.indexOf("<Tasks>"), head = at < 0 ? text : text.slice(0, at);
    const perDay = num(tag(head, "MinutesPerDay")) || 480;
    const resName = {};
    for (const r of blocks(text, "Resource")) resName[tag(r, "UID")] = tag(r, "Name");
    const who = {};
    for (const a of blocks(text, "Assignment")) { const n = resName[tag(a, "ResourceUID")]; if (n) (who[tag(a, "TaskUID")] ||= []).push(n); }
    const rows = [HEAD], links = []; let skipped = 0;
    for (const b of blocks(text, "Task")) {
      // LinkLag and TotalSlack are in tenths of a minute; a lag given as a percentage (LagFormat 19, 20) has no fixed length, so only its sign is kept
      const preds = blocks(b, "PredecessorLink").map(p => { const u = tag(p, "PredecessorUID"), lg = num(tag(p, "LinkLag")), pctLag = /^(19|20|35|36|51|52)$/.test(tag(p, "LagFormat"));
        return u ? { u, s: u + ((LINK[tag(p, "Type")] ?? "") || (lg ? "FS" : "")) + (pctLag ? (lg > 0 ? "+1d" : lg < 0 ? "-1d" : "") : lag(lg / 10 / perDay)) } : null; }).filter(Boolean);
      const base = blocks(b, "Baseline").find(x => tag(x, "Number") === "0") || "";
      // nested blocks repeat <Start>, <Finish> and <Duration>: read the task's own values without them
      const t = b.replace(/<(PredecessorLink|Baseline|ExtendedAttribute|TimephasedData|OutlineCode)>[\s\S]*?<\/\1>/g, "");
      const uid = tag(t, "UID"), name = tag(t, "Name");
      if (uid === "0" || tag(t, "OutlineLevel") === "0") continue; // the project summary row
      if (tag(t, "IsNull") === "1" || !name) continue; // blank rows in the plan
      if (tag(t, "Active") === "0") { skipped++; continue; }
      const d = tag(t, "Duration").match(/^PT(\d+)H(\d+)M/), slack = tag(t, "TotalSlack");
      for (const p of preds) links.push([uid, p.u]);
      rows.push([uid, tag(t, "WBS") || tag(t, "OutlineNumber"), tag(t, "OutlineLevel"), name, "", day(tag(t, "Start")), day(tag(t, "Finish")),
        day(tag(base, "Start")), day(tag(base, "Finish")), day(tag(t, "ActualFinish")), tag(t, "PercentComplete"),
        d ? days((+d[1] * 60 + +d[2]) / perDay) : "", yes(tag(t, "Milestone") === "1"), yes(tag(t, "Summary") === "1"),
        [...new Set(who[uid] || [])].join(", "), preds.map(p => p.s).join(", "), "", slack === "" ? "" : days(num(slack) / 10 / perDay),
        tag(t, "Critical") === "" ? "" : yes(tag(t, "Critical") === "1"), CSTR[tag(t, "ConstraintType")] || ""]);
    }
    withSuccessors(rows, links);
    return { kind: "xml", title: (tag(head, "Title") || tag(head, "Name")).replace(/\.(mpp|xml)$/i, ""), sd: day(tag(head, "StatusDate")), rows, skipped };
  }

  // ---- Primavera P6 XER: tab-separated tables, %T name, %F fields, %R rows ----
  function tables(text) {
    const out = {}; let cur = null, f = [];
    for (const line of text.split(/\r\n?|\n/)) {
      const c = line.split("\t");
      if (c[0] === "%T") { cur = out[c[1]] = []; f = []; }
      else if (c[0] === "%F") f = c.slice(1);
      else if (c[0] === "%R" && cur) { const o = {}; f.forEach((k, i) => o[k] = c[i + 1] ?? ""); cur.push(o); }
    }
    return out;
  }
  const PRED = { PR_FS: "", PR_SS: "SS", PR_FF: "FF", PR_SF: "SF" };
  const XCSTR = { CS_MSO: "Must Start On", CS_MANDSTART: "Must Start On", CS_MEO: "Must Finish On", CS_MANDFIN: "Must Finish On", CS_MSOA: "Start No Earlier Than", CS_MSOB: "Start No Later Than",
    CS_MEOA: "Finish No Earlier Than", CS_MEOB: "Finish No Later Than", CS_ALAP: "As Late As Possible" };

  function fromXer(text) {
    const T = tables(text), all = T.TASK || [];
    // one file can hold several projects (baselines travel as projects): read the one with the most activities
    const count = {}; for (const t of all) count[t.proj_id] = (count[t.proj_id] || 0) + 1;
    const pid = Object.keys(count).sort((a, b) => count[b] - count[a])[0];
    const proj = (T.PROJECT || []).find(p => p.proj_id === pid) || {};
    const wbs = Object.fromEntries((T.PROJWBS || []).map(w => [w.wbs_id, w]));
    // phase = the WBS level just under the project node
    const phase = id => { let w = wbs[id], last = null; for (let i = 0; w && w.proj_node_flag !== "Y" && i < 50; i++) { last = w; w = wbs[w.parent_wbs_id]; } return last ? last.wbs_name : ""; };
    const hrs = Object.fromEntries((T.CALENDAR || []).map(c => [c.clndr_id, num(c.day_hr_cnt) || 8]));
    const res = Object.fromEntries((T.RSRC || []).map(r => [r.rsrc_id, r.rsrc_name]));
    const who = {}; for (const a of T.TASKRSRC || []) if (res[a.rsrc_id]) (who[a.task_id] ||= []).push(res[a.rsrc_id]);
    const code = Object.fromEntries(all.map(t => [t.task_id, t.task_code]));
    const calOf = Object.fromEntries(all.map(t => [t.task_id, hrs[t.clndr_id] || 8]));
    const preds = {}, links = []; for (const p of T.TASKPRED || []) if (code[p.pred_task_id] && code[p.task_id]) { const lg = num(p.lag_hr_cnt) / (calOf[p.pred_task_id] || 8);
      (preds[p.task_id] ||= []).push(code[p.pred_task_id] + ((PRED[p.pred_type] ?? "") || (lg ? "FS" : "")) + lag(lg)); links.push([code[p.task_id], code[p.pred_task_id]]); }
    const rows = [HEAD]; let skipped = 0;
    for (const t of all) {
      if (t.proj_id !== pid) continue;
      if (t.task_type === "TT_LOE" || t.task_type === "TT_WBS") { skipped++; continue; } // level-of-effort and WBS summary bars are not work
      const plan = num(t.target_drtn_hr_cnt), left = num(t.remain_drtn_hr_cnt), fl = t.status_code === "TK_Complete" ? "" : String(t.total_float_hr_cnt ?? "").trim(); // P6 keeps no float for finished work
      const pct = t.status_code === "TK_Complete" ? 100 : t.status_code === "TK_NotStart" ? 0
        : t.complete_pct_type === "CP_Phys" || !plan ? num(t.phys_complete_pct) : Math.max(0, Math.min(100, Math.round((plan - left) / plan * 100)));
      // P6 keeps planned dates in target_*: with no baseline project in the file they are the baseline
      rows.push([t.task_code, wbs[t.wbs_id]?.wbs_short_name || "", "", t.task_name, phase(t.wbs_id), day(t.act_start_date || t.early_start_date || t.target_start_date),
        day(t.act_end_date || t.early_end_date || t.target_end_date), day(t.target_start_date), day(t.target_end_date), day(t.act_end_date), String(pct),
        days(plan / (hrs[t.clndr_id] || 8)), yes(t.task_type === "TT_Mile" || t.task_type === "TT_FinMile"), "No",
        [...new Set(who[t.task_id] || [])].join(", "), (preds[t.task_id] || []).join(", "), "", fl === "" ? "" : days(num(fl) / (hrs[t.clndr_id] || 8)),
        fl === "" ? "" : yes(num(fl) <= 0), XCSTR[t.cstr_type] || ""]);
    }
    withSuccessors(rows, links);
    return { kind: "xer", title: proj.proj_short_name || "", sd: day(proj.last_recalc_date), rows, skipped };
  }

  // sniff by content, not by file name: .xml is also what Jira and others export
  function read(text) {
    const s = String(text || "").replace(/^﻿/, "");
    if (/^ERMHDR\t/.test(s)) return fromXer(s);
    if (/<Project[\s>]/.test(s.slice(0, 2000)) && s.includes("<Tasks>")) return fromXml(s);
    return null;
  }

  root.KeelsonPlanRead = { read, fromXml, fromXer };
  if (typeof module !== "undefined") module.exports = root.KeelsonPlanRead;
})(typeof window !== "undefined" ? window : globalThis);
