/* Keelson Pro app: upload -> check -> reports. Plain browser JS, no build step.
   Reads plans with KeelsonXlsx / KeelsonEngine, builds ReportDocs with KeelsonReports, renders with KeelsonRender,
   exports Word with KeelsonDocx. The only network request it ever makes is the licence check (Lemon Squeezy). */
(function () {
  "use strict";
  if (window.top !== window.self) { document.body.hidden = true; return; } // never run inside someone else's frame
  const CHECKOUT_URL = ""; // Lemon Squeezy checkout link for Keelson Pro; empty shows "coming soon"
  const LICENCE_API = "https://api.lemonsqueezy.com/v1/licenses/validate";
  // set from Lemon Squeezy > Settings > Store / product IDs; while empty, any valid key from any store is accepted
  const STORE_ID = null, PRODUCT_IDS = [];
  const DAY_MS = 864e5, FRESH = 7, GRACE = 14, SNAP_KEEP = 12, HISTORY = 6;
  const SUPPORT = (m => m && !/\{\{/.test(m) ? m : "us")((document.querySelector('meta[name="keelson-support"]') || {}).content);
  const READ_FAIL = "We couldn't read that workbook. Save it again from Excel as .xlsx, or export it as CSV.";
  const CTRL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
  const clean = s => String(s ?? "").replace(CTRL, ""); // typed names end up in Word's XML, which rejects control characters
  const plainErr = msg => Object.assign(new Error(msg), { plain: true });
  const E = window.KeelsonEngine, X = window.KeelsonXlsx, R = window.KeelsonReports, H = window.KeelsonRender;
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const today = () => Math.floor((Date.now() - new Date().getTimezoneOffset() * 6e4) / DAY_MS);
  const fromIso = s => { const m = /^(\d{4})-(\d\d)-(\d\d)/.exec(s || ""); return m ? E.fromYMD(+m[1], +m[2], +m[3]) : null; };
  const F = v => E.fmt(v, true);
  const plural = (n, w) => `${n} ${w}${n === 1 ? "" : "s"}`;
  const reduce = matchMedia("(prefers-reduced-motion: reduce)");
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} },
  };
  const sess = { // unsaved commentary survives a reload (this tab only)
    get(k, d) { try { const v = sessionStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
  };
  const REPORTS = (typeof R?.list === "function" ? R.list() : R?.list || []).map(r => typeof r === "string" ? { id: r, name: r } : r);
  const missing = ["KeelsonEngine", "KeelsonXlsx", "KeelsonReports", "KeelsonRender"].filter(n => !window[n]);

  // ---------- theme (same switch and storage key as index.html) ----------
  const root = document.documentElement, themeBtn = $("theme");
  const SUN = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><circle cx="10" cy="10" r="3.6"/><path d="M10 1.8v2.1M10 16.1v2.1M1.8 10h2.1M16.1 10h2.1M4.2 4.2l1.5 1.5M14.3 14.3l1.5 1.5M4.2 15.8l1.5-1.5M14.3 5.7l1.5-1.5"/></svg>';
  const MOON = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="M16.5 12.6A7 7 0 017.4 3.5a7 7 0 109.1 9.1z"/></svg>';
  try { const t = localStorage.getItem("keelson-theme"); if (t === "light" || t === "dark") root.dataset.theme = t; } catch (e) {}
  const theme = () => root.dataset.theme || (matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark");
  const paintTheme = () => { const d = theme() === "dark"; themeBtn.innerHTML = d ? SUN : MOON; themeBtn.setAttribute("aria-label", d ? "Switch to light mode" : "Switch to dark mode"); };
  themeBtn.addEventListener("click", () => { const n = theme() === "dark" ? "light" : "dark"; root.dataset.theme = n; try { localStorage.setItem("keelson-theme", n); } catch (e) {} paintTheme(); });
  matchMedia("(prefers-color-scheme: light)").addEventListener?.("change", paintTheme);
  paintTheme();

  let toastT = 0;
  function toast(msg) { const t = $("toast"); t.textContent = msg; t.classList.add("on"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("on"), 3200); }

  // ---------- state ----------
  const brandSaved = store.get("keelson-brand", {});
  const S = {
    file: "", sheets: null, sheet: 0, text: "", cols: {}, msKey: "", title: "", sd: today(), k: null,
    client: clean(brandSaved.client), prep: clean(brandSaved.preparedBy),
    brand: { color: brandSaved.color || "#1a7f86", logo: store.get("keelson-logo", null) },
    edits: {}, compare: true, report: REPORTS[0]?.id, dirty: false,
    rules: null, call: null, msExclude: [], msAll: [], msSupported: null,
  };
  const saveBrand = () => store.set("keelson-brand", { color: S.brand.color, client: S.client, preparedBy: S.prep });

  // ---------- licence (a convenience, not security: the reports are built in this tab either way) ----------
  let lic = store.get("keelson-licence", null);
  // a record that was tampered with (or written by a clock set ahead) is dropped, not trusted
  if (lic && (typeof lic !== "object" || typeof lic.key !== "string" || !Number.isFinite(lic.checkedAt) || lic.checkedAt > Date.now())) { lic = null; store.del("keelson-licence"); }
  const licAge = () => lic ? (Date.now() - lic.checkedAt) / DAY_MS : Infinity;
  const isPro = () => !!(lic && lic.valid && (!lic.expiresAt || Date.parse(lic.expiresAt) > Date.now()) && licAge() <= FRESH + GRACE);
  // The server counts as reached only for an answer it really gave: 2xx, or 400/404, with a JSON boolean `valid`.
  // Anything else (429, a captive portal's HTML page, other errors, a blocked request) is "unreachable", so the grace period applies.
  async function validate(key) {
    let res;
    try {
      res = await fetch(LICENCE_API, { method: "POST", headers: { Accept: "application/json" }, body: new URLSearchParams({ license_key: key }), credentials: "omit", referrerPolicy: "no-referrer", cache: "no-store" });
    } catch (e) { return { reach: false, why: await blocked() ? "blocked" : "offline" }; }
    if (res.status === 429) return { reach: false, why: "rate" };
    let j = null; try { j = await res.json(); } catch (e) {}
    if (!(res.ok || res.status === 400 || res.status === 404) || !j || typeof j !== "object" || typeof j.valid !== "boolean") return { reach: false, why: "offline" };
    const lk = j.license_key || {}, meta = j.meta || {};
    const ours = STORE_ID == null || (meta.store_id === STORE_ID && PRODUCT_IDS.includes(meta.product_id));
    const valid = j.valid && ours;
    return { reach: true, rec: { key, valid, status: !ours ? "another product" : lk.status || (valid ? "active" : "invalid"), expiresAt: lk.expires_at || null, checkedAt: Date.now() }, error: j.error || "", other: !ours };
  }
  // the request failed outright: if the server still answers an opaque no-cors request, it is up and this page's origin
  // is refused (CORS), which is not the user's connection; no key is sent with this probe
  const blocked = () => fetch(LICENCE_API, { method: "HEAD", mode: "no-cors", credentials: "omit", referrerPolicy: "no-referrer", cache: "no-store" }).then(() => true, () => false);
  async function revalidate() { // older than 7 days: check again quietly; keep Pro through 14 days of network trouble
    if (!lic || !lic.key || licAge() <= FRESH) return;
    const r = await validate(lic.key);
    if (r.reach) { lic = r.rec; store.set("keelson-licence", lic); paintPro(); refresh(); }
    else { lic.offline = true; paintPro(); }
  }
  const proBtn = $("pro-btn");
  function paintPro() {
    const pro = isPro();
    proBtn.textContent = pro ? "Pro" : "Unlock Pro";
    proBtn.className = pro ? "btn line sm" : "btn fill sm";
    proBtn.setAttribute("aria-label", pro ? "Keelson Pro is active. Licence details" : "Unlock Keelson Pro");
    const act = $("pro-active");
    act.hidden = !lic || !lic.key;
    if (lic && lic.key) {
      const until = new Date(lic.checkedAt + (FRESH + GRACE) * DAY_MS);
      act.innerHTML = pro
        ? `<b>Pro is active on this device</b><span>Key ending ${esc(lic.key.slice(-4))} · checked ${new Date(lic.checkedAt).toLocaleDateString()}${lic.offline ? ` · we couldn't reach the licence server, so Pro stays on until ${until.toLocaleDateString()}` : ""}</span>`
        : `<b style="color:var(--bad)">This key isn't active</b><span>Status: ${esc(lic.status || "unknown")}${lic.expiresAt ? ` · expired ${new Date(lic.expiresAt).toLocaleDateString()}` : ""}. Renew it, or enter another key.</span>`;
    }
    $("pro-buy").hidden = pro;
    $("key-x").hidden = !lic;
    document.querySelector(".perks").hidden = pro;
    $("pro-h").textContent = pro ? "Keelson Pro" : "Send reports that look like yours";
  }
  const dlg = $("pro");
  const openPro = why => { $("pro-why").textContent = why || ""; $("pro-why").hidden = !why || isPro(); $("key-msg").textContent = ""; $("key-msg").className = "status"; paintPro(); dlg.showModal(); ($("pro-buy").hidden ? $("key") : $("buy")).focus(); };
  proBtn.addEventListener("click", () => openPro());
  $("pro-x").addEventListener("click", () => dlg.close());
  dlg.addEventListener("click", e => { if (e.target === dlg) dlg.close(); });
  if (CHECKOUT_URL) $("buy").href = CHECKOUT_URL;
  else { // no checkout yet: say so up front rather than after a click
    $("buy").setAttribute("aria-disabled", "true"); $("buy").removeAttribute("target"); $("buy").textContent = "Checkout coming soon";
    $("buy").setAttribute("aria-describedby", "buy-note"); $("buy-note").textContent = "Pro goes on sale shortly. Got a licence key already? Enter it below.";
    $("buy").addEventListener("click", e => e.preventDefault());
  }
  async function activate() {
    const key = $("key").value.trim(), msg = $("key-msg");
    if (!key) { msg.className = "status bad"; msg.textContent = "Paste your licence key first."; $("key").focus(); return; }
    msg.className = "status"; msg.textContent = "Checking your key…"; $("key-go").disabled = true;
    const r = await validate(key);
    $("key-go").disabled = false;
    if (!r.reach) {
      msg.className = "status warn";
      msg.textContent = r.why === "rate" ? "Too many checks; try again in a minute."
        : r.why === "blocked" ? `Couldn't verify your key from this browser. Email ${SUPPORT} with your key and we'll sort it out.`
        : "We couldn't reach the licence server. Check your connection and try again; your plan is still here.";
      return;
    }
    if (r.other) { msg.className = "status bad"; msg.textContent = "That key is for a different product. Check it's the Keelson Pro key from your receipt email."; return; }
    if (!r.rec.valid) { msg.className = "status bad"; msg.textContent = r.error && !/^license_key/i.test(r.error) ? `That key didn't work: ${r.error}` : "That key isn't valid. Check it matches the one in your receipt email."; return; }
    lic = r.rec; store.set("keelson-licence", lic);
    $("key").value = ""; msg.className = "status ok"; msg.textContent = "Pro is unlocked. Enjoy.";
    paintPro(); refresh();
    setTimeout(() => dlg.open && dlg.close(), 1100);
    toast("Keelson Pro unlocked");
  }
  $("key-go").addEventListener("click", activate);
  $("key").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); activate(); } });
  $("key-x").addEventListener("click", () => { lic = null; store.del("keelson-licence"); paintPro(); refresh(); $("key-msg").textContent = "Key removed from this device."; $("key").focus(); });
  paintPro();
  revalidate();

  // ---------- step 1: upload ----------
  const SAMPLES = (window.KEELSON_SAMPLES || { plans: [] });
  $("samples").innerHTML = SAMPLES.plans.map((p, i) =>
    `<button class="sample" type="button" role="listitem" data-sample="${i}"><span class="ic" aria-hidden="true">${esc(p.label.split(" ").map(w => w[0]).join("").slice(0, 2))}</span><span><b>${esc(p.label)}</b><span>${esc(p.note)}</span></span><span class="go" aria-hidden="true">›</span></button>`).join("");
  $("samples").addEventListener("click", e => {
    const b = e.target.closest("[data-sample]"); if (!b) return;
    const p = SAMPLES.plans[+b.dataset.sample];
    load({ file: p.label + " sample", text: p.text, title: "FieldFlow Service Cloud Implementation", sd: fromIso(SAMPLES.statusDate), msKey: (p.msKeys || [])[0] || "", sample: true });
  });
  const upErr = msg => { const el = $("up-err"); el.hidden = !msg; el.textContent = msg || ""; };
  const dropT = $("drop-t").textContent, dropS = $("drop-s").innerHTML;
  const busy = on => { $("drop-t").innerHTML = on ? `<span class="busy"><span class="spin" aria-hidden="true"></span>Reading your plan…</span>` : dropT; $("drop-s").innerHTML = on ? "" : dropS; };
  async function readFile(f) {
    upErr("");
    if (!f) return;
    const name = f.name || "plan", ext = (name.match(/\.([^.]+)$/) || [, ""])[1].toLowerCase();
    if (f.size > 50 * 1024 * 1024) return upErr("That file is over 50 MB, too large to open here. Save the plan sheet as CSV and try that.");
    if (ext === "xls" || ext === "mpp") return upErr(`.${ext} files can't be read in the browser. In ${ext === "mpp" ? "MS Project" : "Excel"}, use Save As › ${ext === "mpp" ? "Excel Workbook (.xlsx) or CSV" : "Excel Workbook (.xlsx)"}, then drop that file here.`);
    busy(true);
    try {
      const buf = await f.arrayBuffer(), head = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
      const base = clean(name).replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
      if (head[0] === 0x50 && head[1] === 0x4b) {
        const sheets = (await X.sheets(buf)).filter(s => s.rows.length);
        if (!sheets.length) throw plainErr("This workbook has no data in it. Check you saved the plan, then try again.");
        load({ file: name, sheets, title: base });
      } else if (head[0] === 0xd0 && head[1] === 0xcf) throw plainErr("This looks like an older .xls or a password-protected workbook. Save it as .xlsx (without a password) or CSV, then try again.");
      else load({ file: name, text: new TextDecoder().decode(buf).replace(/^﻿/, ""), title: base });
    } catch (e) { console.error("Keelson: could not read", name, e); upErr(e && e.plain ? e.message : READ_FAIL); }
    finally { busy(false); }
  }
  $("file").addEventListener("change", e => { readFile(e.target.files[0]); e.target.value = ""; });
  const drop = $("drop");
  ["dragenter", "dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); if (t === "drop" || !drop.contains(e.relatedTarget)) drop.classList.remove("over"); }));
  drop.addEventListener("drop", e => readFile(e.dataTransfer.files[0]));
  addEventListener("dragover", e => e.preventDefault()); // a file dropped beside the zone must not navigate away
  addEventListener("drop", e => { e.preventDefault(); if (S.view === "upload" && e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]); });
  $("paste-go").addEventListener("click", () => {
    const t = $("paste").value;
    if (!t.trim()) { upErr("Paste your plan into the box first, headers included."); $("paste").focus(); return; }
    upErr(""); load({ file: "Pasted plan", text: t, title: "" });
  });

  // pick the sheet whose header row matches the most known column names
  const score = rows => { const d = E.detect(rows); return Object.values(d.cols).filter(v => v >= 0).length + (d.cols.name >= 0 && d.cols.finish >= 0 ? 100 : 0); };
  const toText = rows => { const w = Math.max(0, ...rows.slice(0, 2000).map(r => r.length)); return rows.slice(0, 2000).map(r => Array.from({ length: w }, (_, i) => String(r[i] ?? "").replace(/[\t\r\n]+/g, " ")).join("\t")).join("\n"); };
  function load(o) {
    S.file = o.file; S.sheets = o.sheets || null; S.cols = {}; S.edits = {}; S.msKey = o.msKey || ""; S.sample = !!o.sample;
    S.sd = o.sd ?? today();
    if (S.sheets) {
      const sc = S.sheets.map(s => score(s.rows));
      S.sheet = sc.indexOf(Math.max(...sc)); S.scores = sc;
      S.text = toText(S.sheets[S.sheet].rows);
    } else S.text = o.text;
    S.title = clean(o.title);
    S.rules = null; S.call = null; S.msExclude = []; S.msSupported = null;
    $("cols").open = false;
    compute();
    if (!S.title || /^(sheet\d*|plan|export|download|untitled)$/i.test(S.title)) S.title = clean(guessTitle()) || S.title || "Project status";
    loadProject();
    compute();
    S.edits = sess.get("keelson-edits", {})[projKey()] || {}; S.dirty = false;
    paintCheck();
    go("check");
  }
  const guessTitle = () => { // a title row above the headers, e.g. "Brightwater — FieldFlow programme tracker"
    const g = E.parseTable(S.text).slice(0, S.k.hdrRow - 1).map(r => r.filter(Boolean)).filter(r => r.length === 1);
    return g.length ? g[0][0].slice(0, 80) : "";
  };
  // milestone keywords: the reports' standard words plus the user's own (comma-separated)
  const msKeys = () => [...(Array.isArray(R?.DEFAULT_MS_KEYS) ? R.DEFAULT_MS_KEYS : []), ...S.msKey.split(",").map(x => x.trim()).filter(Boolean)];
  const msId = m => String(m.key ?? m.name);
  const callOk = () => !S.call || !!(S.call.reason || "").trim();
  function compute() {
    const base = { statusDate: S.sd, cols: S.cols, msKeys: msKeys(), ...(S.rules ? { rules: { ...S.rules } } : {}) };
    const call = S.call && callOk() ? { rag: S.call.rag, reason: S.call.reason.trim() } : null;
    let k = E.compute(S.text, base);
    S.msAll = Array.isArray(k.tasks) && k.tasks.some(t => t.ms) ? k.tasks.filter(t => t.ms) : k.msList || []; // every milestone (msList may be shortened)
    if (S.msSupported == null && S.msAll.length) { // does this engine honour msExclude? (probe once per plan)
      const t = E.compute(S.text, { ...base, msExclude: [msId(S.msAll[0])] });
      S.msSupported = t.nMs < k.nMs;
    }
    const ex = S.msSupported ? S.msExclude.filter(x => S.msAll.some(m => msId(m) === x)) : [];
    if (call || ex.length) k = E.compute(S.text, { ...base, ...(call ? { ragOverride: call } : {}), ...(ex.length ? { msExclude: ex } : {}) });
    if (call && k.ragCalc === undefined) { k.ragCalc = k.rag; k.rag = call.rag; k.ragReason = call.reason; } // engine without ragOverride yet
    S.k = k;
    S.body = E.parseTable(S.text).slice(S.k.hdrRow).filter(r => r.some(c => c)).length; // data rows under the headers
    return S.k;
  }

  // ---------- step 2: check ----------
  const usable = () => S.k && S.k.rows > 0 && S.k.cols.name >= 0 && S.k.cols.finish >= 0;
  function paintCheck() {
    const k = S.k;
    $("check-sub").innerHTML = `From <b>${esc(S.file)}</b>${S.sheets && S.sheets.length > 1 ? `, sheet <b>${esc(S.sheets[S.sheet].name)}</b>` : ""}. Fix anything that looks wrong, then see your reports.`;
    $("sheet-f").hidden = !(S.sheets && S.sheets.length > 1);
    if (S.sheets) {
      $("sheet").innerHTML = S.sheets.map((s, i) => `<option value="${i}">${esc(s.name)}${s.hidden ? " (hidden)" : ""}</option>`).join("");
      $("sheet").value = String(S.sheet);
      const best = S.scores.indexOf(Math.max(...S.scores)), n = S.scores[best] % 100;
      $("sheet-why").textContent = S.sheet === best && n ? `Picked for you: its headers match ${plural(n, "known column")}.` : "";
    }
    $("title").value = S.title; $("client").value = S.client; $("prep").value = S.prep; $("mskey").value = S.msKey;
    $("sd").value = E.iso(S.sd);
    $("call").value = S.call ? S.call.rag : ""; $("call-why").value = S.call ? S.call.reason : "";
    paintFound();
  }
  // ---------- per-project settings: status rules, your call, milestone keywords and exclusions ----------
  const RULE_LABELS = [["amberVar", "Amber when the finish slips more than (days)"], ["redVar", "Red when the finish slips more than (days)"],
    ["amberOver", "Amber when more tasks than this are overdue"], ["redMs", "Red when this many milestones are overdue"],
    ["amberBehind", "Amber when progress is more than (points) behind plan"], ["redBehind", "Red when progress is more than (points) behind plan"]];
  function loadProject() {
    const p = store.get("keelson-projects", {})[projKey()];
    if (!p || typeof p !== "object") return;
    if (p.rules && typeof p.rules === "object") S.rules = Object.fromEntries(RULE_LABELS.map(([k]) => [k, p.rules[k]]).filter(([, v]) => Number.isFinite(v) && v >= 0));
    if (p.call && /^(Green|Amber|Red)$/.test(p.call.rag)) S.call = { rag: p.call.rag, reason: clean(p.call.reason).slice(0, 160) };
    if (typeof p.msKey === "string") S.msKey = clean(p.msKey);
    if (Array.isArray(p.msExclude)) S.msExclude = p.msExclude.map(String);
  }
  function saveProject() {
    const all = store.get("keelson-projects", {}), v = { rules: S.rules, call: S.call, msKey: S.msKey, msExclude: S.msExclude };
    if (!S.rules && !S.call && !S.msKey && !S.msExclude.length) delete all[projKey()]; else all[projKey()] = v;
    store.set("keelson-projects", all);
  }
  function paintRules() {
    const k = S.k, cur = k.rules || {};
    if (!$("rules").children.length) $("rules").innerHTML = RULE_LABELS.map(([key, label]) =>
      `<label for="ru-${key}">${esc(label)}</label><input class="in" type="number" inputmode="numeric" min="0" step="1" id="ru-${key}" data-r="${key}">`).join("");
    RULE_LABELS.forEach(([key]) => { const el = $("ru-" + key); if (document.activeElement !== el) el.value = cur[key] ?? ""; });
    $("rules-reset").hidden = !S.rules;
    $("rulesum").textContent = S.call ? `Your call: ${S.call.rag}` : S.rules ? "Your rules" : "Standard";
    $("call-why-f").hidden = !S.call;
    $("call-err").textContent = S.call && !callOk() ? "Add a short reason so readers know why." : "";
    $("call-why").setAttribute("aria-invalid", String(!callOk()));
  }
  function paintMs() {
    const all = S.msAll, n = S.k.nMs ?? all.length;
    $("mssum").textContent = `${n} found`;
    $("ms-info").textContent = !all.length ? "None found yet. 0-day tasks, a Milestone column and the words below all count."
      : S.msSupported ? "Untick any that aren't really milestones." : "Found from 0-day tasks, a Milestone column and the words below.";
    $("mslist").innerHTML = all.slice(0, 200).map((m, i) => {
      const id = msId(m), out = S.msSupported && S.msExclude.includes(id), d = m.finish != null ? `<span class="d">${esc(F(m.finish))}</span>` : "";
      return S.msSupported
        ? `<li class="${out ? "x" : ""}"><label><input type="checkbox" data-ms="${esc(id)}" ${out ? "" : "checked"}><span class="nm">${esc(m.name)}</span>${d}</label></li>`
        : `<li><span class="nm">${esc(m.name)}</span>${d}</li>`;
    }).join("");
  }
  function paintFound() {
    const k = S.k, ok = usable();
    const called = S.call && callOk() && k.ragCalc && k.ragCalc !== "—";
    $("summary").innerHTML = ok
      ? `<span class="rag ${esc(k.rag)}">${esc(k.rag)}</span><div><b>${Math.round(k.pct * 100)}% complete${k.finish != null ? `, finishing ${F(k.finish)}` : ""}</b><p>${called ? `Your call (calculated ${esc(k.ragCalc)}): ${esc(k.ragReason || S.call.reason)}` : esc(k.why.join(" · "))}</p></div>`
      : `<span class="rag none">—</span><div><b>${S.body ? "One column needs you" : "We couldn't find a plan in this"}</b><p>${S.body ? `Pick the ${k.cols.name < 0 ? "task name" : "finish or due date"} column below.` : "Check the file has a header row with task names and dates."}</p></div>`;
    $("facts").innerHTML = ok ? [[k.nTask, "task"], [k.nMs, "milestone"], [k.nPhases, "phase"]].map(([n, w]) => `<li><b>${n}</b> ${w}${n === 1 ? "" : "s"}</li>`).join("") +
      `<li>Headers in row <b>${k.hdrRow}</b></li>` + (k.evidence ? `<li>Dates read <b>${k.dayFirst ? "day first" : "month first"}</b></li>` : "") : "";
    const c = k.checks, m = [];
    if (ok) {
      if (c.noFinish) m.push(["warn", `${plural(c.noFinish, "task")} ${c.noFinish === 1 ? "has" : "have"} no finish date`, "They're left off timelines and due lists."]);
      if (c.badDates) m.push(["warn", `${plural(c.badDates, "row")} ${c.badDates === 1 ? "has a date" : "have dates"} we couldn't read`, "Dates like 31/12/2026, 2026-12-31 or 31 Dec 2026 work best."]);
      if (c.order) m.push(["warn", `${plural(c.order, "task")} ${c.order === 1 ? "finishes" : "finish"} before ${c.order === 1 ? "it starts" : "they start"}`, "Worth a look in your plan."]);
      if (c.mismatch) m.push(["warn", `${plural(c.mismatch, "task")} ${c.mismatch === 1 ? "has a status that doesn't match its" : "have a status that doesn't match their"} % complete`, "For example Done but 80%. We go by the status."]);
      if (c.capped) m.push(["warn", "Only the first 1,000 rows were read", "Split very large plans by phase for the best reports."]);
      if (k.cols.bfin < 0) m.push(["info", "No baseline column", "Reports can't show slippage against the original plan. Add a Baseline Finish column if you have one."]);
      if (!k.nPhases) m.push(["info", "No phases found", "Add a Phase column, WBS numbers or outline levels to group the work."]);
      if (!k.nMs) m.push(["info", "No milestones found", "Mark 0-day tasks, add a Milestone column, or type a keyword under Columns found."]);
      if (k.nCancel) m.push(["info", `${plural(k.nCancel, "cancelled task")} left out`, ""]);
      if (!m.some(x => x[0] === "warn")) m.unshift(["ok", `${plural(k.rows, "row")} read cleanly`, ""]);
    } else if (S.body) m.push(["bad", `We couldn't find a ${k.cols.name < 0 ? "task name" : "finish or due date"} column`, "Choose it under Columns found."]);
    $("msgs").innerHTML = m.map(([t, a, b]) => `<li class="msg ${t}"><div>${esc(a)}${b ? `<br><span>${esc(b)}</span>` : ""}</div></li>`).join("");
    colMap(k); paintRules(); paintMs();
    if (!ok && S.body) $("cols").open = true;
    const to = $("to-reports");
    to.setAttribute("aria-disabled", String(!ok || !callOk()));
    $("ready-note").textContent = ok ? `${REPORTS.length} reports ready` : "";
  }
  function colMap(k) {
    const found = E.FIELDS.filter(([key]) => k.cols[key] >= 0).length;
    $("colsum").textContent = k.heads.length ? `${found} of ${E.FIELDS.length}` : "";
    $("colinfo").textContent = k.heads.length ? `Headers found in row ${k.hdrRow}${k.evidence ? `. Dates read ${k.dayFirst ? "day first (31/12)" : "month first (12/31)"}` : ""}. Change any match below.` : "";
    const opts = ['<option value="-1">— not used</option>'].concat(k.heads.map((h, i) => h ? `<option value="${i}">${esc(h)}</option>` : "")).join("");
    $("colmap").innerHTML = E.FIELDS.map(([key, label, req]) =>
      `<label for="cm-${key}" class="${k.cols[key] >= 0 ? "ok" : req ? "miss" : ""}">${req ? "<b>" + esc(label) + "</b>" : esc(label)}</label><select class="in" id="cm-${key}" data-k="${key}">${opts}</select>`).join("");
    $("colmap").querySelectorAll("select").forEach(s => { s.value = String(k.cols[s.dataset.k]); });
  }
  $("colmap").addEventListener("change", e => {
    const s = e.target.closest("select"); if (!s) return;
    S.cols[s.dataset.k] = Number(s.value); compute(); paintFound();
    $("cm-" + s.dataset.k)?.focus();
  });
  $("sheet").addEventListener("change", e => { S.sheet = +e.target.value; S.cols = {}; S.text = toText(S.sheets[S.sheet].rows); compute(); paintCheck(); });
  $("title").addEventListener("input", e => { S.title = clean(e.target.value); });
  $("title").addEventListener("change", () => { // a project name used before brings back its settings and unsaved commentary
    if (store.get("keelson-projects", {})[projKey()]) { S.rules = null; S.call = null; S.msExclude = []; loadProject(); compute(); paintCheck(); }
    else if (S.rules || S.call || S.msKey || S.msExclude.length) saveProject();
    const ed = sess.get("keelson-edits", {})[projKey()];
    if (ed && Object.keys(ed).length) S.edits = ed; else saveEdits();
  });
  const syncPeople = () => { $("client").value = $("client2").value = S.client; $("prep").value = $("prep2").value = S.prep; saveBrand(); };
  $("client").addEventListener("input", e => { S.client = clean(e.target.value); syncPeople(); });
  $("prep").addEventListener("input", e => { S.prep = clean(e.target.value); syncPeople(); });
  $("mskey").addEventListener("input", e => { S.msKey = clean(e.target.value); saveProject(); compute(); paintFound(); });
  $("rules").addEventListener("input", e => {
    const el = e.target.closest("[data-r]"); if (!el) return;
    const v = el.value.trim() === "" ? NaN : Number(el.value);
    if (!Number.isFinite(v) || v < 0) return; // half-typed or negative: keep the last good rule
    S.rules = { ...(S.k.rules || {}), ...(S.rules || {}), [el.dataset.r]: Math.round(v) };
    S.rules = Object.fromEntries(RULE_LABELS.map(([k]) => [k, S.rules[k]]).filter(([, x]) => Number.isFinite(x)));
    saveProject(); compute(); paintFound();
  });
  $("rules-reset").addEventListener("click", () => { S.rules = null; saveProject(); compute(); paintFound(); $("ru-amberVar").focus(); });
  $("call").addEventListener("change", e => {
    S.call = e.target.value ? { rag: e.target.value, reason: S.call ? S.call.reason : "" } : null;
    saveProject(); compute(); paintFound();
    if (S.call && !callOk()) $("call-why").focus();
  });
  $("call-why").addEventListener("input", e => { if (!S.call) return; S.call.reason = clean(e.target.value); saveProject(); compute(); paintFound(); });
  $("mslist").addEventListener("change", e => {
    const c = e.target.closest("[data-ms]"); if (!c) return;
    const id = c.dataset.ms;
    S.msExclude = c.checked ? S.msExclude.filter(x => x !== id) : [...new Set([...S.msExclude, id])];
    saveProject(); compute(); paintFound();
    $("mslist").querySelector(`[data-ms="${CSS.escape(id)}"]`)?.focus();
  });
  $("sd").addEventListener("change", e => { const v = fromIso(e.target.value); if (v != null) { S.sd = v; compute(); paintFound(); } });
  $("to-reports").addEventListener("click", e => {
    if (!usable()) { e.preventDefault(); $("cols").open = true; $("colmap").querySelector(".miss + select")?.focus(); }
    else if (!callOk()) { e.preventDefault(); $("rules-d").open = true; $("call-why").focus(); }
  });

  // ---------- snapshots: "since last report" ----------
  const projKey = () => (S.title || "Project").trim().toLowerCase();
  const snaps = () => (store.get("keelson-snapshots", {})[projKey()] || []);
  const earlier = () => { const cur = E.iso(S.sd); return snaps().filter(s => s && s.asOf && s.asOf < cur).sort((a, b) => a.asOf < b.asOf ? -1 : 1); };
  const prevSnap = () => earlier().at(-1) || null;
  // trend points: the previous snapshot's own history (v2), else the earlier snapshots themselves
  const histOf = p => !p ? [] : (Array.isArray(p.history) ? p.history : earlier()).slice(-HISTORY);
  function saveSnap() {
    if (!R.snapshot) return;
    const all = store.get("keelson-snapshots", {}), s = R.snapshot(S.k, prevSnap()); // v2 snapshots carry the trend (last 6) forward
    const list = (all[projKey()] || []).filter(x => x && x.asOf !== s.asOf);
    list.push(s); list.sort((a, b) => a.asOf < b.asOf ? -1 : 1);
    all[projKey()] = list.slice(-SNAP_KEEP);
    store.set("keelson-snapshots", all);
  }
  function paintCompare() {
    const p = prevSnap();
    document.querySelectorAll("[data-compare]").forEach(el => {
      el.hidden = !p;
      if (!p) return;
      el.innerHTML = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h11l-3-3M16 13H5l3 3"/></svg>` +
        `<span>${S.compare ? "Comparing with" : "Not comparing with"} your report of <b>${F(fromIso(p.asOf))}</b></span>` +
        `<label class="switch"><input type="checkbox" role="switch" ${S.compare ? "checked" : ""} data-cmp> Compare</label>`;
    });
  }
  document.addEventListener("change", e => { if (e.target.matches("[data-cmp]")) { S.compare = e.target.checked; paintCompare(); refresh(); } });

  // ---------- report building ----------
  if (H && !document.getElementById("kr-css")) { const st = document.createElement("style"); st.id = "kr-css"; st.textContent = H.css; document.head.append(st); }
  const opts = id => ({ title: S.title || "Project status", client: S.client, preparedBy: S.prep, brand: { ...S.brand }, prev: S.compare ? prevSnap() : null, history: S.compare ? histOf(prevSnap()) : [], edits: { ...(S.edits[id] || {}) } });
  const doc = id => R.build(id, S.k, opts(id));
  const page = (d, edit, forPrint) => H.html(d, { brand: S.brand, watermark: !isPro() && !forPrint, footer: !isPro(), editable: edit });

  // ---------- step 3: gallery ----------
  const ORIENT = id => (REPORTS.find(r => r.id === id) || {}).orientation;
  function paintGallery() {
    const k = S.k;
    $("reports-sub").innerHTML = `<b>${esc(S.title || "Your project")}</b> as of ${F(S.sd)}. Open one to edit the commentary and download it.`;
    $("gallery").innerHTML = REPORTS.map(r => {
      let d; try { d = doc(r.id); } catch (e) { console.error("Keelson: report", r.id, e); d = null; }
      const land = (d?.orientation || r.orientation) === "landscape";
      const locked = !isPro() && r.id !== "onepager";
      return `<li><a class="rc" href="#report/${esc(r.id)}" aria-label="${esc(r.name)}${r.description ? ". " + esc(r.description) : ""}">
        <div class="thumbbox" aria-hidden="true"><div class="thumb ${land ? "landscape" : "portrait"}"><div class="thumb-in" inert style="width:${land ? 1123 : 794}px">${d ? page(d, false) : ""}</div></div></div>
        <div style="display:grid;gap:6px"><h3>${esc(r.name)}${locked ? '<span class="lock">Pro</span>' : ""}</h3>${r.description ? `<p>${esc(r.description)}</p>` : ""}</div>
        <div class="meta"><span>${land ? "Landscape" : "Portrait"}</span><span>${isPro() ? "PDF · Word" : r.id === "onepager" ? "Free PDF" : "Preview"}</span></div></a></li>`;
    }).join("");
    fitThumbs();
  }
  const ro = new ResizeObserver(es => es.forEach(e => fit(e.target)));
  function fit(t) { const inn = t.firstElementChild; if (!inn) return; const s = t.clientWidth / parseFloat(inn.style.width); inn.style.transform = `scale(${s})`; }
  function fitThumbs() { document.querySelectorAll(".thumb").forEach(t => { fit(t); ro.observe(t); }); }

  // ---------- report view ----------
  function paintSwitch() {
    $("switch").innerHTML = REPORTS.map(r => `<a href="#report/${esc(r.id)}"${r.id === S.report ? ' aria-current="page"' : ""}>${esc(r.name)}</a>`).join("");
  }
  function paintReport(keepFocus) {
    const id = S.report, r = REPORTS.find(x => x.id === id);
    $("h-report").textContent = r ? r.name : "Report";
    paintSwitch();
    let d;
    try { d = doc(id); } catch (e) { console.error("Keelson: report", id, e); $("page").innerHTML = `<div class="err">This report couldn't be built from this plan. Try another report, or check the plan's dates under Check.</div>`; return; }
    $("page").innerHTML = page(d, true);
    fitPage();
    $("reset").hidden = !Object.keys(S.edits[id] || {}).length;
    document.querySelector(".stage-h .pencil").hidden = !$("page").querySelector("[data-kr-edit]");
    const pro = isPro(), pdfOk = pro || id === "onepager";
    $("dl-pdf").innerHTML = pdfOk ? "Download PDF" : 'Download PDF <span class="lock">Pro</span>';
    $("dl-docx").innerHTML = pro ? "Download Word" : 'Download Word <span class="lock">Pro</span>';
    $("dl-note").hidden = pro;
    $("dl-note").textContent = pro ? "" : id === "onepager" ? "Free PDF includes a small Keelson footer. Pro removes the watermark and footer and unlocks Word." : "Free plan: the one-pager downloads as PDF. Pro unlocks every report as PDF and Word.";
    paintBranding();
  }
  // wide screens: show the page at its true width, scaled to fit, so it looks exactly like the PDF; phones reflow instead
  function fitPage() {
    const host = $("page"), pg = host.querySelector(".kr-page"); if (!pg) return;
    const nat = pg.classList.contains("kr-landscape") ? 1123 : 794, sc = host.clientWidth / nat;
    if (innerWidth > 760 && sc < 1 && sc >= .7) {
      Object.assign(pg.style, { width: nat + "px", maxWidth: "none", transform: `scale(${sc})`, transformOrigin: "0 0" });
      const h = Math.ceil(pg.offsetHeight * sc) + "px"; if (host.style.height !== h) host.style.height = h;
    } else { Object.assign(pg.style, { width: "", maxWidth: "", transform: "" }); host.style.height = ""; }
    pageRO.observe(pg);
  }
  const pageRO = new ResizeObserver(() => fitPage());
  addEventListener("resize", () => { if (S.view === "report") fitPage(); });
  $("page").addEventListener("input", e => {
    const el = e.target.closest("[data-kr-edit]"); if (!el) return;
    const id = S.report, t = el.innerText.replace(/\n$/, "");
    (S.edits[id] ||= {})[el.dataset.krEdit] = clean(t);
    S.dirty = true; saveEdits();
    $("reset").hidden = false;
  });
  function saveEdits() { const all = sess.get("keelson-edits", {}); all[projKey()] = S.edits; sess.set("keelson-edits", all); }
  const unsaved = () => S.dirty && Object.values(S.edits).some(o => o && Object.keys(o).length);
  addEventListener("beforeunload", e => { if (unsaved()) { e.preventDefault(); e.returnValue = ""; } });
  $("page").addEventListener("focusin", e => { // a placeholder prompt is selected whole, so typing replaces it
    const el = e.target.closest("[data-kr-edit]");
    if (el && el.closest(".kr-ph") && !(S.edits[S.report] || {})[el.dataset.krEdit]) requestAnimationFrame(() => getSelection().selectAllChildren(el));
  });
  $("page").addEventListener("keydown", e => { if (e.key === "Escape" && e.target.closest("[data-kr-edit]")) e.target.blur(); });
  $("page").addEventListener("paste", e => { // plain text only
    const el = e.target.closest("[data-kr-edit]"); if (!el || el.getAttribute("contenteditable") === "plaintext-only") return;
    e.preventDefault(); document.execCommand("insertText", false, e.clipboardData.getData("text/plain"));
  });
  $("reset").addEventListener("click", () => { delete S.edits[S.report]; saveEdits(); paintReport(); toast("Back to the generated commentary"); $("page").querySelector("[data-kr-edit]")?.focus(); });

  // branding
  const PRESETS = [["#1a7f86", "Teal"], ["#1f3a5f", "Navy"], ["#4338ca", "Indigo"], ["#9e1f63", "Berry"], ["#b42318", "Crimson"], ["#344054", "Graphite"]];
  function paintBranding() {
    const c = S.brand.color.toLowerCase();
    $("swatches").innerHTML = PRESETS.map(([h, n]) => `<button type="button" class="sw" style="background:${h}" data-c="${h}" aria-pressed="${h === c}" aria-label="${n}"></button>`).join("") +
      `<input type="color" class="sw-custom" id="colour" value="${esc(c)}" aria-label="Custom colour"${PRESETS.some(p => p[0] === c) ? "" : ' style="box-shadow:0 0 0 2px var(--bg),0 0 0 4px var(--ink)"'}>`;
    $("client2").value = S.client; $("prep2").value = S.prep;
    const has = !!S.brand.logo;
    $("logo-prev").hidden = !has; $("logo-prev").innerHTML = has ? `<img src="${esc(S.brand.logo)}" alt="Your logo">` : "";
    $("logo-x").hidden = !has; $("logo-l").textContent = has ? "Replace logo" : "Add logo";
  }
  function setColour(h) { S.brand.color = h; saveBrand(); refresh(); }
  $("swatches").addEventListener("click", e => { const b = e.target.closest("[data-c]"); if (b) { setColour(b.dataset.c); $("swatches").querySelector(`[data-c="${b.dataset.c}"]`)?.focus(); } });
  $("swatches").addEventListener("change", e => { if (e.target.id === "colour") { setColour(e.target.value); $("colour")?.focus(); } });
  let peopleT = 0;
  const people = () => { saveBrand(); clearTimeout(peopleT); peopleT = setTimeout(() => { $("page").innerHTML = page(doc(S.report), true); fitPage(); }, 250); };
  $("client2").addEventListener("input", e => { S.client = clean(e.target.value); $("client").value = S.client; people(); });
  $("prep2").addEventListener("input", e => { S.prep = clean(e.target.value); $("prep").value = S.prep; people(); });
  $("wipe").addEventListener("click", () => {
    if (!confirm("Remove everything Keelson saved in this browser? That's your theme, branding and logo, licence key, report history, status rules and unsaved commentary. Your plan files aren't affected.")) return;
    for (const st of [localStorage, sessionStorage]) try { Object.keys(st).filter(k => k.startsWith("keelson-")).forEach(k => st.removeItem(k)); } catch (e) {}
    lic = null; S.brand = { color: "#1a7f86", logo: null }; S.client = S.prep = ""; S.edits = {}; S.dirty = false; S.rules = null; S.call = null; S.msExclude = []; S.msKey = "";
    $("client").value = $("prep").value = "";
    if (S.k) compute();
    paintPro(); refresh(); toast("Saved data cleared from this browser");
  });
  $("logo").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    try { S.brand.logo = await shrinkLogo(f); } catch (err) { toast("That image couldn't be read. Try a PNG or JPEG."); return; }
    if (!store.set("keelson-logo", S.brand.logo)) toast("Logo added. It couldn't be saved for next time (browser storage is full or off).");
    refresh(); $("logo-l").focus();
  });
  $("logo-x").addEventListener("click", () => { S.brand.logo = null; store.del("keelson-logo"); refresh(); $("logo-l").focus(); });
  function shrinkLogo(f) { // any image -> PNG data URL at most 600 x 200, so it stores small and embeds anywhere
    return new Promise((ok, no) => {
      const url = URL.createObjectURL(f), img = new Image();
      img.onload = () => {
        const w = img.naturalWidth || 600, h = img.naturalHeight || 200, s = Math.min(1, 600 / w, 200 / h);
        const c = document.createElement("canvas"); c.width = Math.max(1, Math.round(w * s)); c.height = Math.max(1, Math.round(h * s));
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url); ok(c.toDataURL("image/png"));
      };
      img.onerror = () => { URL.revokeObjectURL(url); no(new Error("bad image")); };
      img.src = url;
    });
  }

  // ---------- downloads ----------
  // a file name every OS accepts: no control or reserved characters, no trailing dot/space, no CON/NUL/COM1..., at most 150 characters
  function safeName(s, ext = "") {
    let n = clean(s).replace(/[\u0000-\u001f\u007f]/g, " ").replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim();
    n = n.slice(0, 150 - ext.length).replace(/[. ]+$/, "");
    if (/^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(\.|$)/i.test(n)) n = "_" + n;
    return (n || "Report") + ext;
  }
  const fileBase = ext => { const r = REPORTS.find(x => x.id === S.report); return safeName([S.title || "Project", r ? r.name : S.report, E.iso(S.sd)].join(" - "), ext); };
  $("dl-pdf").addEventListener("click", () => {
    if (!isPro() && S.report !== "onepager") return openPro("PDFs of this report come with Pro. The one-pager is free.");
    const pr = $("print-root"), t = document.title;
    pr.innerHTML = page(doc(S.report), false, true);
    document.body.classList.add("printing"); document.title = fileBase(); // the browser names the PDF after the title
    // the snapshot is saved once the print dialog closes, not on the click (the user may still cancel)
    addEventListener("afterprint", () => { pr.innerHTML = ""; document.body.classList.remove("printing"); document.title = t; saveSnap(); S.dirty = false; paintCompare(); }, { once: true });
    print();
  });
  $("dl-docx").addEventListener("click", async () => {
    if (!isPro()) return openPro("Editable Word documents come with Pro.");
    const D = window.KeelsonDocx;
    if (!D) return toast("Word export isn't available in this build.");
    const b = $("dl-docx"), label = b.innerHTML;
    b.disabled = true; b.innerHTML = `<span class="spin" aria-hidden="true"></span> Preparing Word…`;
    try {
      if (typeof D.load === "function") await D.load();
      const blob = await D.build(doc(S.report), { brand: { ...S.brand }, footer: !isPro(), rasterize });
      const a = document.createElement("a"), url = URL.createObjectURL(blob);
      a.href = url; a.download = fileBase(".docx"); document.body.append(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 4000);
      saveSnap(); S.dirty = false; paintCompare(); toast("Word document saved");
    } catch (e) { console.error("Keelson: Word export", e); toast("The Word file couldn't be made. Try again, or download the PDF instead."); }
    finally { b.disabled = false; b.innerHTML = label; }
  });
  // svg -> Image -> canvas at 2x -> PNG bytes, for the Word export
  async function rasterize(svg, width, height) {
    if (svg && typeof svg === "object") ({ svg, width, height } = svg);
    const vb = /viewBox="[\d.\s-]*?([\d.]+)\s+([\d.]+)"/.exec(svg) || [];
    const w = +width || +(/\bwidth="([\d.]+)"/.exec(svg) || [])[1] || +vb[1] || 600, h = +height || +(/\bheight="([\d.]+)"/.exec(svg) || [])[1] || +vb[2] || 300;
    const src = /xmlns=/.test(svg) ? svg : svg.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
    const url = URL.createObjectURL(new Blob([src], { type: "image/svg+xml" }));
    try {
      const img = new Image(); img.decoding = "sync";
      await new Promise((ok, no) => { img.onload = ok; img.onerror = () => no(new Error("chart image failed")); img.src = url; });
      const c = document.createElement("canvas"); c.width = Math.round(w * 2); c.height = Math.round(h * 2);
      const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
      const png = await new Promise(ok => c.toBlob(ok, "image/png"));
      return new Uint8Array(await png.arrayBuffer());
    } finally { URL.revokeObjectURL(url); }
  }

  // ---------- views and routing (hash, so Back works) ----------
  const VIEWS = { upload: "v-upload", check: "v-check", reports: "v-reports", report: "v-report" };
  function refresh() {
    if (!S.k) return;
    if (S.view === "reports") { paintCompare(); paintGallery(); }
    if (S.view === "report") { paintCompare(); paintReport(); }
  }
  function go(v) { const h = "#" + v; if (location.hash !== h) location.hash = h; else route(); }
  function route() {
    let [v, id] = location.hash.slice(1).split("/");
    if (!VIEWS[v]) v = "upload";
    if (v !== "upload" && !S.k) { history.replaceState(null, "", "#upload"); v = "upload"; }
    if ((v === "reports" || v === "report") && (!usable() || !callOk())) { history.replaceState(null, "", "#check"); v = "check"; }
    if (v === "report") { S.report = REPORTS.some(r => r.id === id) ? id : REPORTS[0]?.id; }
    const changed = S.view !== v || v === "report", first = S.view == null;
    S.view = v;
    Object.entries(VIEWS).forEach(([k, el]) => { $(el).hidden = k !== v; });
    const order = ["upload", "check", "reports"], at = order.indexOf(v === "report" ? "reports" : v);
    document.querySelectorAll("[data-step]").forEach(a => {
      const i = order.indexOf(a.dataset.step), can = i === 0 || (i === 1 && S.k) || (i === 2 && usable());
      a.toggleAttribute("aria-current", i === at); if (i === at) a.setAttribute("aria-current", "step");
      a.setAttribute("aria-disabled", String(!can)); a.tabIndex = can ? 0 : -1;
      a.classList.toggle("done", i < at);
    });
    if (v === "check") paintCheck();
    refresh();
    if (changed && !first) { // move focus to the new view, never on first load (Tab must start at the skip link)
      const h = $(VIEWS[v]).querySelector("h1, h2");
      scrollTo({ top: 0, behavior: reduce.matches ? "auto" : "instant" });
      if (v === "check" && usable()) $("to-reports").focus({ preventScroll: true }); // Enter goes straight on: ten seconds when detection is right
      else h?.focus({ preventScroll: true });
      document.title = { upload: "Keelson Pro · Stakeholder reports from your plan", check: "Check your plan · Keelson Pro", reports: "Your reports · Keelson Pro", report: `${$("h-report").textContent} · Keelson Pro` }[v];
    }
  }
  addEventListener("hashchange", route);
  route();

  if (missing.length) upErr(`Part of the app didn't load (${missing.join(", ")}). Reload the page; if it keeps happening, tell us at keelson.co.in.`);
  window.KeelsonApp = { S, load, isPro, rasterize, safeName }; // for tests
})();
