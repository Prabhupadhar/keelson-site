/* Keelson Pro app: upload -> check -> reports. Plain browser JS, no build step.
   Reads plans with KeelsonXlsx / KeelsonEngine, builds ReportDocs with KeelsonReports, renders with KeelsonRender,
   exports Word with KeelsonDocx and PowerPoint with KeelsonPptx. v2: version history (KeelsonHistory, IndexedDB), What changed (KeelsonDiff), Slip trend
   (KeelsonTrend) and Finish-date confidence (KeelsonForecast). The only network request it ever makes is the licence check (Lemon Squeezy). */
(function () {
  "use strict";
  if (window.top !== window.self) { document.body.hidden = true; return; } // never run inside someone else's frame
  // store settings live in store-config.js (docs/LEMON-SQUEEZY.md); anything malformed counts as not set
  const CFG = window.KEELSON_STORE || {}, CO = CFG.checkout || {};
  const BUY_RE = /^https:\/\/[a-z0-9-]+\.lemonsqueezy\.com\/(checkout\/)?buy\/[0-9a-f-]{36}(\?.*)?$/;
  const buyUrl = k => typeof CO[k] === "string" && BUY_RE.test(CO[k]) ? CO[k] : "";
  // open Pro: every feature free for everyone through the end of CFG.openUntil (YYYY-MM-DD, local time), then back to normal
  const OPEN_END = (m => m ? new Date(+m[1], +m[2] - 1, +m[3] + 1).getTime() : 0)(/^(\d{4})-(\d\d)-(\d\d)$/.exec(CFG.openUntil || ""));
  const openNow = () => Date.now() < OPEN_END;
  const openLabel = () => new Date(OPEN_END - 1).toLocaleDateString(undefined, { day: "numeric", month: "short" });
  const PORTAL = typeof CFG.portal === "string" && /^https:\/\/([a-z0-9-]+\.)*lemonsqueezy\.com(\/[^\s"<>]*)?$/.test(CFG.portal) ? CFG.portal : "";
  const LS_API = "https://api.lemonsqueezy.com/v1/licenses/";
  // no store set = Pro isn't on sale: no key unlocks it (otherwise a key from any Lemon Squeezy store would)
  const STORE_ID = Number.isSafeInteger(CFG.storeId) && CFG.storeId > 0 ? CFG.storeId : null;
  const PRODUCT_IDS = Array.isArray(CFG.proProductIds) ? CFG.proProductIds.filter(Number.isSafeInteger) : [];
  const SELLING = STORE_ID != null && PRODUCT_IDS.length > 0;
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
  const KH = window.KeelsonHistory, KD = window.KeelsonDiff, KT = window.KeelsonTrend, KF = window.KeelsonForecast;
  // v2 reports: built from this plan's saved versions, not from k alone
  const V2 = [
    { id: "changes", name: "What changed", orientation: "portrait", mod: KD,
      description: "Every task added, removed, moved or re-baselined since the last report, with the one sentence your sponsor needs." },
    { id: "trend", name: "Slip trend", orientation: "landscape", mod: KT,
      description: "How milestone dates moved report by report, the phases that keep slipping, and every re-baseline on record." },
    { id: "forecast", name: "Finish-date confidence", orientation: "portrait", mod: KF,
      description: "P50 and P80 finish dates from this plan's own track record, explained in one sentence." },
  ].filter(r => r.mod && KH);
  const ALL = [...REPORTS, ...V2];
  const isV2 = id => V2.some(r => r.id === id);
  const FREE_RULE = `Preview all ${ALL.length} reports · one-pager as PDF free`; // the one free rule, same wording on the website
  const nWord = n => ["Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"][n] || String(n);
  const missing = ["KeelsonEngine", "KeelsonXlsx", "KeelsonReports", "KeelsonRender", "KeelsonHistory", "KeelsonDiff", "KeelsonTrend", "KeelsonForecast"].filter(n => !window[n]);

  // ---------- theme (same switch and storage key as index.html) ----------
  const root = document.documentElement, themeBtn = $("theme");
  const icon = (n, cls = "ic") => `<svg class="${cls}" aria-hidden="true"><use href="#i-${n}"/></svg>`;
  try { const t = localStorage.getItem("keelson-theme"); if (t === "light" || t === "dark") root.dataset.theme = t; } catch (e) {}
  const theme = () => root.dataset.theme || (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
  const paintTheme = () => { const d = theme() === "dark"; themeBtn.innerHTML = icon(d ? "sun" : "moon"); themeBtn.setAttribute("aria-label", d ? "Switch to light mode" : "Switch to dark mode"); };
  themeBtn.addEventListener("click", () => { const n = theme() === "dark" ? "light" : "dark"; root.dataset.theme = n; try { localStorage.setItem("keelson-theme", n); } catch (e) {} paintTheme(); });
  matchMedia("(prefers-color-scheme: dark)").addEventListener?.("change", paintTheme);
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
    hist: [], histMem: false, memSnaps: [], cmp: null, diff: { accepted: [], split: [], exclude: [], filter: "all" },
  };
  const saveBrand = () => store.set("keelson-brand", { color: S.brand.color, client: S.client, preparedBy: S.prep });

  // ---------- working week and holidays (KeelsonEngine.setCalendar; hidden when this engine has none) ----------
  const CAL_OK = typeof E?.setCalendar === "function" && !!E.WEEKENDS;
  const calOf = c => ({ week: CAL_OK && c && Object.hasOwn(E.WEEKENDS, c.week) ? c.week : "sat-sun", hol: clean(c && typeof c.hol === "string" ? c.hol : "").slice(0, 5000) });
  const calDefault = () => calOf(store.get("keelson-calendar", null)); // the last calendar used, for a new project
  S.cal = calDefault(); S.calMine = false; // calMine: this project has its own calendar saved
  // set the engine's calendar from S.cal; holidays are typed one per line, read like the plan's dates
  function applyCal() {
    if (!CAL_OK) return null;
    const lines = S.cal.hol.split(/\r?\n/).map(x => x.trim()).filter(Boolean), df = !!(S.k && S.k.dayFirst);
    const hol = lines.map(l => E.parseDate(l, df)).filter(v => v != null).map(E.iso);
    try { const c = E.setCalendar({ weekend: E.WEEKENDS[S.cal.week], holidays: hol }); return { n: c.holidays.length, ignored: lines.length - c.holidays.length }; }
    catch (e) { console.error("Keelson: calendar", e); return null; }
  }
  const comp = (text, o) => { applyCal(); return E.compute(text, o); };

  // ---------- licence (a convenience, not security: the reports are built in this tab either way) ----------
  let lic = store.get("keelson-licence", null);
  // a record that was tampered with (or written by a clock set ahead) is dropped, not trusted
  if (lic && (typeof lic !== "object" || typeof lic.key !== "string" || !Number.isFinite(lic.checkedAt) || lic.checkedAt > Date.now() || (lic.instanceId != null && typeof lic.instanceId !== "string"))) { lic = null; store.del("keelson-licence"); }
  const licAge = () => lic ? (Date.now() - lic.checkedAt) / DAY_MS : Infinity;
  const isPro = () => openNow() || !!(SELLING && lic && lic.valid && (!lic.expiresAt || Date.parse(lic.expiresAt) > Date.now()) && licAge() <= FRESH + GRACE);
  // The server counts as reached only for an answer it really gave: 2xx, or 400/404/422, with the JSON boolean the call
  // promises (valid / activated / deactivated). Anything else (429, a captive portal's HTML page, other errors, a blocked
  // request) is "unreachable", so the grace period applies.
  const FLAG = { validate: "valid", activate: "activated", deactivate: "deactivated" };
  async function ls(op, params) {
    let res;
    try {
      res = await fetch(LS_API + op, { method: "POST", headers: { Accept: "application/json" }, body: new URLSearchParams(params), credentials: "omit", referrerPolicy: "no-referrer", cache: "no-store" });
    } catch (e) { return { reach: false, why: await blocked() ? "blocked" : "offline" }; }
    if (res.status === 429) return { reach: false, why: "rate" };
    let j = null; try { j = await res.json(); } catch (e) {}
    if (!(res.ok || [400, 404, 422].includes(res.status)) || !j || typeof j !== "object" || typeof j[FLAG[op]] !== "boolean") return { reach: false, why: "offline" };
    const meta = j.meta || {}, test = !!(j.license_key && j.license_key.test_mode); // a test-mode key (free with the test card) only unlocks the test page
    const mine = SELLING && meta.store_id === STORE_ID && PRODUCT_IDS.includes(meta.product_id) && (!test || CFG.testMode === true);
    return { reach: true, j, ok: j[FLAG[op]] && mine, mine, other: !mine && meta.store_id != null, instance: (j.instance && typeof j.instance.id === "string" && j.instance.id) || "", error: String(j.error || "") };
  }
  const record = (key, r, instanceId) => {
    const lk = r.j.license_key || {};
    return { key, instanceId: instanceId || null, valid: r.ok, status: r.other ? "another product" : lk.status || (r.ok ? "active" : "invalid"), expiresAt: lk.expires_at || null, checkedAt: Date.now() };
  };
  const instanceName = () => { const u = navigator.userAgent; return `Keelson · ${/Edg\//.test(u) ? "Edge" : /OPR\//.test(u) ? "Opera" : /Firefox\//.test(u) ? "Firefox" : /Chrom(e|ium)\//.test(u) ? "Chrome" : /Safari\//.test(u) ? "Safari" : "Browser"} · ${new Date().toISOString().slice(0, 10)}`; };
  const release = (key, id) => { if (key && id) ls("deactivate", { license_key: key, instance_id: id }); }; // best effort: frees the activation slot
  const LIMIT = /activation limit/i;
  // the request failed outright: if the server still answers an opaque no-cors request, it is up and this page's origin
  // is refused (CORS), which is not the user's connection; no key is sent with this probe
  const blocked = () => fetch(LS_API + "validate", { method: "HEAD", mode: "no-cors", credentials: "omit", referrerPolicy: "no-referrer", cache: "no-store" }).then(() => true, () => false);
  async function revalidate() { // older than 7 days: check again quietly; keep Pro through 14 days of network trouble
    if (!SELLING || !lic || !lic.key || licAge() <= FRESH) return;
    const key = lic.key;
    let inst = lic.instanceId, tried = !!lic.instTried;
    if (!inst && !tried) { // a key saved before per-browser activations: claim a slot once; at the limit it still works without one
      const a = await ls("activate", { license_key: key, instance_name: instanceName() });
      if (a.reach) { tried = true; if (a.ok && a.instance) inst = a.instance; else if (a.instance) release(key, a.instance); }
    }
    const r = await ls("validate", inst ? { license_key: key, instance_id: inst } : { license_key: key });
    if (!lic || lic.key !== key) return; // removed or replaced meanwhile
    if (r.reach) { lic = { ...record(key, r, inst), ...(tried ? { instTried: true } : {}) }; paintPro(); refresh(); }
    else { Object.assign(lic, { offline: true, instanceId: inst || null }, tried ? { instTried: true } : {}); paintPro(); }
    store.set("keelson-licence", { ...lic, offline: undefined });
  }
  const proBtn = $("pro-btn");
  function paintPro() {
    const pro = isPro();
    proBtn.textContent = openNow() ? `Pro free until ${openLabel()}` : pro ? "Pro" : "Unlock Pro";
    proBtn.className = pro ? "btn line sm" : "btn fill sm";
    proBtn.setAttribute("aria-label", pro ? "Keelson Pro is active. Licence details" : "Unlock Keelson Pro");
    const act = $("pro-active");
    act.hidden = !lic || !lic.key;
    if (lic && lic.key) {
      const until = new Date(lic.checkedAt + (FRESH + GRACE) * DAY_MS);
      act.innerHTML = pro
        ? `<b>Pro is active on this device</b><span>Key ending ${esc(lic.key.slice(-4))} · checked ${new Date(lic.checkedAt).toLocaleDateString()}${lic.offline ? ` · we couldn't reach the licence server, so Pro stays on until ${until.toLocaleDateString()}` : ""}</span>`
        : `<b style="color:var(--bad)">This key isn't active</b><span>Status: ${esc(SELLING ? lic.status || "unknown" : "Pro isn't on sale")}${lic.expiresAt ? ` · expired ${new Date(lic.expiresAt).toLocaleDateString()}` : ""}. ${PORTAL ? `<a href="${esc(PORTAL)}" target="_blank" rel="noopener">Renew</a> it` : "Renew it"}, or enter another key.</span>`;
    }
    $("pro-buy").hidden = pro;
    $("key-x").hidden = !lic;
    $("portal").hidden = !(lic && lic.key && PORTAL);
    document.querySelector(".perks").hidden = pro;
    $("pro-h").textContent = openNow() ? `Pro is free for everyone until ${openLabel()}` : pro ? "Keelson Pro" : proCtx || "Send reports that look like yours";
    $("pro-thumb").hidden = pro || !proCtx;
  }
  const dlg = $("pro");
  let proCtx = ""; // "Your <project> <report> is ready" when a locked download opened the dialog
  // a locked download: say whose report is waiting and show it (the watermarked page they were viewing, inert, scaled down)
  const lockedPro = () => {
    const r = ALL.find(x => x.id === S.report);
    proCtx = `Your ${S.title || "project"} ${r ? r.name : "report"} is ready`;
    let d = null; try { d = doc(S.report); } catch (e) {}
    const land = d && d.orientation === "landscape";
    $("pro-thumb").innerHTML = d ? `<div class="thumb ${land ? "landscape" : "portrait"}"><div class="thumb-in" inert style="width:${land ? 1123 : 794}px">${page(d, false)}</div></div>` : "";
    openPro(`Download it without the watermark, plus Word and PowerPoint export and all ${nWord(ALL.length).toLowerCase()} reports — $12 a month or $96 a year.`, true);
    const t = $("pro-thumb").firstElementChild; if (t) fit(t);
  };
  const openPro = (why, ctx) => { if (!ctx) proCtx = ""; $("pro-why").textContent = why || ""; $("pro-why").hidden = !why || isPro(); $("key-msg").textContent = ""; $("key-msg").className = "status"; paintPro(); dlg.showModal(); ($("pro-buy").hidden ? $("key") : $("buy")).focus(); };
  proBtn.addEventListener("click", () => openPro());
  $("pro-x").addEventListener("click", () => dlg.close());
  dlg.addEventListener("click", e => { if (e.target === dlg) dlg.close(); });
  if (PORTAL) $("portal").href = PORTAL;
  const plans = [["proMonthly", "Monthly · $12"], ["proYearly", "Yearly · $96 (save 33%)"]].filter(([k]) => buyUrl(k));
  if (plans.length) plans.forEach(([k, label], i) => { const a = $(i ? "buy-y" : "buy"); a.href = buyUrl(k); a.textContent = label; a.hidden = false; $("buy-note").hidden = true; });
  else { // no checkout yet: a waitlist email, said up front; purchase terms only make sense once it is on sale
    $("buy-terms").hidden = true;
    $("buy").href = `mailto:${/@/.test(SUPPORT) ? SUPPORT : "support@keelson.co.in"}?subject=${encodeURIComponent("Keelson Pro waitlist")}`; $("buy").removeAttribute("target");
    $("buy").setAttribute("aria-describedby", "buy-note"); $("buy-note").textContent = "Pro isn't on sale yet. Join the waitlist and we'll email you when it is.";
  }
  async function activate() {
    const key = $("key").value.trim(), msg = $("key-msg");
    if (!SELLING) { msg.className = "status warn"; msg.textContent = "Pro isn't on sale yet, so there are no licence keys to enter. Join the waitlist and we'll email you when it is."; return; }
    if (!key) { msg.className = "status bad"; msg.textContent = "Paste your licence key first."; $("key").focus(); return; }
    msg.className = "status"; msg.textContent = "Checking your key…"; $("key-go").disabled = true;
    const r = await ls("activate", { license_key: key, instance_name: instanceName() });
    $("key-go").disabled = false;
    if (!r.reach) {
      msg.className = "status warn";
      msg.textContent = r.why === "rate" ? "Too many checks; try again in a minute."
        : r.why === "blocked" ? `Couldn't verify your key from this browser. Email ${SUPPORT} with your key and we'll sort it out.`
        : "We couldn't reach the licence server. Check your connection and try again; your plan is still here.";
      return;
    }
    if (r.other) { release(key, r.instance); msg.className = "status bad"; msg.textContent = "That key is for a different product. Check it's the Keelson Pro key from your receipt email."; return; }
    if (!r.ok) {
      const st = (r.j.license_key || {}).status;
      msg.className = "status bad";
      if (LIMIT.test(r.error)) { msg.textContent = `This key is already active on its maximum number of browsers. Remove it from one of them (Unlock Pro › Remove the key), or email ${SUPPORT}.`; return; }
      if (r.mine && (st === "expired" || st === "disabled" || /expired|disabled/i.test(r.error))) { // keep it so the dialog shows the renew link
        lic = record(key, r); if (!/expired|disabled/.test(lic.status)) lic.status = /expired/i.test(r.error) ? "expired" : "disabled";
        store.set("keelson-licence", lic); $("key").value = ""; msg.textContent = `That key is ${lic.status}.`; paintPro(); return;
      }
      msg.textContent = r.error && !/^license_key/i.test(r.error) ? `That key didn't work: ${r.error}` : "That key isn't valid. Check it matches the one in your receipt email.";
      return;
    }
    if (lic && lic.key && lic.instanceId && lic.instanceId !== r.instance) release(lic.key, lic.instanceId); // replacing a key frees the old slot
    lic = record(key, r, r.instance); store.set("keelson-licence", lic);
    $("key").value = ""; msg.className = "status ok"; msg.textContent = "Pro is unlocked.";
    paintPro(); refresh();
    setTimeout(() => dlg.open && dlg.close(), 1100);
    toast("Keelson Pro unlocked");
  }
  $("key-go").addEventListener("click", activate);
  $("key").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); activate(); } });
  $("key-x").addEventListener("click", () => { if (lic) release(lic.key, lic.instanceId); lic = null; store.del("keelson-licence"); paintPro(); refresh(); $("key-msg").textContent = "Key removed from this device."; $("key").focus(); });
  paintPro();
  revalidate();

  // ---------- step 1: upload ----------
  const SAMPLES = (window.KEELSON_SAMPLES || { plans: [] }), SH = window.KEELSON_SAMPLE_HISTORY;
  const sampleBtn = (id, label, note) => `<button class="sample" type="button" role="listitem" data-sample="${id}"><span><b>${esc(label)}</b><span class="n">${esc(note)}</span></span>${icon("arrow", "ic sm go")}</button>`;
  $("samples").innerHTML = (SH && KH ? sampleBtn("history", SH.label, SH.note) : "") + SAMPLES.plans.map((p, i) => sampleBtn(i, p.label, p.note)).join("");
  $("samples").addEventListener("click", e => {
    const b = e.target.closest("[data-sample]"); if (!b) return;
    upErr("");
    try { openSample(b); } catch (err) { console.error(err); upErr("Couldn't open the sample. Reload the page (Ctrl+F5, or Cmd+Shift+R on a Mac) and try again."); }
  });
  function openSample(b) {
    if (b.dataset.sample === "history") { // the current plan plus five earlier weekly versions, kept in memory
      const cur = SH.versions[SH.versions.length - 1];
      return load({ file: cur.file, text: cur.text, title: SH.title, sd: fromIso(cur.asOf), msKey: (SH.msKeys || [])[0] || "", sample: true, history: SH.versions.slice(0, -1) });
    }
    const p = SAMPLES.plans[+b.dataset.sample];
    load({ file: p.label + " sample", text: p.text, title: p.id === "messy" ? "FieldFlow messy tracker (sample)" : "FieldFlow Service Cloud Implementation", sd: fromIso(SAMPLES.statusDate), msKey: (p.msKeys || [])[0] || "", sample: true });
  }
  $("try-sample").addEventListener("click", () => { const b = $("samples").querySelector("[data-sample]"); // the first sample: the one with history
    if (b) b.click(); else upErr("The sample plans didn't load. Reload the page (Ctrl+F5, or Cmd+Shift+R on a Mac) and try again."); });
  const upErr = msg => { const el = $("up-err"); el.hidden = !msg; el.textContent = msg || ""; };
  const dropT = $("drop-t").textContent, dropS = $("drop-s").innerHTML;
  const busy = on => { $("drop-t").innerHTML = on ? `<span class="busy"><span class="spin" aria-hidden="true"></span>Reading your plan…</span>` : dropT; $("drop-s").innerHTML = on ? "" : dropS; };
  // any plan file -> {file, title, sheets} (xlsx) or {file, title, text}; plain errors say what to do
  async function parsePlan(f) {
    const name = f.name || "plan", ext = (name.match(/\.([^.]+)$/) || [, ""])[1].toLowerCase();
    if (f.size > 50 * 1024 * 1024) throw plainErr("That file is over 50 MB, too large to open here. Save the plan sheet as CSV and try that.");
    if (ext === "xls" || ext === "mpp") throw plainErr(`.${ext} files can't be read in the browser. In ${ext === "mpp" ? "MS Project" : "Excel"}, use Save As › ${ext === "mpp" ? "Excel Workbook (.xlsx) or CSV" : "Excel Workbook (.xlsx)"}, then drop that file here.`);
    const buf = await f.arrayBuffer(), head = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
    const title = clean(name).replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
    if (head[0] === 0x50 && head[1] === 0x4b) {
      const sheets = (await X.sheets(buf)).filter(s => s.rows.length);
      if (!sheets.length) throw plainErr("This workbook has no data in it. Check you saved the plan, then try again.");
      return { file: name, sheets, title };
    }
    if (head[0] === 0xd0 && head[1] === 0xcf) throw plainErr("This looks like an older .xls or a password-protected workbook. Save it as .xlsx (without a password) or CSV, then try again.");
    return { file: name, text: new TextDecoder().decode(buf).replace(/^\ufeff/, ""), title };
  }
  // the plan text of a parsed file: its best sheet for a workbook
  const planText = p => { if (!p.sheets) return { text: p.text, sheet: "" }; const sc = p.sheets.map(s => score(s.rows)), i = bestSheet(p.sheets, sc); return { text: toText(p.sheets[i].rows), sheet: p.sheets[i].name }; };
  async function readFile(f) {
    upErr("");
    if (!f) return;
    busy(true);
    let p;
    try { p = await parsePlan(f); }
    catch (e) { console.error("Keelson: could not read", f.name, e); upErr(e && e.plain ? e.message : READ_FAIL); busy(false); return; }
    try { load(p); } // the file was read; a failure here is ours, so say so and keep the error text for support
    catch (e) { console.error("Keelson: could not open", f.name, e); upErr(`The file was read, but the app hit an error opening it (${String(e && e.message || e).slice(0, 120)}). Reload the page (Ctrl+F5, or Cmd+Shift+R on a Mac) and try again; if it happens again, email ${SUPPORT} with this message.`); }
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
  // on a tie, a visible sheet over a hidden one, and the current plan over "Previous week", "Old", "Archive" or "Copy of"
  const STALE = /\b(old|previous|prev|last week|archive[ds]?|backup|copy|superseded)\b/i;
  const bestSheet = (sheets, sc) => { const top = sc.map((v, i) => i).filter(i => sc[i] === Math.max(...sc)); return top.find(i => !sheets[i].hidden && !STALE.test(sheets[i].name)) ?? top.find(i => !sheets[i].hidden) ?? top[0]; };
  const score = rows => { const d = E.detect(rows); return Object.values(d.cols).filter(v => v >= 0).length + (d.cols.name >= 0 && d.cols.finish >= 0 ? 100 : 0); };
  const toText = rows => { const w = Math.max(0, ...rows.slice(0, 2000).map(r => r.length)); return rows.slice(0, 2000).map(r => Array.from({ length: w }, (_, i) => String(r[i] ?? "").replace(/[\t\r\n]+/g, " ")).join("\t")).join("\n"); };
  function load(o) {
    S.file = o.file; S.sheets = o.sheets || null; S.cols = {}; S.edits = {}; S.msKey = o.msKey || ""; S.sample = !!o.sample;
    S.cmp = null; S.diff = { accepted: [], split: [], exclude: [], filter: "all" }; S.hist = []; S.memSnaps = [];
    S.sd = o.sd ?? today();
    if (S.sheets) {
      const sc = S.sheets.map(s => score(s.rows));
      S.sheet = bestSheet(S.sheets, sc); S.scores = sc;
      S.text = toText(S.sheets[S.sheet].rows);
    } else S.text = o.text;
    S.title = clean(o.title);
    S.rules = null; S.call = null; S.msExclude = []; S.msSupported = null; S.cal = calDefault(); S.calMine = false;
    $("cols").open = false;
    compute();
    if (!S.title || /^(sheet\d*|plan|export|download|untitled)$/i.test(S.title)) S.title = clean(guessTitle()) || S.title || "Project status";
    loadProject();
    compute();
    S.edits = sess.get("keelson-edits", {})[projKey()] || {}; S.dirty = false; withNotes();
    // a sample's earlier versions live in memory until a report is downloaded; a real plan uses this browser's saved history
    if (KH) { if (S.histMem) { KH.use(null); S.histMem = false; } if (o.history) memHistory(o.history).then(loadHist); else loadHist(); }
    paintCheck();
    go(confident() ? "reports" : "check"); // nothing to fix: straight to the reports ("Columns detected · Change" leads back)
  }
  const guessTitle = () => { // a title row above the headers, e.g. "Brightwater — FieldFlow programme tracker"
    const g = E.parseTable(S.text).slice(0, S.k.hdrRow - 1).map(r => r.filter(Boolean)).filter(r => r.length === 1);
    return g.length ? g[0][0].slice(0, 80) : "";
  };
  // milestone keywords: the reports' standard words plus the user's own (comma-separated)
  const msKeys = () => [...(Array.isArray(R?.DEFAULT_MS_KEYS) ? R.DEFAULT_MS_KEYS : []), ...S.msKey.split(",").map(x => x.trim()).filter(Boolean)];
  const msId = m => String(m.key ?? m.name);
  if (Array.isArray(R?.DEFAULT_MS_KEYS) && R.DEFAULT_MS_KEYS.length) $("mskey-def").textContent = `Already counted: names containing ${R.DEFAULT_MS_KEYS.slice(0, -1).join(", ")} or ${R.DEFAULT_MS_KEYS.at(-1)}.`;
  const callOk = () => !S.call || !!(S.call.reason || "").trim();
  function compute() {
    const base = { statusDate: S.sd, cols: S.cols, msKeys: msKeys(), ...(S.rules ? { rules: { ...S.rules } } : {}) };
    const call = S.call && callOk() ? { rag: S.call.rag, reason: S.call.reason.trim() } : null;
    let k = comp(S.text, base);
    S.msAll = Array.isArray(k.tasks) && k.tasks.some(t => t.ms) ? k.tasks.filter(t => t.ms) : k.msList || []; // every milestone (msList may be shortened)
    if (S.msSupported == null && S.msAll.length) { // does this engine honour msExclude? (probe once per plan)
      const t = comp(S.text, { ...base, msExclude: [msId(S.msAll[0])] });
      S.msSupported = t.nMs < k.nMs;
    }
    const ex = S.msSupported ? S.msExclude.filter(x => S.msAll.some(m => msId(m) === x)) : [];
    if (call || ex.length) k = comp(S.text, { ...base, ...(call ? { ragOverride: call } : {}), ...(ex.length ? { msExclude: ex } : {}) });
    if (call && k.ragCalc === undefined) { k.ragCalc = k.rag; k.rag = call.rag; k.ragReason = call.reason; } // engine without ragOverride yet
    S.k = k;
    S.body = E.parseTable(S.text).slice(S.k.hdrRow).filter(r => r.some(c => c)).length; // data rows under the headers
    // no date settles day-first vs month-first: ambiguous when reading them the other way gives different dates
    const dates = x => JSON.stringify((x.tasks || []).map(t => [t.start, t.finish, t.bfin]));
    S.ambig = !k.evidence && usable() && dates(E.compute(S.text, { ...base, dateOrder: "day" })) !== dates(k);
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
      const best = bestSheet(S.sheets, S.scores), n = S.scores[best] % 100;
      $("sheet-why").textContent = S.sheet === best && n ? `Picked for you: its headers match ${plural(n, "known column")}.` : "";
    }
    $("title").value = S.title; $("client").value = S.client; $("prep").value = S.prep; $("mskey").value = S.msKey;
    $("sd").value = E.iso(S.sd);
    $("call").value = S.call ? S.call.rag : ""; $("call-why").value = S.call ? S.call.reason : "";
    $("cal-week").value = S.cal.week; if (document.activeElement !== $("cal-hol")) $("cal-hol").value = S.cal.hol;
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
    if (p.cal && typeof p.cal === "object") { S.cal = calOf(p.cal); S.calMine = true; }
  }
  function saveProject() {
    const all = store.get("keelson-projects", {}), v = { rules: S.rules, call: S.call, msKey: S.msKey, msExclude: S.msExclude, ...(S.calMine ? { cal: S.cal } : {}) };
    if (!S.rules && !S.call && !S.msKey && !S.msExclude.length && !S.calMine) delete all[projKey()]; else all[projKey()] = v;
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
    $("ms-info").textContent = !all.length ? "None found yet. 0-day tasks, a Milestone column and the words listed below all count."
      : S.msSupported ? "Untick any that aren't really milestones." : "Found from 0-day tasks, a Milestone column and the words below.";
    $("mslist").innerHTML = all.slice(0, 200).map((m, i) => {
      const id = msId(m), out = S.msSupported && S.msExclude.includes(id), d = m.finish != null ? `<span class="d">${esc(F(m.finish))}</span>` : "";
      return S.msSupported
        ? `<li class="${out ? "x" : ""}"><label><input type="checkbox" data-ms="${esc(id)}" ${out ? "" : "checked"}><span class="nm">${esc(m.name)}</span>${d}</label></li>`
        : `<li><span class="nm">${esc(m.name)}</span>${d}</li>`;
    }).join("");
  }
  // what the Check step flags: [kind, text, help]; any "warn" or "bad" means the read needs a look before the reports
  function checkMsgs() {
    const k = S.k, c = k.checks, m = [];
    if (usable()) {
      const unread = c.capped ? 0 : S.body - k.rows - (k.nCancel || 0);
      if (S.sheets && S.scores && S.scores.filter(x => x === S.scores[S.sheet]).length > 1) m.push(["warn", `More than one sheet looks like a plan. We used ${S.sheets[S.sheet].name}`, "Pick another under Sheet if that's wrong."]);
      if (unread > 0) m.push(["warn", `${plural(unread, "row")} ${unread === 1 ? "has" : "have"} no task name, so ${unread === 1 ? "it was" : "they were"} left out`, "Check the Task name column under Columns found."]);
      if (S.ambig) m.push(["warn", `Dates could be day first or month first. We read them ${k.dayFirst ? "day first (31/12)" : "month first (12/31)"}`, "Check one date you know in the reports."]);
      if (c.noFinish) m.push(["warn", `${plural(c.noFinish, "task")} ${c.noFinish === 1 ? "has" : "have"} no finish date`, "They're left off timelines and due lists."]);
      if (c.badDates) m.push(["warn", `${plural(c.badDates, "row")} ${c.badDates === 1 ? "has a date" : "have dates"} we couldn't read`, "Dates like 31/12/2026, 2026-12-31 or 31 Dec 2026 work best."]);
      if (c.order) m.push(["warn", `${plural(c.order, "task")} ${c.order === 1 ? "finishes" : "finish"} before ${c.order === 1 ? "it starts" : "they start"}`, "Worth a look in your plan."]);
      if (c.mismatch) m.push(["warn", `${plural(c.mismatch, "task")} ${c.mismatch === 1 ? "has a status that doesn't match its" : "have a status that doesn't match their"} % complete`, "For example Done but 80%. We go by the status."]);
      if (c.capped) m.push(["warn", "Only the first 1,000 rows were read", "Split very large plans by phase for the best reports."]);
      if (k.cols.bfin < 0) m.push(["info", "No baseline column", "Reports can't show slippage against the original plan. Add a Baseline Finish column if you have one."]);
      if (!k.nPhases) m.push(["info", "No phases found", "Add a Phase column, WBS numbers or outline levels to group the work."]);
      if (!k.nMs) m.push(["info", "No milestones found", "Mark 0-day tasks, add a Milestone column, or type a keyword under Milestones."]);
      if (k.nCancel) m.push(["info", `${plural(k.nCancel, "cancelled task")} left out`, ""]);
      if (!m.some(x => x[0] === "warn")) m.unshift(["ok", `${plural(k.rows, "row")} read cleanly`, ""]);
    } else if (S.body) m.push(["bad", `We couldn't find a ${k.cols.name < 0 ? "task name" : "finish or due date"} column`, "Choose it under Columns found."]);
    return m;
  }
  // a confident read: required columns found, nothing warned, milestones found, no status call waiting for a reason
  const confident = () => usable() && callOk() && S.k.nMs > 0 && !checkMsgs().some(x => x[0] === "warn" || x[0] === "bad");
  function paintFound() {
    const k = S.k, ok = usable();
    const called = S.call && callOk() && k.ragCalc && k.ragCalc !== "—";
    $("summary").innerHTML = ok
      ? `<span class="rag ${esc(k.rag)}">${esc(k.rag)}</span><div><b>${Math.round(k.pct * 100)}% complete${k.finish != null ? `, finishing ${F(k.finish)}` : ""}</b><p>${called ? `Your call (calculated ${esc(k.ragCalc)}): ${esc(k.ragReason || S.call.reason)}` : esc(k.why.join(" · "))}</p></div>`
      : `<span class="rag none">—</span><div><b>${S.body ? "One column needs you" : "We couldn't find a plan in this"}</b><p>${S.body ? `Pick the ${k.cols.name < 0 ? "task name" : "finish or due date"} column below.` : "Check the file has a header row with task names and dates."}</p></div>`;
    $("facts").innerHTML = ok ? [[k.nTask, "task"], [k.nMs, "milestone"], [k.nPhases, "phase"]].map(([n, w]) => `<li><b>${n}</b> ${w}${n === 1 ? "" : "s"}</li>`).join("") +
      `<li>Headers in row <b>${k.hdrRow}</b></li>` + (k.evidence ? `<li>Dates read <b>${k.dayFirst ? "day first" : "month first"}</b></li>` : "") : "";
    const m = checkMsgs();
    $("msgs").innerHTML = m.map(([t, a, b]) => `<li class="msg ${t}"><div>${esc(a)}${b ? `<br><span>${esc(b)}</span>` : ""}</div></li>`).join("");
    const HC = window.KeelsonHealth, hc = ok ? HC.run(k, S.hist.filter(v => v.asOf < E.iso(k.statusDate)).pop() || null) : [];
    $("health-d").hidden = !hc.length; if (S.sample && /messy/i.test(S.file)) $("health-d").open = true; $("healthsum").textContent = HC.summary(hc); $("health").innerHTML = HC.html(hc, esc);
    colMap(k); paintRules(); paintMs(); paintCal();
    if (!ok && S.body) $("cols").open = true;
    const to = $("to-reports");
    to.setAttribute("aria-disabled", String(!ok || !callOk()));
    $("ready-note").textContent = ok ? `${nWord(ALL.length)} reports ready` : "";
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
    if (store.get("keelson-projects", {})[projKey()]) { S.rules = null; S.call = null; S.msExclude = []; S.cal = calDefault(); S.calMine = false; loadProject(); compute(); paintCheck(); }
    else if (S.rules || S.call || S.msKey || S.msExclude.length || S.calMine) saveProject();
    const ed = sess.get("keelson-edits", {})[projKey()];
    if (ed && Object.keys(ed).length) S.edits = ed; else saveEdits();
    withNotes(); paintCrumb();
    if (!S.histMem) loadHist();
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
  // working week: saved with this project, and remembered as the default for the next new one
  $("cal-f").hidden = !CAL_OK;
  function paintCal() {
    const r = applyCal(); if (!r) return;
    $("cal-hol-n").textContent = !S.cal.hol.trim() ? "" : `${plural(r.n, "holiday")} used${r.ignored ? ` · ${plural(r.ignored, "line")} ignored (not a date, or a repeat)` : ""}`;
  }
  const calChanged = () => { S.calMine = true; store.set("keelson-calendar", S.cal); saveProject(); compute(); paintFound(); };
  $("cal-week").addEventListener("change", e => { if (Object.hasOwn(E.WEEKENDS || {}, e.target.value)) { S.cal.week = e.target.value; calChanged(); } });
  $("cal-hol").addEventListener("input", e => { S.cal.hol = clean(e.target.value).slice(0, 5000); calChanged(); });
  $("sd").addEventListener("change", e => { const v = fromIso(e.target.value); if (v != null) { S.sd = v; compute(); paintFound(); } });
  $("to-reports").addEventListener("click", e => {
    if (!usable()) { e.preventDefault(); $("cols").open = true; $("colmap").querySelector(".miss + select")?.focus(); }
    else if (!callOk()) { e.preventDefault(); $("rules-d").open = true; $("call-why").focus(); }
  });

  // ---------- snapshots: "since last report" ----------
  const projKey = () => (S.title || "Project").trim().toLowerCase();
  // "since last report" in the weekly: snapshots saved on download (localStorage), else the sample's in-memory ones, else a
  // snapshot read off a saved version (backfilled or imported; it has no per-phase %, so the weekly skips that column)
  const baseKey = t => String(t.key ?? "").replace(/\s*\(?#\d+\)?$/, "");
  function snapOf(v) {
    const n = {}, tasks = {};
    for (const t of v.tasks) n[baseKey(t)] = (n[baseKey(t)] || 0) + 1;
    for (const t of v.tasks) if (t.name && n[baseKey(t)] === 1 && t.finish != null) tasks[baseKey(t)] = E.iso(t.finish);
    const d = x => x == null ? null : E.iso(x);
    return { v: 2, asOf: v.asOf, rag: v.rag, pct: v.pct, finish: d(v.finish), bfinish: d(v.bfinish), var: v.var, overT: v.overT, overMs: v.overMs,
      nMs: v.nMs, nMsDone: v.nMsDone, nTask: v.nTask, nTaskDone: v.nTaskDone, phases: {}, tasks };
  }
  const snaps = () => {
    const by = {};
    for (const v of S.hist) by[v.asOf] = snapOf(v);
    for (const x of [...S.memSnaps, ...(store.get("keelson-snapshots", {})[projKey()] || [])]) if (x && x.asOf) by[x.asOf] = x;
    return Object.values(by);
  };
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
      el.innerHTML = icon("swap") +
        `<span>${S.compare ? "Comparing with" : "Not comparing with"} your report of <b>${F(fromIso(p.asOf))}</b></span>` +
        `<label class="switch"><input type="checkbox" role="switch" ${S.compare ? "checked" : ""} data-cmp> Compare</label>`;
    });
  }
  document.addEventListener("change", e => { if (e.target.matches("[data-cmp]")) { S.compare = e.target.checked; paintCompare(); refresh(); } });

  // ---------- version history (KeelsonHistory: IndexedDB, or memory for a sample) ----------
  const sheetName = () => S.sheets ? S.sheets[S.sheet].name : "";
  let histSeq = 0;
  async function loadHist() {
    if (!KH || !S.k) return;
    const seq = ++histSeq;
    let l = [];
    try { l = await KH.list(projKey()); } catch (e) { console.error("Keelson: version history", e); }
    if (seq !== histSeq) return;
    S.hist = l; paintHistBtn(); refresh();
    if ($("hist").open) paintHist();
  }
  async function memHistory(versions) { // the sample's earlier weeks: computed here, stored in a memory store, never in IndexedDB
    KH.use(KH.memoryStore()); S.histMem = true; S.memSnaps = [];
    for (const v of versions) {
      const k = comp(v.text, { statusDate: fromIso(v.asOf), msKeys: msKeys() });
      await KH.put(projKey(), KH.makeVersion(k, { project: projKey(), asOf: v.asOf, file: v.file, savedAt: v.asOf + "T09:00:00.000Z" }));
      if (R.snapshot) S.memSnaps.push(R.snapshot(k, S.memSnaps[S.memSnaps.length - 1] || null));
    }
  }
  // Word or PowerPoint download or Save version: this plan becomes a saved version (a sample's history is kept from then on). Never on PDF
  // print: afterprint also fires when the print dialog is cancelled.
  async function saveVersion() {
    if (!KH) return false;
    let ok = true;
    try {
      if (S.histMem) { const vs = await KH.list(projKey()); KH.use(null); S.histMem = false; for (const v of vs) await KH.put(projKey(), v); }
      await KH.save(projKey(), S.k, { file: S.file, sheet: sheetName() });
    } catch (e) { ok = false; console.error("Keelson: version not saved", e); toast("This version couldn't be saved in this browser (storage is full or off)."); }
    await loadHist();
    return ok;
  }
  async function savePoint() { saveSnap(); const ok = await saveVersion(); S.dirty = false; paintCompare(); return ok; }
  document.addEventListener("click", async e => {
    if (!e.target.closest("[data-save]") || !S.k) return;
    if (await savePoint()) toast(`Saved the version of ${F(S.sd)}${KH && KH.storeKind === "memory" ? " for this tab" : ""}`);
  });
  // this plan as a Version (not saved), and the versions the v2 reports read
  let cvK = null, cvV = null;
  const curVersion = () => { if (cvK !== S.k) { cvK = S.k; cvV = KH.makeVersion(S.k, { project: projKey(), file: S.file, sheet: sheetName() }); } return cvV; };
  const prevVersion = () => { if (S.cmp) return S.cmp.version; const a = curVersion().asOf; return S.hist.filter(v => v.asOf < a).pop() || null; };
  const trendVersions = () => { const c = curVersion(); return [...S.hist.filter(v => v.asOf !== c.asOf), c]; };
  const paintHistBtn = () => { $("hist-n").textContent = S.hist.length ? `Version history · ${S.hist.length}` : "Version history"; };

  // ---------- report building ----------
  if (H && !document.getElementById("kr-css")) { const st = document.createElement("style"); st.id = "kr-css"; st.textContent = H.css; document.head.append(st); }
  const opts = id => ({ title: S.title || "Project status", client: S.client, preparedBy: S.prep, brand: { ...S.brand }, prev: S.compare ? prevSnap() : null, history: S.compare ? histOf(prevSnap()) : [], edits: { ...(S.edits[id] || {}) }, actions: id === "weekly" ? actsFor() : null });
  // the forecast runs 5,000 resamples: its result and its page are kept until the plan (S.k) or the saved versions (S.hist) change
  // ponytail: KF.report() can't take a precomputed result, so a new plan still computes twice (page + tool panel)
  const fc = { k: null, h: null, res: null, key: "", doc: null };
  const fcFresh = () => { if (fc.k !== S.k || fc.h !== S.hist) Object.assign(fc, { k: S.k, h: S.hist, res: null, key: "", doc: null }); return fc; };
  function forecastRes() {
    const c = fcFresh();
    if (!c.res) try { c.res = KF.forecast(S.k, { versions: trendVersions() }); } catch (e) { c.res = { ok: false, reason: "The forecast couldn't be worked out for this plan." }; }
    return c.res;
  }
  function doc(id) {
    applyCal();
    if (!isV2(id)) return R.build(id, S.k, opts(id));
    const o = opts(id), base = { title: o.title, client: o.client, preparedBy: o.preparedBy, edits: o.edits };
    if (id === "trend") return KT.report(trendVersions(), base);
    if (id === "forecast") { const c = fcFresh(), key = JSON.stringify(base); if (c.key !== key || !c.doc) Object.assign(c, { key, doc: KF.report(S.k, { ...base, versions: trendVersions() }) }); return c.doc; }
    const p = prevVersion();
    if (p) return KD.report(p, curVersion(), { ...base, renames: { accepted: S.diff.accepted, split: S.diff.split }, exclude: S.diff.exclude });
    return { id: "changes", name: "What changed", title: o.title, subtitle: ["What changed in the plan", o.client && `for ${o.client}`].filter(Boolean).join(" ") + (o.preparedBy ? ` · Prepared by ${o.preparedBy}` : ""),
      asOf: E.iso(S.sd), rag: S.k.rag || "—", orientation: "portrait", blocks: [
        { type: "note", text: "Nothing to compare with yet. This report sets this plan beside your last saved report, or beside another export of the same plan." },
        { type: "heading", text: "What it will show" },
        { type: "list", items: ["Tasks added, removed or cancelled.", "Finish dates that moved later or earlier, in working days, and milestone moves.", "Baselines that changed, flagged in red.",
          "Owner changes, renames, phase moves and progress that went backwards.", "One sentence that sums it up, for your sponsor."] },
        { type: "heading", text: "How to start" },
        { type: "list", items: ["Compare with another file: drop last week's export of this plan above the page.", "Or press Save version today (a Word or PowerPoint download saves one too): next week this one compares itself automatically.",
          "Or add older exports under Version history; each is placed by its status date."] }] };
  }
  // re-baseline notes are an audit trail, so they are kept per project in localStorage (unsaved commentary is only per tab)
  const NOTES = "keelson-notes";
  function withNotes() { const n = store.get(NOTES, {})[projKey()]; if (n && typeof n === "object") S.edits.trend = { ...n, ...(S.edits.trend || {}) }; }
  const page = (d, edit, forPrint) => H.html(d, { brand: S.brand, watermark: !isPro() && !forPrint, footer: !isPro(), editable: edit });

  // ---------- step 3: gallery ----------
  // what a v2 card still needs before it says much, shown on the card
  function need(id, d) {
    if (id === "changes") return prevVersion() ? "" : "Needs an earlier version";
    if (id === "trend") { const n = trendVersions().length; return n < 3 ? `Needs 3 reports · you have ${n}` : ""; }
    if (id === "forecast") return d && d.blocks[0] && d.blocks[0].type === "note" && /will not estimate/.test(d.blocks[0].text) ? "Not enough history to forecast" : "";
    return "";
  }
  function card(r) {
    let d; try { d = doc(r.id); } catch (e) { console.error("Keelson: report", r.id, e); d = null; }
    const land = (d?.orientation || r.orientation) === "landscape", v2 = isV2(r.id);
    const locked = !isPro() && r.id !== "onepager", nd = v2 ? need(r.id, d) : "";
    return `<li><a class="rc${v2 ? " v2" : ""}" href="#report/${esc(r.id)}" aria-label="${esc(r.name)}${r.description ? ". " + esc(r.description) : ""}${nd ? ". " + esc(nd) : ""}">
      <div class="thumbbox" aria-hidden="true"><div class="thumb ${land ? "landscape" : "portrait"}"><div class="thumb-in" inert style="width:${land ? 1123 : 794}px">${d ? page(d, false) : ""}</div></div></div>
      <div class="rc-b"><h3><span>${esc(r.name)}</span>${locked ? '<span class="lock">Pro</span>' : ""}</h3>${r.description ? `<p>${esc(r.description)}</p>` : ""}</div>
      <div class="meta"><span>${land ? "Landscape" : "Portrait"}</span><span>${isPro() ? "PDF · Word · PowerPoint" : r.id === "onepager" ? "Free PDF" : "Preview"}</span>${nd ? `<span class="need">${esc(nd)}</span>` : ""}</div></a>` +
      (r.id === "changes" && nd ? `<button class="btn fill block" type="button" id="g-cmp">${icon("plus")}Add last week's export to see what changed</button>` : "") + `</li>`;
  }
  const LEAD = ["onepager", "changes", "forecast"]; // the three reports most PMs want first; the rest sit under More reports
  // "Columns detected: … · Change": a wrong mapping is one click from Check
  function paintDetected() {
    const used = E.FIELDS.filter(([key]) => S.k.cols[key] >= 0).map(([, l]) => l.split(" / ")[0]);
    $("detected").innerHTML = `${icon("check")}<span>Columns detected: ${esc(used.slice(0, 6).join(", "))}${used.length > 6 ? `, +${used.length - 6} more` : ""}` +
      `${CAL_OK && E.calText ? ` · Working week: <span style="white-space:nowrap">${esc(E.calText())}</span>` : ""}</span> <span style="white-space:nowrap">· <a href="#check" id="detected-change" aria-label="Change columns${CAL_OK ? " or working week" : ""}">Change</a></span>`;
  }
  $("detected").addEventListener("click", e => { if (e.target.closest("#detected-change")) { $("cols").open = true; toCols = true; } });
  let toCols = false;
  function paintGallery() {
    $("loop-hint").hidden = isPro();
    $("reports-sub").innerHTML = `${nWord(ALL.length)} reports for <b>${esc(S.title || "Your project")}</b> as of ${F(S.sd)}. Open one to edit the commentary and download it.${isPro() ? "" : ` ${esc(FREE_RULE)}.`}`;
    const n = trendVersions().length, lead = LEAD.map(id => ALL.find(r => r.id === id)).filter(Boolean), more = ALL.filter(r => !lead.includes(r));
    $("gallery").innerHTML = `<h3 class="ghead" id="g-lead">Start here</h3><ul class="gallery three" aria-labelledby="g-lead">${lead.map(card).join("")}</ul>` +
      (V2.length ? `<p class="hint gnote">What changed, Slip trend and Finish-date confidence use this project's saved versions: ${n === 1 ? "only this plan so far" : `${n} reports from ${F(fromIso(trendVersions()[0].asOf))}`}. Press Save version (or download a Word or PowerPoint report) to keep this week's. Add older exports under Version history.</p>` : "") +
      `<h3 class="ghead" id="g-more">More reports</h3><ul class="gallery" aria-labelledby="g-more">${more.map(card).join("")}</ul>`;
    paintDetected();
    fitThumbs();
    paintChase(); paintActs();
  }
  // chase list: who to ask for an update, from this plan (and the last saved version for stale progress); copying is Pro
  function paintChase() {
    const C = window.KeelsonChase, g = S.k ? C.build(S.k, prevVersion(), AC ? actsLog() : []) : [];
    S.chase = g; $("chase").hidden = !usable(); $("chase-all").hidden = !g.length;
    if (!g.length) { $("chase-sub").textContent = "Nobody to chase this week: nothing overdue, nothing due in the next 5 working days, nothing stuck since the last version."; $("chase-list").innerHTML = ""; return; }
    const n = g.reduce((t, x) => t + x.items.length, 0), what = { overdue: "Overdue", soon: "Due in the next 5 working days", stale: "No progress since the last version", action: "Overdue action from the log" };
    $("chase-sub").textContent = `${plural(n, "task")} to ask about, ${g.length} ${g.length === 1 ? "person" : "people"}. Copy a message and paste it into email, Teams or Slack. Keelson sends nothing.`;
    $("chase-list").innerHTML = g.map((x, i) => `<div class="ch"><div class="ch-h"><b>${esc(x.owner)}</b><span>${x.items.length}</span><button class="btn ghost sm" type="button" data-chase="${i}">Copy message</button></div>
      <ul>${x.items.map(t => `<li class="${t.kind}">${esc(t.name)}<small>${esc(what[t.kind])} · due ${F(t.finish)}${t.kind === "overdue" ? ` · ${t.late} wd late` : ""}</small></li>`).join("")}</ul></div>`).join("");
  }
  function copyChase(i) {
    if (!isPro()) return openPro("Copy ready-to-send chase messages, plus Word and PowerPoint export — $12 a month or $96 a year.");
    const C = window.KeelsonChase, txt = i == null ? C.all(S.chase, S.title) : C.text(S.chase[i], S.title);
    navigator.clipboard.writeText(txt).then(() => toast(i == null ? "All chase messages copied" : `Message for ${S.chase[i].owner} copied`), () => toast("Couldn't copy. Select the text and copy it yourself."));
  }
  $("chase-all").addEventListener("click", () => copyChase(null));
  $("chase-list").addEventListener("click", e => { const b = e.target.closest("[data-chase]"); if (b) copyChase(+b.dataset.chase); });

  // ---------- actions, decisions and risks: a log per project in this browser, from minutes or tables (actions.js); Pro ----------
  const AC = window.KeelsonActions, ACTS = "keelson-actions";
  const actsFor = () => { const l = AC ? actsLog() : []; return l.length ? { ...AC.summary(l, S.sd), sd: S.sd } : null; }; // for the Weekly report
  const actsLog = () => { const l = store.get(ACTS, {})[projKey()]; return Array.isArray(l) ? l : []; };
  const actsPut = log => { const all = store.get(ACTS, {}); all[projKey()] = log; if (!store.set(ACTS, all)) toast("Couldn't save the log in this browser (storage full or blocked). Download the CSV to keep it."); };
  let actsDraft = null; // items read and waiting for the PM to confirm
  const isoOr = v => v == null ? "" : E.iso(v), dayIn = v => v ? E.parseDate(v) : null;
  function paintActs() {
    if (!AC) return;
    $("acts").hidden = !usable(); $("acts-sample").hidden = !(S.sample && SAMPLES.minutes);
    if (!usable()) return;
    const log = actsLog(), sm = AC.summary(log, S.sd), all = $("acts-filter").value === "all";
    $("acts-sub").textContent = log.length
      ? `${plural(sm.openActions, "open action")} · ${sm.overdue.length} overdue · ${plural(sm.decisions.length, "decision")} this week · ${plural(sm.risks.length, "open risk")}. Kept in this browser for ${S.title || "this project"}.`
      : "Turn meeting minutes into a log that carries forward week to week. Overdue actions join the chase list and the Weekly report.";
    $("acts-tools").hidden = !log.length; $("acts-csv").hidden = !log.length;
    const rows = log.filter(x => all || (x.type === "Decision" ? (x.raised ?? x.added) >= S.sd - 7 : x.status === "Open")) // decisions: this week's.sort((a, b) => (a.status !== "Open") - (b.status !== "Open") || (a.due ?? 1e9) - (b.due ?? 1e9));
    $("acts-list").innerHTML = !log.length ? "" : !rows.length ? `<p class="hint">Nothing open and no decisions this week. Choose Everything to see the whole log.</p>` :
      `<table class="at"><thead><tr><th class="c-type">Type</th><th>Item</th><th class="c-own">Owner</th><th class="c-due">Due</th><th class="c-st">Status</th><th class="c-x"><span class="sr">Remove</span></th></tr></thead><tbody>${rows.map(x => {
        const late = x.status === "Open" && x.due != null && x.due < S.sd, slip = x.firstDue != null && x.due != null && x.due > x.firstDue ? ` · moved ${E.wdiff(x.firstDue, x.due)} wd` : "";
        return `<tr class="${x.status === "Open" ? "" : "done"}"><td class="c-type"><span class="tag">${esc(x.type)}</span></td><td>${esc(x.text)}${x.score ? ` <span class="hint">(score ${esc(x.score)})</span>` : ""}</td><td class="c-own">${esc(x.owner) || (x.type === "Decision" ? "" : '<span class="hint">No owner</span>')}</td>
          <td class="c-due${late ? " late" : ""}">${x.due != null ? F(x.due) : ""}${late ? " · late" : ""}<span class="hint">${slip}</span></td>
          <td class="c-st">${x.type === "Decision" ? "" : x.status === "Open" ? `<button class="link" type="button" data-act="done" data-id="${esc(x.id)}">${x.type === "Risk" ? "Close" : "Mark done"}</button>` : `${esc(x.status)} · <button class="link" type="button" data-act="open" data-id="${esc(x.id)}">Reopen</button>`}</td>
          <td class="c-x"><button class="link" type="button" data-act="del" data-id="${esc(x.id)}" aria-label="Remove ${esc(x.text)}">✕</button></td></tr>`; }).join("")}</tbody></table>`;
  }
  function paintDraft() {
    const r = actsDraft;
    if (!r) { $("acts-review").innerHTML = ""; return; }
    const opts = v => AC.TYPES.map(t => `<option${t === v ? " selected" : ""}>${t}</option>`).join("");
    $("acts-review").innerHTML = !r.items.length ? "" : `<table class="at"><thead><tr><th class="c-x"><span class="sr">Keep</span></th><th class="c-type">Type</th><th>Item</th><th class="c-own">Owner</th><th class="c-due">Due</th></tr></thead><tbody>${r.items.map((x, i) =>
      `<tr><td class="c-x"><input type="checkbox" checked data-i="${i}" data-f="keep" aria-label="Keep ${esc(x.text)}"></td><td class="c-type"><select class="in" data-i="${i}" data-f="type" aria-label="Type">${opts(x.type)}</select></td>
       <td><input class="in" data-i="${i}" data-f="text" value="${esc(x.text)}" aria-label="Item"></td><td class="c-own"><input class="in" data-i="${i}" data-f="owner" value="${esc(x.owner)}" aria-label="Owner" placeholder="No owner"></td>
       <td class="c-due"><input class="in" type="date" data-i="${i}" data-f="due" value="${isoOr(x.due)}" aria-label="Due"></td></tr>`).join("")}</tbody></table>
      <div class="acts-save"><button class="btn fill sm" type="button" id="acts-save">Save to the log</button><button class="btn ghost sm" type="button" id="acts-cancel">Cancel</button></div>`;
  }
  function readActs(text, from) {
    const md = dayIn($("acts-date").value) ?? S.sd, r = AC.read(text, { meetingDate: md, dayFirst: S.k ? S.k.dayFirst !== false : true });
    actsDraft = { ...r, items: r.items.map(x => ({ ...x, keep: true })) };
    const sk = r.skipped.length;
    $("acts-msg").className = "status " + (r.items.length ? "ok" : "bad");
    $("acts-msg").textContent = r.items.length
      ? `Found ${plural(r.items.length, "item")}${from ? ` in ${from}` : ""}${r.kind === "table" ? " (read as a table)" : ""}${sk ? `; ${plural(sk, "line")} without a marker skipped` : ""}. Check them, then save.`
      : r.kind === "notes" ? "Nothing found. These minutes may be discussion only. Start lines with ACTION:, DECISION:, RISK: or ISSUE:, put bullets under a heading like “Action items” or “Next steps”, or use “Use your own AI tool” below to pull them out." : "No rows found under the headers.";
    paintDraft();
  }
  const actsPro = () => { if (isPro()) return true; openPro("Keep an actions, decisions and risks log from your meeting minutes, plus Word and PowerPoint export — $12 a month or $96 a year."); return false; };
  $("acts-new").addEventListener("click", () => { const p = $("acts-add"); p.hidden = !p.hidden; if (!p.hidden) { if (!$("acts-date").value) $("acts-date").value = isoOr(S.sd); $("acts-text").focus(); } });
  $("acts-sample").addEventListener("click", () => { const m = SAMPLES.minutes; if (!m) return; $("acts-text").value = m.text; $("acts-date").value = m.meetingDate; $("acts-text").focus(); });
  $("acts-read").addEventListener("click", () => { if (actsPro()) readActs($("acts-text").value); });
  $("acts-file").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f || !actsPro()) return;
    try {
      const p = await parsePlan(f), tsv = g => g.map(r => r.map(c => String(c ?? "").replace(/[\t\r\n]+/g, " ")).join("\t")).join("\n");
      const text = p.text ?? (p.sheets.map(sh => tsv(sh.rows)).find(t => AC.parseTable(t)) || tsv(p.sheets[0].rows));
      readActs(text, f.name);
    } catch (err) { $("acts-msg").className = "status bad"; $("acts-msg").textContent = err.plain ? err.message : "Couldn't read that file. Save it as .xlsx or CSV and try again."; }
  });
  $("acts-file-l").addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("acts-file").click(); } });
  $("acts-review").addEventListener("input", e => { const t = e.target, i = t.dataset.i; if (i == null || !actsDraft) return; const x = actsDraft.items[+i];
    if (t.dataset.f === "keep") x.keep = t.checked; else if (t.dataset.f === "due") x.due = dayIn(t.value); else x[t.dataset.f] = t.dataset.f === "type" ? t.value : t.value.slice(0, AC.LIM.text); });
  $("acts-review").addEventListener("click", e => {
    if (e.target.id === "acts-cancel") { actsDraft = null; paintDraft(); $("acts-msg").textContent = ""; return; }
    if (e.target.id !== "acts-save" || !actsDraft || !actsPro()) return;
    const keep = actsDraft.items.filter(x => x.keep && x.text.trim()).map(x => ({ ...x, text: x.text.trim(), owner: x.owner.trim() }));
    const r = AC.merge(actsLog(), keep, S.sd); actsPut(r.log);
    actsDraft = null; paintDraft(); $("acts-text").value = "";
    $("acts-msg").className = "status ok"; $("acts-msg").textContent = `Saved: ${r.added} new, ${r.updated} updated${r.same ? `, ${r.same} already in the log` : ""}.`;
    paintActs(); paintChase();
  });
  $("acts-list").addEventListener("click", e => {
    const b = e.target.closest("[data-act]"); if (!b || !actsPro()) return;
    const id = b.dataset.id, log = actsLog(), x = log.find(y => y.id === id); if (!x) return;
    if (b.dataset.act === "del" && !confirm(`Remove “${x.text}” from the log?`)) return;
    actsPut(b.dataset.act === "del" ? AC.remove(log, id) : AC.setStatus(log, id, b.dataset.act === "done" ? "Done" : "Open", S.sd));
    paintActs(); paintChase();
  });
  $("acts-filter").addEventListener("change", paintActs);
  $("acts-csv").addEventListener("click", () => { if (actsPro()) saveBlob(new Blob([AC.toCSV(actsLog())], { type: "text/csv" }), `${(S.title || "project").replace(/[^\p{L}\p{N} _-]+/gu, "").trim() || "project"} actions ${E.iso(S.sd)}.csv`); });
  $("acts-tpl").addEventListener("click", e => { e.preventDefault(); saveBlob(new Blob([AC.TEMPLATE], { type: "text/csv" }), "keelson-actions-template.csv"); });
  $("acts-prompt").textContent = AC ? AC.PROMPT : "";
  $("acts-prompt-copy").addEventListener("click", () => navigator.clipboard.writeText(AC.PROMPT).then(() => toast("Prompt copied"), () => toast("Couldn't copy. Select the prompt and copy it yourself.")));
  // one version only: last week's export goes straight into What changed (read here, compared, never uploaded)
  $("gallery").addEventListener("click", e => { if (e.target.closest("#g-cmp")) $("g-cmp-file").click(); });
  $("g-cmp-file").addEventListener("change", async e => { const f = e.target.files[0]; e.target.value = ""; if (f && await compareFile(f)) go("report/changes"); });
  const ro = new ResizeObserver(es => es.forEach(e => fit(e.target)));
  function fit(t) { const inn = t.firstElementChild; if (!inn) return; const s = t.clientWidth / parseFloat(inn.style.width); inn.style.transform = `scale(${s})`; }
  function fitThumbs() { document.querySelectorAll(".thumb").forEach(t => { fit(t); ro.observe(t); }); }

  // ---------- report view ----------
  function paintSwitch() {
    $("switch").innerHTML = ALL.map(r => `<a href="#report/${esc(r.id)}"${isV2(r.id) ? ' class="v2"' : ""}${r.id === S.report ? ' aria-current="page"' : ""}>${esc(r.name)}</a>`).join("");
  }
  function paintReport(keepFocus) {
    const id = S.report, r = ALL.find(x => x.id === id);
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
    $("dl-pptx").innerHTML = pro ? "Download PowerPoint" : 'Download PowerPoint <span class="lock">Pro</span>';
    $("dl-note").hidden = pro;
    $("dl-note").textContent = pro ? "" : `${FREE_RULE}${id === "onepager" ? ", with a small Keelson footer" : ""}. Pro removes the watermark and footer and unlocks every report as PDF, Word and PowerPoint.`;
    paintBranding();
    paintTool();
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
  function saveEdits() {
    const all = sess.get("keelson-edits", {}); all[projKey()] = S.edits; sess.set("keelson-edits", all);
    const n = store.get(NOTES, {}); if (S.edits.trend && Object.keys(S.edits.trend).length) n[projKey()] = S.edits.trend; else delete n[projKey()]; store.set(NOTES, n);
  }
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
  $("wipe").addEventListener("click", async () => {
    if (!confirm("Remove everything Keelson saved in this browser? That's your theme, branding and logo, licence key, report history, version history of every project, status rules, actions logs and unsaved commentary. Your plan files aren't affected.")) return;
    for (const st of [localStorage, sessionStorage]) try { Object.keys(st).filter(k => k.startsWith("keelson-")).forEach(k => st.removeItem(k)); } catch (e) {}
    let cleared = true;
    if (KH) { // delete the whole "keelson" IndexedDB database (and a sample's in-memory versions)
      try { if (S.histMem) { KH.use(null); S.histMem = false; } cleared = (await KH.clear()) === true; } catch (e) { cleared = false; console.error("Keelson: clearing version history", e); }
      S.hist = []; S.memSnaps = []; S.cmp = null; paintHistBtn();
    }
    if (lic) release(lic.key, lic.instanceId);
    lic = null; S.brand = { color: "#1a7f86", logo: null }; S.client = S.prep = ""; S.edits = {}; S.dirty = false; S.rules = null; S.call = null; S.msExclude = []; S.msKey = ""; S.cal = calOf(null); S.calMine = false;
    $("client").value = $("prep").value = "";
    if (S.k) compute();
    paintPro(); refresh(); toast(cleared ? "Saved data cleared" : "Version history couldn't be cleared. Close other Keelson tabs, then try again.");
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
  const fileBase = ext => { const r = ALL.find(x => x.id === S.report); return safeName([S.title || "Project", r ? r.name : S.report, E.iso(S.sd)].join(" - "), ext); };
  $("dl-pdf").addEventListener("click", () => {
    if (!isPro() && S.report !== "onepager") return lockedPro();
    const pr = $("print-root"), t = document.title;
    pr.innerHTML = page(doc(S.report), false, true);
    document.body.classList.add("printing"); document.title = fileBase(); // the browser names the PDF after the title
    // nothing is saved here: afterprint fires whether the PDF was saved or the dialog cancelled (Save version does that)
    addEventListener("afterprint", () => { pr.innerHTML = ""; document.body.classList.remove("printing"); document.title = t; }, { once: true });
    print();
  });
  // Word and PowerPoint: Pro only; the library loads on first use; a download also saves this version
  const office = (id, mod, kind, ext, done) => $(id).addEventListener("click", async () => {
    if (!isPro()) return lockedPro();
    const D = window[mod];
    if (!D) return toast(`${kind} export isn't available in this build.`);
    const b = $(id), label = b.innerHTML;
    b.disabled = true; b.innerHTML = `<span class="spin" aria-hidden="true"></span> Preparing ${kind}…`;
    try {
      if (typeof D.load === "function") await D.load();
      const blob = await D.build(doc(S.report), { brand: { ...S.brand }, footer: !isPro(), rasterize });
      saveBlob(blob, fileBase(ext));
      await savePoint(); toast(done);
    } catch (e) { console.error(`Keelson: ${kind} export`, e); toast(`The ${kind} file couldn't be made. Try again, or download the PDF instead.`); }
    finally { b.disabled = false; b.innerHTML = label; }
  });
  office("dl-docx", "KeelsonDocx", "Word", ".docx", "Word document saved");
  office("dl-pptx", "KeelsonPptx", "PowerPoint", ".pptx", "PowerPoint deck saved");
  // svg -> Image -> canvas at 2x -> PNG bytes, for the Word and PowerPoint exports
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

  // ---------- v2 tool panel above the page: compare source, change review, trend versions, forecast inputs ----------
  const GROUPS = [["all", "All"], ["milestoneMoved", "Milestones"], ["baselineMoved", "Baseline"], ["finishLater", "Later"], ["finishEarlier", "Earlier"],
    ["added", "Added"], ["removed", "Removed", ["removed", "cancelled"]], ["startMoved", "Start moved"], ["ownerChanged", "Owner"],
    ["progress", "Progress", ["reopened", "pctBackwards"]], ["renamed", "Renamed"], ["phaseMoved", "Phase"]];
  const WHAT = { milestoneMoved: "Milestone moved", baselineMoved: "Baseline changed", finishLater: "Finish later", finishEarlier: "Finish earlier", added: "Added",
    removed: "Removed", cancelled: "Cancelled", startMoved: "Start moved", ownerChanged: "Owner changed", reopened: "Reopened", pctBackwards: "% complete down",
    renamed: "Renamed", phaseMoved: "Moved phase" };
  const inGroup = (g, c) => g === "all" || ((GROUPS.find(x => x[0] === g) || [])[2] || [g]).includes(c.type);
  const pctT = v => v == null ? "—" : Math.round(v * 100) + "%";
  const Ds = v => v == null ? "—" : E.fmt(v, E.iso(v).slice(0, 4) !== E.iso(S.sd).slice(0, 4));
  const span = c => ["ownerChanged", "renamed", "phaseMoved"].includes(c.type) ? `${c.from || "—"} → ${c.to || "—"}`
    : c.type === "reopened" ? `Done → ${pctT(c.to)}` : c.type === "pctBackwards" ? `${pctT(c.from)} → ${pctT(c.to)}`
    : c.type === "added" ? `Due ${Ds(c.to)}` : ["removed", "cancelled"].includes(c.type) ? `Was due ${Ds(c.from)}` : `${Ds(c.from)} → ${Ds(c.to)}`;
  const toolHead = (r, title, sub, extra = "") => `<div class="tool-h"><div class="grow"><h3>${title}</h3><p>${sub}</p></div>${extra}</div>`;
  const cmpDrop = () => `<label class="minidrop" for="cmp-file" id="cmp-drop"><input type="file" id="cmp-file" accept=".xlsx,.csv,.tsv,.txt">
      ${icon("swap", "ic sm")}<span><b>Compare with another file</b><span>${S.cmp ? "Drop a different export to compare with instead." : "Last week's export of this plan, as .xlsx or CSV. Read here, never uploaded."}</span></span></label>`;
  function toolChanges(r) {
    const p = prevVersion(), cur = curVersion();
    const src = S.cmp ? `Comparing with <b>${esc(S.cmp.name)}</b>` : p ? `Comparing with your saved version of <b>${esc(F(fromIso(p.asOf)))}</b>` : "There is no earlier version of this project yet.";
    const date = S.cmp ? `<div class="field" style="flex:0 1 180px"><label for="cmp-date">Its status date</label><input class="in" type="date" id="cmp-date" value="${E.iso(S.cmp.asOf)}"></div>` : "";
    const back = S.cmp && S.hist.some(v => v.asOf < cur.asOf) ? `<button class="link" type="button" data-act="cmp-clear">Compare with the last saved version instead</button>` : "";
    const head = toolHead(r, "What changed", src, date) + `<div class="row">${cmpDrop()}${back}</div>`;
    if (!p) return head;
    const rn = { accepted: S.diff.accepted, split: S.diff.split };
    const all = KD.diff(p, cur, { renames: rn }), inc = KD.diff(p, cur, { renames: rn, exclude: S.diff.exclude });
    const groups = GROUPS.filter(([g]) => g === "all" || all.changes.some(c => inGroup(g, c)));
    if (!groups.some(([g]) => g === S.diff.filter)) S.diff.filter = "all";
    const strip = `<div class="strip" role="group" aria-label="Show changes by type">${groups.map(([g, label]) =>
      `<button type="button" class="cnt" data-f="${g}" aria-pressed="${S.diff.filter === g}">${label} <b>${inc.changes.filter(c => inGroup(g, c)).length}</b></button>`).join("")}</div>`;
    const renames = all.renames.length ? `<div class="renames">${all.renames.map(x => {
      const acc = S.diff.accepted.includes(x.key), spl = S.diff.split.includes(x.key);
      return `<div class="rn"><span class="grow">${acc ? "Confirmed rename" : spl ? "Kept as two tasks" : "Same task, renamed?"} <b>${esc(x.from)}</b> → <b>${esc(x.to)}</b>${x.phase ? ` <span style="color:var(--muted)">· ${esc(x.phase)}</span>` : ""}</span>` +
        (acc || spl ? `<button type="button" class="btn line sm" data-rn="${esc(x.key)}" data-act="undo">Undo</button>`
          : `<button type="button" class="btn line sm" data-rn="${esc(x.key)}" data-act="accept">Confirm rename</button><button type="button" class="btn line sm" data-rn="${esc(x.key)}" data-act="split">Treat as two tasks</button>`) + `</div>`;
    }).join("")}</div>` : "";
    const rows = all.changes.filter(c => inGroup(S.diff.filter, c)), CAP = 200;
    const table = !all.changes.length ? `<p class="hint">No changes of a working day or more between the two versions.</p>` :
      `<div class="chg" id="chg"><table><thead><tr><th scope="col"><span class="sr">Include in the client version</span></th><th scope="col">Change</th><th scope="col">Task</th><th scope="col">From → to</th><th scope="col" style="text-align:right">Moved</th></tr></thead><tbody>${rows.slice(0, CAP).map(c => {
        const out = S.diff.exclude.includes(c.id);
        return `<tr${out ? ' class="out"' : ""}><td><label class="cb"><input type="checkbox" data-x="${esc(c.id)}" ${out ? "" : "checked"} aria-label="Include in the client version: ${esc(WHAT[c.type])}, ${esc(c.name)}"></label></td>` +
          `<td class="ty">${esc(WHAT[c.type])}</td><td>${c.milestone ? "<b>" + esc(c.name) + "</b>" : esc(c.name)}${c.phase ? `<br><span style="color:var(--muted);font-size:13px">${esc(c.phase)}</span>` : ""}</td>` +
          `<td>${esc(span(c))}</td><td class="num">${c.wd == null ? "" : (c.wd > 0 ? "+" : "") + c.wd + " wd"}</td></tr>`;
      }).join("")}</tbody></table></div>${rows.length > CAP ? `<p class="hint">Showing ${CAP} of ${rows.length}. The report lists the rest.</p>` : ""}`;
    const nOut = S.diff.exclude.length;
    return head + strip + renames + table + `<p class="hint">Untick a change to leave it out of the client version${nOut ? ` (${plural(nOut, "change")} left out)` : ""}. Click the summary on the page to rewrite it.</p>`;
  }
  function toolTrend(r) {
    const vs = trendVersions(), n = vs.length, cur = curVersion();
    const dots = `<ul class="vdots" aria-label="Reports in this trend">${vs.map(v => `<li class="${v.asOf === cur.asOf ? "cur" : ""}">${esc(F(fromIso(v.asOf)))}${v.asOf === cur.asOf ? " · this plan" : v.backfilled ? " · older file" : ""}</li>`).join("")}</ul>`;
    return toolHead(r, "Slip trend", n < 3 ? `Needs 3 reports; you have ${n}.` : `${n} reports, ${esc(F(fromIso(vs[0].asOf)))} to ${esc(F(fromIso(vs[n - 1].asOf)))}.`) + dots +
      (n < 3 ? `<div class="callout warn">${icon("alert")}<div>The milestone trend, slipping phases and drift rate need 3 reports. Add older exports of this plan and Keelson places each one by its status date, or press Save version each week.
          <div class="row" style="margin-top:10px"><button class="btn fill sm" type="button" data-act="hist-add">${icon("plus")}Add older files</button></div></div></div>`
        : `<div class="callout">${icon("pencil")}<div>Every re-baseline stays in the log on the page until it has a note. Click a note to record why it happened and who agreed it; notes are kept with this project on this device.</div></div>`);
  }
  function toolForecast(r) {
    const f = forecastRes();
    const c = S.k.cols, has = [["bstart", "Baseline start"], ["bfin", "Baseline finish"], ["afin", "Actual finish"]].map(([key, label]) =>
      `<span class="chip ${c[key] >= 0 ? "chip-ok" : "chip-bad"}"><span class="dot"></span>${label}${c[key] >= 0 ? "" : " missing"}</span>`).join("");
    return toolHead(r, "Finish-date confidence", f.ok ? `P50 <b>${esc(F(f.p50))}</b> and P80 <b>${esc(F(f.p80))}</b>, from ${plural(f.n, "completed task")}. The plan date is ${esc(F(f.planFinish))}.`
      : "Keelson only forecasts when the plan's own history is enough to go on.") + `<div class="row">${has}</div>` +
      (f.ok ? `<div class="callout">${icon("shield")}<div>These dates come from the plan and can't be edited. The page explains how they're worked out, so a sponsor can check them.</div></div>`
        : `<div class="callout warn">${icon("alert")}<div>${esc(f.reason)}</div></div>`);
  }
  function paintTool() {
    const el = $("tool"), r = V2.find(x => x.id === S.report);
    el.hidden = !r;
    if (!r) { el.innerHTML = ""; return; }
    const sc = $("chg") ? $("chg").scrollTop : 0, fid = document.activeElement && el.contains(document.activeElement) ? focusKey(document.activeElement) : null;
    el.innerHTML = r.id === "changes" ? toolChanges(r) : r.id === "trend" ? toolTrend(r) : toolForecast(r);
    if ($("chg")) $("chg").scrollTop = sc;
    if (fid) el.querySelector(fid)?.focus();
  }
  const focusKey = a => a.id ? "#" + CSS.escape(a.id) : a.dataset.x ? `[data-x="${CSS.escape(a.dataset.x)}"]` : a.dataset.f ? `[data-f="${CSS.escape(a.dataset.f)}"]` : a.dataset.rn ? `[data-rn="${CSS.escape(a.dataset.rn)}"]` : null;
  const repaint = () => { paintReport(); };
  $("tool").addEventListener("click", e => {
    const f = e.target.closest("[data-f]"), rn = e.target.closest("[data-rn]"), act = e.target.closest("[data-act]");
    if (f) { S.diff.filter = f.dataset.f; paintTool(); return; }
    if (rn) {
      const key = rn.dataset.rn, a = act.dataset.act, d = S.diff;
      d.accepted = d.accepted.filter(x => x !== key); d.split = d.split.filter(x => x !== key);
      if (a === "accept") d.accepted.push(key); if (a === "split") d.split.push(key);
      repaint(); $("tool").querySelector(`[data-rn="${CSS.escape(key)}"]`)?.focus(); return;
    }
    if (!act) return;
    if (act.dataset.act === "cmp-clear") { S.cmp = null; S.diff = { accepted: [], split: [], exclude: [], filter: "all" }; repaint(); }
    if (act.dataset.act === "hist-add") openHist(true);
  });
  $("tool").addEventListener("change", e => {
    const x = e.target.closest("[data-x]");
    if (x) { const id = x.dataset.x; S.diff.exclude = x.checked ? S.diff.exclude.filter(i => i !== id) : [...S.diff.exclude, id]; repaint(); return; }
    if (e.target.id === "cmp-file") { const f = e.target.files[0]; e.target.value = ""; if (f) compareFile(f); }
    if (e.target.id === "cmp-date" && S.cmp) { const d = fromIso(e.target.value); if (d != null) { S.cmp.asOf = d; S.cmp.version = cmpVersion(S.cmp.text, d, S.cmp.name, S.cmp.sheet); repaint(); } }
  });
  ["dragenter", "dragover"].forEach(t => $("tool").addEventListener(t, e => { const z = e.target.closest("#cmp-drop"); if (z) { e.preventDefault(); z.classList.add("over"); } }));
  $("tool").addEventListener("dragleave", e => { const z = e.target.closest("#cmp-drop"); if (z && !z.contains(e.relatedTarget)) z.classList.remove("over"); });
  $("tool").addEventListener("drop", e => { const z = e.target.closest("#cmp-drop"); if (!z) return; e.preventDefault(); e.stopPropagation(); z.classList.remove("over"); if (e.dataTransfer.files[0]) compareFile(e.dataTransfer.files[0]); });
  const dayOf = ms => Math.floor((ms - new Date(ms).getTimezoneOffset() * 6e4) / DAY_MS);
  const guessAsOf = f => { const d = f.lastModified ? dayOf(f.lastModified) : null; return d != null && d < S.sd ? d : S.sd - 7; }; // the file's date, if it is older than this plan
  function cmpVersion(text, asOf, file, sheet) {
    const k = comp(text, { statusDate: asOf, msKeys: msKeys() });
    if (!(k.rows > 0 && k.cols.name >= 0 && k.cols.finish >= 0)) throw plainErr("We couldn't find task names and finish dates in that file. Check it's an export of this plan.");
    return KH.makeVersion(k, { project: projKey(), file, sheet, asOf: E.iso(asOf) });
  }
  async function compareFile(f) {
    try {
      const pp = await parsePlan(f), { text, sheet } = planText(pp), asOf = guessAsOf(f);
      S.cmp = { name: pp.file, text, sheet, asOf, version: cmpVersion(text, asOf, pp.file, sheet) };
      S.diff = { accepted: [], split: [], exclude: [], filter: "all" };
      repaint(); toast(`Comparing with ${pp.file}`); $("cmp-date")?.focus(); return true;
    } catch (e) { console.error("Keelson: compare file", e); toast(e && e.plain ? e.message : READ_FAIL); return false; }
  }

  // ---------- version history dialog ----------
  const hd = $("hist");
  const RAGC = { Red: "chip-bad", Amber: "chip-warn", Green: "chip-ok" };
  function paintHist() {
    const cur = S.k ? curVersion().asOf : null, n = S.hist.length;
    $("hist-sub").innerHTML = `<b>${esc(S.title || "This project")}</b> · ${plural(n, "saved version")}${S.histMem ? " · a sample's history, kept in memory until you save a version" : ""}`;
    const mem = KH.storeKind === "memory" && !S.histMem; $("hist-mem").hidden = !mem; $("hist-note").hidden = mem;
    $("hlist").innerHTML = n ? S.hist.slice().reverse().map(v => `<li><span class="when">${esc(F(fromIso(v.asOf)))}</span>` +
      `<span class="what"><span class="chip ${RAGC[v.rag] || "chip-mute"}"><span class="dot"></span>${esc(v.rag === "—" ? "No status" : v.rag)}</span>${v.asOf === cur ? "This plan's date · " : ""}${v.backfilled ? "Added later · " : ""}${esc(v.file || "Saved report")}</span>` +
      `<button type="button" class="icon-x" data-del="${esc(v.asOf)}" aria-label="Delete the version of ${esc(F(fromIso(v.asOf)))}">${icon("trash")}</button></li>`).join("")
      : `<li class="hempty" style="display:block">No saved versions yet. Press Save version (a Word or PowerPoint download saves one too); you can also add older exports below.</li>`;
    $("h-export").disabled = $("h-clear").disabled = !n;
  }
  function openHist(add) { if (!S.k) return; bfReset(); paintHist(); hd.showModal(); (add ? $("bf-file") : $("hist-h")).focus(); }
  document.addEventListener("click", e => { if (e.target.closest("[data-hist]")) openHist(); });
  $("hist-x").addEventListener("click", () => hd.close());
  hd.addEventListener("click", e => { if (e.target === hd) hd.close(); });
  $("hlist").addEventListener("click", async e => {
    const b = e.target.closest("[data-del]"); if (!b) return;
    try { await KH.remove(projKey(), b.dataset.del); } catch (err) { console.error(err); }
    await loadHist(); paintHist(); toast(`Deleted the version of ${F(fromIso(b.dataset.del))}`); ($("hlist").querySelector("[data-del]") || $("hist-h")).focus();
  });
  let bf = null;
  const bfMsg = (m, c = "") => { $("bf-msg").textContent = m; $("bf-msg").className = "status " + c; };
  function bfReset() { bf = null; $("bf-row").hidden = true; $("bf-t").textContent = "Add an older file"; $("bf-s").textContent = "Last month's export, or any earlier one. It's placed by its status date."; bfMsg(""); }
  $("bf-file").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    bfMsg("Reading…");
    try {
      const pp = await parsePlan(f);
      bf = { file: pp.file, ...planText(pp) };
      $("bf-t").textContent = pp.file; $("bf-s").textContent = "Pick the status date it was exported for, then add it.";
      $("bf-row").hidden = false; $("bf-date").value = E.iso(guessAsOf(f)); bfMsg(""); $("bf-date").focus();
    } catch (err) { console.error("Keelson: older file", err); bfMsg(err && err.plain ? err.message : READ_FAIL, "bad"); }
  });
  $("bf-cancel").addEventListener("click", () => { bfReset(); $("bf-file").focus(); });
  $("bf-go").addEventListener("click", async () => {
    const d = fromIso($("bf-date").value);
    if (!bf) return;
    if (d == null) { bfMsg("Pick the date this file's figures were as of.", "bad"); $("bf-date").focus(); return; }
    const k = comp(bf.text, { statusDate: d, msKeys: msKeys() });
    if (!(k.rows > 0 && k.cols.name >= 0 && k.cols.finish >= 0)) { bfMsg("We couldn't find task names and finish dates in that file.", "bad"); return; }
    try { await KH.put(projKey(), KH.fromFile(k, E.iso(d), { project: projKey(), file: bf.file, sheet: bf.sheet })); }
    catch (err) { console.error("Keelson: older file", err); bfMsg("That version couldn't be saved in this browser.", "bad"); return; }
    bfReset(); await loadHist(); paintHist();
    bfMsg(`Added the version of ${F(d)}${d >= S.sd ? ". It's not older than this plan, so What changed still compares with the one before." : "."}`, "ok"); $("bf-file").focus();
  });
  function saveBlob(blob, name) {
    const a = document.createElement("a"), url = URL.createObjectURL(blob);
    a.href = url; a.download = name; document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  $("h-export").addEventListener("click", async () => {
    try { saveBlob(await KH.exportJSON(projKey()), KH.fileName(S.title || projKey())); toast("History file saved"); }
    catch (e) { console.error("Keelson: export history", e); toast("The history file couldn't be made."); }
  });
  $("h-import").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = ""; if (!f) return;
    try {
      const pv = await KH.previewImport(f, { project: projKey() }), name = S.title || pv.project;
      if (pv.wouldExceed) { const m = `Importing would take "${name}" past ${KH.MAX} saved versions. Remove some versions or import into another project.`; bfMsg(m, "bad"); toast(m); return; }
      if (pv.mismatch && !confirm(`This file is the history of '${pv.fileProject}'. Add it to '${name}'?`)) return;
      if (pv.replaces > 0) {
        if (confirm(`This replaces ${plural(pv.replaces, "saved version")}. Export your current history first?\n\nOK saves a copy first. Cancel lets you replace them without a copy.`)) saveBlob(await KH.exportJSON(projKey()), KH.fileName(name));
        else if (!confirm(`Replace ${plural(pv.replaces, "saved version")} of '${name}' without a copy?`)) return;
      }
      const r = await KH.importJSON(f, { project: projKey() }); await loadHist(); paintHist(); toast(`Imported ${plural(r.count, "version")}${r.skipped ? ` (${r.skipped} skipped)` : ""}`);
    }
    catch (err) { console.error("Keelson: import history", err); toast(err && err.plain ? err.message : "That file couldn't be imported."); }
  });
  $("h-clear").addEventListener("click", async () => {
    if (!confirm(`Delete all ${plural(S.hist.length, "saved version")} of "${S.title}" from this browser? Export a .keelson.json file first if you might need them.`)) return;
    try { await KH.clear(projKey()); } catch (e) { console.error(e); }
    S.cmp = null; await loadHist(); paintHist(); toast("This project's history is cleared"); $("hist-h").focus();
  });
  function paintCrumb() {
    const c = $("crumb"), r = ALL.find(x => x.id === S.report);
    c.hidden = !S.k || S.view === "upload";
    if (!c.hidden) c.innerHTML = `<b>${esc(S.title || "Project")}</b>${S.view === "report" && r ? " / " + esc(r.name) : ""}`;
  }

  // ---------- views and routing (hash, so Back works) ----------
  const VIEWS = { upload: "v-upload", check: "v-check", reports: "v-reports", report: "v-report", dash: "v-dash" };
  // the Dashboard tab (studio.js): mounted for the current plan, unmounted when you leave it (its canvases animate)
  let dash = null;
  function paintDash() {
    dash?.destroy(); dash = null;
    if (!window.KeelsonStudio) { $("dash").textContent = "The dashboard didn't load. Reload the page."; return; }
    const prev = KH ? prevVersion() : null;
    dash = window.KeelsonStudio.mount($("dash"), { k: S.k, prev, title: S.title, client: S.client, screen: S.dashScreen, brand: S.brand, fc: S.k ? forecastRes() : null,
      hist: S.hist.map(v => ({ day: fromIso(v.asOf), pct: v.pct })) });
  }
  function refresh() {
    if (!S.k) return;
    if (S.view === "reports") { paintCompare(); paintGallery(); }
    if (S.view === "report") { paintCompare(); paintReport(); }
    if (S.view === "dash") paintDash();
  }
  function go(v) { const h = "#" + v; if (location.hash !== h) location.hash = h; else route(); }
  function route() {
    let [v, id] = location.hash.slice(1).split("/");
    if (!VIEWS[v]) v = "upload";
    if (v !== "upload" && !S.k) { history.replaceState(null, "", "#upload"); v = "upload"; }
    if ((v === "reports" || v === "report" || v === "dash") && (!usable() || !callOk())) { history.replaceState(null, "", "#check"); v = "check"; }
    if (v === "report") { S.report = ALL.some(r => r.id === id) ? id : REPORTS[0]?.id; }
    if (v === "dash") S.dashScreen = id;
    if (v !== "dash" && dash) { dash.destroy(); dash = null; }
    const changed = S.view !== v || v === "report", first = S.view == null;
    S.view = v; document.body.dataset.view = v;
    Object.entries(VIEWS).forEach(([k, el]) => { $(el).hidden = k !== v; });
    const order = ["upload", "check", "reports", "dash"], at = order.indexOf(v === "report" ? "reports" : v);
    document.querySelectorAll("[data-step]").forEach(a => {
      const i = order.indexOf(a.dataset.step), can = i === 0 || (i === 1 && S.k) || (i >= 2 && usable());
      a.toggleAttribute("aria-current", i === at); if (i === at) a.setAttribute("aria-current", "step");
      a.setAttribute("aria-disabled", String(!can)); a.tabIndex = can ? 0 : -1;
      a.classList.toggle("done", i < at);
      a.querySelector("i").innerHTML = i < at ? icon("check") : String(i + 1);
      a.setAttribute("aria-label", `Step ${i + 1} of ${order.length}: ${a.querySelector("span").textContent}${i < at ? ", done" : ""}`);
    });
    if (v === "check") paintCheck();
    refresh(); paintCrumb();
    if (v !== "check") toCols = false;
    if (first && v === "upload" && cta) { $("file").focus({ preventScroll: true }); $("drop").scrollIntoView({ block: "nearest" }); } // the website's main button
    if (changed && !first) { // move focus to the new view, never on first load (Tab must start at the skip link)
      const h = $(VIEWS[v]).querySelector("h1, h2");
      scrollTo({ top: 0, behavior: reduce.matches ? "auto" : "instant" });
      if (v === "check" && toCols) $("cols").querySelector("summary").focus({ preventScroll: true });
      else if (v === "check" && usable()) $("to-reports").focus({ preventScroll: true }); // Enter goes straight on: ten seconds when detection is right
      else h?.focus({ preventScroll: true });
      toCols = false;
      document.title = { upload: "Keelson Pro · Stakeholder reports from your plan", check: "Check your plan · Keelson Pro", reports: "Your reports · Keelson Pro", report: `${$("h-report").textContent} · Keelson Pro`, dash: "Dashboard · Keelson Pro" }[v];
    }
  }
  addEventListener("hashchange", route);
  const buyer = location.hash === "#pro", cta = location.hash === "#upload"; // Lemon Squeezy sends buyers back to app.html#pro after checkout
  if (buyer) history.replaceState(null, "", "#upload");
  route();
  if (buyer) openPro();

  if (missing.length) upErr(`Part of the app didn't load (${missing.join(", ")}). Reload the page; if it keeps happening, tell us at keelson.co.in.`);
  // Word and PowerPoint (Pro) load their libraries on first use; for Pro, fetch both once the page is idle so they still work with Wi-Fi off
  (window.requestIdleCallback || (f => setTimeout(f, 1500)))(() => isPro() && [window.KeelsonDocx, window.KeelsonPptx].forEach(m => m?.load().catch(() => {})));
  window.KeelsonApp = { S, load, confident, isPro, rasterize, safeName, loadHist, curVersion }; // for tests
})();
