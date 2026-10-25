/* Keelson version history: one compact Version per project per report date (format: docs/PRO-ARCHITECTURE.md section 5).
   IndexedDB db "keelson", store "versions", key [project, asOf]; in-memory when IndexedDB is unavailable (private mode).
   Nothing leaves the browser: export/import is a local .keelson.json file. Tested by tests/web/history.test.js.
   Versions keep dates, not the work calendar: What changed and trend count working days with the calendar active now
   (KeelsonEngine.setCalendar), so every saved version is compared on the same calendar. The stored headline var is as saved. */
(function (root) {
  const E = root.KeelsonEngine || (typeof require !== "undefined" ? require("./engine.js") : null);
  const MAX = 52, FORMAT = "keelson-history";
  const LIM = { bytes: 25 * 1024 * 1024, versions: 1000, tasks: 5000, text: 300, project: 200 };
  const ISO = /^\d{4}-\d{2}-\d{2}$/;

  // ---------- pure core ----------
  const str = (v, n = LIM.text) => String(v ?? "").replace(/[\u0000-\u001F\u007F￾￿]/g, " ").trim().slice(0, n);
  const D0 = E.fromYMD(1950, 1, 1), D1 = E.fromYMD(2200, 12, 31); // same plausible range as parsed plans
  const day = v => Number.isInteger(v) && v >= D0 && v <= D1 ? v : null;
  const num = v => typeof v === "number" && Number.isFinite(v) ? v : null;
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, num(v) ?? 0));
  const okIso = s => { if (typeof s !== "string" || !ISO.test(s)) return false; const d = E.fromYMD(...s.split("-").map(Number)); return E.iso(d) === s && day(d) != null; };

  const task = t => ({ key: str(t.key, 2 * LIM.text + 8), uid: str(t.uid), name: str(t.name), phase: str(t.phase), owner: str(t.owner),
    start: day(t.start), finish: day(t.finish), bfin: day(t.bfin), bstart: day(t.bstart), afin: day(t.afin),
    pct: num(t.pct) == null ? null : Math.min(1, Math.max(0, t.pct)), done: t.done === true, ms: t.ms === true });
  const cancelled = c => ({ uid: str(c.uid), name: str(c.name), phase: str(c.phase), owner: str(c.owner) });

  // a Version from engine output; meta = {project, asOf, file, sheet, backfilled, savedAt}
  function makeVersion(k, meta = {}) {
    const asOf = meta.asOf || E.iso(k.statusDate);
    if (!okIso(asOf)) throw new Error("A version needs an as-of date (YYYY-MM-DD).");
    const n = v => Math.round(clamp(v, 0, LIM.tasks)); // counts
    return { v: 1, project: str(meta.project ?? "", LIM.project), asOf, savedAt: str(meta.savedAt || new Date().toISOString(), 40),
      file: str(meta.file), sheet: str(meta.sheet), backfilled: meta.backfilled === true,
      rag: ["Red", "Amber", "Green"].includes(k.rag) ? k.rag : "—", pct: clamp(k.pct, 0, 1), finish: day(k.finish), bfinish: day(k.bfinish), var: clamp(k.var, -1e5, 1e5),
      overT: n(k.overT), overMs: n(k.overMs), nMsDone: n(k.nMsDone), nMs: n(k.nMs), nTask: n(k.nTask), nTaskDone: n(k.nTaskDone),
      idCol: !!(k.cols && k.cols.uid >= 0),
      tasks: (k.tasks || []).slice(0, LIM.tasks).map(task), cancelled: (k.cancelled || []).slice(0, LIM.tasks).map(cancelled) };
  }
  // an older file dropped in to backfill: compute k with statusDate = that file's as-of date, then tag it
  const fromFile = (k, asOf, meta = {}) => makeVersion(k, { ...meta, asOf: asOf || E.iso(k.statusDate), backfilled: true });

  // a stored or imported version, re-checked field by field (imported files are untrusted)
  function clean(x, project) {
    if (!x || typeof x !== "object" || !okIso(x.asOf) || !Array.isArray(x.tasks) || x.tasks.length > LIM.tasks) return null;
    const obj = a => Array.isArray(a) ? a.filter(t => t && typeof t === "object") : [];
    return makeVersion({ ...x, statusDate: null, cols: { uid: x.idCol ? 0 : -1 }, tasks: obj(x.tasks), cancelled: obj(x.cancelled) },
      { project, asOf: x.asOf, file: x.file, sheet: x.sheet, backfilled: x.backfilled, savedAt: typeof x.savedAt === "string" ? x.savedAt : "" });
  }

  // ---------- stores: all(project) -> records, put(rec), del(project, asOf), projects() ----------
  function memoryStore() {
    const m = new Map(); // project -> Map(asOf -> rec)
    return {
      kind: "memory",
      all: async p => [...(m.get(p) || new Map()).values()].map(r => JSON.parse(JSON.stringify(r))),
      put: async r => { if (!m.has(r.project)) m.set(r.project, new Map()); m.get(r.project).set(r.asOf, JSON.parse(JSON.stringify(r))); },
      putMany: async rs => { for (const r of rs) { if (!m.has(r.project)) m.set(r.project, new Map()); m.get(r.project).set(r.asOf, JSON.parse(JSON.stringify(r))); } },
      del: async (p, a) => { m.get(p)?.delete(a); if (m.get(p)?.size === 0) m.delete(p); },
      projects: async () => [...m.keys()],
      destroy: async () => { m.clear(); },
    };
  }
  function idbStore(idb) {
    let dbp = null;
    const db = () => dbp || (dbp = new Promise((ok, no) => {
      const rq = idb.open("keelson", 1);
      rq.onupgradeneeded = () => { if (!rq.result.objectStoreNames.contains("versions")) rq.result.createObjectStore("versions", { keyPath: ["project", "asOf"] }); };
      rq.onsuccess = () => ok(rq.result); rq.onerror = () => no(rq.error); rq.onblocked = () => no(new Error("IndexedDB blocked"));
    }));
    const go = (mode, f) => db().then(d => new Promise((ok, no) => {
      const tx = d.transaction("versions", mode), st = tx.objectStore("versions"); let out;
      const rq = f(st); if (rq) rq.onsuccess = () => { out = rq.result; };
      tx.oncomplete = () => ok(out); tx.onerror = tx.onabort = () => no(tx.error);
    }));
    const range = p => IDBKeyRange.bound([p, ""], [p, "￿"]);
    return {
      kind: "indexeddb", ready: db,
      all: p => go("readonly", st => st.getAll(range(p))),
      put: r => go("readwrite", st => st.put(r)),
      putMany: rs => go("readwrite", st => { for (const r of rs) st.put(r); }), // one transaction: all or nothing
      del: (p, a) => go("readwrite", st => st.delete([p, a])),
      projects: () => go("readonly", st => st.getAllKeys()).then(ks => [...new Set(ks.map(k => k[0]))]),
      // close our connection, then delete the whole database; rejects if blocked (another tab holds it open) or on error
      destroy: async () => {
        if (dbp) { const d = await dbp; d.close(); dbp = null; }
        await new Promise((ok, no) => { const rq = idb.deleteDatabase("keelson");
          rq.onsuccess = () => ok(); rq.onerror = () => no(rq.error); rq.onblocked = () => no(new Error("Close other Keelson tabs, then clear history again.")); });
      },
    };
  }
  let store = null;
  async function S() {
    if (store) return store;
    const idb = typeof indexedDB !== "undefined" ? indexedDB : null;
    if (idb) { const s = idbStore(idb); try { await s.ready(); return (store = s); } catch (e) { /* private mode: fall through */ } }
    return (store = memoryStore());
  }
  const use = s => { store = s || null; return api; };

  // ---------- API (all async) ----------
  const byAsOf = (a, b) => a.asOf < b.asOf ? -1 : a.asOf > b.asOf ? 1 : 0;
  const proj = p => { const s = str(p, LIM.project); if (!s) throw new Error("A project name is needed to keep its history."); return s; };
  async function list(project) { return (await (await S()).all(proj(project))).sort(byAsOf); }
  async function put(project, version) {
    const p = proj(project), s = await S(), v = clean(version, p);
    if (!v) throw new Error("That is not a Keelson version.");
    await s.put(v);
    const all = (await s.all(p)).sort(byAsOf);
    for (const old of all.slice(0, Math.max(0, all.length - MAX))) await s.del(p, old.asOf); // keep the latest 52 by as-of date
    return v;
  }
  const save = (project, k, meta = {}) => put(project, makeVersion(k, { ...meta, project }));
  async function get(project, asOf) { return (await list(project)).find(v => v.asOf === asOf) || null; }
  async function latest(project, before) { const a = (await list(project)).filter(v => !before || v.asOf < before); return a[a.length - 1] || null; }
  async function remove(project, asOf) { await (await S()).del(proj(project), asOf); }
  // clear(project): delete that project's versions. clear() with no project: delete every project's versions, close the
  // connection and delete the "keelson" database; resolves true, rejects if any step fails (so the app never toasts a false success).
  async function clear(project) {
    const s = await S();
    if (project === undefined) { await s.destroy(); store = null; return true; }
    const p = proj(project); for (const v of await s.all(p)) await s.del(p, v.asOf);
    return true;
  }
  async function projects() { return (await (await S()).projects()).sort(); }

  async function exportJSON(project) {
    const p = proj(project), body = JSON.stringify({ format: FORMAT, v: 1, project: p, versions: await list(p) });
    return new Blob([body], { type: "application/json" });
  }
  const fileName = project => (str(project, 80).replace(/[\\/:*?"<>|]+/g, "-") || "project") + ".keelson.json";

  // parse + check a history file; returns {d, p (target project), vs (clean versions, one per as-of, sorted)}
  async function readFile(fileOrText, opts) {
    const bad = m => Object.assign(new Error(m), { plain: true });
    const size = typeof fileOrText === "string" ? fileOrText.length : fileOrText && fileOrText.size;
    if (!(size > 0)) throw bad("This history file is empty.");
    if (size > LIM.bytes) throw bad("This history file is too large to be a Keelson history (over 25 MB).");
    let d; try { d = JSON.parse(typeof fileOrText === "string" ? fileOrText : await fileOrText.text()); } catch (e) { throw bad("This isn't a Keelson history file (.keelson.json): it could not be read."); }
    if (!d || d.format !== FORMAT) throw bad("This isn't a Keelson history file (.keelson.json).");
    if (d.v !== 1) throw bad("This history file was made by a newer version of Keelson. Update Keelson and try again.");
    if (!Array.isArray(d.versions) || d.versions.length > LIM.versions) throw bad("This history file has no usable versions.");
    const p = proj(opts.project || d.project);
    const ok = d.versions.map(x => clean(x, p)).filter(Boolean);
    if (!ok.length) throw bad("This history file has no usable versions.");
    const vs = [...new Map(ok.map(v => [v.asOf, v])).values()].sort(byAsOf); // a repeated as-of in the file: the later entry wins
    return { d, p, vs, skipped: d.versions.length - ok.length };
  }
  const summary = (d, p, vs, have) => {
    const old = new Set(have.map(v => v.asOf)), replaces = vs.filter(v => old.has(v.asOf)).length;
    return { fileProject: str(d.project, LIM.project), project: p, mismatch: str(d.project, LIM.project) !== p, incoming: vs.length, replaces,
      wouldExceed: old.size + vs.length - replaces > MAX };
  };

  /**
   * previewImport(fileOrText, {project}) -> {fileProject, project, mismatch, incoming, replaces, wouldExceed}. Writes nothing.
   *   fileProject: the project named in the file; project: where it would go (opts.project, else the file's own name);
   *   mismatch: the two differ ("This file is the history of 'X'. Add it to 'Y'?"); incoming: distinct as-of dates in the file;
   *   replaces: saved versions with the same as-of date that would be overwritten ("This replaces N saved versions");
   *   wouldExceed: saved + new dates would pass the 52-version cap, so importJSON will refuse.
   * Throws the same plain errors as importJSON for files that are not Keelson history.
   */
  async function previewImport(fileOrText, opts = {}) {
    const { d, p, vs } = await readFile(fileOrText, opts);
    return summary(d, p, vs, await (await S()).all(p));
  }
  // importJSON(fileOrText, {project}) -> {project, count: distinct as-of dates written, skipped: unusable entries}.
  // Same as-of dates are replaced; never prunes saved versions: if the result would pass 52 versions it rejects and writes nothing.
  async function importJSON(fileOrText, opts = {}) {
    const { d, p, vs, skipped } = await readFile(fileOrText, opts), s = await S();
    if (summary(d, p, vs, await s.all(p)).wouldExceed)
      throw Object.assign(new Error(`Importing would take "${p}" past ${MAX} saved versions. Remove some versions or import into another project.`), { plain: true });
    await s.putMany(vs);
    return { project: p, count: vs.length, skipped };
  }

  const api = { MAX, makeVersion, fromFile, clean, save, put, list, get, latest, remove, clear, projects, exportJSON, importJSON, previewImport, fileName, use, memoryStore,
    get storeKind() { return store ? store.kind : null; } };
  root.KeelsonHistory = api;
  if (typeof module !== "undefined") module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
