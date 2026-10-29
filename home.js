// Homepage motion (home.css): the app window straightens on scroll, the pinned story follows the step in view,
// and blocks below the first screen rise in. Everything stays still with reduced motion.
(() => {
  const still = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const app = document.getElementById("x-app");
  const tilt = () => { const r = app.getBoundingClientRect(), p = Math.min(1, Math.max(0, 1 - (r.top - 60) / (innerHeight * .7)));
    app.style.setProperty("--tilt", `${10 * (1 - p)}deg`); app.style.setProperty("--sc", .96 + .04 * p); };
  // pinned story: the step nearest the middle of the screen drives the visual
  const steps = [...document.querySelectorAll(".x-step")], vs = [...document.querySelectorAll(".x-v")];
  const story = () => { const mid = innerHeight / 2; let best = 0, d = Infinity;
    steps.forEach((s, i) => { const r = s.getBoundingClientRect(), k = Math.abs(r.top + r.height / 2 - mid); if (k < d) { d = k; best = i; } });
    steps.forEach((s, i) => s.classList.toggle("on", i === best)); vs.forEach((v, i) => v.classList.toggle("on", i === best)); };
  addEventListener("scroll", () => { if (app && !still) tilt(); story(); }, { passive: true }); if (app && !still) tilt(); story();
  const rv = [...document.querySelectorAll(".x-rv")];
  if (still || !("IntersectionObserver" in window)) return rv.forEach(el => el.classList.add("in"));
  rv.forEach(el => { if (el.getBoundingClientRect().top > innerHeight) el.classList.add("pre"); });
  const io = new IntersectionObserver(es => es.forEach(e => { e.target.classList.toggle("in", e.isIntersecting); if (e.isIntersecting) e.target.classList.remove("pre"); }), { threshold: .15 });
  rv.forEach(el => io.observe(el));
})();
