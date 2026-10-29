// Homepage motion (home.css): the app window straightens on scroll and blocks below the first screen rise in. Everything stays still with reduced motion.
(() => {
  const still = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const app = document.getElementById("x-app");
  const tilt = () => { const r = app.getBoundingClientRect(), p = Math.min(1, Math.max(0, 1 - (r.top - 60) / (innerHeight * .7)));
    app.style.setProperty("--tilt", `${10 * (1 - p)}deg`); app.style.setProperty("--sc", .96 + .04 * p); };
  if (app && !still) { addEventListener("scroll", tilt, { passive: true }); tilt(); }
  const rv = [...document.querySelectorAll(".x-rv")];
  if (still || !("IntersectionObserver" in window)) return rv.forEach(el => el.classList.add("in"));
  rv.forEach(el => { if (el.getBoundingClientRect().top > innerHeight) el.classList.add("pre"); });
  const io = new IntersectionObserver(es => es.forEach(e => { e.target.classList.toggle("in", e.isIntersecting); if (e.isIntersecting) e.target.classList.remove("pre"); }), { threshold: .15 });
  rv.forEach(el => io.observe(el));
})();
