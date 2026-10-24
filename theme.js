// Apply the visitor's saved theme before first paint (loaded in <head> on every site page).
try { const t = localStorage.getItem("keelson-theme"); if (t === "light" || t === "dark") document.documentElement.dataset.theme = t; } catch (e) {}
