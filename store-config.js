// Lemon Squeezy store settings. Empty values keep the waitlist / notify-me emails. See docs/LEMON-SQUEEZY.md.
window.KEELSON_STORE = Object.freeze({
  openUntil: "2026-11-09", // Pro free for everyone through this date (local time); "" turns it off
  storeId: null,              // number, LS > Settings > Stores
  proProductIds: [],          // numbers: the Keelson Pro product id(s) whose keys unlock Pro
  checkout: Object.freeze({ proMonthly: "", proYearly: "", packSaas: "" }), // LS share links https://keelson.lemonsqueezy.com/checkout/buy/<uuid>
  portal: "https://app.lemonsqueezy.com/my-orders",  // customer self-service (invoices, cancel, card)
});
