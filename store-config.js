// Lemon Squeezy store settings. Empty values keep the waitlist / notify-me emails. See docs/LEMON-SQUEEZY.md.
window.KEELSON_STORE = Object.freeze({
  openUntil: "2026-11-09", // Pro free for everyone through this date (local time); "" turns it off
  storeId: 505056,            // number, LS > Settings > Stores (same store in test and live)
  proProductIds: [1483975],   // Keelson Pro, live mode (never the test product 1468657)
  checkout: Object.freeze({ proMonthly: "https://keelson.lemonsqueezy.com/checkout/buy/acf0d5ff-3d83-4278-98f6-b4f08def2935", proYearly: "https://keelson.lemonsqueezy.com/checkout/buy/acf0d5ff-3d83-4278-98f6-b4f08def2935", packSaas: "https://keelson.lemonsqueezy.com/checkout/buy/d65f24fb-0244-4ca2-a405-70aa69c8b93b" }), // one Pro link; the buyer picks Monthly or Yearly // LS share links https://keelson.lemonsqueezy.com/checkout/buy/<uuid>
  portal: "https://app.lemonsqueezy.com/my-orders",  // customer self-service (invoices, cancel, card)
});
