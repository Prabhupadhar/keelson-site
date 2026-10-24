// TEST MODE store settings, used only by app-test.html (a hidden copy of the app for test purchases). See docs/LEMON-SQUEEZY.md.
window.KEELSON_STORE = Object.freeze({
  storeId: 505056,
  proProductIds: [1468657],   // Keelson Pro, test mode
  checkout: Object.freeze({
    proMonthly: "https://keelson.lemonsqueezy.com/checkout/buy/5e87c293-f02a-4245-bebd-07ae1d3e5258", // one link; the buyer picks Monthly or Yearly
    proYearly: "https://keelson.lemonsqueezy.com/checkout/buy/5e87c293-f02a-4245-bebd-07ae1d3e5258",
    packSaas: "",
  }),
  portal: "https://app.lemonsqueezy.com/my-orders",
});
