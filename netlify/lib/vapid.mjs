// The public key is public (it's also in site/index.html), so it lives here instead of in a setting that can get mistyped.
export const PUBLIC_KEY = "BE1IU63-h9iDYqIuGcx_SMC9R25Ads4TmqfPCiCeYqefPzJhebCYgOl9yVg2y8-C-GJ7xzj2noPQeWHXj23Ceqg";
const tidy = (v) => String(v || "").trim().replace(/^["']|["']$/g, "").replace(/\s+/g, "");
export function privateKey() { return tidy(process.env.VAPID_PRIVATE_KEY); }
