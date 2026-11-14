/** Small, pure helpers. No DOM access here, so they are easy to test. */

/** Escape text before putting it inside HTML. Every user-supplied value goes through this. */
export function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[ch]);
}

/** Only allow http(s) links, so nobody can sneak in a javascript: URL. */
export function safeUrl(value) {
  try {
    const url = new URL(String(value || "").trim());
    return url.protocol === "https:" || url.protocol === "http:" ? url.href : "";
  } catch {
    return "";
  }
}

export const money = (n) =>
  n == null || Number.isNaN(n) ? "—" : "$" + Math.round(n).toLocaleString("en-US");

export const shortMoney = (n) =>
  n >= 1000 ? "$" + Math.round(n / 1000) + "K" : money(n);

/** "$400K 9/17" -> 400000. Returns null for "Original Price" or "$383K Org.". */
export function parseOriginal(text) {
  const m = String(text || "").match(/\$\s*([\d.]+)\s*K/i);
  return m ? Math.round(parseFloat(m[1]) * 1000) : null;
}

/** How much the asking price has dropped from the original, in dollars (0 if none). */
export function priceDrop(home) {
  const original = parseOriginal(home.originalPrice);
  return original && home.price && original > home.price ? original - home.price : 0;
}

/** Turn "$385,000" or "385k" into 385000. */
export function parsePrice(text) {
  const s = String(text || "").toLowerCase().replace(/[$,\s]/g, "");
  if (!s) return null;
  const n = parseFloat(s);
  if (Number.isNaN(n)) return null;
  return Math.round(s.endsWith("k") ? n * 1000 : n);
}

const GRADES = ["F", "D-", "D", "D+", "C-", "C", "C+", "B-", "B", "B+", "A-", "A", "A+"];
/** Letter grade -> sortable number (A+ = 12). Unknown grades sort last. */
export const gradeRank = (g) => GRADES.indexOf(String(g || "").trim().toUpperCase());
export const gradeTier = (g) => {
  const r = gradeRank(g);
  if (r < 0) return "none";
  return r >= 10 ? "good" : r >= 7 ? "warn" : "bad"; // A range, B range, C and below
};

export const slug = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");

/** Links every card gets. A real Redfin URL wins over the search fallback. */
export function linksFor(home, state = "NC") {
  const place = `${home.address}, ${home.city}, ${state} ${home.zip}`;
  return {
    redfin: safeUrl(home.redfinUrl) ||
      "https://www.google.com/search?q=" + encodeURIComponent(`redfin ${place}`),
    redfinIsDirect: Boolean(safeUrl(home.redfinUrl)),
    safety: `https://crimegrade.org/safest-places-in-${slug(home.city)}-${state.toLowerCase()}/`,
    map: "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(place),
  };
}

/** Deterministic pastel colors per home, so each little house looks a bit different. */
export function paletteFor(id) {
  let h = 0;
  for (const ch of String(id)) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
  const hue = h % 360;
  return {
    body: `hsl(${hue} 62% 82%)`,
    roof: `hsl(${(hue + 160) % 360} 45% 42%)`,
    door: `hsl(${(hue + 40) % 360} 55% 55%)`,
    sky: `hsl(${(hue + 190) % 360} 70% 90%)`,
  };
}

export function median(nums) {
  const a = nums.filter((n) => typeof n === "number").sort((x, y) => x - y);
  if (!a.length) return null;
  const mid = a.length >> 1;
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}
