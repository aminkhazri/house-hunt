/**
 * One-time update from the Nov 14, 2026 Redfin check.
 * Run applyTop10() once from the editor. It:
 *   - writes rank 1–10 plus "why we like it" / "keep in mind" notes,
 *   - updates prices and status from Redfin,
 *   - clears the rank on every other home.
 * After that, the Homes tab is the list: change a rank, notes or price right
 * in the sheet and the family page shows it on the next refresh.
 */
const TOP10 = [
  { id: "41-bluejack-court", price: 360000, deadline: "Offers due by Monday, Nov 16",
    why: "The most house for the money: big home on 1.27 acres at the end of a quiet cul-de-sac, fully fenced, and no city taxes. The price has come down a lot.",
    watch: "A road (Covered Bridge Rd) runs behind the house, so check how noisy it is." },
  { id: "117-cinnabar-street", price: 364900,
    why: "Brand-new home on a corner lot, next to a top-rated elementary school (Riverwood, 10 out of 10). The builder has already lowered the price.",
    watch: "Ask what the HOA includes, such as the golf club." },
  { id: "399-florence-drive", price: 399000,
    why: "The most room on the list, with 5 bedrooms. It's on a quiet cul-de-sac with trees in the back yard.",
    watch: "Near the top of the budget, and it's an older home in the neighborhood." },
  { id: "944-walker-ridge-way", price: 379990,
    why: "Brand-new home with nice landscaping and a private back yard. It has a basement for storage or future rooms.",
    watch: "The basement is unfinished." },
  { id: "479-curling-creek-drive", price: 385000,
    why: "Almost new, with a bedroom and full bathroom on the first floor. It backs up to private woods, with a playground just up the street.",
    watch: "HOA is $82 a month." },
  { id: "273-florence-drive", price: 325000,
    why: "The lowest price on the list for a big 4-bedroom home, with trees in the back yard. The price has dropped $40,000.",
    watch: "It used to be a rental, so the kitchen cabinets look worn. Plan for some updates." },
  { id: "909-walker-ridge-way", price: 404990,
    why: "The biggest new home on the list, on the same street as #4.",
    watch: "The highest price here, and the seller hasn't lowered it since February." },
  { id: "113-cinnabar-street", price: 374900,
    why: "Brand-new home on a corner lot, a few doors from #2, with the same top-rated schools.",
    watch: "A little smaller than #2 and costs about $10,000 more." },
  { id: "112-copenhaver-drive", price: 325000,
    why: "One of the two lowest prices, for a fairly new home in a well-kept neighborhood.",
    watch: "Only 3 bedrooms, and the smallest home on the list." },
  { id: "17-wild-sage-circle", price: 377700,
    why: "Newer home on a nice corner lot, with room to fence in the yard.",
    watch: "Only 3 bedrooms, and Angier is a bit farther out." },
];

// Homes Redfin no longer shows as for sale (Nov 14, 2026).
const REDFIN_STATUS = {
  "51-tall-oak-court": "Pending", "450-windgate-drive": "Pending",
  "82-green-willows-drive": "Sold", "42-carlton-st": "Sold", "4712-forest-highland-drive": "Sold",
  "201-sarazen-drive": "Sold", "137-w-cultivator-way": "Sold", "5-e-smoketree-court": "Sold",
  "95-beacon-way": "Sold", "4503-wingate-song-court": "Sold", "650-monterey-pine-drive": "Sold",
  "140-gasper-court": "Sold", "72-montpelier-lane": "Sold",
  "118-zoai-place": "Off market", "117-hansen-run": "Off market", "219-albereto-drive": "Off market",
  "124-s-gallaspy-drive": "Off market", "306-travis-lane": "Off market", "329-florence-drive": "Off market",
};

function applyTop10() {
  const sheet = sheet_(); // also adds the rank/why/watch/deadline columns if missing
  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const c = (name) => header.indexOf(name);
  const byId = {};
  TOP10.forEach((h, i) => { byId[h.id] = Object.assign({ rank: i + 1 }, h); });

  let updated = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = String(row[c("id")]);
    const top = byId[id];
    row[c("rank")] = top ? top.rank : "";
    row[c("why")] = top ? top.why : row[c("why")];
    row[c("watch")] = top ? top.watch : row[c("watch")];
    row[c("deadline")] = top ? (top.deadline || "") : row[c("deadline")];
    if (top) { row[c("price")] = top.price; row[c("status")] = "Active"; updated++; }
    if (REDFIN_STATUS[id]) row[c("status")] = REDFIN_STATUS[id];
  }
  sheet.getRange(1, 1, values.length, header.length).setValues(values);
  Logger.log("Ranked %s homes; %s homes marked sold/pending/off market.", updated, Object.keys(REDFIN_STATUS).length);
}
