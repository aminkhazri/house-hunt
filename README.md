# 🏡 Rachel, Amin & Aurora's Home Hunt

**Live page:** https://aminkhazri.github.io/house-hunt/

One bright, simple page for our family of four, with a **Top 5 Homes** strip at the top, to see the homes we picked, ranked with **#1 at the top**. It was built for readers who aren't comfortable with technology: large text, one table, and three buttons per home (**Open in Redfin** opens the Redfin listing, **Map** opens Google Maps, **Remove** takes it off our list). On phones and tablets each row becomes its own block with the same labels.

## Where the list lives

The page reads the **Homes** tab of our Google Sheet through the Apps Script web app (`apps-script/Code.gs`). A home is on the family page when its **`rank`** cell has a number and **`removed`** isn't `TRUE`. The page shows those homes sorted by rank.

| You want to… | Do this |
|---|---|
| Remove a home | Tap **Remove**, say why (saved in the Sheet's `removedReason` column), and type the family passcode. The Sheet's `removed` cell becomes `TRUE`. **Undo** or "Bring it back" reverses it. |
| Change the order | In **Top 5 Homes**, **press and hold** a home, then **drag** it to a new spot (keyboard: focus a card and use the arrow keys). Or change the `rank` numbers in the Sheet. Every move renumbers the list 1, 2, 3… |
| Add a home | Tap **＋ Add a home from Redfin**, paste the home's Redfin link, check the preview, and tap **Add to our list**. It goes to the bottom of the list. (Or add a row in the Sheet and give it a `rank`.) |
| Change the notes | Edit `why`, `watch` or `deadline` in the Sheet. |

Refresh the page to see changes.

### How "Add from Redfin" works

1. The page sends the link to the Apps Script web app (`?action=lookup&url=…`). The browser can't read Redfin itself, because of CORS, but Google's servers can.
2. `redfinLookup_()` fetches the listing and reads Redfin's page summary (`<meta name="description">`, e.g. *"For Sale: 5 beds, 3.5 baths · 2339 sq. ft. · 479 Curling Crk Dr, Clayton, NC 27527 · $385,000 · MLS# …"*), the photo (`og:image`) and the year built (`"yearBuilt":2023`).
3. The page shows a preview. **Add to our list** sends the link and the passcode to `addFromRedfin_()`, which reads Redfin again on the server (it never trusts what the browser says), gives the home the next rank, and borrows the safety grade most homes in the same zip have. If the home was removed before, it brings it back instead of making a copy.

## Excel on your computer (live link)

The web app also serves the current list as CSV at `…/exec?action=csv` (add `&include=removed` to also list removed homes). Excel can load that address and refresh it on its own:

1. In Excel: **Data → From Web**, paste the CSV address, and click **OK → Load**.
2. Right-click the table → **Table → External Data Properties → Query Properties** (on Mac: **Data → Refresh All → Connection Properties**).
3. Tick **Refresh every 5 minutes** and **Refresh data when opening the file**, then click **OK**.

Every add or remove in the Sheet or on the page then shows up in Excel automatically. Excel *reads* the list; changes happen on the page or in the Sheet.

## Links from the Redfin app

The Redfin phone app shares short links like `https://redf.in/rFnamM`, sometimes inside a sentence. The page pulls the link out of whatever was pasted, and the Apps Script follows the short link's redirects (`followRedirects: false`, reading each `Location` header) until it reaches the full redfin.com address. Tracking bits like `?utm_campaign=` are dropped before saving.

## Changing the order

In the **Top 5** or the full **All our homes** list, press and hold a home (on its number, photo or address) and drag it up or down. While you drag near the top or bottom of the screen, the page scrolls for you. Or tap the ▲ ▼ buttons under a home's number to move it one spot. Either way the Sheet's `rank` column is renumbered 1, 2, 3…

## Speed

The slow part of loading is Google's Apps Script (1.5–3 seconds per request). Three things keep the page fast:

1. **Saved copy on each device** (stale-while-revalidate): the page stores the last list in `localStorage`, shows it instantly, then asks the Sheet for the latest and redraws only if something changed. It waits if someone is mid-drag or has a box open.
2. **Saved copy in the script**: `listJson_()` keeps the finished answer in `CacheService` for 5 minutes. Every change made through the page clears it (`clearListCache_()`), and so does editing the Sheet by hand (the `onEdit` trigger).
3. **Smaller answer**: only homes that are or were on our ranked list are sent, along with their comments.

## Daily price check

Redfin has no feed we can subscribe to, so the Apps Script checks for us. `refreshPrices()` runs every morning from a **time-driven trigger** (Apps Script editor → Triggers). For each home on the page it re-reads the Redfin listing and compares it with the Sheet:

- **Price changed** → the `price` cell is updated, the old price goes in `lastPrice`, the date in `priceChangedAt`, and the page shows "▼ Dropped $5,000 · Nov 17" (or "▲ Went up").
- **Status changed** (Active / Pending / Sold / Off market) → the `status` cell is updated and the page shows a tag such as "Sale pending".
- Every change also gets a line in the **Price history** tab.

Guards: a price is only trusted while the home is for sale, a new price below half or above double the old one is ignored, and a page that doesn't look like a listing is skipped rather than read as "off the market".

## Safety grades

Each home's safety comes from **CrimeGrade's page for its zip code** (`https://crimegrade.org/safest-places-in-ZIP/`). The Apps Script reads that page, pulls out the overall grade, the safety score (0–1000, higher is safer), the "safer than X% of US neighborhoods" figure and the violent/property grades, and keeps them in Script Properties (`SAFETY_<zip>`) for a week before checking again. The page shows the grade, the score, how the area ranks against the other areas on our list, and a **CrimeGrade ↗** link. If CrimeGrade can't be reached, the page falls back to the grade in the Sheet's `safety` column.

## Comments

Under each home there's a **💬 Add a comment** button. Comments go into a **Comments** tab in the Google Sheet (one row each: home id, name, comment, time) and show as small bubbles under the home. To delete one, delete its row in the Comments tab. The Excel feed has a "Family comments" column.

## The family passcode

Only **Remove** needs the family passcode, and it asks every time. Adding a home, changing the order, commenting, **Undo** and **Bring it back** work without it. The Apps Script enforces this on the server (`body.action !== "remove"` skips the check), so changing the page can't get around it. Nothing is saved on the device.

## How the top 10 were picked

1. **Only homes still for sale.** Each listing's status and current price were checked on Redfin (Nov 14, 2026); sold, pending and off-market homes were dropped.
2. **A score out of 100:** safety grade (40), value in $ per sq ft (25), price (15), school ratings (10), and size and age (10).
3. **Our notes count too:** real problems (foundation issues, "forget it") cost 15 points, smaller worries (noise, small yard) cost 5, and pluses (corner lot, cul-de-sac, wooded yard) add a little.

---

# The explorer (portfolio version)

`/explorer/` is the full version of the project: every home we looked at, with search, filters, and adding and removing homes saved to a Google Sheet through Apps Script. The family page doesn't need any of it; it's kept here as a portfolio piece.


A small web app for a family house hunt. Reviewers can browse the shortlist, filter it, open each home on Redfin, check the area's safety map, and **add or remove homes**. All changes save to a shared Google Sheet.

**Stack:** HTML, CSS and vanilla JavaScript (ES modules), hosted free on GitHub Pages. The backend is a Google Sheet behind a Google Apps Script web app. No frameworks, no build step, no server to pay for.

---

## What it does

- **Home cards** show the price, any drop from the original price, beds/baths/sqft, HOA, the estimated mortgage, a color-coded safety grade, and school scores drawn as small bars.
- **One-tap links** on each card: the home's Redfin listing, the CrimeGrade safety map, and Google Maps. Cards show the listing photo, or a drawn house if there isn't one.
- **Search, filters and sort**: by status (for sale / pending / sold / saved), city, max price and beds; sort by price, safety, biggest price drop, size or newest build.
- **Add a home** from a form with validation (zip format, price parsing like `$385K`, safe links only).
- **Remove a home**: there's a confirm step and an *Undo* toast, and removed homes move to a "Removed homes" drawer where anyone can put them back.
- **♥ Save** favorites (kept per person, in their own browser).
- Touches like button ripples, staggered card entrances, confetti when a home is added, dark mode, a mobile layout, and support for reduced-motion settings.

## How it fits together

```
 Browser (GitHub Pages)                    Google
┌──────────────────────────┐   GET  ?action=list    ┌──────────────────────┐     ┌──────────────┐
│ index.html               │ ─────────────────────► │ Apps Script web app  │ ──► │ Google Sheet │
│ js/app.js   (UI)         │                        │ apps-script/Code.gs  │     │  "Homes" tab │
│ js/store.js (data layer) │ ◄───────────────────── │ doGet / doPost       │ ◄── │              │
│ js/utils.js (pure fns)   │  POST {action,passcode}│ + passcode check     │     └──────────────┘
└──────────────────────────┘                        └──────────────────────┘
```

| File | Job |
|---|---|
| `explorer/index.html` | Page structure, dialogs and the add form |
| `explorer/css/styles.css` | Design tokens (colors, fonts), layout, components, animations, dark mode |
| `explorer/js/config.js` | **The only file you edit**: paste your Apps Script URL here |
| `explorer/js/app.js` | State, rendering, filtering/sorting, event handling |
| `explorer/js/store.js` | Data layer: `DemoStore` (browser only) and `SheetStore` (Google Sheet), both with the same interface |
| `explorer/js/utils.js` | Small pure helpers: formatting, safe links, grade ranking, price-drop math |
| `explorer/js/effects.js` | Ripples, confetti, toasts |
| `explorer/data/homes.json` | Starter list (57 homes) used in demo mode |
| `apps-script/Code.gs` | The backend API |
| `apps-script/Seed.gs` | Optional offline copy of the starter list (otherwise `setup()` downloads `explorer/data/homes.json`) |

---

## Setup

You can test the site before connecting Google: with `scriptUrl` left empty it runs in **demo mode**, where changes save only in your browser.

### 1. Put it on GitHub Pages

1. Create a new public repo on GitHub, for example `house-hunt`.
2. Upload everything in this folder (keep the folder structure), or push it with git:
   ```bash
   git init && git add . && git commit -m "House hunt site"
   git branch -M main
   git remote add origin https://github.com/<your-username>/house-hunt.git
   git push -u origin main
   ```
3. In the repo, go to **Settings → Pages → Build and deployment**, set **Source: Deploy from a branch**, **Branch: main / (root)**, then **Save**.
4. After a minute or so the site is live at `https://<your-username>.github.io/house-hunt/` (the family page) and `/house-hunt/explorer/` (the explorer).

> **Concept: static hosting.** GitHub Pages only serves files and can't run code on its servers. That's why saving changes needs a separate backend, the Apps Script in step 2.

### 2. Turn the Google Sheet into an API

1. Open the **Rachel House Inventory** Google Sheet, then go to **Extensions → Apps Script**.
2. Rename the project "House Hunt API". Replace the contents of `Code.gs` with `apps-script/Code.gs` from this repo.
3. *(Optional)* Click **+ → Script**, name it `Seed`, and paste in `apps-script/Seed.gs`. Without it, `setup()` downloads the same starter list from `explorer/data/homes.json` in this repo.
4. Pick **setup** in the function dropdown and click **Run**. Google asks you to authorize the script: choose your account, then **Advanced → Go to House Hunt API → Allow**. (The warning appears because the script is yours and hasn't been reviewed by Google; that's normal for personal scripts.)
5. Open **Execution log**. It shows your **family passcode**, for example `porch-4821`. A new **Homes** tab now exists in the sheet with all 57 homes.
   - To change the passcode: ⚙️ **Project Settings → Script Properties → PASSCODE**.
6. Click **Deploy → New deployment → ⚙️ → Web app** and set:
   - **Execute as:** Me
   - **Who has access:** Anyone
   - Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

> **Concept: soft delete.** "Remove" never deletes a row. It sets `removed = TRUE` and a timestamp, so nothing is lost and anything can be put back from the page or the sheet.
>
> **Concept: locking.** `LockService` makes simultaneous writes wait their turn, so two reviewers adding homes at the same moment can't overwrite each other.

### 3. Connect the site to the sheet

Edit `explorer/js/config.js`:

```js
export const CONFIG = {
  scriptUrl: "https://script.google.com/macros/s/AKfy.../exec",
  ...
};
```

Commit and push. The badge above the cards changes from **Demo** to **Live**.

### 4. Share it

Send reviewers the GitHub Pages link and the passcode. The first time they add or remove a home, the page asks for the passcode and remembers it on that device.

### Updating the backend later

If you change `Code.gs`, go to **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**. The URL stays the same, so `config.js` doesn't need to change.

---

## Security notes

- **Reading** the list is open to anyone with the Apps Script URL. Don't put anything private in the sheet.
- **Writing** requires the passcode, which is checked on the server, so someone can't skip it by editing the page.
- Everything sent to the server is cleaned up there: text is trimmed and length-limited, numbers are coerced, and only `http(s)` links are accepted.
- Everything shown on the page goes through `esc()` to block HTML injection, and links go through `safeUrl()` so `javascript:` URLs never become clickable.
- The passcode keeps random visitors out. It isn't strong account security; for that you'd add Google sign-in.

## Things worth pointing out in a portfolio

- **Separation of concerns:** the UI (`app.js`) never knows where data lives. It calls `store.load() / add() / remove() / restore()`, and the demo and live stores both implement that interface (a simple strategy pattern).
- **Pure functions** in `utils.js` hold the logic that's easy to get wrong (price parsing, grade ranking, link building). They can be tested without a browser.
- **Event delegation:** one click listener on the grid handles every card's buttons, including cards added later.
- **CORS without a proxy:** writes are sent as `text/plain` so the browser treats them as "simple requests" and skips the preflight that Apps Script can't answer.
- **Accessibility:** native `<dialog>`s, labeled controls, visible focus rings, `aria-live` regions for counts and toasts, a skip link, and full support for `prefers-reduced-motion`.
- **Design tokens:** every color and font comes from CSS custom properties, which is how dark mode can be a single block.

## Ideas for later

- A comments thread or 👍/👎 votes per home.
- A map view with every home pinned.

## Search helpers (added Nov 17, 2026)

Everything below runs from one morning job, `morningCheck()`, scheduled by running `setupTriggers()` once in the Apps Script editor (it replaces the old `refreshPrices` trigger, 6–7 am). Each step runs separately, so one failing doesn't stop the others.

| Feature | Where it shows | How it works |
|---|---|---|
| **New-listing radar** | "🆕 New homes that match" cards at the top, with **＋ Add**, **Redfin** and **Not for us** | `radarCheck_()` downloads Redfin's search results as CSV (`/stingray/api/gis-csv`, the same file as Redfin's "Download all" link) for each town in Settings, filtered by our must-haves. New addresses go into the **New listings** tab. The first run only records what's already for sale, so the box starts empty except for homes listed in the last 3 days. New matches get their listing page read for a photo and the garage size, and homes with too small a garage are left out. |
| **Morning email** | Your inbox | `buildDigest_()` collects price drops and increases, status changes (pending, sold) and new matches, and `MailApp` sends them to the Sheet's owner plus the addresses in Settings. On days with no news, no email is sent, except on Thursdays and Fridays when an open house is coming up (a weekend preview). Run `sendTestEmail()` to see a sample. |
| **Might negotiate** | A yellow 🤝 badge under the address, with the reasons | The page flags homes that have been for sale about twice as long as is usual in their town (at least 60 days), have had 2+ price cuts, or cost 10%+ less per sq ft than is typical there. (Typical = the middle value of all matching listings the radar saw that morning, saved in Script Properties as `MARKET`.) |
| **Family votes** | ♥ family score under each home, "Waiting on …", and a **♥ Vote** button | One row per person per home in the **Votes** tab. Voters are listed in Settings. |
| **Must-have ticks** | Green ✓ and red ✗ under the size | Compares beds, baths, sq ft, garage and price against Settings. |
| **Monthly cost** | "≈ $2,430/mo" under the price; tap it for the breakdown | 30-year loan payment + Redfin's tax-rate and insurance estimates for that home + HOA (+ about 0.5% a year mortgage insurance under 20% down). The rate starts at today's 30-year fixed rate that Redfin shows. The down payment and rate can be changed in the breakdown, for everyone. |
| **Map and drive times** | **🗺 Show the map** (Leaflet + OpenStreetMap, loaded only when opened); "🚗 RDU Airport 32 min" under the address | The location comes from the listing page. Drive times come from Apps Script's built-in Google Maps service (`Maps.newDirectionFinder`) and are only recalculated when the places in Settings change. |
| **Tour notes** | **📝 Tour notes** button; finished tours show as green chips under the home | A checklist (Yes / No / Not sure), notes, photos (shrunk to 1600 px on the phone before upload) and voice notes (recorded in the browser with `MediaRecorder`, 3 minutes max). Files go to a Drive folder, "Home Hunt – tour photos & voice notes", shared as *anyone with the link can view* so the page can show them. Answers go in the **Tours** tab. |
| **Open houses** | "🗓 Open houses coming up", grouped by day, with **＋ Add to calendar** and **🚗 Plan the route** | Read from the listing page every morning. The route link opens Google Maps with that day's stops in time order. |
| **Bigger text** | **A+ Bigger text** in the header | Makes everything about 20% larger, and remembers the choice on that device. |
| **Home screen** | **📲 Put it on my home screen** | A web app manifest and icons, so the page opens with one tap like an app. The button shows the steps for iPhone or Android. |

### The Settings tab

The must-haves, budget, radar towns and price cap, down payment, rate, family voters, email recipients and drive-time places are in the **Settings** tab, in plain words. Change a value there and the page and the morning job use it (cached for one minute).
