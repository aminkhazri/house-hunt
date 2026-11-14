# 🏡 Rachel, Amin & Aurora's Home Hunt

**Live page:** https://aminkhazri.github.io/house-hunt/

One bright, simple page for our family of four, with a **Top 5 Homes** strip at the top, to see the homes we picked, ranked with **#1 at the top**. It was built for readers who aren't comfortable with technology: large text, one table, and three buttons per home (**See photos** opens the Redfin listing, **Map** opens Google Maps, **Remove** takes it off our list). On phones and tablets each row becomes its own block with the same labels.

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

## Changing the order

In the **Top 5** or the full **All our homes** list, press and hold a home (on its number, photo or address) and drag it up or down. While you drag near the top or bottom of the screen, the page scrolls for you. Or tap the ▲ ▼ buttons under a home's number to move it one spot. Either way the Sheet's `rank` column is renumbered 1, 2, 3…

## Comments

Under each home there's a **💬 Add a comment** button. Comments go into a **Comments** tab in the Google Sheet (one row each: home id, name, comment, time) and show as small bubbles under the home. To delete one, delete its row in the Comments tab. The Excel feed has a "Family comments" column.

## The family passcode

Every change goes through the Apps Script, which checks the passcode on the server. **Remove** and **Add** ask for it every time. **Add a comment** asks too, unless you typed it in the last 5 minutes. Dragging in the Top 5, **Undo** and **Bring it back** reuse a passcode typed in the last 5 minutes, so they don't nag right after you typed it. Nothing is saved on the device: close the page and it's forgotten.

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
