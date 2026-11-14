import { CONFIG } from "./config.js";
import { createStore, storage, AuthError } from "./store.js";
import { enableRipples, confetti, toast } from "./effects.js";
import {
  esc, money, shortMoney, priceDrop, parsePrice, gradeRank, gradeTier,
  linksFor, paletteFor, median, slug, safeUrl,
} from "./utils.js";

/* ==========================================================================
   State
   ========================================================================== */
const $ = (sel) => document.querySelector(sel);

const state = {
  homes: [],
  favorites: new Set(storage.get("househunt.favorites", [])),
  filters: { q: "", status: "Active", city: "", maxPrice: "", beds: "", sort: "price-asc" },
};

const store = createStore(CONFIG, askPasscode);

/* ==========================================================================
   Filtering & sorting (pure: takes data, returns data)
   ========================================================================== */
const SORTS = {
  "price-asc": (a, b) => (a.price ?? Infinity) - (b.price ?? Infinity),
  "price-desc": (a, b) => (b.price ?? 0) - (a.price ?? 0),
  safety: (a, b) => gradeRank(b.safety) - gradeRank(a.safety) || (a.price ?? 0) - (b.price ?? 0),
  drop: (a, b) => priceDrop(b) - priceDrop(a),
  sqft: (a, b) => (b.sqft ?? 0) - (a.sqft ?? 0),
  newest: (a, b) => (b.yearBuilt ?? 0) - (a.yearBuilt ?? 0),
};

function visibleHomes() {
  const f = state.filters;
  const q = f.q.trim().toLowerCase();
  return state.homes
    .filter((h) => !h.removed)
    .filter((h) => {
      if (f.status === "Favorites") return state.favorites.has(h.id);
      return f.status === "All" || h.status === f.status;
    })
    .filter((h) => !f.city || h.city === f.city)
    .filter((h) => !f.maxPrice || (h.price ?? 0) <= Number(f.maxPrice))
    .filter((h) => !f.beds || (h.beds ?? 0) >= Number(f.beds))
    .filter((h) => !q || [h.address, h.city, h.zip, h.neighborhood, h.notes, h.elementary, h.middle, h.high]
      .join(" ").toLowerCase().includes(q))
    .sort(SORTS[f.sort] || SORTS["price-asc"]);
}

/* ==========================================================================
   Rendering
   ========================================================================== */
function houseSvg(id) {
  const c = paletteFor(id);
  return `
    <svg class="house" viewBox="0 0 120 100" aria-hidden="true">
      <rect x="84" y="14" width="10" height="22" fill="${c.roof}"/>
      <path d="M8 50 60 10l52 40" fill="${c.roof}"/>
      <rect x="18" y="44" width="84" height="56" rx="3" fill="${c.body}"/>
      <rect x="51" y="66" width="18" height="34" rx="3" fill="${c.door}"/>
      <circle cx="65" cy="84" r="1.6" fill="#fff"/>
      <rect x="28" y="58" width="16" height="16" rx="2" fill="#fff" opacity=".85"/>
      <rect x="76" y="58" width="16" height="16" rx="2" fill="#fff" opacity=".85"/>
      <path d="M36 58v16M28 66h16M84 58v16M76 66h16" stroke="${c.body}" stroke-width="2"/>
    </svg>`;
}

const ICONS = {
  redfin: '<svg viewBox="0 0 24 24"><path d="M3 11 12 4l9 7"/><path d="M5 10v10h14V10"/></svg>',
  shield: '<svg viewBox="0 0 24 24"><path d="M12 3 4 6v6c0 5 3.5 8 8 9 4.5-1 8-4 8-9V6z"/><path d="m9 12 2 2 4-4"/></svg>',
  pin: '<svg viewBox="0 0 24 24"><path d="M12 21s-7-6.2-7-11a7 7 0 0 1 14 0c0 4.8-7 11-7 11z"/><circle cx="12" cy="10" r="2.5"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/></svg>',
  heart: '<svg viewBox="0 0 24 24"><path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.6A4 4 0 0 1 19 10c0 5.6-7 10-7 10z"/></svg>',
};

function schoolRow(label, name, score) {
  if (!name && score == null) return "";
  const v = Math.max(0, Math.min(10, Number(score) || 0));
  return `
    <div class="school" title="${esc(label)}: ${esc(name)} (${score ?? "?"}/10)">
      <span>${esc(label)} · ${esc(name || "—")}</span>
      <span class="meter"><i style="--v:${v}"></i></span>
      <b>${score ?? "–"}</b>
    </div>`;
}

function cardHtml(h, i) {
  const links = linksFor(h, CONFIG.state);
  const drop = priceDrop(h);
  const fav = state.favorites.has(h.id);
  const c = paletteFor(h.id);
  const photo = safeUrl(h.photoUrl);
  const facts = [
    h.beds != null && `${h.beds} bd`,
    h.baths != null && `${h.baths} ba`,
    h.sqft && `${Number(h.sqft).toLocaleString()} sqft`,
    h.yearBuilt && `Built ${h.yearBuilt}`,
    h.hoa != null && (Number(h.hoa) ? `HOA $${h.hoa}/mo` : "No HOA"),
    h.mortgage && `≈ ${money(h.mortgage)}/mo`,
  ].filter(Boolean);
  const notes = h.notes ? `<p class="notes">${esc(h.notes)}</p>${h.notes.length > 140 ? '<button type="button" class="more" data-act="more">Read more</button>' : ""}` : "";

  return `
  <article class="card" data-id="${esc(h.id)}" data-status="${esc(h.status)}" style="--i:${Math.min(i, 16)}; --sky:${c.sky}; --grass:color-mix(in srgb, ${c.roof} 18%, #9fd49a)">
    <div class="card__art">
      ${photo ? `<img src="${esc(photo)}" alt="Photo of ${esc(h.address)}" loading="lazy">` : `<span class="sun"></span>${houseSvg(h.id)}`}
      <span class="ribbon" data-status="${esc(h.status)}">${h.status === "Active" ? esc(h.list || "For sale") : esc(h.status)}</span>
      <button type="button" class="fav" data-act="fav" aria-pressed="${fav}" aria-label="${fav ? "Unsave" : "Save"} ${esc(h.address)}">${ICONS.heart}</button>
    </div>
    <div class="card__body">
      <div class="price-row">
        <span class="price">${money(h.price)}</span>
        ${drop ? `<span class="drop" title="Originally ${esc(h.originalPrice)}">↓ ${shortMoney(drop)}</span>` : ""}
      </div>
      <div>
        <div class="address">${esc(h.address)}</div>
        <div class="place">${esc([h.city, h.zip].filter(Boolean).join(", "))}${h.neighborhood ? " · " + esc(h.neighborhood) : ""}</div>
      </div>
      <ul class="facts">${facts.map((f) => `<li>${esc(f)}</li>`).join("")}</ul>
      <div class="safety-row">
        <span class="grade" data-tier="${gradeTier(h.safety)}">${esc(h.safety || "?")}</span>
        <span>${h.safety ? "Safety grade" : "No safety grade yet"}</span>
      </div>
      <div class="schools">
        ${schoolRow("Elem", h.elementary, h.elementaryScore)}
        ${schoolRow("Middle", h.middle, h.middleScore)}
        ${schoolRow("High", h.high, h.highScore)}
      </div>
      ${notes}
      ${h.addedBy && h.addedBy !== "Rachel's list" ? `<p class="meta">Added by ${esc(h.addedBy)}</p>` : ""}
    </div>
    <div class="card__actions">
      <a class="btn btn--primary" href="${esc(links.redfin)}" target="_blank" rel="noopener" title="${links.redfinIsDirect ? "Open the Redfin listing" : "Search Redfin for this address"}">${ICONS.redfin}Redfin</a>
      <a class="btn" href="${esc(links.safety)}" target="_blank" rel="noopener" title="Crime map for ${esc(h.city)}">${ICONS.shield}Safety</a>
      <a class="btn" href="${esc(links.map)}" target="_blank" rel="noopener" title="Open in Google Maps">${ICONS.pin}Map</a>
      <button type="button" class="btn btn--remove" data-act="remove" aria-label="Remove ${esc(h.address)}">${ICONS.trash}</button>
    </div>
  </article>`;
}

function renderStats() {
  const active = state.homes.filter((h) => !h.removed && h.status === "Active");
  const cities = new Set(active.map((h) => h.city));
  const drops = active.filter((h) => priceDrop(h) > 0).length;
  const stats = [
    ["For sale", active.length],
    ["Median price", shortMoney(median(active.map((h) => h.price)) ?? 0)],
    ["Price drops", drops],
    ["Towns", cities.size],
  ];
  $("#stats").innerHTML = stats.map(([k, v]) => `<div><dt>${k}</dt><dd>${esc(v)}</dd></div>`).join("");
}

function renderCityOptions() {
  const cities = [...new Set(state.homes.map((h) => h.city).filter(Boolean))].sort();
  const sel = $("#city");
  const current = sel.value;
  sel.innerHTML = '<option value="">All cities</option>' +
    cities.map((c) => `<option ${c === current ? "selected" : ""}>${esc(c)}</option>`).join("");
  $("#city-options").innerHTML = cities.map((c) => `<option value="${esc(c)}">`).join("");
}

function renderRemoved() {
  const removed = state.homes.filter((h) => h.removed);
  $("#removed-count").textContent = removed.length;
  $("#removed").hidden = removed.length === 0;
  $("#removed-list").innerHTML = removed.map((h) => `
    <li data-id="${esc(h.id)}">
      <span><b>${esc(h.address)}</b>, ${esc(h.city)} · ${money(h.price)}</span>
      <button type="button" class="btn" data-act="restore">Put back</button>
    </li>`).join("");
}

function render() {
  const homes = visibleHomes();
  $("#grid").innerHTML = homes.map(cardHtml).join("");
  $("#empty").hidden = homes.length > 0;
  $("#result-count").textContent = `Showing ${homes.length} of ${state.homes.filter((h) => !h.removed).length} homes`;
  renderStats();
  renderRemoved();
}

function renderMode() {
  const pill = $("#mode-pill");
  pill.dataset.live = String(store.live);
  pill.textContent = store.live
    ? "Live · changes are shared with everyone"
    : "Demo · changes are saved in this browser only";
}

/* ==========================================================================
   Actions
   ========================================================================== */
async function load() {
  $("#result-count").textContent = "Loading homes…";
  try {
    state.homes = await store.load();
    renderCityOptions();
    render();
  } catch (err) {
    $("#result-count").textContent = "";
    $("#empty").hidden = false;
    $("#empty").textContent = `Couldn't load the list: ${err.message}`;
  }
}

async function removeHome(id) {
  const home = state.homes.find((h) => h.id === id);
  if (!home) return;
  const ok = await confirmDialog(`${home.address} will move to "Removed homes" at the bottom of the page. You can put it back anytime.`);
  if (!ok) return;

  const card = document.querySelector(`.card[data-id="${CSS.escape(id)}"]`);
  try {
    await store.remove(id);
    home.removed = true;
    card?.classList.add("is-leaving");
    setTimeout(render, card ? 330 : 0);
    toast(`Removed ${home.address}`, { action: { label: "Undo", onClick: () => restoreHome(id) } });
  } catch (err) {
    toast(err instanceof AuthError ? err.message + " Try again." : err.message);
  }
}

async function restoreHome(id) {
  const home = state.homes.find((h) => h.id === id);
  if (!home) return;
  try {
    await store.restore(id);
    home.removed = false;
    render();
    flash(id);
    toast(`${home.address} is back on the list`);
  } catch (err) {
    toast(err.message);
  }
}

function toggleFavorite(id, btn) {
  state.favorites.has(id) ? state.favorites.delete(id) : state.favorites.add(id);
  storage.set("househunt.favorites", [...state.favorites]);
  const on = state.favorites.has(id);
  btn.setAttribute("aria-pressed", String(on));
  btn.classList.remove("is-bursting");
  void btn.offsetWidth; // restart the animation
  btn.classList.add("is-bursting");
  if (state.filters.status === "Favorites" && !on) render();
}

/** Scroll a card into view and give it a quick highlight. */
function flash(id) {
  const card = document.querySelector(`.card[data-id="${CSS.escape(id)}"]`);
  if (!card) return;
  card.classList.add("is-new");
  card.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => card.classList.remove("is-new"), 2200);
}

/* ---------- Add form ---------- */
function readForm(form) {
  const data = Object.fromEntries(new FormData(form));
  const errors = [];
  const price = parsePrice(data.price);
  const num = (v) => (v === "" || v == null ? null : Number(v));
  const redfinUrl = data.redfinUrl ? safeUrl(data.redfinUrl) : "";

  if (!data.address.trim()) errors.push(["address", "Add the street address."]);
  if (!data.city.trim()) errors.push(["city", "Add the city."]);
  if (!/^\d{5}$/.test(data.zip.trim())) errors.push(["zip", "Zip should be 5 digits."]);
  if (!price || price < 1000) errors.push(["price", "Add the asking price, like 385000 or $385K."]);
  if (data.redfinUrl && !redfinUrl) errors.push(["redfinUrl", "The Redfin link should start with https://"]);

  const id = slug(`${data.address}-${data.zip}`);
  if (state.homes.some((h) => h.id === id || slug(h.address) === slug(data.address) && h.zip === data.zip.trim())) {
    errors.push(["address", "That home is already on the list."]);
  }

  const by = data.addedBy.trim();
  if (by) storage.set("househunt.name", by);

  return {
    errors,
    home: {
      id,
      address: data.address.trim(),
      city: data.city.trim(),
      zip: data.zip.trim(),
      county: "",
      neighborhood: data.neighborhood.trim(),
      price,
      originalPrice: "",
      status: "Active",
      mls: "",
      daysOnMarket: "",
      mortgage: null,
      safety: data.safety,
      yearBuilt: num(data.yearBuilt),
      sqft: num(data.sqft),
      beds: num(data.beds),
      baths: num(data.baths),
      hoa: num(data.hoa),
      notes: data.notes.trim(),
      elementary: "", elementaryScore: null,
      middle: "", middleScore: null,
      high: "", highScore: null,
      list: "Added " + new Date().toLocaleDateString("en-US", { month: "short", day: "numeric" }),
      redfinUrl,
      photoUrl: "",
      addedBy: by || "A reviewer",
      addedAt: new Date().toISOString(),
      removed: false,
    },
  };
}

async function submitAdd(e) {
  e.preventDefault();
  const form = e.currentTarget;
  form.querySelectorAll("[aria-invalid]").forEach((el) => el.removeAttribute("aria-invalid"));
  const { errors, home } = readForm(form);
  if (errors.length) {
    errors.forEach(([name]) => form.elements[name]?.setAttribute("aria-invalid", "true"));
    $("#add-error").textContent = errors[0][1];
    form.elements[errors[0][0]]?.focus();
    return;
  }

  const submit = form.querySelector('[type="submit"]');
  submit.disabled = true;
  submit.textContent = "Adding…";
  try {
    const saved = await store.add(home);
    state.homes.unshift({ ...home, ...saved, removed: false });
    $("#add-dialog").close();
    form.reset();
    // Make sure the new home is visible, then celebrate it.
    Object.assign(state.filters, { status: "Active", q: "", city: "", maxPrice: "", beds: "" });
    syncControls();
    renderCityOptions();
    render();
    flash(home.id);
    const fab = $("#add-btn").getBoundingClientRect();
    confetti(fab.left + fab.width / 2, fab.top);
    toast(`Added ${home.address}`);
  } catch (err) {
    $("#add-error").textContent = err.message;
  } finally {
    submit.disabled = false;
    submit.textContent = "Add to list";
  }
}

/* ---------- Dialog helpers ---------- */
function confirmDialog(text) {
  const dlg = $("#confirm-dialog");
  $("#confirm-text").textContent = text;
  dlg.returnValue = "";
  dlg.showModal();
  return new Promise((resolve) => {
    dlg.addEventListener("close", () => resolve(dlg.returnValue === "ok"), { once: true });
  });
}

function askPasscode() {
  const dlg = $("#pass-dialog");
  const form = $("#pass-form");
  $("#pass-error").textContent = "";
  form.reset();
  dlg.showModal();
  return new Promise((resolve) => {
    const onSubmit = (e) => {
      e.preventDefault();
      const value = $("#pass-input").value.trim();
      if (!value) { $("#pass-error").textContent = "Type the passcode first."; return; }
      dlg.close();
      resolve(value);
    };
    form.addEventListener("submit", onSubmit);
    dlg.addEventListener("close", () => { form.removeEventListener("submit", onSubmit); resolve(null); }, { once: true });
  });
}

/* ==========================================================================
   Wiring
   ========================================================================== */
function syncControls() {
  const f = state.filters;
  $("#q").value = f.q;
  $("#city").value = f.city;
  $("#max-price").value = f.maxPrice;
  $("#beds").value = f.beds;
  $("#sort").value = f.sort;
  document.querySelectorAll("#status-filter button").forEach((b) =>
    b.setAttribute("aria-checked", String(b.dataset.status === f.status)));
}

function bindEvents() {
  let typing;
  $("#q").addEventListener("input", (e) => {
    clearTimeout(typing);
    typing = setTimeout(() => { state.filters.q = e.target.value; render(); }, 120);
  });
  for (const [id, key] of [["#city", "city"], ["#max-price", "maxPrice"], ["#beds", "beds"], ["#sort", "sort"]]) {
    $(id).addEventListener("change", (e) => { state.filters[key] = e.target.value; render(); });
  }
  $("#status-filter").addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;
    state.filters.status = btn.dataset.status;
    syncControls();
    render();
  });

  // One listener for every card button (event delegation).
  $("#grid").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-act]");
    if (!btn) return;
    const id = btn.closest(".card").dataset.id;
    if (btn.dataset.act === "remove") removeHome(id);
    if (btn.dataset.act === "fav") toggleFavorite(id, btn);
    if (btn.dataset.act === "more") {
      const notes = btn.previousElementSibling;
      notes.classList.toggle("is-open");
      btn.textContent = notes.classList.contains("is-open") ? "Show less" : "Read more";
    }
  });
  // If a listing photo fails to load (moved, blocked, offline), fall back to the drawn house.
  $("#grid").addEventListener("error", (e) => {
    if (e.target.tagName !== "IMG") return;
    const art = e.target.closest(".card__art");
    const id = e.target.closest(".card")?.dataset.id;
    e.target.replaceWith(document.createRange().createContextualFragment(`<span class="sun"></span>${houseSvg(id)}`));
    art?.classList.add("no-photo");
  }, true);

  $("#removed-list").addEventListener("click", (e) => {
    const btn = e.target.closest('[data-act="restore"]');
    if (btn) restoreHome(btn.closest("li").dataset.id);
  });

  $("#add-btn").addEventListener("click", () => {
    $("#add-error").textContent = "";
    $("#f-by").value = storage.get("househunt.name", "");
    $("#add-dialog").showModal();
    $("#f-address").focus();
  });
  $("#add-form").addEventListener("submit", submitAdd);

  // Any [data-close] button closes its dialog; clicking the backdrop does too.
  document.querySelectorAll("dialog").forEach((dlg) => {
    dlg.addEventListener("click", (e) => {
      if (e.target.closest("[data-close]") || e.target === dlg) dlg.close();
    });
  });
}

enableRipples();
bindEvents();
renderMode();
load();
