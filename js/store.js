/**
 * Data layer. Both stores share one interface, so the UI never cares where
 * the homes live:
 *
 *   load()        -> Promise<Home[]>   (every home, removed ones flagged)
 *   add(home)     -> Promise<Home>
 *   remove(id)    -> Promise<void>
 *   restore(id)   -> Promise<void>
 *   live          -> true when changes are shared with everyone
 */

export class AuthError extends Error {}

/* ---------- Safe localStorage (it can be missing or throw) ---------- */
export const storage = {
  get(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw == null ? fallback : JSON.parse(raw);
    } catch {
      return fallback;
    }
  },
  set(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* private mode etc. */ }
  },
  remove(key) {
    try { localStorage.removeItem(key); } catch { /* ignore */ }
  },
};

/* ---------- Demo: starter list + changes kept in this browser ---------- */
export class DemoStore {
  live = false;
  #key = "househunt.demo.v1";

  constructor(seedUrl) {
    this.seedUrl = seedUrl;
  }

  #changes() {
    return storage.get(this.#key, { added: [], removed: {} });
  }

  async load() {
    const res = await fetch(this.seedUrl, { cache: "no-cache" });
    if (!res.ok) throw new Error(`Couldn't load the starter list (${res.status}).`);
    const seed = await res.json();
    const { added, removed } = this.#changes();
    return [...added, ...seed].map((h) => ({ ...h, removed: Boolean(removed[h.id]) }));
  }

  async add(home) {
    const changes = this.#changes();
    changes.added.unshift(home);
    storage.set(this.#key, changes);
    return home;
  }

  async remove(id) {
    const changes = this.#changes();
    changes.removed[id] = new Date().toISOString();
    storage.set(this.#key, changes);
  }

  async restore(id) {
    const changes = this.#changes();
    delete changes.removed[id];
    storage.set(this.#key, changes);
  }
}

/* ---------- Live: Google Sheet via an Apps Script web app ---------- */
export class SheetStore {
  live = true;

  /** @param {string} url  @param {() => Promise<string|null>} askPasscode */
  constructor(url, askPasscode) {
    this.url = url;
    this.askPasscode = askPasscode;
  }

  async load() {
    const res = await fetch(`${this.url}?action=list`);
    const data = await this.#read(res);
    return data.homes;
  }

  add(home) { return this.#write({ action: "add", home }).then((d) => d.home); }
  remove(id) { return this.#write({ action: "remove", id }); }
  restore(id) { return this.#write({ action: "restore", id }); }

  async #write(payload) {
    let passcode = storage.get("househunt.passcode", "");
    if (!passcode) passcode = await this.askPasscode();
    if (!passcode) throw new Error("Cancelled.");

    // text/plain keeps this a "simple" request, so the browser skips the CORS
    // preflight that Apps Script can't answer.
    const res = await fetch(this.url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ ...payload, passcode }),
    });
    try {
      const data = await this.#read(res);
      storage.set("househunt.passcode", passcode);
      return data;
    } catch (err) {
      if (err instanceof AuthError) storage.remove("househunt.passcode");
      throw err;
    }
  }

  async #read(res) {
    if (!res.ok) throw new Error(`The sheet didn't answer (${res.status}). Try again in a moment.`);
    const data = await res.json();
    if (data.ok) return data;
    if (data.error === "bad_passcode") throw new AuthError("That passcode didn't work.");
    throw new Error(data.error || "Something went wrong saving to the sheet.");
  }
}

export function createStore(config, askPasscode) {
  return config.scriptUrl
    ? new SheetStore(config.scriptUrl, askPasscode)
    : new DemoStore(config.seedUrl);
}
