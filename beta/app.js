/**
 * Zeekr electricity tracker - application logic.
 *
 * A single-page app with no build step and no third-party libraries.
 * The page is static (GitHub Pages); all user data lives in one JSON file,
 * `zeekr-data.json`, inside the app's own folder in the user's OneDrive.
 *
 * File layout
 *   index.html  - markup only
 *   styles.css  - presentation only
 *   app.js      - this file
 *
 * Sections in this file
 *   1. Configuration
 *   2. Application state
 *   3. Generic helpers (DOM, dates, numbers, formatting)
 *   4. Data model (validation, normalisation, serialisation)
 *   5. Calculations (periods, efficiency, totals, calendar months)
 *   6. Charts (hand-drawn SVG)
 *   7. Rendering
 *   8. Entry forms
 *   9. Import and export (JSON + CSV, merge instead of replace)
 *  10. OneDrive storage (OAuth 2.0 PKCE + Microsoft Graph)
 *  11. Start-up
 *
 * Data model (also the shape of zeekr-data.json)
 *   {
 *     app: "zeekr-tracker", version: 2, updated: ISO timestamp,
 *     readings: [{ date, meter, odo, price, soc }],
 *     public:   [{ id, date, kwh, cost, provider, note }],
 *     settings: { fuelPrice, fuelL100, batteryKwh },
 *     loan:     { buyDate, carPrice, altPrice, schedules: [...], ... }  (see section 5b)
 *   }
 *
 *   reading - one look at the home charger's meter.
 *     date   "YYYY-MM-DD". Unique: it is the reading's identity.
 *     meter  cumulative kWh shown on the charger.
 *     odo    car odometer in km, or null when it was not noted.
 *     price  home price per kWh (ILS) for the period that ENDS at this
 *            reading. Storing it per reading is what makes a price change
 *            apply from that reading onwards and never retroactively.
 *     soc    battery charge in percent (0-100) shown in the car at the
 *            reading, or null. Optional; when both ends of a span have it,
 *            the change in stored energy corrects that span's efficiency.
 *
 *   public charge - one charging session away from home.
 *     id        opaque unique string.
 *     date      "YYYY-MM-DD".
 *     kwh       energy charged.
 *     cost      total amount paid for the session (ILS).
 *     provider  free text: the company that operates the charger.
 *     note      free text.
 *
 *   Public charges are a separate list. Adding one never touches a reading;
 *   it is simply counted in whichever period its date falls into.
 */
(function () {
  "use strict";

  /* ======================================================================
   * 1. Configuration
   * ==================================================================== */

  /** Shown in Settings; bump together with the ?v= query in index.html. */
  const APP_VERSION = "3.8";

  /** Application (client) ID of the Microsoft Entra app registration. Public, not a secret. */
  const CLIENT_ID = "66f6788f-925a-471b-b98f-87d662b76db6";

  /** Microsoft identity platform endpoint for personal accounts. */
  const AUTH_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/";

  /** Least privilege: the app can only see its own folder (OneDrive/Apps/Zeekr Tracker). */
  const SCOPE = "Files.ReadWrite.AppFolder offline_access";

  /**
   * Must match a redirect URI registered for the app, character for character.
   * Only the folder addresses are registered (…/Zeekr_Tracker/ and
   * …/Zeekr_Tracker/beta/), so a page opened as …/index.html still signs in
   * through its folder address.
   */
  const REDIRECT_URI = location.origin + location.pathname.replace(/index\.html$/, "");

  /**
   * Two copies of this app are published from one repository:
   *   /Zeekr_Tracker/        the main version, working on the real readings
   *   /Zeekr_Tracker/beta/   the trial version, where changes are tried first
   * The trial version always works against a separate file, so nothing done
   * there can touch the real readings. Which copy this is follows from the
   * address alone, so it cannot be switched by mistake.
   */
  const IS_BETA = location.pathname.includes("/beta/");
  const REAL_FILE_NAME = "zeekr-data.json";
  const DATA_FILE_NAME = IS_BETA ? "zeekr-test.json" : REAL_FILE_NAME;

  /** Microsoft Graph address of a file inside the app folder (OneDrive/Apps/Zeekr Tracker). */
  const APP_FOLDER_URL = "https://graph.microsoft.com/v1.0/me/drive/special/approot:/";
  /** The file this copy of the app reads and writes. */
  const FILE_URL = APP_FOLDER_URL + DATA_FILE_NAME;

  /**
   * localStorage / sessionStorage keys. Both copies share one origin, so they
   * share the sign-in token (sign in once) but keep separate offline caches.
   */
  const CACHE_KEYS = { main: "zk-data", beta: "zk-data-test" };
  const STORE = { token: "zk-tok", cache: IS_BETA ? CACHE_KEYS.beta : CACHE_KEYS.main, tab: "zk-tab", pkce: "zk-pkce" };

  const DATA_FORMAT_VERSION = 2;
  /** batteryKwh: Zeekr X, 66 kWh per the manufacturer's specification page. */
  const DEFAULT_SETTINGS = { fuelPrice: 7.3, fuelL100: 7.5, batteryKwh: 66 };
  /** Sanity limits for the battery capacity setting, in kWh. */
  const MIN_BATTERY_KWH = 10, MAX_BATTERY_KWH = 200;
  const MS_PER_DAY = 864e5;
  /** A reading older than this many days turns the status pill into a reminder. */
  const STALE_AFTER_DAYS = 35;
  /** A month needs at least this many days with odometer data to get its own efficiency figure. */
  const MIN_DAYS_FOR_MONTH_EFFICIENCY = 7;
  /** Charts show at most this many most-recent periods. */
  const CHART_MAX_POINTS = 14;
  /**
   * Daily mean temperature for the consumption-versus-temperature chart, from
   * Open-Meteo (free, no key). Only these fixed coordinates are sent: Beit
   * Yehoshua, where the car is kept. Kept in localStorage, never in the data file.
   */
  const WEATHER = {
    place: "בית יהושע",
    latitude: 32.262,
    longitude: 34.863,
    archiveUrl: "https://archive-api.open-meteo.com/v1/archive",
    recentUrl: "https://api.open-meteo.com/v1/forecast",
    storeKey: "zk-temp",
  };
  /**
   * Official petrol prices by effective date, kept up to date by a GitHub
   * workflow (see scripts/update_fuel_prices.py). Shared by both copies of
   * the app, so the trial copy reads it from the folder above.
   */
  const FUEL_PRICES_URL = (IS_BETA ? "../" : "") + "data/fuel-prices.json";
  const FUEL_PRICES_KEY = "zk-fuel-prices";
  /**
   * Charging companies offered in the public-charge form, in Ofek's order;
   * the first is the default. `name` is what gets saved. `aliases` let names
   * typed before the list existed (or in English) count as the same company;
   * the saved names themselves are never rewritten.
   * Symbols: the companies' official app icons (Paz: the icon on paz.co.il).
   */
  const PROVIDERS = [
    { id: "paz", name: "פז", aliases: ["paz", "פז charge", "פזcharge", "yellow"] },
    { id: "sonol", name: "סונול", aliases: ["sonol", "סונול evi", "sonol evi", "evi"] },
    { id: "on", name: "ON EV", aliases: ["on", "on-ev", "onev", "און", "אפקון", "afcon", "afcon ev"] },
    { id: "delek", name: "דלק", aliases: ["delek"] },
    { id: "gnrgy", name: "ג׳ינרג׳י", aliases: ["gnrgy", "gnrgy go", "ג'ינרג'י", "גינרגי", "ג׳ינרגי"] },
    { id: "evedge", name: "אי.וי אדג׳", aliases: ["ev edge", "evedge", "אי וי אדג'", "איוי אדג"] },
    { id: "evmeter", name: "אי.וי מטר", aliases: ["ev meter", "evmeter", "אי וי מטר", "איוי מטר"] },
    { id: "tesla", name: "טסלה", aliases: ["tesla", "supercharger", "טסלה סופרצ'רג'ר", "סופרצ'רג'ר"] },
  ];
  const OTHER_PROVIDER = "other";
  const providerIcon = (id) => "icons/providers/" + id + ".png?v=1";

  /** A month needs this many days with temperature data to appear in the chart. */
  const MIN_DAYS_FOR_MONTH_TEMP = 7;
  /** Sanity limit for a home price per kWh, in ILS. */
  const MAX_PRICE_PER_KWH = 5;

  /* ======================================================================
   * 2. Application state
   * ==================================================================== */

  const state = {
    /** @type {Reading[]} sorted by date */
    readings: [],
    /** @type {PublicCharge[]} sorted by date */
    publicCharges: [],
    settings: { ...DEFAULT_SETTINGS },
    /** Inputs of the car-financing comparison (section 5b). */
    loan: null,
    /** True once the OneDrive file was loaded, so saving is safe. */
    connected: false,
    loading: false,
    /** eTag of the OneDrive file as last seen; sent back to detect concurrent edits. */
    etag: null,
  };

  /**
   * @typedef {{date:string, meter:number, odo:(number|null), price:(number|null), soc:(number|null)}} Reading
   * @typedef {{id:string, date:string, kwh:number, cost:number, provider:string, note:string}} PublicCharge
   * @typedef {{readings:Reading[], publicCharges:PublicCharge[], settings:{fuelPrice:number, fuelL100:number, batteryKwh:number}}} AppData
   */

  /* ======================================================================
   * 3. Generic helpers
   * ==================================================================== */

  /** Shorthand for document.getElementById. */
  const $ = (id) => document.getElementById(id);

  /** Escapes text for safe insertion into HTML. */
  function escapeHtml(value) {
    const map = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };
    return String(value ?? "").replace(/[&<>"']/g, (ch) => map[ch]);
  }

  /** Safe JSON wrapper around localStorage (it can throw in private windows). */
  const storage = {
    get(key) {
      try { return JSON.parse(localStorage.getItem(key)); } catch (e) { return null; }
    },
    set(key, value) {
      try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* storage unavailable */ }
    },
    remove(key) {
      try { localStorage.removeItem(key); } catch (e) { /* storage unavailable */ }
    },
  };

  /** Converts "YYYY-MM-DD" to a UTC timestamp. */
  const toTime = (isoDate) => Date.parse(isoDate + "T00:00:00Z");

  /** Whole days from one ISO date to another. */
  const daysBetween = (from, to) => Math.round((toTime(to) - toTime(from)) / MS_PER_DAY);

  /** Today's local date as "YYYY-MM-DD". */
  function todayIso() {
    const d = new Date();
    const pad = (v) => String(v).padStart(2, "0");
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }

  /** "2026-09-02" -> "2.9.26" */
  function formatDate(isoDate) {
    const [y, m, d] = isoDate.split("-");
    return +d + "." + +m + "." + y.slice(2);
  }

  /** Formats a number for display; missing or invalid values become a dash. */
  function formatNumber(value, decimals = 0) {
    if (value == null || !isFinite(value)) return "–";
    return value.toLocaleString("he-IL", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  }

  const byDate = (a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0);

  /**
   * Parses a date written by a person or by Excel.
   * Accepts YYYY-MM-DD and day-first forms: D/M/YYYY, D.M.YYYY, D-M-YYYY (2-digit years mean 20YY).
   * @returns {string|null} ISO date, or null when the text is not a real date.
   */
  function parseDate(text) {
    const s = String(text ?? "").trim();
    let y, m, d;
    let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (match) {
      [, y, m, d] = match.map(Number);
    } else {
      match = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})$/.exec(s);
      if (!match) return null;
      [, d, m, y] = match.map(Number);
      if (y < 100) y += 2000;
    }
    const date = new Date(Date.UTC(y, m - 1, d));
    const isReal = date.getUTCFullYear() === y && date.getUTCMonth() === m - 1 && date.getUTCDate() === d;
    return isReal ? date.toISOString().slice(0, 10) : null;
  }

  /**
   * Parses a number that may come from a spreadsheet: tolerates spaces, the
   * shekel sign, thousands separators and a decimal comma.
   * @returns {number|null} null for an empty cell, NaN for text that is not a number.
   */
  function parseNumber(value) {
    if (typeof value === "number") return value;
    let s = String(value ?? "").replace(/[\s₪]/g, "");
    if (s === "") return null;
    if (s.includes(",") && s.includes(".")) s = s.replace(/,/g, "");
    else if (/^-?\d{1,3}(,\d{3})+$/.test(s)) s = s.replace(/,/g, "");
    else s = s.replace(",", ".");
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
  }

  /* ======================================================================
   * 4. Data model
   * ==================================================================== */

  let idCounter = 0;
  /** Creates a unique id for a new public charge. */
  function newChargeId() {
    idCounter += 1;
    return "p" + Date.now().toString(36) + idCounter.toString(36);
  }

  /**
   * Validates one raw reading (from the form, a JSON file or a CSV row).
   * @returns {{value: Reading}|{error: string}} error is a Hebrew message for the user.
   */
  function normalizeReading(raw) {
    const date = parseDate(raw && raw.date);
    const meter = parseNumber(raw && raw.meter);
    const odo = parseNumber(raw && raw.odo);
    const price = parseNumber(raw && raw.price);
    const rawSoc = raw && raw.soc;
    const soc = parseNumber(typeof rawSoc === "string" ? rawSoc.replace("%", "") : rawSoc);
    if (!date) return { error: "תאריך חסר או לא תקין" };
    if (meter == null || !isFinite(meter) || meter < 0) return { error: "קריאת מונה חסרה או לא תקינה" };
    if (odo != null && (!isFinite(odo) || odo < 0)) return { error: "קילומטראז׳ לא תקין" };
    if (price != null && (!isFinite(price) || price < 0 || price > MAX_PRICE_PER_KWH)) return { error: "מחיר לקוט״ש לא תקין" };
    if (soc != null && (!isFinite(soc) || soc < 0 || soc > 100)) return { error: "אחוז סוללה צריך להיות בין 0 ל‑100" };
    return { value: { date, meter, odo, price: price > 0 ? price : null, soc } };
  }

  /**
   * Validates one raw public charge.
   * @returns {{value: PublicCharge}|{error: string}}
   */
  function normalizePublicCharge(raw) {
    const date = parseDate(raw && raw.date);
    const kwh = parseNumber(raw && raw.kwh);
    const cost = parseNumber(raw && raw.cost);
    if (!date) return { error: "תאריך חסר או לא תקין" };
    if (kwh == null || !isFinite(kwh) || kwh <= 0) return { error: "קוט״ש חסר או לא תקין" };
    if (cost == null || !isFinite(cost) || cost < 0) return { error: "עלות חסרה או לא תקינה" };
    return {
      value: {
        id: raw.id != null && String(raw.id).trim() !== "" ? String(raw.id).trim() : "",
        date, kwh, cost,
        provider: String(raw.provider ?? "").trim().slice(0, 60),
        note: String(raw.note ?? "").trim().slice(0, 120),
      },
    };
  }

  /**
   * Turns the content of a data file (or the cache) into clean application data.
   * Invalid entries are dropped; older files without `provider` are accepted.
   * @returns {AppData}
   */
  function normalizeData(raw) {
    const source = raw || {};
    const readingsByDate = new Map();
    for (const item of source.readings || []) {
      const result = normalizeReading(item);
      if (result.value) readingsByDate.set(result.value.date, result.value);
    }
    const publicCharges = [];
    const usedIds = new Set();
    for (const item of source.public || source.publicCharges || []) {
      const result = normalizePublicCharge(item);
      if (!result.value) continue;
      if (!result.value.id || usedIds.has(result.value.id)) result.value.id = newChargeId();
      usedIds.add(result.value.id);
      publicCharges.push(result.value);
    }
    const settings = { ...DEFAULT_SETTINGS };
    const rawSettings = source.settings || {};
    if (rawSettings.fuelPrice > 0) settings.fuelPrice = Number(rawSettings.fuelPrice);
    if (rawSettings.fuelL100 > 0) settings.fuelL100 = Number(rawSettings.fuelL100);
    const battery = Number(rawSettings.batteryKwh);
    if (battery >= MIN_BATTERY_KWH && battery <= MAX_BATTERY_KWH) settings.batteryKwh = battery;
    return {
      readings: [...readingsByDate.values()].sort(byDate),
      publicCharges: publicCharges.sort(byDate),
      settings,
      loan: normalizeLoan(source.loan),
    };
  }

  /** Converts application data to the JSON structure stored in zeekr-data.json, entries ordered by date. */
  function serializeData(data) {
    data = { ...data, readings: [...data.readings].sort(byDate), publicCharges: [...data.publicCharges].sort(byDate) };
    return {
      app: "zeekr-tracker",
      version: DATA_FORMAT_VERSION,
      updated: new Date().toISOString(),
      readings: data.readings.map((r) => ({ date: r.date, meter: r.meter, odo: r.odo ?? null, price: r.price ?? null, soc: r.soc ?? null })),
      public: data.publicCharges.map((p) => ({ id: p.id, date: p.date, kwh: p.kwh, cost: p.cost, provider: p.provider || "", note: p.note || "" })),
      settings: { ...data.settings },
      loan: data.loan ? { ...data.loan, schedules: (data.loan.schedules || []).map((s) => ({ ...s })) } : null,
    };
  }

  /** A detached copy of the current data, safe to modify before saving. */
  function snapshot() {
    return {
      readings: state.readings.map((r) => ({ ...r })),
      publicCharges: state.publicCharges.map((p) => ({ ...p })),
      settings: { ...state.settings },
      loan: { ...state.loan, schedules: state.loan.schedules.map((s) => ({ ...s })) },
    };
  }

  /** Replaces the in-memory data and refreshes the offline cache. */
  function applyData(data) {
    const clean = normalizeData(serializeData(data));
    state.readings = clean.readings;
    state.publicCharges = clean.publicCharges;
    state.settings = clean.settings;
    state.loan = clean.loan;
    storage.set(STORE.cache, serializeData(clean));
  }

  /**
   * The home price per kWh to use for the period ending at readings[index].
   * A reading without a price inherits the nearest earlier price, and failing
   * that the nearest later one, so an imported row with an empty price cell
   * still gets a sensible cost.
   */
  function priceForReading(readings, index) {
    for (let i = index; i >= 0; i--) if (readings[i].price > 0) return readings[i].price;
    for (let i = index + 1; i < readings.length; i++) if (readings[i].price > 0) return readings[i].price;
    return 0;
  }

  /* ======================================================================
   * 5. Calculations
   * ==================================================================== */

  /**
   * Derives everything the screens show from the raw readings and charges.
   *
   * period - the stretch between two consecutive readings. Its home energy is
   *   the meter difference; its public energy and cost are the public charges
   *   dated after the opening reading and up to the closing one.
   *
   * span - one or more periods joined so that both ends have an odometer
   *   value. Efficiency (kWh per 100 km) can only be computed per span, which
   *   is how a reading with a missing odometer is handled.
   *
   * Battery correction: the meter shows energy that went INTO the car, not
   * energy used for driving. When both ends of a span have a battery
   * percentage, the change in stored energy (percent change x capacity) is
   * taken off the charged energy to get the energy used (kwhUsed). Costs are
   * never corrected: what was paid was paid.
   *
   * @param {Reading[]} readings sorted by date
   * @param {PublicCharge[]} publicCharges
   * @param {number} batteryKwh usable battery capacity
   */
  function buildModel(readings, publicCharges, batteryKwh) {
    const periods = [];
    for (let i = 1; i < readings.length; i++) {
      const open = readings[i - 1];
      const close = readings[i];
      const charges = publicCharges.filter((p) => p.date > open.date && p.date <= close.date);
      const kwhHome = close.meter - open.meter;
      const price = priceForReading(readings, i);
      periods.push({
        from: open.date,
        to: close.date,
        days: daysBetween(open.date, close.date),
        kwhHome,
        kwhPublic: charges.reduce((sum, p) => sum + p.kwh, 0),
        price,
        costHome: kwhHome * price,
        costPublic: charges.reduce((sum, p) => sum + p.cost, 0),
        km: open.odo != null && close.odo != null ? close.odo - open.odo : null,
        efficiency: null,   // filled in below when a span closes on this period
        mergedCount: 0,     // how many periods the efficiency figure covers
        kmMerged: null,     // km of the whole span when mergedCount > 1
      });
    }

    const spans = [];
    let anchor = readings.findIndex((r) => r.odo != null);
    if (anchor >= 0) {
      let kwh = 0, cost = 0, count = 0;
      for (let i = anchor + 1; i < readings.length; i++) {
        const period = periods[i - 1];
        kwh += period.kwhHome + period.kwhPublic;
        cost += period.costHome + period.costPublic;
        count += 1;
        if (readings[i].odo == null) continue;
        const km = readings[i].odo - readings[anchor].odo;
        if (km > 0) {
          const socFrom = readings[anchor].soc, socTo = readings[i].soc;
          const socAdjust = socFrom != null && socTo != null ? -((socTo - socFrom) / 100) * batteryKwh : 0;
          const kwhUsed = Math.max(0, kwh + socAdjust);
          const span = {
            from: readings[anchor].date,
            to: readings[i].date,
            days: daysBetween(readings[anchor].date, readings[i].date),
            km, kwh, cost,
            socAdjust: kwhUsed - kwh,   // kWh added to (or taken off) the charged energy
            kwhUsed,
            efficiency: (kwhUsed / km) * 100,
          };
          spans.push(span);
          period.efficiency = span.efficiency;
          period.mergedCount = count;
          if (count > 1) period.kmMerged = km;
        }
        anchor = i;
        kwh = 0; cost = 0; count = 0;
      }
    }

    /** Totals over spans: the basis for per-km figures. */
    const driven = { km: 0, kwh: 0, kwhUsed: 0, cost: 0, days: 0, corrected: 0 };
    for (const s of spans) {
      driven.km += s.km; driven.kwh += s.kwh; driven.kwhUsed += s.kwhUsed; driven.cost += s.cost; driven.days += s.days;
      if (s.socAdjust) driven.corrected += 1;
    }

    /** Totals over every closed period: the basis for money and energy figures. */
    const spent = { home: 0, public: 0, total: 0, days: 0, kwh: 0 };
    for (const p of periods) {
      spent.home += p.costHome;
      spent.public += p.costPublic;
      spent.days += p.days;
      spent.kwh += p.kwhHome + p.kwhPublic;
    }
    spent.total = spent.home + spent.public;

    /**
     * Public charges that no closed period covers yet: those dated after the
     * latest reading (or all of them while there is no reading). They are
     * reported separately and deliberately kept out of `spent`: the home
     * energy for the same stretch is still unknown, so adding them would make
     * the per-day averages and the yearly forecast count cost without days.
     * They join a period automatically when the next reading is saved.
     */
    const lastDate = readings.length ? readings[readings.length - 1].date : null;
    const pending = { count: 0, kwh: 0, cost: 0 };
    for (const p of publicCharges) {
      if (lastDate && p.date <= lastDate) continue;
      pending.count += 1; pending.kwh += p.kwh; pending.cost += p.cost;
    }

    return { periods, spans, driven, spent, pending };
  }

  const MONTH_NAMES = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
  const MONTH_SHORT = ["ינו׳", "פבר׳", "מרץ", "אפר׳", "מאי", "יוני", "יולי", "אוג׳", "ספט׳", "אוק׳", "נוב׳", "דצמ׳"];

  /** "2026-03" -> "מרץ 2026" */
  function monthName(key) {
    const [y, m] = key.split("-");
    return MONTH_NAMES[+m - 1] + " " + y;
  }

  /** "2026-03" -> "מרץ", with the year added on January and on the first label of a chart. */
  function monthLabel(key, withYear) {
    const [y, m] = key.split("-");
    return MONTH_SHORT[+m - 1] + (withYear || m === "01" ? " " + y.slice(2) : "");
  }

  /** First day of the month after the given ISO date. */
  function startOfNextMonth(isoDate) {
    const [y, m] = isoDate.split("-").map(Number);
    return m === 12 ? (y + 1) + "-01-01" : y + "-" + String(m + 1).padStart(2, "0") + "-01";
  }

  /** The ISO date one day earlier. */
  const dayBefore = (isoDate) => new Date(toTime(isoDate) - MS_PER_DAY).toISOString().slice(0, 10);

  /**
   * Turns periods into calendar months.
   *
   * A period runs from the day of its opening reading up to the day before its
   * closing reading. Its home energy, home cost and distance are spread evenly
   * over those days, and each day is counted in the month it falls in. So a
   * reading taken on the 2nd or the 6th of a month does not shift usage into
   * the wrong month, and two readings in one month simply give that month two
   * pieces that add up.
   *
   * Distance comes from spans (stretches with an odometer value at both ends),
   * spread over the span's days. Efficiency for a month uses only the days that
   * a span covers, and is shown only when there are enough of them.
   *
   * Public charges are not spread: each is counted in the month of its date.
   * A charge dated on a closing reading that falls on the 1st, when that month
   * has no covered days yet, goes to the month before, where its period's days are.
   *
   * @returns {{key:string, days:number, daysInMonth:number, partial:boolean,
   *   kwhHome:number, kwhPublic:number, costHome:number, costPublic:number,
   *   km:number, odoDays:number, efficiency:(number|null)}[]} oldest first
   */
  function buildMonths(readings, publicCharges, model) {
    const { periods, spans } = model;
    const months = new Map();
    const monthOf = (isoDate) => isoDate.slice(0, 7);
    const get = (key) => {
      if (!months.has(key)) months.set(key, { key, days: 0, odoDays: 0, kwhHome: 0, kwhPublic: 0, costHome: 0, costPublic: 0, km: 0, effKwh: 0 });
      return months.get(key);
    };

    for (const period of periods) {
      if (period.days <= 0) continue;
      const span = spans.find((s) => s.from <= period.from && period.to <= s.to);
      const kwhPerDay = period.kwhHome / period.days;
      const costPerDay = period.costHome / period.days;
      const kmPerDay = span ? span.km / span.days : 0;
      const adjustPerDay = span && span.days ? span.socAdjust / span.days : 0;
      // Walk the period one month-slice at a time.
      for (let cursor = period.from; cursor < period.to;) {
        const next = startOfNextMonth(cursor);
        const end = next < period.to ? next : period.to;
        const days = daysBetween(cursor, end);
        const month = get(monthOf(cursor));
        month.days += days;
        month.kwhHome += kwhPerDay * days;
        month.costHome += costPerDay * days;
        if (span) {
          month.odoDays += days;
          month.km += kmPerDay * days;
          month.effKwh += (kwhPerDay + adjustPerDay) * days;
        }
        cursor = end;
      }
    }

    if (readings.length > 1) {
      const first = readings[0].date;
      const last = readings[readings.length - 1].date;
      for (const charge of publicCharges) {
        if (charge.date <= first || charge.date > last) continue; // outside tracking, or still pending
        let key = monthOf(charge.date);
        if (!months.has(key)) key = monthOf(dayBefore(charge.date));
        const month = get(key);
        month.kwhPublic += charge.kwh;
        month.costPublic += charge.cost;
        if (spans.some((s) => charge.date > s.from && charge.date <= s.to)) month.effKwh += charge.kwh;
      }
    }

    return [...months.values()]
      .sort((a, b) => (a.key < b.key ? -1 : 1))
      .map((m) => {
        const [y, mm] = m.key.split("-").map(Number);
        const daysInMonth = new Date(Date.UTC(y, mm, 0)).getUTCDate();
        const efficiency = m.km > 0 && m.odoDays >= MIN_DAYS_FOR_MONTH_EFFICIENCY ? (m.effKwh / m.km) * 100 : null;
        return { ...m, daysInMonth, partial: m.days < daysInMonth, efficiency };
      });
  }

  /** Average price paid per kWh in a month (home and public together), or null. */
  function monthPricePerKwh(m) {
    const kwh = m.kwhHome + m.kwhPublic;
    return kwh > 0 ? (m.costHome + m.costPublic) / kwh : null;
  }

  /**
   * Money saved against the petrol car in one month: the petrol for the
   * month's km at that month's petrol price, minus the electricity used for
   * them at the month's average price per kWh. Null when the month has no km.
   */
  function monthSaving(m) {
    if (!(m.km > 0)) return null;
    const price = monthPricePerKwh(m);
    return price == null ? null : m.km * monthFuelPerKm(m.key) - m.effKwh * price;
  }

  /**
   * Calendar years built from the months.
   * @returns {{year:string, days:number, daysInYear:number, partial:boolean, km:number,
   *   kwhHome:number, kwhPublic:number, costHome:number, costPublic:number,
   *   efficiency:(number|null), saving:number}[]} oldest first
   */
  function buildYears(months) {
    const years = new Map();
    for (const m of months) {
      const key = m.key.slice(0, 4);
      if (!years.has(key)) years.set(key, { year: key, days: 0, km: 0, effKwh: 0, kwhHome: 0, kwhPublic: 0, costHome: 0, costPublic: 0, saving: 0 });
      const y = years.get(key);
      y.days += m.days; y.km += m.km; y.effKwh += m.effKwh;
      y.kwhHome += m.kwhHome; y.kwhPublic += m.kwhPublic; y.costHome += m.costHome; y.costPublic += m.costPublic;
      y.saving += monthSaving(m) || 0;
    }
    return [...years.values()].map((y) => {
      const daysInYear = daysBetween(y.year + "-01-01", (+y.year + 1) + "-01-01");
      return { ...y, daysInYear, partial: y.days < daysInYear, efficiency: y.km > 0 ? (y.effKwh / y.km) * 100 : null };
    });
  }

  /** The listed company a saved name belongs to, or null. */
  function findProvider(name) {
    const key = simplify(name);
    if (!key) return null;
    return PROVIDERS.find((p) => simplify(p.name) === key || p.aliases.some((a) => simplify(a) === key)) || null;
  }

  /** A company name with its symbol in front, for tables. */
  function providerLabel(name) {
    const provider = findProvider(name);
    if (!provider) return escapeHtml(name);
    return `<span class="prov"><img src="${providerIcon(provider.id)}" alt="" width="20" height="20">${escapeHtml(provider.name)}</span>`;
  }

  /**
   * Public charges grouped by charging company. Names that differ only in
   * spaces or letter case are one company, shown with its most used spelling.
   * Every public charge is counted, including those not yet in a period.
   * @returns {{name:string, count:number, kwh:number, cost:number}[]} biggest spend first
   */
  function buildProviders(publicCharges) {
    const groups = new Map();
    for (const p of publicCharges) {
      const listed = findProvider(p.provider);
      const name = listed ? listed.name : (p.provider || "").replace(/\s+/g, " ").trim();
      const key = listed ? "#" + listed.id : name.toLowerCase();
      if (!groups.has(key)) groups.set(key, { spellings: new Map(), count: 0, kwh: 0, cost: 0 });
      const g = groups.get(key);
      g.spellings.set(name, (g.spellings.get(name) || 0) + 1);
      g.count += 1; g.kwh += p.kwh; g.cost += p.cost;
    }
    return [...groups.values()].map((g) => {
      const name = [...g.spellings.entries()].sort((a, b) => b[1] - a[1])[0][0];
      return { name, count: g.count, kwh: g.kwh, cost: g.cost };
    }).sort((a, b) => b.cost - a.cost);
  }

  /* ---------- Official petrol prices ---------- */

  /** [{from: "YYYY-MM-DD", price}] sorted by date; empty until loaded. */
  let fuelPrices = sanitizeFuelPrices(storage.get(FUEL_PRICES_KEY));

  function sanitizeFuelPrices(raw) {
    const list = (raw && Array.isArray(raw.prices) ? raw.prices : [])
      .filter((p) => p && parseDate(p.from) && Number(p.price) > 0)
      .map((p) => ({ from: parseDate(p.from), price: Number(p.price) }));
    return list.sort((a, b) => (a.from < b.from ? -1 : 1));
  }

  /** Loads the price list (fresh when online, the stored copy otherwise). */
  async function loadFuelPrices(onDone) {
    try {
      const response = await fetch(FUEL_PRICES_URL, { cache: "no-cache" });
      if (!response.ok) return;
      const json = await response.json();
      const list = sanitizeFuelPrices(json);
      if (!list.length) return;
      storage.set(FUEL_PRICES_KEY, json);
      fuelPrices = list;
      onDone();
    } catch (err) { /* offline: keep the stored copy */ }
  }

  /**
   * Petrol price per liter on a date: the official price in effect then, or
   * the price in Settings for dates before the list starts.
   */
  function fuelPriceOn(isoDate) {
    let price = null;
    for (const p of fuelPrices) { if (p.from <= isoDate) price = p.price; else break; }
    return price ?? state.settings.fuelPrice;
  }

  /** The month's petrol price, averaged over its days (prices can change mid-month). */
  function monthFuelPrice(monthKey) {
    const [y, m] = monthKey.split("-").map(Number);
    const days = new Date(Date.UTC(y, m, 0)).getUTCDate();
    let sum = 0;
    for (let d = 1; d <= days; d++) sum += fuelPriceOn(monthKey + "-" + String(d).padStart(2, "0"));
    return sum / days;
  }

  /** Petrol cost per km in a month for the comparison car. */
  const monthFuelPerKm = (monthKey) => (state.settings.fuelL100 / 100) * monthFuelPrice(monthKey);

  /* ---------- Weather (daily mean temperature) ---------- */

  const weather = {
    /** { "YYYY-MM-DD": °C } */
    days: (storage.get(WEATHER.storeKey) || {}).days || {},
    fetchedOn: (storage.get(WEATHER.storeKey) || {}).fetchedOn || null,
    from: (storage.get(WEATHER.storeKey) || {}).from || null,
    /** "idle" | "loading" | "failed" */
    status: "idle",
  };

  async function fetchDailyTemps(url, params) {
    const query = new URLSearchParams({
      latitude: WEATHER.latitude, longitude: WEATHER.longitude,
      daily: "temperature_2m_mean", timezone: "Asia/Jerusalem", ...params,
    });
    const response = await fetch(url + "?" + query);
    if (!response.ok) throw new Error("weather " + response.status);
    const json = await response.json();
    const result = {};
    const dates = (json.daily && json.daily.time) || [];
    const temps = (json.daily && json.daily.temperature_2m_mean) || [];
    dates.forEach((d, i) => { if (temps[i] != null) result[d] = temps[i]; });
    return result;
  }

  /**
   * Makes sure the temperatures from `from` to today are loaded, at most once
   * a day. The archive lags a few days behind, so the last weeks come from
   * the forecast service's recent past. Calls onDone when new data arrived.
   */
  async function ensureWeather(from, onDone) {
    const today = todayIso();
    if (weather.status === "loading") return;
    if (weather.fetchedOn === today && weather.from && weather.from <= from) return;
    weather.status = "loading";
    try {
      const archiveEnd = new Date(toTime(today) - 7 * MS_PER_DAY).toISOString().slice(0, 10);
      const days = { ...weather.days };
      if (from <= archiveEnd) Object.assign(days, await fetchDailyTemps(WEATHER.archiveUrl, { start_date: from, end_date: archiveEnd }));
      Object.assign(days, await fetchDailyTemps(WEATHER.recentUrl, { past_days: 30, forecast_days: 1 }));
      for (const d of Object.keys(days)) if (d > today) delete days[d];
      weather.days = days;
      weather.fetchedOn = today;
      weather.from = from;
      weather.status = "idle";
      storage.set(WEATHER.storeKey, { days, fetchedOn: today, from });
      onDone();
    } catch (err) {
      weather.status = "failed";
      onDone();
    }
  }

  /**
   * Mean temperature over the days of a month that the readings cover.
   * @returns {number|null}
   */
  function monthTemperature(monthKey, firstDate, lastDate) {
    let sum = 0, count = 0;
    for (const [date, temp] of Object.entries(weather.days)) {
      if (date.slice(0, 7) !== monthKey || date < firstDate || date >= lastDate) continue;
      sum += temp; count += 1;
    }
    return count >= MIN_DAYS_FOR_MONTH_TEMP ? sum / count : null;
  }

  /* ======================================================================
   * 5b. Car financing: was the loan worth it?
   *
   * Compares two paths from the day the Zeekr was bought:
   *   Zeekr path  - the Zeekr, paid with own cash, a loan against the
   *                 provident fund (kupat gemel) and the Bituach Leumi
   *                 standing loan; running costs are the real electricity
   *                 from this app.
   *   Alternative - a second-hand petrol car paid in cash; its fuel is the
   *                 same km at that month's official petrol price.
   *
   * Every shekel paid out could have stayed in the provident fund instead, so
   * every outflow is grown at the fund's yearly return up to the month shown
   * ("lost return"). The cost of a path at month t is
   *   Zeekr:  grown outflows + loan balance + standing loan owed - car value
   *   Alt.:   grown outflows - car value
   * and the advantage of the Zeekr is the alternative's cost minus the Zeekr's.
   *
   * Loan schedules: the provident-fund loan is prime-linked, so each change in
   * the rate is a new schedule (from date, balance, yearly rate, months left);
   * the monthly payment of each schedule is a regular annuity (Spitzer).
   * ==================================================================== */

  /** Upper bound for the comparison horizon, in months. */
  const LOAN_MAX_MONTHS = 180;
  const DEFAULT_LOAN = {
    buyDate: "", carPrice: null, carDepPct: null, carIns: null, carMaint: null,
    altPrice: 80000, altL100: null, altDepPct: null, altIns: null, altMaint: null,
    returnPct: null, standAmount: null, standDropPct: null, standFreeMonths: 42,
    kmMonth: null, schedules: [],
  };
  const LOAN_NUMBERS = ["carPrice", "carDepPct", "carIns", "carMaint", "altPrice", "altL100", "altDepPct", "altIns", "altMaint",
    "returnPct", "standAmount", "standDropPct", "standFreeMonths", "kmMonth"];

  /** Cleans the `loan` object of the data file; unknown or invalid values become null. */
  function normalizeLoan(raw) {
    const loan = { ...DEFAULT_LOAN, schedules: [] };
    if (!raw || typeof raw !== "object") return loan;
    if (typeof raw.buyDate === "string" && /^\d{4}-\d{2}-\d{2}$/.test(raw.buyDate)) loan.buyDate = raw.buyDate;
    for (const key of LOAN_NUMBERS) {
      if (raw[key] === null || raw[key] === "" || raw[key] === undefined) { loan[key] = key in raw ? null : DEFAULT_LOAN[key]; continue; }
      const n = Number(raw[key]);
      loan[key] = isFinite(n) && n >= 0 ? n : null;
    }
    for (const s of Array.isArray(raw.schedules) ? raw.schedules : []) {
      const from = typeof s.from === "string" && /^\d{4}-\d{2}-\d{2}$/.test(s.from) ? s.from : null;
      const balance = Number(s.balance), ratePct = Number(s.ratePct), months = Math.round(Number(s.months));
      const payment = s.payment == null || s.payment === "" ? null : Number(s.payment);
      if (from && balance > 0 && ratePct >= 0 && ratePct < 50 && months > 0 && months <= 600 && (payment == null || payment > 0)) {
        loan.schedules.push(payment ? { from, balance, ratePct, months, payment } : { from, balance, ratePct, months });
      }
    }
    loan.schedules.sort((a, b) => (a.from < b.from ? -1 : a.from > b.from ? 1 : 0));
    return loan;
  }

  /** The ISO date n months after the given one; the day is kept or clamped to the month's last day. */
  function addMonths(isoDate, n) {
    const [y, m, d] = isoDate.split("-").map(Number);
    const target = new Date(Date.UTC(y, m - 1 + n, 1));
    const last = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
    target.setUTCDate(Math.min(d, last));
    return target.toISOString().slice(0, 10);
  }

  /** Whole months from one ISO date to a later one. */
  function monthsBetween(from, to) {
    const [y1, m1, d1] = from.split("-").map(Number);
    const [y2, m2, d2] = to.split("-").map(Number);
    return (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
  }

  /**
   * Monthly payment of a schedule: the fixed payment when one is given (a grace
   * period, where mostly interest is paid), otherwise a regular (Spitzer) loan.
   */
  const schedulePayment = (s) => (s.payment > 0 ? s.payment : annuityPayment(s.balance, s.ratePct, s.months));

  /** Balance of a schedule after k of its payments; with a fixed payment, what is left at the end stays owed. */
  function scheduleBalance(s, k) {
    if (!(s.payment > 0)) return annuityBalance(s.balance, s.ratePct, s.months, k);
    k = Math.min(k, s.months);
    const r = s.ratePct / 100 / 12, g = Math.pow(1 + r, k);
    return Math.max(0, r ? s.balance * g - (s.payment * (g - 1)) / r : s.balance - s.payment * k);
  }

  /** Monthly payment of a regular (Spitzer) loan. */
  function annuityPayment(balance, yearlyPct, months) {
    const r = yearlyPct / 100 / 12;
    return r ? (balance * r) / (1 - Math.pow(1 + r, -months)) : balance / months;
  }

  /** Balance left after k payments of a regular loan. */
  function annuityBalance(balance, yearlyPct, months, k) {
    if (k >= months) return 0;
    const r = yearlyPct / 100 / 12;
    const p = annuityPayment(balance, yearlyPct, months);
    return r ? balance * Math.pow(1 + r, k) - (p * (Math.pow(1 + r, k) - 1)) / r : balance - p * k;
  }

  /** Inputs the comparison cannot run without, as Hebrew field names. */
  function loanMissing(loan) {
    const missing = [];
    if (!loan.buyDate) missing.push("תאריך הקנייה");
    if (!(loan.carPrice > 0)) missing.push("מחיר הזיקר");
    if (!(loan.altPrice > 0)) missing.push("מחיר הרכב החלופי");
    if (!loan.schedules.length) missing.push("לוח סילוקין אחד לפחות");
    return missing;
  }

  /**
   * Month-by-month comparison from the purchase to the later of today and the
   * end of the last loan schedule.
   * @param {object} loan normalised loan inputs
   * @param {object[]} months calendar months from buildMonths
   * @param {string} today ISO date
   */
  function buildLoanModel(loan, months, today) {
    const missing = loanMissing(loan);
    if (missing.length) return { missing };

    // Driving and energy per month: the real month when it is fully measured, otherwise the average.
    const byKey = new Map(months.map((m) => [m.key, m]));
    const measured = months.filter((m) => !m.partial && m.km > 0);
    const sum = (list, f) => list.reduce((s, m) => s + f(m), 0);
    const measuredKm = sum(measured, (m) => m.km);
    const costPerKm = measuredKm ? sum(measured, (m) => m.costHome + m.costPublic) / measuredKm : null;
    const kmAvg = measured.length ? measuredKm / measured.length : null;
    const kmEstimate = loan.kmMonth > 0 ? loan.kmMonth : kmAvg;

    const schedules = loan.schedules;
    const principal0 = schedules[0].balance;
    const ownCash = loan.carPrice - (loan.standAmount || 0) - principal0;
    const loanEnd = schedules.reduce((end, s) => { const e = addMonths(s.from, s.months); return e > end ? e : end; }, loan.buyDate);
    const todayIndex = Math.max(0, monthsBetween(loan.buyDate, today));
    const horizon = Math.min(LOAN_MAX_MONTHS, Math.max(todayIndex, monthsBetween(loan.buyDate, loanEnd), 12));
    const growth = 1 + (loan.returnPct || 0) / 100 / 12;
    const depZ = 1 - (loan.carDepPct || 0) / 100, depA = 1 - (loan.altDepPct || 0) / 100;
    const standDrop = 1 - (loan.standDropPct || 0) / 100;

    /** Loan state at a date: the active schedule, payments made under it, and the balance. */
    const loanAt = (date) => {
      let active = null;
      for (const s of schedules) if (s.from <= date) active = s;
      if (!active) return { payment: 0, interest: 0, balance: principal0, schedule: null };
      const k = monthsBetween(active.from, date);
      const balance = scheduleBalance(active, k);
      if (k < 1 || k > active.months) return { payment: 0, interest: 0, balance, schedule: active };
      const before = scheduleBalance(active, k - 1);
      return { payment: schedulePayment(active), interest: before * (active.ratePct / 100 / 12), balance, schedule: active };
    };

    const z = { fv: ownCash, out: ownCash, payments: 0, interest: 0, energy: 0, ins: 0, maint: 0 };
    const a = { fv: loan.altPrice, out: loan.altPrice, fuel: 0, ins: 0, maint: 0 };
    const rows = [];
    let km = 0, estimated = 0;
    for (let i = 0; i <= horizon; i++) {
      const date = addMonths(loan.buyDate, i);
      const key = date.slice(0, 7);
      const state_ = loanAt(date);
      let row = { i, date, key, past: i <= todayIndex, km: 0, energy: 0, fuel: 0, payment: 0, interest: 0, estimate: false };
      if (i > 0) {
        z.fv *= growth; a.fv *= growth;
        const m = byKey.get(key);
        const real = m && !m.partial && m.km > 0;
        const monthKm = real ? m.km : kmEstimate || 0;
        const energy = real ? m.costHome + m.costPublic : monthKm * (costPerKm || 0);
        const fuel = monthKm * ((loan.altL100 || state.settings.fuelL100) / 100) * monthFuelPrice(key);
        const insZ = (loan.carIns || 0) / 12, maintZ = (loan.carMaint || 0) / 12;
        const insA = (loan.altIns || 0) / 12, maintA = (loan.altMaint || 0) / 12;
        const outZ = state_.payment + energy + insZ + maintZ;
        const outA = fuel + insA + maintA;
        z.fv += outZ; z.out += outZ; z.payments += state_.payment; z.interest += state_.interest;
        z.energy += energy; z.ins += insZ; z.maint += maintZ;
        a.fv += outA; a.out += outA; a.fuel += fuel; a.ins += insA; a.maint += maintA;
        if (i <= todayIndex) { km += monthKm; if (!real) estimated += 1; }
        row = { ...row, km: monthKm, energy, fuel, payment: state_.payment, interest: state_.interest, estimate: !real };
      }
      const valueZ = loan.carPrice * Math.pow(depZ, i / 12);
      const valueA = loan.altPrice * Math.pow(depA, i / 12);
      const standOwed = loan.standFreeMonths > 0 && i >= loan.standFreeMonths ? 0 : (loan.standAmount || 0) * Math.pow(standDrop, i / 12);
      const costZ = z.fv + state_.balance + standOwed - valueZ;
      const costA = a.fv - valueA;
      row = { ...row, balance: state_.balance, standOwed, valueZ, valueA, costZ, costA, advantage: costA - costZ, rate: state_.schedule ? state_.schedule.ratePct : null };
      // Components of the totals at this month (they add up to costZ and costA exactly).
      row.partsZ = {
        depreciation: loan.carPrice - valueZ,
        interest: z.payments + state_.balance - principal0,   // interest and any fee added on a new schedule
        energy: z.energy, ins: z.ins, maint: z.maint,
        lostReturn: z.fv - z.out,
        standing: standOwed - (loan.standAmount || 0),         // negative: the part of the standing loan that is no longer owed
      };
      row.partsA = { depreciation: loan.altPrice - valueA, fuel: a.fuel, ins: a.ins, maint: a.maint, lostReturn: a.fv - a.out };
      rows.push(row);
    }

    // The break-even month: from here to the end of the horizon the Zeekr stays ahead.
    let breakEven = null;
    for (let i = rows.length - 1; i >= 1 && rows[i].advantage >= 0; i--) {
      if (rows[i - 1].advantage < 0) { breakEven = rows[i]; break; }
    }
    return {
      missing, rows, todayIndex, horizon, breakEven, ownCash, principal0, loanEnd, costPerKm, kmAvg, kmEstimate,
      kmSoFar: km, estimatedMonths: estimated, firstPayment: schedulePayment(schedules[0]),
      now: rows[Math.min(todayIndex, rows.length - 1)], end: rows[rows.length - 1],
    };
  }

  /* ======================================================================
   * 6. Charts
   * ==================================================================== */

  const CHART_HEIGHT = 210;

  function svgOpen(width, height) {
    return `<svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" role="img" font-family="Assistant,sans-serif">`;
  }

  /** Horizontal grid line with its value label on the left axis. */
  function gridLine(x1, x2, y, label) {
    return `<line x1="${x1}" x2="${x2}" y1="${y}" y2="${y}" stroke="var(--line)" stroke-width="1"/>` +
      `<text x="${x1 - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${label}</text>`;
  }

  /**
   * Stacked bar chart: home cost at the bottom, public cost on top.
   * A faded bar marks a month the readings do not fully cover.
   * @param {HTMLElement} el
   * @param {{label:string, home:number, pub:number, faded?:boolean}[]} items
   */
  function drawCostChart(el, items) {
    if (!items.length) { el.innerHTML = '<p class="empty">יופיע כאן אחרי שתי מדידות.</p>'; return; }
    items = items.slice(-CHART_MAX_POINTS);
    const width = Math.max(280, el.clientWidth || 600);
    const left = 38, bottom = 26, top = 18;
    const plotW = width - left - 8, plotH = CHART_HEIGHT - bottom - top;
    const max = Math.max(...items.map((it) => it.home + it.pub)) * 1.12 || 1;
    const step = plotW / items.length;
    const barW = Math.min(46, step * 0.62);
    const baseY = top + plotH;

    let svg = svgOpen(width, CHART_HEIGHT);
    for (let g = 0; g <= 3; g++) svg += gridLine(left, width - 8, baseY - (plotH * g) / 3, Math.round((max * g) / 3));
    items.forEach((it, i) => {
      const x = left + step * i + (step - barW) / 2;
      const hHome = (plotH * it.home) / max;
      const hPub = (plotH * it.pub) / max;
      const faded = it.faded ? ' opacity="0.45"' : "";
      svg += `<rect x="${x}" y="${baseY - hHome}" width="${barW}" height="${hHome}" rx="3" fill="var(--accent)"${faded}/>`;
      if (hPub > 0) svg += `<rect x="${x}" y="${baseY - hHome - hPub}" width="${barW}" height="${hPub}" rx="3" fill="var(--public)"${faded}/>`;
      if (step > 34) svg += `<text x="${x + barW / 2}" y="${baseY - hHome - hPub - 5}" text-anchor="middle" font-size="11" font-weight="500" fill="var(--fg)">${Math.round(it.home + it.pub)}</text>`;
      svg += `<text x="${x + barW / 2}" y="${CHART_HEIGHT - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${it.label}</text>`;
    });
    el.innerHTML = svg + "</svg>";
  }

  /**
   * Line chart with an optional dashed line for the overall average.
   * @param {HTMLElement} el
   * @param {{label:string, value:number}[]} points
   * @param {number|null} average
   * @param {{decimals?:number, empty?:string, fromZero?:boolean}} [options]
   */
  function drawLineChart(el, points, average, options = {}) {
    const decimals = options.decimals ?? 1;
    if (!points.length) { el.innerHTML = `<p class="empty">${options.empty || "יופיע כאן אחרי שתי מדידות עם קילומטראז׳."}</p>`; return; }
    points = points.slice(-CHART_MAX_POINTS);
    const width = Math.max(280, el.clientWidth || 600);
    const left = 44, bottom = 26, top = 22;
    const plotW = width - left - 24, plotH = CHART_HEIGHT - bottom - top;
    const values = points.map((p) => p.value).concat(average != null ? [average] : []);
    const min = Math.min(...values), max = Math.max(...values);
    const pad = Math.max(2, (max - min) * 0.12);
    let low = options.fromZero ? Math.min(0, min) : Math.max(0, Math.floor(min - pad));
    let high = Math.ceil(max + pad);
    if (high <= low) high = low + 1;
    const xAt = (i) => left + 12 + (points.length > 1 ? (plotW * i) / (points.length - 1) : plotW / 2);
    const yAt = (v) => top + plotH - (plotH * (v - low)) / (high - low);

    let svg = svgOpen(width, CHART_HEIGHT);
    for (let g = 0; g <= 3; g++) {
      const v = low + ((high - low) * g) / 3;
      svg += gridLine(left, width - 8, yAt(v), formatNumber(v, 0));
    }
    if (average != null) svg += `<line x1="${left}" x2="${width - 8}" y1="${yAt(average)}" y2="${yAt(average)}" stroke="var(--muted)" stroke-width="1.5" stroke-dasharray="5 4"/>`;
    if (points.length > 1) {
      const path = points.map((p, i) => (i ? "L" : "M") + xAt(i) + " " + yAt(p.value)).join(" ");
      const floor = yAt(Math.max(low, 0));
      svg += `<path d="${path} L${xAt(points.length - 1)} ${floor} L${xAt(0)} ${floor} Z" fill="var(--accent-soft)" stroke="none"/>`;
      svg += `<path d="${path}" fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round"/>`;
    }
    // On a narrow screen, label every n-th month (always the latest) so labels do not overlap.
    const every = Math.max(1, Math.ceil(36 / (plotW / Math.max(1, points.length - 1))));
    points.forEach((p, i) => {
      const isLast = i === points.length - 1;
      const labelled = (points.length - 1 - i) % every === 0;
      svg += `<circle cx="${xAt(i)}" cy="${yAt(p.value)}" r="${isLast ? 5.5 : 4}" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"/>`;
      if (labelled) svg += `<text x="${xAt(i)}" y="${yAt(p.value) - 10}" text-anchor="middle" font-size="11" font-weight="${isLast ? 700 : 500}" fill="var(--fg)">${formatNumber(p.value, decimals)}</text>`;
      if (labelled) svg += `<text x="${xAt(i)}" y="${CHART_HEIGHT - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${p.label}</text>`;
    });
    el.innerHTML = svg + "</svg>";
  }

  /**
   * Consumption and temperature over time: two lines on one month axis,
   * consumption (kWh per 100 km, left scale) and the month's mean temperature
   * (right scale). Tapping a month (or hovering / clicking on a computer)
   * shows its name and both values in a small label.
   * @param {HTMLElement} el
   * @param {{label:string, name:string, temp:(number|null), value:number}[]} points oldest first
   */
  function drawTempTimeChart(el, points) {
    points = points.slice(-CHART_MAX_POINTS);
    const width = Math.max(280, el.clientWidth || 600);
    const left = 34, right = 38, top = 14, bottom = 26;
    const plotW = width - left - right, plotH = CHART_HEIGHT - top - bottom;
    const values = points.map((p) => p.value);
    const temps = points.map((p) => p.temp).filter((t) => t != null);
    const vLow = Math.max(0, Math.floor(Math.min(...values) - 1)), vHigh = Math.ceil(Math.max(...values) + 1);
    const tLow = Math.floor((Math.min(...temps) - 2) / 5) * 5, tHigh = Math.ceil((Math.max(...temps) + 2) / 5) * 5;
    const xAt = (i) => left + (points.length > 1 ? (plotW * i) / (points.length - 1) : plotW / 2);
    const yV = (v) => top + plotH - (plotH * (v - vLow)) / (vHigh - vLow || 1);
    const yT = (t) => top + plotH - (plotH * (t - tLow)) / (tHigh - tLow || 1);
    const TEMP = "var(--temp, #e8743b)";

    let svg = svgOpen(width, CHART_HEIGHT);
    for (let g = 0; g <= 4; g++) {
      const y = top + (plotH * g) / 4;
      svg += `<line x1="${left}" x2="${width - right}" y1="${y}" y2="${y}" stroke="var(--line)" stroke-width="1"/>`;
      svg += `<text x="${left - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--accent)">${formatNumber(vHigh - ((vHigh - vLow) * g) / 4, 0)}</text>`;
      svg += `<text x="${width - right + 6}" y="${y + 4}" font-size="11" fill="${TEMP}">${formatNumber(tHigh - ((tHigh - tLow) * g) / 4, 0)}°</text>`;
    }
    // Temperature line, broken where a month has no weather data.
    let path = "", pen = false;
    points.forEach((p, i) => {
      if (p.temp == null) { pen = false; return; }
      path += (pen ? "L" : "M") + xAt(i) + " " + yT(p.temp) + " ";
      pen = true;
    });
    if (path) svg += `<path d="${path}" fill="none" stroke="${TEMP}" stroke-width="2.5" stroke-linejoin="round"/>`;
    const vPath = points.map((p, i) => (i ? "L" : "M") + xAt(i) + " " + yV(p.value)).join(" ");
    svg += `<path d="${vPath}" fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round"/>`;
    const every = Math.max(1, Math.ceil(40 / (plotW / Math.max(1, points.length - 1))));
    points.forEach((p, i) => {
      svg += `<circle cx="${xAt(i)}" cy="${yV(p.value)}" r="3.5" fill="var(--accent)"/>`;
      if (p.temp != null) svg += `<circle cx="${xAt(i)}" cy="${yT(p.temp)}" r="3" fill="${TEMP}"/>`;
      if ((points.length - 1 - i) % every === 0) svg += `<text x="${xAt(i)}" y="${CHART_HEIGHT - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${p.label}</text>`;
    });
    // One tall, finger-wide target per month.
    const slot = points.length > 1 ? plotW / (points.length - 1) : plotW;
    points.forEach((p, i) => {
      const tempText = p.temp != null ? formatNumber(p.temp, 1) + " מעלות" : "אין נתון טמפרטורה";
      svg += `<rect class="slot" data-i="${i}" x="${xAt(i) - slot / 2}" y="${top}" width="${slot}" height="${plotH}" fill="transparent" tabindex="0" role="button" ` +
        `aria-label="${p.name}: ${formatNumber(p.value, 1)} קוט״ש ל‑100 ק״מ, ${tempText}"><title>${p.name}: ${formatNumber(p.value, 1)} קוט״ש ל‑100 ק״מ, ${tempText}</title></rect>`;
    });
    el.innerHTML = svg + '<g class="dot-tip" pointer-events="none"></g></svg>';

    const tip = el.querySelector(".dot-tip");
    const show = (i) => {
      const p = points[i];
      if (!p) { tip.innerHTML = ""; return; }
      const cx = xAt(i);
      const boxW = 176, boxH = 42;
      const x = Math.min(width - right - boxW, Math.max(left, cx - boxW / 2));
      const y = top + 2;
      const second = formatNumber(p.value, 1) + " קוט״ש ל‑100 · " + (p.temp != null ? formatNumber(p.temp, 1) + " מעלות" : "אין טמפרטורה");
      tip.innerHTML = `<line x1="${cx}" x2="${cx}" y1="${top}" y2="${top + plotH}" stroke="var(--muted)" stroke-width="1" stroke-dasharray="3 3"/>` +
        `<rect x="${x}" y="${y}" width="${boxW}" height="${boxH}" rx="8" fill="var(--fg)"/>` +
        `<text x="${x + boxW / 2}" y="${y + 16}" text-anchor="middle" font-size="12" font-weight="700" fill="var(--surface)">${p.name}</text>` +
        `<text x="${x + boxW / 2}" y="${y + 33}" text-anchor="middle" font-size="11" fill="var(--surface)" direction="rtl">${second}</text>`;
    };
    let pinned = -1;
    el.querySelectorAll(".slot").forEach((r) => {
      const i = +r.dataset.i;
      r.addEventListener("click", (e) => { e.stopPropagation(); pinned = pinned === i ? -1 : i; show(pinned); });
      r.addEventListener("mouseenter", () => show(i));
      r.addEventListener("mouseleave", () => show(pinned));
      r.addEventListener("focus", () => show(i));
    });
    el.querySelector("svg").addEventListener("click", () => { pinned = -1; show(-1); });
  }

  /* ======================================================================
   * 7. Rendering
   * ==================================================================== */

  /** "X ₪, מתוכם Y בית ו‑Z ציבורי" - the home / public split of an amount. */
  function describeSplit(home, pub) {
    return formatNumber(home + pub) + " ₪, מתוכם " + formatNumber(home) + " ₪ טעינה ביתית ו‑" + formatNumber(pub) + " ₪ טעינה ציבורית";
  }

  function renderStatus() {
    const pill = $("status");
    const readings = state.readings;
    if (!readings.length) {
      if (state.loading) pill.textContent = "טוען נתונים מ‑OneDrive…";
      else pill.textContent = state.connected ? "עוד אין מדידות. התחל בלשונית הזנה" : "לא מחובר ל‑OneDrive";
      pill.className = "pill" + (!state.connected && !state.loading ? " warn" : "");
      return;
    }
    const last = readings[readings.length - 1];
    const age = daysBetween(last.date, todayIso());
    pill.textContent = "מדידה אחרונה " + formatDate(last.date) + " · לפני " + age + " ימים";
    pill.className = "pill" + (age > STALE_AFTER_DAYS ? " warn" : "");
  }

  function renderOverview(model, months) {
    const { driven, spent, pending, spans } = model;
    const efficiency = driven.km ? (driven.kwhUsed / driven.km) * 100 : null;
    const costPerKm = driven.km ? driven.cost / driven.km : null;
    // Per-km comparison and forecasts use today's petrol price; amounts so far use each month's price.
    const currentFuelPrice = fuelPriceOn(todayIso());
    const fuelPerKm = (state.settings.fuelL100 / 100) * currentFuelPrice;
    const petrolSoFar = months.reduce((sum, m) => sum + (m.km > 0 ? m.km * monthFuelPerKm(m.key) : 0), 0);
    const savedByMonth = months.reduce((sum, m) => sum + (monthSaving(m) || 0), 0);
    const saved = driven.km ? savedByMonth : null;

    $("k-eff").textContent = formatNumber(efficiency, 1);
    $("k-cpk").textContent = formatNumber(costPerKm == null ? null : costPerKm * 100, 1);
    // Km driven on electricity for the money one liter of petrol costs today.
    $("k-kpl").textContent = formatNumber(costPerKm ? currentFuelPrice / costPerKm : null, 0);
    $("k-kpl-unit").textContent = "בכסף של ליטר, " + formatNumber(currentFuelPrice, 2) + " ₪";
    $("k-month").textContent = formatNumber(spent.days ? (spent.total / spent.days) * 30 : null, 0);
    $("k-save").textContent = formatNumber(saved, 0);

    const shown = months.slice(-CHART_MAX_POINTS);
    drawCostChart($("c-cost"), shown.map((m, i) => ({ label: monthLabel(m.key, i === 0), home: m.costHome, pub: m.costPublic, faded: m.partial })));
    const lastFull = [...months].reverse().find((m) => !m.partial);
    let splitText = lastFull
      ? "החודש המלא האחרון, " + monthName(lastFull.key) + ": " + describeSplit(lastFull.costHome, lastFull.costPublic) + "."
      : "";
    if (shown.some((m) => m.partial)) splitText += (splitText ? " " : "") + "עמודה בהירה היא חודש שהמדידות לא מכסות את כל ימיו.";
    if (pending.count) {
      splitText += (splitText ? " " : "") + (pending.count === 1
        ? "מאז המדידה האחרונה נרשמה טעינה ציבורית אחת בסך " + formatNumber(pending.cost) + " ₪. היא תיכנס"
        : "מאז המדידה האחרונה נרשמו " + pending.count + " טעינות ציבוריות בסך " + formatNumber(pending.cost) + " ₪. הן ייכנסו") +
        " לתקופה הבאה כשתוזן מדידה.";
    }
    $("cost-split").textContent = splitText;

    const effMonths = months.filter((m) => m.efficiency != null).slice(-CHART_MAX_POINTS);
    drawLineChart($("c-eff"), effMonths.map((m, i) => ({ label: monthLabel(m.key, i === 0), value: m.efficiency })), efficiency || 0);
    const trend = $("trend");
    trend.textContent = "";
    trend.className = "";
    const lastFullEff = [...months].reverse().find((m) => !m.partial && m.efficiency != null);
    if (lastFullEff && effMonths.length > 1 && efficiency) {
      const change = ((lastFullEff.efficiency - efficiency) / efficiency) * 100;
      trend.className = "pill " + (change <= 0 ? "good" : "warn");
      trend.textContent = monthName(lastFullEff.key) + ": " + formatNumber(Math.abs(change), 0) + "% " + (change <= 0 ? "מתחת לממוצע" : "מעל הממוצע");
    }

    renderExtraCharts(model, months, fuelPerKm);
    renderTemperature(months);
    renderProviders(model);

    const facts = $("facts");
    if (driven.km) {
      const kmPerDay = driven.km / driven.days;
      const yearKm = kmPerDay * 365;
      const yearCost = (spent.total / spent.days) * 365;
      const rows = [
        ["ק״מ שנסעת במעקב", formatNumber(driven.km) + " ק״מ"],
        ["חשמל שנצרך", formatNumber(spent.kwh) + " קוט״ש"],
        ["סה״כ שילמת על חשמל", formatNumber(spent.total) + " ₪"],
        ["מתוכם טעינה ביתית", formatNumber(spent.home) + " ₪"],
        ["מתוכם טעינה ציבורית", formatNumber(spent.public) + " ₪"],
        ...(pending.count ? [["טעינות ציבוריות מאז המדידה האחרונה", formatNumber(pending.cost) + " ₪"]] : []),
        ["אותם ק״מ בבנזין", formatNumber(petrolSoFar) + " ₪"],
        ["עלות לק״מ: חשמל מול בנזין", formatNumber(costPerKm * 100, 0) + " מול " + formatNumber(fuelPerKm * 100, 0) + " אג׳"],
        ["קצב נסיעה", formatNumber(kmPerDay * 30) + " ק״מ בחודש"],
        ["תחזית ק״מ לשנה", formatNumber(yearKm) + " ק״מ"],
        ["תחזית עלות חשמל לשנה", formatNumber(yearCost) + " ₪"],
        ["תחזית חיסכון לשנה", formatNumber(yearKm * fuelPerKm - yearKm * costPerKm) + " ₪"],
      ];
      facts.innerHTML = rows.map(([label, value]) => `<div><dt>${label}</dt><dd>${value}</dd></div>`).join("");
      $("facts-note").textContent = (driven.corrected ? "הצריכה מתוקנת לפי אחוז הסוללה ב‑" + driven.corrected + " מתוך " + spans.length + " תקופות. " : "") + "ההשוואה מניחה רכב בנזין שצורך " + formatNumber(state.settings.fuelL100, 1) +
        " ליטר ל‑100 ק״מ. " + (fuelPrices.length
          ? "מחיר הבנזין בכל חודש הוא המחיר הרשמי של משרד האנרגיה, וכרגע " + formatNumber(currentFuelPrice, 2) + " ₪ לליטר."
          : "מחיר הבנזין הוא " + formatNumber(currentFuelPrice, 2) + " ₪ לליטר.") +
        " אפשר לשנות בהגדרות. הצריכה נמדדת במונה העמדה ולכן כוללת את הפסדי הטעינה.";
    } else {
      facts.innerHTML = '<p class="empty span-all">יופיע כאן אחרי שתי מדידות עם קילומטראז׳.</p>';
      $("facts-note").textContent = "";
    }

    // Efficiency by season, shown once at least two seasons have data.
    const seasons = { "קיץ": { kwh: 0, km: 0 }, "חורף": { kwh: 0, km: 0 }, "עונות מעבר": { kwh: 0, km: 0 } };
    for (const s of spans) {
      const month = new Date((toTime(s.from) + toTime(s.to)) / 2).getUTCMonth() + 1;
      const name = month >= 6 && month <= 9 ? "קיץ" : month === 12 || month <= 2 ? "חורף" : "עונות מעבר";
      seasons[name].kwh += s.kwhUsed;
      seasons[name].km += s.km;
    }
    const withData = Object.keys(seasons).filter((name) => seasons[name].km > 0);
    $("season-panel").hidden = withData.length < 2;
    $("season").innerHTML = withData
      .map((name) => `<div><dt>${name}</dt><dd>${formatNumber((seasons[name].kwh / seasons[name].km) * 100, 1)} קוט״ש ל‑100 ק״מ</dd></div>`)
      .join("");
  }

  /** Cost per 100 km, cumulative savings and the public share, month by month. */
  function renderExtraCharts(model, months, fuelPerKm) {
    const { driven } = model;
    const withEff = months.filter((m) => m.efficiency != null);
    const cost100 = withEff.map((m) => ({ key: m.key, value: (m.efficiency * monthPricePerKwh(m)) || 0 }));
    drawLineChart($("c-cost100"), cost100.slice(-CHART_MAX_POINTS).map((p, i) => ({ label: monthLabel(p.key, i === 0), value: p.value })),
      driven.km ? (driven.cost / driven.km) * 100 : null, { decimals: 0 });

    // Per month: that month's petrol price divided by that month's electricity cost per km.
    const kpl = withEff
      .map((m) => ({ key: m.key, costPerKm: (m.efficiency * monthPricePerKwh(m)) / 100 }))
      .filter((p) => p.costPerKm > 0)
      .map((p) => ({ key: p.key, value: monthFuelPrice(p.key) / p.costPerKm }));
    const kplAverage = driven.km && driven.cost ? fuelPriceOn(todayIso()) / (driven.cost / driven.km) : null;
    drawLineChart($("c-kpl"), kpl.slice(-CHART_MAX_POINTS).map((p, i) => ({ label: monthLabel(p.key, i === 0), value: p.value })),
      kplAverage, { decimals: 0 });

    let running = 0;
    const cumulative = [];
    for (const m of months) {
      const saving = monthSaving(m);
      if (saving == null) continue;
      running += saving;
      cumulative.push({ key: m.key, value: running });
    }
    drawLineChart($("c-saving"), cumulative.slice(-CHART_MAX_POINTS).map((p, i) => ({ label: monthLabel(p.key, i === 0), value: p.value })),
      null, { decimals: 0, fromZero: true });

    const share = months.filter((m) => m.costHome + m.costPublic > 0)
      .map((m) => ({ key: m.key, value: (m.costPublic / (m.costHome + m.costPublic)) * 100 }));
    const totalShare = model.spent.total ? (model.spent.public / model.spent.total) * 100 : null;
    drawLineChart($("c-share"), share.slice(-CHART_MAX_POINTS).map((p, i) => ({ label: monthLabel(p.key, i === 0), value: p.value })),
      totalShare, { decimals: 0, fromZero: true, empty: "יופיע כאן אחרי שתי מדידות." });
    $("share-note").textContent = totalShare != null
      ? "בכל התקופה " + formatNumber(totalShare, 0) + "% מההוצאה על חשמל היו טעינה ציבורית."
      : "";
  }

  /** Consumption against the month's average temperature in Beit Yehoshua. */
  function renderTemperature(months) {
    const el = $("c-temp");
    const note = $("temp-note");
    if (state.readings.length < 2) {
      el.innerHTML = '<p class="empty">יופיע כאן אחרי שתי מדידות עם קילומטראז׳.</p>';
      note.textContent = "";
      return;
    }
    const first = state.readings[0].date;
    const last = state.readings[state.readings.length - 1].date;
    ensureWeather(first, () => { if (!$("v-overview").hidden) renderTemperature(lastMonths); });
    const withEff = months.filter((m) => m.efficiency != null);
    const points = withEff.map((m, i) => ({ label: monthLabel(m.key, i === 0), name: monthName(m.key), temp: monthTemperature(m.key, first, last), value: m.efficiency }));
    if (points.some((p) => p.temp != null)) {
      drawTempTimeChart(el, points);
      note.textContent = "נגיעה בחודש, או ריחוף במחשב, מראה את הצריכה ואת הטמפרטורה שלו. הטמפרטורה היא הממוצע היומי בבית יהושע, מהשירות Open‑Meteo.";
    } else if (weather.status === "loading") {
      el.innerHTML = '<p class="empty">טוען נתוני מזג אוויר…</p>';
      note.textContent = "";
    } else if (weather.status === "failed" && !Object.keys(weather.days).length) {
      el.innerHTML = '<p class="empty">לא הצלחתי לטעון נתוני מזג אוויר. זה יעבוד כשיהיה חיבור לאינטרנט.</p>';
      note.textContent = "";
    } else {
      el.innerHTML = '<p class="empty">יופיע כאן כשיהיה חודש עם צריכה ידועה.</p>';
      note.textContent = "";
    }
  }

  /** Public charging per company, inside a collapsed section. */
  function renderProviders(model) {
    const providers = buildProviders(state.publicCharges);
    $("prov-panel").hidden = !providers.length;
    if (!providers.length) return;
    const total = providers.reduce((sum, p) => sum + p.cost, 0);
    const homeKwh = model.periods.reduce((sum, p) => sum + p.kwhHome, 0);
    const homePrice = homeKwh ? model.spent.home / homeKwh : null;
    const header = "<tr><th>חברה</th><th>טעינות</th><th>קוט״ש</th><th>סה״כ ₪</th><th>₪ לקוט״ש</th><th>מההוצאה הציבורית</th></tr>";
    const rows = providers.map((p) =>
      `<tr><td>${p.name ? providerLabel(p.name) : "ללא שם חברה"}</td><td>${p.count}</td><td>${formatNumber(p.kwh, 1)}</td>` +
      `<td>${formatNumber(p.cost)}</td><td>${formatNumber(p.kwh ? p.cost / p.kwh : null, 2)}</td>` +
      `<td>${total ? formatNumber((p.cost / total) * 100, 0) + "%" : "–"}</td></tr>`).join("");
    const homeRow = homePrice != null
      ? `<tr class="muted-row"><td>טעינה ביתית, להשוואה</td><td>–</td><td>${formatNumber(homeKwh, 1)}</td><td>${formatNumber(model.spent.home)}</td><td>${formatNumber(homePrice, 2)}</td><td>–</td></tr>`
      : "";
    $("t-providers").innerHTML = header + rows + homeRow;
    $("prov-summary").textContent = "ניתוח לפי חברת טעינה (" + providers.length + (providers.length === 1 ? " חברה)" : " חברות)");
  }

  /** One row per calendar year. */
  function renderYears(months) {
    const years = buildYears(months);
    $("years-panel").hidden = !years.length;
    if (!years.length) return;
    const header = "<tr><th>שנה</th><th>ימים</th><th>ק״מ</th><th>קוט״ש</th><th>עלות בית ₪</th><th>עלות ציבורי ₪</th>" +
      "<th>סה״כ ₪</th><th>קוט״ש/100</th><th>חיסכון מול בנזין ₪</th><th>שינוי בעלות ליום</th></tr>";
    const rows = years.map((y, i) => {
      const total = y.costHome + y.costPublic;
      const previous = years[i - 1];
      let change = "–";
      if (previous && previous.days && y.days) {
        const perDayNow = total / y.days, perDayBefore = (previous.costHome + previous.costPublic) / previous.days;
        if (perDayBefore > 0) {
          const pct = ((perDayNow - perDayBefore) / perDayBefore) * 100;
          change = (pct > 0 ? "עלייה של " : "ירידה של ") + formatNumber(Math.abs(pct), 0) + "%";
        }
      }
      return `<tr><td>${y.year}</td><td>${y.partial ? y.days + " מתוך " + y.daysInYear + " (חלקית)" : y.days}</td>` +
        `<td>${y.km ? formatNumber(y.km) : "–"}</td><td>${formatNumber(y.kwhHome + y.kwhPublic)}</td>` +
        `<td>${formatNumber(y.costHome)}</td><td>${y.costPublic ? formatNumber(y.costPublic) : "–"}</td>` +
        `<td><strong>${formatNumber(total)}</strong></td><td>${y.efficiency != null ? formatNumber(y.efficiency, 1) : "–"}</td>` +
        `<td>${y.km ? formatNumber(y.saving) : "–"}</td><td>${change}</td></tr>`;
    }).reverse().join("");
    $("t-years").innerHTML = header + rows;
  }

  function renderHistory(model, months) {
    const { periods } = model;
    renderYears(months);

    if (months.length) {
      const header = "<tr><th>חודש</th><th>ימים</th><th>ק״מ</th><th>קוט״ש בית</th><th>קוט״ש ציבורי</th>" +
        "<th>עלות בית ₪</th><th>עלות ציבורי ₪</th><th>סה״כ ₪</th><th>קוט״ש/100</th></tr>";
      const rows = months.slice().reverse().map((m) =>
        `<tr><td>${monthName(m.key)}</td><td>${m.partial ? m.days + " מתוך " + m.daysInMonth + " (חלקי)" : m.days}</td>` +
        `<td>${m.km ? formatNumber(m.km) : "–"}</td><td>${formatNumber(m.kwhHome, 1)}</td>` +
        `<td>${m.kwhPublic ? formatNumber(m.kwhPublic, 1) : "–"}</td><td>${formatNumber(m.costHome)}</td>` +
        `<td>${m.costPublic ? formatNumber(m.costPublic) : "–"}</td><td><strong>${formatNumber(m.costHome + m.costPublic)}</strong></td>` +
        `<td>${m.efficiency != null ? formatNumber(m.efficiency, 1) : "–"}</td></tr>`).join("");
      $("t-months").innerHTML = header + rows;
    } else {
      $("t-months").innerHTML = '<tr><td class="empty">עוד אין שתי מדידות להשוואה.</td></tr>';
    }

    const deleteButton = (target) => `<button class="btn ghost small danger" data-del="${escapeHtml(target)}">מחק</button>`;

    if (periods.length) {
      const header = "<tr><th>עד תאריך</th><th>ימים</th><th>ק״מ</th><th>קוט״ש בית</th><th>קוט״ש ציבורי</th>" +
        "<th>עלות בית ₪</th><th>עלות ציבורי ₪</th><th>סה״כ ₪</th><th>קוט״ש/100</th><th>₪ ל‑30 יום</th></tr>";
      const rows = periods.slice().reverse().map((p) => {
        const total = p.costHome + p.costPublic;
        const km = p.km != null ? formatNumber(p.km) : p.kmMerged ? formatNumber(p.kmMerged) + " *" : "חסר";
        const eff = p.efficiency != null ? formatNumber(p.efficiency, 1) + (p.mergedCount > 1 ? " *" : "") : "–";
        return `<tr><td>${formatDate(p.to)}</td><td>${p.days}</td><td>${km}</td><td>${formatNumber(p.kwhHome, 1)}</td>` +
          `<td>${p.kwhPublic ? formatNumber(p.kwhPublic, 1) : "–"}</td><td>${formatNumber(p.costHome)}</td>` +
          `<td>${p.costPublic ? formatNumber(p.costPublic) : "–"}</td><td><strong>${formatNumber(total)}</strong></td>` +
          `<td>${eff}</td><td>${formatNumber(p.days ? (total / p.days) * 30 : null)}</td></tr>`;
      }).join("");
      const footnote = periods.some((p) => p.mergedCount > 1)
        ? '<tr><td colspan="10" class="footnote">* חסר קילומטראז׳ במדידה קודמת, ולכן הק״מ והצריכה מחושבים על כמה תקופות יחד.</td></tr>'
        : "";
      $("t-periods").innerHTML = header + rows + footnote;
    } else {
      $("t-periods").innerHTML = '<tr><td class="empty">עוד אין שתי מדידות להשוואה.</td></tr>';
    }

    if (state.readings.length) {
      const rows = state.readings.map((r, i) =>
        `<tr><td>${formatDate(r.date)}</td><td>${formatNumber(r.meter, 2)}</td><td>${r.odo != null ? formatNumber(r.odo) : "חסר"}</td>` +
        `<td>${formatNumber(priceForReading(state.readings, i), 4)}</td><td>${r.soc != null ? formatNumber(r.soc) + "%" : "–"}</td><td>${deleteButton("readings/" + r.date)}</td></tr>`);
      $("t-readings").innerHTML = "<tr><th>תאריך</th><th>מונה קוט״ש</th><th>קילומטראז׳</th><th>₪ לקוט״ש</th><th>סוללה</th><th></th></tr>" + rows.reverse().join("");
    } else {
      $("t-readings").innerHTML = '<tr><td class="empty">עוד אין מדידות.</td></tr>';
    }

    if (state.publicCharges.length) {
      // Where each charge is counted: the period that contains it, or why it is in none yet.
      const countedIn = (charge) => {
        const period = periods.find((p) => charge.date > p.from && charge.date <= p.to);
        if (period) return "תקופה עד " + formatDate(period.to);
        const last = state.readings[state.readings.length - 1];
        return !last || charge.date > last.date ? "ממתינה למדידה הבאה" : "לפני תחילת המעקב";
      };
      const rows = state.publicCharges.slice().reverse().map((p) =>
        `<tr><td>${formatDate(p.date)}</td><td>${p.provider ? providerLabel(p.provider) : "–"}</td><td>${formatNumber(p.kwh, 1)}</td><td>${formatNumber(p.cost, 2)}</td>` +
        `<td>${formatNumber(p.kwh ? p.cost / p.kwh : null, 2)}</td><td>${countedIn(p)}</td><td>${escapeHtml(p.note)}</td><td>${deleteButton("public/" + p.id)}</td></tr>`);
      $("t-public").innerHTML = "<tr><th>תאריך</th><th>חברה</th><th>קוט״ש</th><th>₪</th><th>₪ לקוט״ש</th><th>נספרת ב</th><th>הערה</th><th></th></tr>" + rows.join("");
    } else {
      $("t-public").innerHTML = '<tr><td class="empty">עוד לא נרשמו טעינות ציבוריות. מוסיפים בלשונית הזנה.</td></tr>';
    }
    $("export").hidden = !state.readings.length && !state.publicCharges.length;
  }

  /** Settings line about the official petrol prices. */
  function renderFuelPricesInfo() {
    const el = $("fuel-auto");
    if (!el) return;
    if (!fuelPrices.length) { el.textContent = "המחירים הרשמיים עוד לא נטענו. עד שייטענו, החישוב משתמש במחיר שכאן."; return; }
    const last = fuelPrices[fuelPrices.length - 1];
    el.textContent = "מחיר הבנזין מתעדכן אוטומטית לפי ההודעות של משרד האנרגיה. המחיר העדכני הוא " +
      formatNumber(last.price, 2) + " ₪ לליטר, מיום " + formatDate(last.from) +
      ". המחיר שכאן משמש רק לתאריכים שלפני " + formatDate(fuelPrices[0].from) + ".";
  }

  /** Keeps form defaults in sync with the data without overwriting what the user is typing. */
  function renderFormDefaults() {
    const priceInput = $("r-price");
    if (!priceInput.dataset.touched && document.activeElement !== priceInput) priceInput.value = latestPrice();
    if (document.activeElement !== $("s-price")) $("s-price").value = state.settings.fuelPrice;
    if (document.activeElement !== $("s-l100")) $("s-l100").value = state.settings.fuelL100;
    if (document.activeElement !== $("s-battery")) $("s-battery").value = state.settings.batteryKwh;
  }

  /** The months of the latest render, for redraws that only need them (weather arriving). */
  let lastMonths = [];

  /** Redraws every screen from the current state. */
  function render() {
    const model = buildModel(state.readings, state.publicCharges, state.settings.batteryKwh);
    const months = buildMonths(state.readings, state.publicCharges, model);
    lastMonths = months;
    renderStatus();
    renderOverview(model, months);
    renderHistory(model, months);
    renderLoan(months);
    renderFormDefaults();
    renderFuelPricesInfo();
    renderReadingPreview();
  }

  /* ---------- Car financing screen ---------- */

  const money = (v) => (v == null || !isFinite(v) ? "–" : (v < 0 ? "−" : "") + formatNumber(Math.abs(v), 0) + " ₪");

  /**
   * Cumulative cost of both paths over the months: solid up to today, dashed
   * after it (forecast). A dot marks the month the Zeekr becomes worth it.
   */
  function drawLoanChart(el, model) {
    const rows = model.rows;
    const width = Math.max(280, el.clientWidth || 600);
    const height = 230, left = 44, right = 12, top = 16, bottom = 26;
    const plotW = width - left - right, plotH = height - top - bottom;
    const values = rows.flatMap((r) => [r.costZ, r.costA]);
    let low = Math.min(0, ...values), high = Math.max(...values);
    if (high <= low) high = low + 1;
    const span = high - low;
    high += span * 0.08;
    const xAt = (i) => left + (plotW * i) / Math.max(1, rows.length - 1);
    const yAt = (v) => top + plotH - (plotH * (v - low)) / (high - low);
    const path = (from, to, f) => rows.slice(from, to + 1).map((r, j) => (j ? "L" : "M") + xAt(from + j).toFixed(1) + " " + yAt(f(r)).toFixed(1)).join(" ");

    let svg = svgOpen(width, height);
    for (let g = 0; g <= 3; g++) {
      const v = low + ((high - low) * g) / 3;
      svg += gridLine(left, width - right, yAt(v), formatNumber(v / 1000, 0) + "K");
    }
    const t = Math.min(model.todayIndex, rows.length - 1);
    for (const [f, color] of [[(r) => r.costA, "var(--public)"], [(r) => r.costZ, "var(--accent)"]]) {
      svg += `<path d="${path(0, t, f)}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linejoin="round"/>`;
      if (t < rows.length - 1) svg += `<path d="${path(t, rows.length - 1, f)}" fill="none" stroke="${color}" stroke-width="2" stroke-dasharray="5 4"/>`;
    }
    svg += `<line x1="${xAt(t)}" x2="${xAt(t)}" y1="${top}" y2="${top + plotH}" stroke="var(--muted)" stroke-width="1" stroke-dasharray="3 3"/>`;
    svg += `<text x="${xAt(t)}" y="${top + 10}" text-anchor="${t > rows.length * 0.8 ? "end" : "start"}" dx="${t > rows.length * 0.8 ? -4 : 4}" font-size="11" fill="var(--muted)">היום</text>`;
    if (model.breakEven) {
      const b = model.breakEven;
      svg += `<circle cx="${xAt(b.i)}" cy="${yAt(b.costZ)}" r="5.5" fill="var(--good)" stroke="var(--surface)" stroke-width="2"/>`;
    }
    // Year labels along the bottom.
    for (const r of rows) {
      if (r.i % 12 !== 0) continue;
      svg += `<text x="${xAt(r.i)}" y="${height - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${r.i ? "שנה " + r.i / 12 : "קנייה"}</text>`;
    }
    el.innerHTML = svg + "</svg>";
  }

  const factsHtml = (rows) => rows.map(([label, value]) => `<div><dt>${label}</dt><dd>${value}</dd></div>`).join("");
  const pct = (v) => (v == null ? "–" : formatNumber(v, 2) + "%");

  function renderLoan(months) {
    const loan = state.loan;
    const model = buildLoanModel(loan, months, todayIso());
    const ready = !model.missing.length;
    $("loan-missing").hidden = ready;
    $("loan-missing").textContent = ready ? "" : "כדי לחשב חסרים: " + model.missing.join(", ") + ". אפשר להזין אותם למטה, בנתוני הבסיס.";
    for (const id of ["loan-summary", "loan-compare", "loan-chart-panel", "loan-detail-panel"]) $(id).hidden = !ready;
    renderLoanForm();
    if (!ready) return;

    const now = model.now, end = model.end;
    const adv = now.advantage;
    $("loan-adv").textContent = money(Math.abs(adv));
    $("loan-adv").className = "loan-big " + (adv >= 0 ? "good" : "bad");
    $("loan-adv-label").textContent = adv >= 0 ? "לטובת הזיקר עד היום" : "לטובת רכב היד השנייה עד היום";
    const pill = $("loan-pill");
    pill.textContent = adv >= 0 ? "כדאי כרגע" : "עוד לא כדאי";
    pill.className = "pill " + (adv >= 0 ? "good" : "warn");
    $("loan-progress").style.width = Math.min(100, (model.todayIndex / Math.max(1, model.horizon)) * 100) + "%";
    $("loan-month").textContent = "חודש " + model.todayIndex + " מתוך " + model.horizon;
    $("loan-forecast").textContent = "תחזית לסוף התקופה, " + formatDate(end.date) + ": " + money(Math.abs(end.advantage)) + (end.advantage >= 0 ? " לטובת הזיקר." : " לטובת רכב היד השנייה.") +
      (model.breakEven ? " נקודת המפגש: " + monthName(model.breakEven.key) + "." : "");

    const pz = now.partsZ, pa = now.partsA;
    $("loan-z").innerHTML = factsHtml([
      ["ירידת ערך", money(pz.depreciation)],
      ["ריבית ועמלות", money(pz.interest)],
      ["חשמל", money(pz.energy)],
      ["ביטוח", money(pz.ins)],
      ["טיפולים", money(pz.maint)],
      ["תשואה שאבדה", money(pz.lostReturn)],
      ["הלוואה עומדת שנשחקה", money(pz.standing)],
      ["<b>סה״כ עלות</b>", "<b>" + money(now.costZ) + "</b>"],
    ]);
    $("loan-a").innerHTML = factsHtml([
      ["ירידת ערך", money(pa.depreciation)],
      ["דלק", money(pa.fuel)],
      ["ביטוח", money(pa.ins)],
      ["טיפולים", money(pa.maint)],
      ["תשואה שאבדה", money(pa.lostReturn)],
      ["<b>סה״כ עלות</b>", "<b>" + money(now.costA) + "</b>"],
    ]);

    drawLoanChart($("c-loan"), model);
    $("loan-chart-note").textContent = "קו מלא עד היום, מקווקו אחרי היום הוא תחזית לפי הממוצע." + (model.breakEven ? " הנקודה הירוקה היא נקודת המפגש." : "");

    // Base data, then derived figures, then the formula and the month table.
    $("loan-base").innerHTML = factsHtml([
      ["תאריך הקנייה", formatDate(loan.buyDate)],
      ["מחיר הזיקר", money(loan.carPrice)],
      ["הלוואה על חשבון הגמל", money(model.principal0)],
      ["הלוואה עומדת, ביטוח לאומי", money(loan.standAmount || 0)],
      ["מזומן במעמד הקנייה", money(model.ownCash)],
      ["ירידת ערך שנתית, זיקר", pct(loan.carDepPct)],
      ["ביטוח שנתי, זיקר", money(loan.carIns)],
      ["טיפולים בשנה, זיקר", money(loan.carMaint)],
      ["מחיר הרכב החלופי", money(loan.altPrice)],
      ["צריכת דלק, חלופי", formatNumber(loan.altL100 || state.settings.fuelL100, 1) + " ל׳ ל‑100 ק״מ"],
      ["ירידת ערך שנתית, חלופי", pct(loan.altDepPct)],
      ["ביטוח שנתי, חלופי", money(loan.altIns)],
      ["טיפולים בשנה, חלופי", money(loan.altMaint)],
      ["תשואה שנתית בגמל", pct(loan.returnPct)],
      ["שחיקת החוב העומד בשנה", pct(loan.standDropPct)],
      ["חודשים עד שאין החזר עומד", loan.standFreeMonths > 0 ? formatNumber(loan.standFreeMonths) : "אין"],
    ]);
    $("loan-derived").innerHTML = factsHtml([
      ["החזר חודשי בלוח הראשון", money(model.firstPayment)],
      ["סוף ההלוואה", formatDate(model.loanEnd)],
      ["יתרת הלוואה היום", money(now.balance)],
      ["חוב עומד היום", money(now.standOwed)],
      ["שווי הזיקר היום", money(now.valueZ)],
      ["שווי החלופי היום", money(now.valueA)],
      ["ק״מ בחודש לתחזית", formatNumber(model.kmEstimate, 0) + (loan.kmMonth > 0 ? " ק״מ, לפי ההגדרה" : " ק״מ, ממוצע המדידות")],
      ["עלות חשמל לק״מ", model.costPerKm ? formatNumber(model.costPerKm * 100, 1) + " אג׳" : "אין עדיין מדידות"],
      ["חודשים לפי הערכה עד היום", formatNumber(model.estimatedMonths) + " מתוך " + formatNumber(model.todayIndex)],
    ]);
    $("loan-sched").innerHTML = "<tr><th>מתאריך</th><th>יתרה</th><th>ריבית שנתית</th><th>חודשים</th><th>החזר חודשי</th><th>יתרה בסוף</th></tr>" +
      loan.schedules.map((s) => `<tr><td>${formatDate(s.from)}</td><td>${money(s.balance)}</td><td>${pct(s.ratePct)}</td><td>${s.months}</td><td>${money(schedulePayment(s))}${s.payment > 0 ? ", קבוע" : ""}</td><td>${money(scheduleBalance(s, s.months))}</td></tr>`).join("");
    $("loan-table").innerHTML = "<tr><th>חודש</th><th>ק״מ</th><th>חשמל</th><th>דלק בחלופי</th><th>החזר</th><th>מתוכו ריבית</th><th>יתרה</th><th>חוב עומד</th><th>שווי זיקר</th><th>שווי חלופי</th><th>עלות זיקר</th><th>עלות חלופי</th><th>הפרש</th></tr>" +
      model.rows.map((r) => `<tr class="${r.past ? "" : "muted-row"}"><td>${monthLabel(r.key, true)}${r.estimate && r.i ? " *" : ""}</td><td>${formatNumber(r.km)}</td><td>${formatNumber(r.energy)}</td><td>${formatNumber(r.fuel)}</td><td>${formatNumber(r.payment)}</td><td>${formatNumber(r.interest)}</td><td>${formatNumber(r.balance)}</td><td>${formatNumber(r.standOwed)}</td><td>${formatNumber(r.valueZ)}</td><td>${formatNumber(r.valueA)}</td><td>${formatNumber(r.costZ)}</td><td>${formatNumber(r.costA)}</td><td class="${r.advantage >= 0 ? "pos" : "neg"}">${formatNumber(r.advantage)}</td></tr>`).join("");
  }

  /* ---------- Base data form ---------- */

  const LOAN_FIELDS = {
    "l-buy": "buyDate", "l-car": "carPrice", "l-car-dep": "carDepPct", "l-car-ins": "carIns", "l-car-maint": "carMaint",
    "l-alt": "altPrice", "l-alt-l100": "altL100", "l-alt-dep": "altDepPct", "l-alt-ins": "altIns", "l-alt-maint": "altMaint",
    "l-return": "returnPct", "l-stand": "standAmount", "l-stand-drop": "standDropPct", "l-stand-free": "standFreeMonths", "l-km": "kmMonth",
  };

  /** Fills the form from the saved data, unless the user is editing it. */
  function renderLoanForm() {
    const form = $("f-loan");
    if (form.dataset.dirty) return;
    for (const [id, key] of Object.entries(LOAN_FIELDS)) $(id).value = state.loan[key] ?? "";
    $("l-sched").innerHTML = "";
    for (const s of state.loan.schedules) addScheduleRow(s);
    if (!state.loan.schedules.length) addScheduleRow({ from: state.loan.buyDate || "", balance: "", ratePct: "", months: "" });
  }

  function addScheduleRow(s) {
    const row = document.createElement("div");
    row.className = "sched-row";
    row.innerHTML =
      `<label>מתאריך<input type="date" class="s-from" value="${escapeHtml(s.from || "")}"></label>` +
      `<label>יתרה (₪)<input type="number" inputmode="decimal" step="any" class="s-bal" value="${s.balance ?? ""}"></label>` +
      `<label>ריבית שנתית (%)<input type="number" inputmode="decimal" step="any" class="s-rate" value="${s.ratePct ?? ""}"></label>` +
      `<label>חודשים<input type="number" inputmode="numeric" step="1" class="s-months" value="${s.months ?? ""}"></label>` +
      `<label>החזר קבוע, ריק לשפיצר<input type="number" inputmode="decimal" step="any" class="s-pay" value="${s.payment ?? ""}"></label>` +
      `<button type="button" class="btn ghost small danger s-del" aria-label="מחק לוח">מחק</button>`;
    row.querySelector(".s-del").addEventListener("click", () => { row.remove(); $("f-loan").dataset.dirty = "1"; });
    $("l-sched").appendChild(row);
  }

  async function onSubmitLoan(event) {
    event.preventDefault();
    if (!state.connected) return showMessage("l-msg", NOT_CONNECTED);
    const raw = {};
    for (const [id, key] of Object.entries(LOAN_FIELDS)) {
      const value = $(id).value.trim();
      raw[key] = key === "buyDate" ? value : value === "" ? null : parseNumber(value);
    }
    raw.schedules = [...document.querySelectorAll("#l-sched .sched-row")].map((row) => ({
      from: row.querySelector(".s-from").value,
      balance: parseNumber(row.querySelector(".s-bal").value),
      ratePct: parseNumber(row.querySelector(".s-rate").value),
      months: parseNumber(row.querySelector(".s-months").value),
      payment: parseNumber(row.querySelector(".s-pay").value),
    })).filter((s) => s.from || s.balance);
    const loan = normalizeLoan(raw);
    if (raw.schedules.length !== loan.schedules.length) return showMessage("l-msg", "בכל לוח סילוקין צריך תאריך, יתרה, ריבית ומספר חודשים.");
    try {
      delete $("f-loan").dataset.dirty;
      await mutate((draft) => { draft.loan = loan; });
      showMessage("l-msg", "נשמר.", true);
    } catch (err) {
      showMessage("l-msg", describeSaveError(err));
    }
  }

  /* ======================================================================
   * 8. Entry forms
   * ==================================================================== */

  /** Reads a numeric input: null when empty. */
  function inputNumber(id) {
    const text = $(id).value.trim();
    return text === "" ? null : Number(text);
  }

  /** Shows a result line under a form. */
  function showMessage(id, text, isSuccess) {
    const el = $(id);
    el.textContent = text;
    el.className = "msg full " + (isSuccess ? "ok" : "err");
  }

  /** The most recent home price, used to prefill the next reading. */
  function latestPrice() {
    return state.readings.length ? priceForReading(state.readings, state.readings.length - 1) || "" : "";
  }

  /** The readings just before and just after a date. */
  function neighbours(date) {
    let previous = null, next = null;
    for (const r of state.readings) {
      if (r.date < date) previous = r;
      else if (r.date > date && !next) next = r;
    }
    return { previous, next };
  }

  /** Live summary under the reading form: what the new reading means since the previous one. */
  function renderReadingPreview() {
    const el = $("r-preview");
    const date = $("r-date").value;
    const meter = inputNumber("r-meter");
    const odo = inputNumber("r-odo");
    const price = inputNumber("r-price");
    if (!date || meter == null) { el.textContent = "מלא את השדות כדי לראות את הסיכום מאז המדידה הקודמת."; return; }
    const { previous } = neighbours(date);
    if (!previous) { el.textContent = "זו תהיה המדידה הראשונה, נקודת ההתחלה של המעקב."; return; }

    const kwh = meter - previous.meter;
    const costHome = kwh * (price > 0 ? price : latestPrice() || 0);
    const charges = state.publicCharges.filter((p) => p.date > previous.date && p.date <= date);
    const costPublic = charges.reduce((sum, p) => sum + p.cost, 0);
    let text = "מאז " + formatDate(previous.date) + " (" + daysBetween(previous.date, date) + " ימים): " +
      formatNumber(kwh, 1) + " קוט״ש בבית. עלות: " + describeSplit(costHome, costPublic);
    // Distance and efficiency are measured from the last reading that has an
    // odometer value, which may be further back than the previous reading.
    const anchor = [...state.readings].reverse().find((r) => r.date < date && r.odo != null);
    if (odo != null && anchor) {
      const km = odo - anchor.odo;
      let kwhSinceAnchor = meter - anchor.meter +
        state.publicCharges.filter((p) => p.date > anchor.date && p.date <= date).reduce((sum, p) => sum + p.kwh, 0);
      const soc = inputNumber("r-soc");
      if (soc != null && anchor.soc != null) kwhSinceAnchor = Math.max(0, kwhSinceAnchor - ((soc - anchor.soc) / 100) * state.settings.batteryKwh);
      text += " · " + formatNumber(km) + " ק״מ";
      if (km > 0) text += " · " + formatNumber((kwhSinceAnchor / km) * 100, 1) + " קוט״ש ל‑100 ק״מ";
      if (km > 0 && soc != null && anchor.soc != null) text += " (מתוקן לפי אחוז הסוללה)";
      if (anchor !== previous) text += " (מחושב מאז " + formatDate(anchor.date) + ", המדידה האחרונה עם קילומטראז׳)";
    }
    el.textContent = text;
  }

  /**
   * The charging-company list: a button that opens a listbox of companies
   * with their symbols. Keyboard: arrows move, Enter or Space picks, Escape
   * closes. "אחר" shows a text field for a company not in the list.
   */
  const providerPicker = (() => {
    const button = () => $("p-provider-btn");
    const list = () => $("p-provider-list");
    const other = () => $("p-provider-other");
    let selected = PROVIDERS[0].id;
    const options = [...PROVIDERS.map((p) => ({ id: p.id, name: p.name, icon: providerIcon(p.id) })), { id: OTHER_PROVIDER, name: "אחר", icon: null }];

    const optionHtml = (o) => (o.icon ? `<img src="${o.icon}" alt="" width="28" height="28">` : '<span class="picker-blank" aria-hidden="true">…</span>') +
      `<span class="picker-name">${escapeHtml(o.name)}</span>`;

    function paint() {
      const current = options.find((o) => o.id === selected);
      button().innerHTML = optionHtml(current) + '<span class="picker-caret" aria-hidden="true">▾</span>';
      $("p-provider").value = selected === OTHER_PROVIDER ? "" : current.name;
      other().hidden = selected !== OTHER_PROVIDER;
      for (const li of list().children) li.setAttribute("aria-selected", String(li.dataset.id === selected));
    }

    function open() {
      list().hidden = false;
      button().setAttribute("aria-expanded", "true");
      // Keep the whole list on screen, clear of the floating bottom bar.
      list().scrollIntoView({ block: "nearest", behavior: "smooth" });
      const current = list().querySelector('[aria-selected="true"]') || list().firstElementChild;
      current.focus();
    }

    function close(focusButton) {
      list().hidden = true;
      button().setAttribute("aria-expanded", "false");
      if (focusButton) button().focus();
    }

    function choose(id) {
      selected = id;
      paint();
      close(id !== OTHER_PROVIDER);
      if (id === OTHER_PROVIDER) other().focus();
    }

    function init() {
      list().innerHTML = options.map((o) => `<li role="option" tabindex="-1" data-id="${o.id}" aria-selected="false">${optionHtml(o)}</li>`).join("");
      button().addEventListener("click", () => (list().hidden ? open() : close(true)));
      button().addEventListener("keydown", (e) => { if (e.key === "ArrowDown" || e.key === "ArrowUp") { e.preventDefault(); open(); } });
      list().addEventListener("click", (e) => { const li = e.target.closest("li"); if (li) choose(li.dataset.id); });
      list().addEventListener("keydown", (e) => {
        const items = [...list().children];
        const index = items.indexOf(document.activeElement);
        if (e.key === "ArrowDown") { e.preventDefault(); items[Math.min(items.length - 1, index + 1)].focus(); }
        else if (e.key === "ArrowUp") { e.preventDefault(); items[Math.max(0, index - 1)].focus(); }
        else if (e.key === "Home") { e.preventDefault(); items[0].focus(); }
        else if (e.key === "End") { e.preventDefault(); items[items.length - 1].focus(); }
        else if (e.key === "Enter" || e.key === " ") { e.preventDefault(); if (index >= 0) choose(items[index].dataset.id); }
        else if (e.key === "Escape" || e.key === "Tab") close(e.key === "Escape");
      });
      document.addEventListener("click", (e) => { if (!list().hidden && !e.target.closest(".picker")) close(false); });
      paint();
    }

    return {
      init,
      /** The company name to save: a listed name, the typed one, or null when "אחר" is empty. */
      value() {
        if (selected !== OTHER_PROVIDER) return $("p-provider").value;
        const typed = other().value.trim();
        return typed ? typed : null;
      },
      /** Back to the default company after a save. */
      reset() { other().value = ""; selected = PROVIDERS[0].id; paint(); },
    };
  })();

  const NOT_CONNECTED = "צריך להתחבר ל‑OneDrive קודם.";

  async function onSubmitReading(event) {
    event.preventDefault();
    const date = $("r-date").value;
    const meter = inputNumber("r-meter");
    const odo = inputNumber("r-odo");
    const price = inputNumber("r-price");
    const soc = inputNumber("r-soc");
    if (!state.connected) return showMessage("r-msg", NOT_CONNECTED);
    if (soc != null && !(soc >= 0 && soc <= 100)) return showMessage("r-msg", "אחוז הסוללה צריך להיות בין 0 ל‑100.");
    if (!date || meter == null || !isFinite(meter)) return showMessage("r-msg", "צריך תאריך וקריאת מונה.");
    if (date > todayIso()) return showMessage("r-msg", "התאריך בעתיד.");
    if (!(price > 0) || price > MAX_PRICE_PER_KWH) return showMessage("r-msg", "צריך מחיר לקוט״ש בשקלים, למשל 0.5451.");

    const { previous, next } = neighbours(date);
    if (previous && meter < previous.meter) return showMessage("r-msg", "המונה נמוך מהמדידה הקודמת (" + formatNumber(previous.meter, 2) + "). בדוק את המספר.");
    if (next && meter > next.meter) return showMessage("r-msg", "המונה גבוה מהמדידה שאחריה (" + formatNumber(next.meter, 2) + ").");
    const previousWithOdo = [...state.readings].reverse().find((r) => r.date < date && r.odo != null);
    if (odo != null && previousWithOdo && odo < previousWithOdo.odo) return showMessage("r-msg", "הקילומטראז׳ נמוך מהמדידה הקודמת (" + formatNumber(previousWithOdo.odo) + ").");
    const nextWithOdo = state.readings.find((r) => r.date > date && r.odo != null);
    if (odo != null && nextWithOdo && odo > nextWithOdo.odo) return showMessage("r-msg", "הקילומטראז׳ גבוה מהמדידה שאחריה (" + formatNumber(nextWithOdo.odo) + ").");

    try {
      // A reading is identified by its date, so saving the same date again corrects it.
      await mutate((draft) => {
        draft.readings = draft.readings.filter((r) => r.date !== date);
        draft.readings.push({ date, meter, odo, price, soc });
      });
      delete $("r-price").dataset.touched;
      showMessage("r-msg", "המדידה נשמרה." + (odo == null ? " בלי קילומטראז׳ הצריכה תחושב יחד עם התקופה הבאה." : ""), true);
      $("r-meter").value = "";
      $("r-odo").value = "";
      $("r-soc").value = "";
      renderReadingPreview();
    } catch (err) {
      showMessage("r-msg", describeSaveError(err));
    }
  }

  async function onSubmitPublicCharge(event) {
    event.preventDefault();
    const date = $("p-date").value;
    const kwh = inputNumber("p-kwh");
    const cost = inputNumber("p-cost");
    const provider = providerPicker.value();
    if (provider == null) return showMessage("p-msg", "כתוב את שם החברה, או בחר חברה מהרשימה.");
    const note = $("p-note").value.trim();
    if (!state.connected) return showMessage("p-msg", NOT_CONNECTED);
    if (!date || !(kwh > 0) || cost == null || cost < 0) return showMessage("p-msg", "צריך תאריך, קוט״ש גדול מאפס וסכום.");
    try {
      // Always appended as a new entry: it never replaces a reading or another charge.
      await mutate((draft) => { draft.publicCharges.push({ id: newChargeId(), date, kwh, cost, provider, note }); });
      showMessage("p-msg", "הטעינה נשמרה.", true);
      for (const id of ["p-kwh", "p-cost", "p-note"]) $(id).value = "";
      providerPicker.reset();
    } catch (err) {
      showMessage("p-msg", describeSaveError(err));
    }
  }

  async function onSubmitFuelSettings(event) {
    event.preventDefault();
    const fuelPrice = inputNumber("s-price");
    const fuelL100 = inputNumber("s-l100");
    if (!state.connected) return showMessage("s-msg", NOT_CONNECTED);
    if (!(fuelPrice > 0) || !(fuelL100 > 0)) return showMessage("s-msg", "שני הערכים צריכים להיות גדולים מאפס.");
    try {
      await mutate((draft) => { draft.settings = { ...draft.settings, fuelPrice, fuelL100 }; });
      showMessage("s-msg", "נשמר.", true);
    } catch (err) {
      showMessage("s-msg", describeSaveError(err));
    }
  }

  async function onSubmitCarSettings(event) {
    event.preventDefault();
    const batteryKwh = inputNumber("s-battery");
    if (!state.connected) return showMessage("c-msg", NOT_CONNECTED);
    if (!(batteryKwh >= MIN_BATTERY_KWH && batteryKwh <= MAX_BATTERY_KWH)) return showMessage("c-msg", "קיבולת הסוללה צריכה להיות בין " + MIN_BATTERY_KWH + " ל‑" + MAX_BATTERY_KWH + " קוט״ש.");
    try {
      await mutate((draft) => { draft.settings = { ...draft.settings, batteryKwh }; });
      showMessage("c-msg", "נשמר.", true);
    } catch (err) {
      showMessage("c-msg", describeSaveError(err));
    }
  }

  /**
   * Trial version only: replaces the test file with a fresh copy of the real
   * readings, so features are tried on realistic data. The real file is only
   * read here, never written.
   */
  async function onClickCopyRealData(event) {
    const button = event.currentTarget;
    const label = "העתק את הנתונים האמיתיים לקובץ הבדיקות";
    if (!IS_BETA) return;
    if (!state.connected) return showMessage("t-msg", NOT_CONNECTED);
    if (!button.dataset.armed) {
      button.dataset.armed = "1";
      button.textContent = "זה יחליף את נתוני הבדיקה. להעתיק?";
      setTimeout(() => { delete button.dataset.armed; button.textContent = label; }, 3500);
      return;
    }
    delete button.dataset.armed;
    button.textContent = label;
    try {
      const accessToken = await getAccessToken();
      if (!accessToken) { requireSignIn(); throw authError(); }
      const real = await readAppFile(APP_FOLDER_URL + REAL_FILE_NAME, accessToken);
      if (real.status === 404) return showMessage("t-msg", "קובץ הנתונים האמיתי עוד לא קיים ב‑OneDrive.");
      if (real.status !== 200) throw authError();
      const copy = normalizeData(real.content);
      await mutate((draft) => { Object.assign(draft, copy); });
      showMessage("t-msg", "הועתקו " + copy.readings.length + " מדידות ו‑" + copy.publicCharges.length + " טעינות ציבוריות מהקובץ האמיתי.", true);
    } catch (err) {
      showMessage("t-msg", err && (err.auth || err.conflict) ? describeSaveError(err) : "ההעתקה נכשלה. בדוק את החיבור ונסה שוב.");
    }
  }

  /** Empties the test file. Only reachable in the trial version; asks for a second click. */
  async function onClickResetTestData(event) {
    const button = event.currentTarget;
    if (!IS_BETA) return;
    if (!state.connected) return showMessage("t-msg", NOT_CONNECTED);
    if (!button.dataset.armed) {
      button.dataset.armed = "1";
      button.textContent = "בטוח? רוקן";
      setTimeout(() => { delete button.dataset.armed; button.textContent = "רוקן את קובץ הבדיקות"; }, 3500);
      return;
    }
    delete button.dataset.armed;
    button.textContent = "רוקן את קובץ הבדיקות";
    try {
      await mutate((draft) => { draft.readings = []; draft.publicCharges = []; draft.settings = { ...DEFAULT_SETTINGS }; draft.loan = null; });
      showMessage("t-msg", "קובץ הבדיקות רוקן.", true);
    } catch (err) {
      showMessage("t-msg", describeSaveError(err));
    }
  }

  /** Delete buttons ask for a second click instead of opening a dialog. */
  async function onClickDelete(event) {
    const button = event.target.closest("[data-del]");
    if (!button || !state.connected) return;
    if (!button.dataset.armed) {
      button.dataset.armed = "1";
      button.textContent = "בטוח? מחק";
      setTimeout(() => {
        if (button.isConnected) { delete button.dataset.armed; button.textContent = "מחק"; }
      }, 3500);
      return;
    }
    button.disabled = true;
    const separator = button.dataset.del.indexOf("/");
    const list = button.dataset.del.slice(0, separator);
    const key = button.dataset.del.slice(separator + 1);
    try {
      await mutate((draft) => {
        if (list === "public") draft.publicCharges = draft.publicCharges.filter((p) => p.id !== key);
        else draft.readings = draft.readings.filter((r) => r.date !== key);
      });
    } catch (err) {
      button.disabled = false;
      button.textContent = "נכשל, נסה שוב";
      delete button.dataset.armed;
    }
  }

  const TABS = ["overview", "loan", "add", "history", "settings"];

  function showTab(tab) {
    document.querySelectorAll(".tabs button").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.tab === tab)));
    for (const name of TABS) $("v-" + name).hidden = name !== tab;
    try { localStorage.setItem(STORE.tab, tab); } catch (e) { /* storage unavailable */ }
    if (tab === "overview" || tab === "loan") render(); // charts measure their container, so draw once it is visible
  }

  /** Switches the entry screen between the monthly reading and the public charge forms. */
  function showEntryForm(which) {
    $("seg-reading").setAttribute("aria-pressed", String(which === "reading"));
    $("seg-public").setAttribute("aria-pressed", String(which === "public"));
    $("p-reading").hidden = which !== "reading";
    $("p-public").hidden = which !== "public";
  }

  /* ======================================================================
   * 9. Import and export
   * ==================================================================== */

  /**
   * CSV columns. `key` is the internal field, `header` is what the export
   * writes, and `aliases` are other spellings the import accepts. Headers are
   * compared after stripping everything except letters and digits, so
   * "מחיר לקוט״ש ₪" and "מחיר לקוטש" are the same column.
   */
  const CSV_COLUMNS = [
    { key: "type", header: "סוג", aliases: ["type"] },
    { key: "date", header: "תאריך", aliases: ["date"] },
    { key: "meter", header: "מונה קוט״ש", aliases: ["מונה", "meter", "meter_kwh"] },
    { key: "odo", header: "קילומטראז׳", aliases: ["קמ", "odo", "odometer", "odometer_km", "km"] },
    { key: "price", header: "מחיר לקוט״ש", aliases: ["מחיר", "price", "price_per_kwh"] },
    { key: "kwh", header: "קוט״ש", aliases: ["kwh"] },
    { key: "cost", header: "עלות", aliases: ["סכום", "cost"] },
    { key: "provider", header: "חברה", aliases: ["provider", "company"] },
    { key: "note", header: "הערה", aliases: ["note"] },
    { key: "id", header: "מזהה", aliases: ["id"] },
    { key: "soc", header: "אחוז סוללה", aliases: ["סוללה", "soc", "battery"] },
  ];
  const TYPE_READING = "מדידה";
  const TYPE_PUBLIC = "טעינה ציבורית";

  /** Lower-cases and strips everything except Hebrew letters, Latin letters and digits. */
  const simplify = (text) => String(text ?? "").toLowerCase().replace(/[^a-z0-9א-ת]/g, "");

  /** Builds the CSV text for the given data (UTF-8 with BOM so Excel shows Hebrew correctly). */
  function toCsv(data) {
    const rows = [CSV_COLUMNS.map((c) => c.header)];
    for (const r of data.readings) rows.push([TYPE_READING, r.date, r.meter, r.odo ?? "", r.price ?? "", "", "", "", "", "", r.soc ?? ""]);
    for (const p of data.publicCharges) rows.push([TYPE_PUBLIC, p.date, "", "", "", p.kwh, p.cost, p.provider || "", p.note || "", p.id, ""]);
    const quote = (cell) => '"' + String(cell).replace(/"/g, '""') + '"';
    return "﻿" + rows.map((row) => row.map(quote).join(",")).join("\r\n");
  }

  /**
   * Minimal RFC 4180 CSV reader: quoted fields, doubled quotes, CRLF or LF.
   * The delimiter (comma, semicolon or tab) is detected from the header line,
   * because Excel picks it according to the computer's regional settings.
   * @returns {string[][]} one entry per line, blank lines included, so that
   *   a row's index + 1 is its line number in the file.
   */
  function parseCsv(text) {
    text = String(text).replace(/^﻿/, "");
    const headerLine = text.split(/\r?\n/, 1)[0];
    const delimiter = [",", ";", "\t"].reduce((best, d) => (headerLine.split(d).length > headerLine.split(best).length ? d : best), ",");
    const rows = [];
    let row = [], cell = "", inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; }
        else if (ch === '"') inQuotes = false;
        else cell += ch;
      } else if (ch === '"') inQuotes = true;
      else if (ch === delimiter) { row.push(cell); cell = ""; }
      else if (ch === "\n" || ch === "\r") {
        if (ch === "\r" && text[i + 1] === "\n") i++;
        row.push(cell); rows.push(row); row = []; cell = "";
      } else cell += ch;
    }
    if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }

  const isBlankRow = (cells) => cells.every((c) => c.trim() === "");

  /**
   * Reads an import file (JSON or CSV) into validated entries.
   * Nothing is thrown for bad rows: they are collected in `rejected` so the
   * user sees exactly which lines were skipped and why.
   * @returns {{readings:Reading[], publicCharges:PublicCharge[], settings:(object|null),
   *            rejected:{where:string, reason:string}[]}}
   * @throws {Error} with a Hebrew message when the file as a whole is unusable.
   */
  function parseImportText(text) {
    const result = { readings: [], publicCharges: [], settings: null, rejected: [] };
    const accept = (kind, raw, where) => {
      const outcome = kind === "reading" ? normalizeReading(raw) : normalizePublicCharge(raw);
      if (outcome.error) result.rejected.push({ where, reason: outcome.error });
      else (kind === "reading" ? result.readings : result.publicCharges).push(outcome.value);
    };

    const trimmed = String(text).replace(/^﻿/, "").trim();
    if (trimmed.startsWith("{")) {
      let json;
      try { json = JSON.parse(trimmed); } catch (e) { throw new Error("קובץ ה‑JSON לא תקין: " + e.message); }
      const readings = Array.isArray(json.readings) ? json.readings : [];
      const charges = Array.isArray(json.public) ? json.public : [];
      if (!Array.isArray(json.readings) && !Array.isArray(json.public)) throw new Error("בקובץ ה‑JSON חסרות הרשימות readings ו‑public.");
      readings.forEach((raw, i) => accept("reading", raw, "readings[" + i + "]"));
      charges.forEach((raw, i) => accept("public", raw, "public[" + i + "]"));
      if (json.settings && (json.settings.fuelPrice > 0 || json.settings.fuelL100 > 0 || json.settings.batteryKwh > 0)) result.settings = json.settings;
      return result;
    }

    const rows = parseCsv(trimmed);
    if (rows.filter((r) => !isBlankRow(r)).length < 2) throw new Error("בקובץ אין שורות נתונים. השורה הראשונה צריכה להיות שורת כותרות.");
    const columnOf = {};
    rows[0].forEach((header, index) => {
      const name = simplify(header);
      const column = CSV_COLUMNS.find((c) => simplify(c.header) === name || c.aliases.some((a) => simplify(a) === name));
      if (column && !(column.key in columnOf)) columnOf[column.key] = index;
    });
    if (!("date" in columnOf)) throw new Error("לא נמצאה עמודת תאריך. בדוק את שורת הכותרות.");
    if (!("meter" in columnOf) && !("kwh" in columnOf)) throw new Error("לא נמצאה עמודת מונה קוט״ש או עמודת קוט״ש.");

    rows.slice(1).forEach((cells, i) => {
      if (isBlankRow(cells)) return;
      const where = "שורה " + (i + 2);
      const raw = {};
      for (const key of Object.keys(columnOf)) raw[key] = (cells[columnOf[key]] ?? "").trim();
      // The type column is optional: without it, a row with a meter value is a reading.
      const type = simplify(raw.type);
      let kind;
      if (["מדידה", "reading", "home", "בית"].includes(type)) kind = "reading";
      else if (["טעינהציבורית", "ציבורי", "ציבורית", "public"].includes(type)) kind = "public";
      else if (type === "") kind = raw.meter ? "reading" : raw.kwh ? "public" : null;
      if (!kind) { result.rejected.push({ where, reason: type ? "סוג לא מוכר: " + raw.type : "שורה בלי מונה ובלי קוט״ש" }); return; }
      accept(kind, raw, where);
    });
    return result;
  }

  /** Two public charges with the same date, energy and cost are treated as the same session. */
  const chargeSignature = (p) => p.date + "|" + p.kwh + "|" + p.cost;

  /**
   * Merges imported entries into the current data without discarding anything.
   *
   *   Reading        - matched by date. Same date: the file's values replace the
   *                    stored ones. New date: added, even when another reading
   *                    exists in the same calendar month.
   *   Public charge  - matched by id, otherwise by date + kWh + cost. A match is
   *                    updated (provider, note); anything else is added.
   *   Everything not mentioned in the file is kept as it is.
   *
   * @param {AppData} current
   * @param {ReturnType<typeof parseImportText>} incoming
   * @returns {{data: AppData, counts: object, warnings: string[]}}
   */
  function planMerge(current, incoming) {
    const counts = {
      readings: { added: 0, updated: 0, unchanged: 0 },
      publicCharges: { added: 0, updated: 0, unchanged: 0 },
      settingsChanged: false,
    };
    const same = (a, b, keys) => keys.every((k) => (a[k] ?? null) === (b[k] ?? null));

    const readings = new Map(current.readings.map((r) => [r.date, { ...r }]));
    for (let item of incoming.readings) {
      const existing = readings.get(item.date);
      // Older files have no battery column: an empty value keeps the stored one.
      if (existing && item.soc == null && existing.soc != null) item = { ...item, soc: existing.soc };
      if (!existing) counts.readings.added += 1;
      else if (same(existing, item, ["meter", "odo", "price", "soc"])) { counts.readings.unchanged += 1; continue; }
      else counts.readings.updated += 1;
      readings.set(item.date, { ...item });
    }

    const charges = current.publicCharges.map((p) => ({ ...p }));
    for (const item of incoming.publicCharges) {
      let index = item.id ? charges.findIndex((p) => p.id === item.id) : -1;
      if (index < 0) index = charges.findIndex((p) => chargeSignature(p) === chargeSignature(item));
      if (index < 0) {
        charges.push({ ...item, id: item.id && !charges.some((p) => p.id === item.id) ? item.id : newChargeId() });
        counts.publicCharges.added += 1;
      } else if (same(charges[index], item, ["date", "kwh", "cost", "provider", "note"])) {
        counts.publicCharges.unchanged += 1;
      } else {
        charges[index] = { ...item, id: charges[index].id };
        counts.publicCharges.updated += 1;
      }
    }

    const settings = { ...current.settings };
    if (incoming.settings) {
      for (const key of ["fuelPrice", "fuelL100", "batteryKwh"]) {
        const value = Number(incoming.settings[key]);
        if (key === "batteryKwh" && !(value >= MIN_BATTERY_KWH && value <= MAX_BATTERY_KWH)) continue;
        if (value > 0 && value !== settings[key]) { settings[key] = value; counts.settingsChanged = true; }
      }
    }

    const data = { readings: [...readings.values()].sort(byDate), publicCharges: charges.sort(byDate), settings, loan: current.loan };

    // The meter and the odometer only ever go up; a drop after the merge points at a typo in the file.
    const warnings = [];
    for (let i = 1; i < data.readings.length; i++) {
      if (data.readings[i].meter < data.readings[i - 1].meter) {
        warnings.push("קריאת המונה ב‑" + formatDate(data.readings[i].date) + " נמוכה מהקריאה שלפניה (" + formatDate(data.readings[i - 1].date) + ").");
      }
    }
    const withOdo = data.readings.filter((r) => r.odo != null);
    for (let i = 1; i < withOdo.length; i++) {
      if (withOdo[i].odo < withOdo[i - 1].odo) {
        warnings.push("הקילומטראז׳ ב‑" + formatDate(withOdo[i].date) + " נמוך מהקילומטראז׳ ב‑" + formatDate(withOdo[i - 1].date) + ".");
      }
    }
    return { data, counts, warnings };
  }

  /** Offers a generated file to the browser's download flow. */
  function downloadFile(name, text, mimeType) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([text], { type: mimeType }));
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 5000);
  }

  /** A small, valid CSV the user can open in Excel and fill in. */
  function templateCsv() {
    return toCsv({
      readings: [{ date: "2026-10-01", meter: 4350.2, odo: 56200, price: 0.5451, soc: 80 }],
      publicCharges: [{ id: "", date: "2026-10-12", kwh: 32.5, cost: 61, provider: "שם החברה", note: "נסיעה לצפון" }],
    });
  }

  /** The parsed file waiting for the user's confirmation, or null. */
  let pendingImport = null;

  function setImportMessage(text, kind) {
    const el = $("d-msg");
    el.textContent = text;
    el.className = "msg" + (kind ? " " + kind : "");
  }

  /** Describes "N חדשות, M יעודכנו, K ללא שינוי". */
  function describeCounts(c) {
    return c.added + " חדשות, " + c.updated + " יעודכנו, " + c.unchanged + " ללא שינוי";
  }

  /** Step 1 of an import: read the file and show what would change. Nothing is saved yet. */
  async function onImportFileChosen(event) {
    const file = event.target.files[0];
    event.target.value = ""; // allow choosing the same file again
    if (!file) return;
    pendingImport = null;
    $("imp-box").hidden = true;
    let parsed;
    try {
      parsed = parseImportText(await file.text());
    } catch (err) {
      return setImportMessage(err.message, "err");
    }
    if (!parsed.readings.length && !parsed.publicCharges.length) {
      const first = parsed.rejected[0];
      return setImportMessage("לא נמצאה בקובץ אף רשומה תקינה." + (first ? " לדוגמה, " + first.where + ": " + first.reason + "." : ""), "err");
    }

    const plan = planMerge(snapshot(), parsed);
    const lines = [
      "מדידות: " + describeCounts(plan.counts.readings) + ".",
      "טעינות ציבוריות: " + describeCounts(plan.counts.publicCharges) + ".",
    ];
    if (plan.counts.settingsChanged) lines.push("ההגדרות, רכב הבנזין וקיבולת הסוללה, יעודכנו לפי הקובץ.");
    lines.push("רשומות קיימות שלא מופיעות בקובץ נשארות ללא שינוי.");
    for (const warning of plan.warnings) lines.push("שים לב: " + warning);
    if (parsed.rejected.length) {
      const shown = parsed.rejected.slice(0, 5).map((r) => r.where + " (" + r.reason + ")").join(", ");
      lines.push(parsed.rejected.length + " שורות לא ייובאו: " + shown + (parsed.rejected.length > 5 ? " ועוד." : "."));
    }
    $("imp-text").innerHTML = "<strong>" + escapeHtml(file.name) + "</strong>" + lines.map((l) => "<br>" + escapeHtml(l)).join("");
    pendingImport = parsed;
    $("imp-box").hidden = false;
    setImportMessage("");
  }

  /** Step 2 of an import: merge and save. The merge is recomputed so it uses the latest data. */
  async function onImportConfirmed() {
    if (!pendingImport) return;
    if (!state.connected) return setImportMessage(NOT_CONNECTED, "err");
    const parsed = pendingImport;
    try {
      let plan;
      await mutate((draft) => {
        plan = planMerge(draft, parsed);
        Object.assign(draft, plan.data);
      });
      pendingImport = null;
      $("imp-box").hidden = true;
      const changed = plan.counts.readings.added + plan.counts.readings.updated + plan.counts.publicCharges.added + plan.counts.publicCharges.updated;
      const summary = changed === 0 ? "לא היו שינויים לשמור."
        : changed === 1 ? "רשומה אחת נוספה או עודכנה ונשמרה ב‑OneDrive."
        : changed + " רשומות נוספו או עודכנו ונשמרו ב‑OneDrive.";
      setImportMessage("הייבוא הושלם: " + summary, "ok");
    } catch (err) {
      setImportMessage(describeSaveError(err), "err");
    }
  }

  function onImportCancelled() {
    pendingImport = null;
    $("imp-box").hidden = true;
    setImportMessage("");
  }

  /* ======================================================================
   * 10. OneDrive storage
   *
   * Sign-in uses the OAuth 2.0 authorization code flow with PKCE, the flow
   * meant for pages that cannot keep a secret. The page never sees the
   * user's password: Microsoft's own sign-in page handles it and returns a
   * short-lived access token plus a refresh token, both kept in localStorage
   * on this device only.
   * ==================================================================== */

  /** base64url without padding, as PKCE requires. */
  function base64Url(buffer) {
    return btoa(String.fromCharCode(...new Uint8Array(buffer))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  /** Sends the browser to Microsoft's sign-in page. */
  async function beginSignIn() {
    const verifier = base64Url(crypto.getRandomValues(new Uint8Array(48)));
    const stateToken = base64Url(crypto.getRandomValues(new Uint8Array(12)));
    const challenge = base64Url(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
    sessionStorage.setItem(STORE.pkce, JSON.stringify({ verifier, stateToken }));
    location.assign(AUTH_URL + "authorize?" + new URLSearchParams({
      client_id: CLIENT_ID,
      response_type: "code",
      redirect_uri: REDIRECT_URI,
      response_mode: "query",
      scope: SCOPE,
      code_challenge: challenge,
      code_challenge_method: "S256",
      state: stateToken,
    }));
  }

  /** Calls the token endpoint and stores the result. Errors carry `auth = true`. */
  async function requestToken(params) {
    const response = await fetch(AUTH_URL + "token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: CLIENT_ID, scope: SCOPE, ...params }),
    });
    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      const error = new Error(body.error_description || body.error || "token");
      error.auth = true;
      throw error;
    }
    const previous = storage.get(STORE.token) || {};
    const token = {
      access: body.access_token,
      expires: Date.now() + (body.expires_in - 60) * 1000,
      refresh: body.refresh_token || previous.refresh,
    };
    storage.set(STORE.token, token);
    return token;
  }

  /** @returns {Promise<string|null>} a valid access token, refreshed when needed; null when sign-in is required. */
  async function getAccessToken() {
    const token = storage.get(STORE.token);
    if (!token) return null;
    if (token.access && token.expires > Date.now()) return token.access;
    if (!token.refresh) return null;
    try {
      return (await requestToken({ grant_type: "refresh_token", refresh_token: token.refresh })).access;
    } catch (err) {
      if (err.auth) { storage.remove(STORE.token); return null; }
      throw err;
    }
  }

  /** Shows the sign-in panel, optionally with an explanation. */
  function requireSignIn(message) {
    state.connected = false;
    state.loading = false;
    $("signin").hidden = false;
    if (message) $("signin-text").textContent = message;
    render();
  }

  function authError() {
    const error = new Error("auth");
    error.auth = true;
    return error;
  }

  /**
   * Reads one JSON file from the app folder.
   * Graph's /content endpoint redirects in a way browsers refuse, so the file
   * is fetched in two steps: its metadata (with a pre-authenticated download
   * URL), then the URL itself.
   * @returns {Promise<{status:number, etag:(string|null), content:(object|null)}>}
   *   status 200 with the parsed content, or 401 / 404 with no content.
   * @throws on any other failure, so a broken read is never mistaken for an empty file.
   */
  async function readAppFile(fileUrl, accessToken) {
    // cache: "no-store" - the download URL is short-lived, so a cached answer is useless.
    const meta = await fetch(fileUrl + "?select=id,eTag,@microsoft.graph.downloadUrl", {
      headers: { Authorization: "Bearer " + accessToken },
      cache: "no-store",
    });
    if (meta.status === 401 || meta.status === 404) return { status: meta.status, etag: null, content: null };
    if (!meta.ok) throw new Error("graph " + meta.status);
    const item = await meta.json();
    const downloadUrl = item["@microsoft.graph.downloadUrl"];
    if (!downloadUrl) throw new Error("missing downloadUrl");
    const file = await fetch(downloadUrl, { cache: "no-store" });
    if (!file.ok) throw new Error("download " + file.status);
    return { status: 200, etag: item.eTag || null, content: await file.json() };
  }

  /** Loads this copy's data file from OneDrive into the page. */
  async function loadFromOneDrive() {
    const accessToken = await getAccessToken();
    if (!accessToken) return requireSignIn();
    state.loading = true;
    render();

    const file = await readAppFile(FILE_URL, accessToken);
    if (file.status === 401) { storage.remove(STORE.token); return requireSignIn(); }
    state.etag = file.etag; // null on first use, when the file does not exist yet
    if (file.status === 200) applyData(normalizeData(file.content));
    state.connected = true;
    state.loading = false;
    $("signin").hidden = true;
    render();
  }

  /**
   * Writes the whole data file. The If-Match header makes OneDrive reject the
   * write (412) when the file changed since it was loaded, so two devices
   * cannot silently overwrite each other.
   */
  async function saveToOneDrive(data) {
    const accessToken = await getAccessToken();
    if (!accessToken) { requireSignIn(); throw authError(); }
    const headers = { Authorization: "Bearer " + accessToken, "Content-Type": "application/json" };
    if (state.etag) headers["If-Match"] = state.etag;
    const response = await fetch(FILE_URL + ":/content", { method: "PUT", headers, body: JSON.stringify(serializeData(data), null, 1) });
    if (response.status === 412) {
      await loadFromOneDrive().catch(() => {});
      const error = new Error("conflict");
      error.conflict = true;
      throw error;
    }
    if (response.status === 401) { storage.remove(STORE.token); requireSignIn(); throw authError(); }
    if (!response.ok) throw new Error("graph " + response.status);
    state.etag = (await response.json()).eTag || null;
  }

  /**
   * The only way data changes: copy the data, let `change` edit the copy,
   * save it to OneDrive, and only then show it. A failed save leaves the
   * screen and the cache untouched.
   * @param {(draft: AppData) => void} change
   */
  async function mutate(change) {
    const draft = snapshot();
    change(draft);
    await saveToOneDrive(draft);
    applyData(draft);
    render();
  }

  /** Hebrew explanation for a failed save. */
  function describeSaveError(err) {
    if (err && err.conflict) return "הקובץ עודכן ממכשיר אחר. טענתי את הגרסה החדשה, נסה שוב.";
    if (err && err.auth) return "ההתחברות פגה. התחבר שוב ל‑OneDrive.";
    return "השמירה נכשלה. בדוק את החיבור לאינטרנט ונסה שוב.";
  }

  /**
   * Signs this browser out. The token is shared by the main and the trial
   * copy, so both offline caches are removed: no readings stay on the device.
   */
  function signOut() {
    storage.remove(STORE.token);
    storage.remove(CACHE_KEYS.main);
    storage.remove(CACHE_KEYS.beta);
    location.reload();
  }

  /** Completes the redirect back from Microsoft's sign-in page, if this page load is one. */
  async function finishSignInRedirect() {
    const query = new URLSearchParams(location.search);
    if (!query.has("code") && !query.has("error")) return true;
    const pkce = JSON.parse(sessionStorage.getItem(STORE.pkce) || "{}");
    sessionStorage.removeItem(STORE.pkce);
    history.replaceState(null, "", REDIRECT_URI); // drop the one-time code from the address bar
    if (query.has("error")) { requireSignIn("ההתחברות לא הושלמה (" + query.get("error") + "). אפשר לנסות שוב."); return false; }
    if (!pkce.verifier || pkce.stateToken !== query.get("state")) { requireSignIn("ההתחברות לא הושלמה. נסה שוב."); return false; }
    try {
      await requestToken({ grant_type: "authorization_code", code: query.get("code"), redirect_uri: REDIRECT_URI, code_verifier: pkce.verifier });
      return true;
    } catch (err) {
      requireSignIn("ההתחברות נכשלה: " + err.message);
      return false;
    }
  }

  /* ======================================================================
   * 11. Start-up
   * ==================================================================== */

  function bindEvents() {
    document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
    $("seg-reading").addEventListener("click", () => showEntryForm("reading"));
    $("seg-public").addEventListener("click", () => showEntryForm("public"));

    for (const id of ["r-date", "r-meter", "r-odo", "r-price", "r-soc"]) $(id).addEventListener("input", renderReadingPreview);
    $("r-price").addEventListener("input", () => { $("r-price").dataset.touched = "1"; });
    $("f-reading").addEventListener("submit", onSubmitReading);
    $("f-public").addEventListener("submit", onSubmitPublicCharge);
    $("f-fuel").addEventListener("submit", onSubmitFuelSettings);
    $("f-car").addEventListener("submit", onSubmitCarSettings);
    $("f-loan").addEventListener("submit", onSubmitLoan);
    $("f-loan").addEventListener("input", () => { $("f-loan").dataset.dirty = "1"; });
    $("l-add-sched").addEventListener("click", () => { addScheduleRow({ from: "", balance: "", ratePct: "", months: "" }); $("f-loan").dataset.dirty = "1"; });
    providerPicker.init();
    document.addEventListener("click", onClickDelete);

    $("export").addEventListener("click", () => downloadFile("zeekr-electric-" + todayIso() + ".csv", toCsv(state), "text/csv"));
    $("dl-json").addEventListener("click", () => downloadFile("zeekr-data-" + todayIso() + ".json", JSON.stringify(serializeData(state), null, 2), "application/json"));
    $("dl-template").addEventListener("click", () => downloadFile("zeekr-import-template.csv", templateCsv(), "text/csv"));
    $("imp-file").addEventListener("change", onImportFileChosen);
    $("imp-go").addEventListener("click", onImportConfirmed);
    $("imp-cancel").addEventListener("click", onImportCancelled);

    $("signin-btn").addEventListener("click", () => beginSignIn().catch(() => {
      $("signin-text").textContent = "לא הצלחתי לפתוח את דף ההתחברות. הדף חייב לרוץ בכתובת https.";
    }));
    $("signout").addEventListener("click", signOut);
    $("test-reset").addEventListener("click", onClickResetTestData);
    $("test-copy").addEventListener("click", onClickCopyRealData);

    let resizeTimer;
    window.addEventListener("resize", () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(render, 150); });
  }

  async function start() {
    state.loan = normalizeLoan(null);
    bindEvents();
    $("r-date").value = $("p-date").value = todayIso();
    $("app-version").textContent = APP_VERSION;
    $("head-version").textContent = "v" + APP_VERSION;
    for (const el of document.querySelectorAll(".file-name")) el.textContent = DATA_FILE_NAME;
    $("beta-banner").hidden = !IS_BETA;
    $("beta-panel").hidden = !IS_BETA;
    $("other-version").textContent = IS_BETA ? "לגרסה הראשית" : "לגרסת הניסוי";
    $("other-version").href = IS_BETA ? "../" : "beta/";

    let tab = "overview";
    try { tab = localStorage.getItem(STORE.tab) || tab; } catch (e) { /* storage unavailable */ }
    showTab(TABS.includes(tab) ? tab : "overview");

    // Paint the last known data immediately; OneDrive replaces it when it answers.
    const cached = storage.get(STORE.cache);
    if (cached) applyData(normalizeData(cached));
    render();
    loadFuelPrices(render);

    if (!(await finishSignInRedirect())) return;
    try {
      await loadFromOneDrive();
    } catch (err) {
      state.loading = false;
      render();
      $("status").textContent = "לא מצליח להגיע ל‑OneDrive כרגע" + (state.readings.length ? " · מוצג עותק שמור" : "");
      $("status").className = "pill warn";
    }
  }

  // The pure functions are exposed for automated checks and for debugging in the console.
  const api = { scheduleBalance, buildLoanModel, normalizeLoan, annuityPayment, addMonths, monthsBetween, parseDate, parseNumber, parseCsv, parseImportText, planMerge, normalizeData, serializeData, buildModel, buildMonths, toCsv };
  if (typeof window !== "undefined") window.ZeekrTracker = api;
  if (typeof document !== "undefined") start();
  else if (typeof globalThis !== "undefined") globalThis.ZeekrTracker = api;
})();
