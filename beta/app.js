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
 *   5. Calculations (periods, efficiency, totals)
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
 *     readings: [{ date, meter, odo, price }],
 *     public:   [{ id, date, kwh, cost, provider, note }],
 *     settings: { fuelPrice, fuelL100 }
 *   }
 *
 *   reading - one look at the home charger's meter.
 *     date   "YYYY-MM-DD". Unique: it is the reading's identity.
 *     meter  cumulative kWh shown on the charger.
 *     odo    car odometer in km, or null when it was not noted.
 *     price  home price per kWh (ILS) for the period that ENDS at this
 *            reading. Storing it per reading is what makes a price change
 *            apply from that reading onwards and never retroactively.
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
  const APP_VERSION = "2.3";

  /** Application (client) ID of the Microsoft Entra app registration. Public, not a secret. */
  const CLIENT_ID = "66f6788f-925a-471b-b98f-87d662b76db6";

  /** Microsoft identity platform endpoint for personal accounts. */
  const AUTH_URL = "https://login.microsoftonline.com/consumers/oauth2/v2.0/";

  /** Least privilege: the app can only see its own folder (OneDrive/Apps/Zeekr Tracker). */
  const SCOPE = "Files.ReadWrite.AppFolder offline_access";

  /** Must match the redirect URI registered for the app, character for character. */
  const REDIRECT_URI = location.origin + location.pathname;

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
  const STORE = { token: "zk-tok", cache: IS_BETA ? "zk-data-test" : "zk-data", tab: "zk-tab", pkce: "zk-pkce" };

  const DATA_FORMAT_VERSION = 2;
  const DEFAULT_SETTINGS = { fuelPrice: 7.3, fuelL100: 7.5 };
  const MS_PER_DAY = 864e5;
  /** A reading older than this many days turns the status pill into a reminder. */
  const STALE_AFTER_DAYS = 35;
  /** Charts show at most this many most-recent periods. */
  const CHART_MAX_POINTS = 14;
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
    /** True once the OneDrive file was loaded, so saving is safe. */
    connected: false,
    loading: false,
    /** eTag of the OneDrive file as last seen; sent back to detect concurrent edits. */
    etag: null,
  };

  /**
   * @typedef {{date:string, meter:number, odo:(number|null), price:(number|null)}} Reading
   * @typedef {{id:string, date:string, kwh:number, cost:number, provider:string, note:string}} PublicCharge
   * @typedef {{readings:Reading[], publicCharges:PublicCharge[], settings:{fuelPrice:number, fuelL100:number}}} AppData
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

  /** "2026-09-02" -> "2.9" (chart axis labels) */
  function formatDateShort(isoDate) {
    const [, m, d] = isoDate.split("-");
    return +d + "." + +m;
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
    if (!date) return { error: "תאריך חסר או לא תקין" };
    if (meter == null || !isFinite(meter) || meter < 0) return { error: "קריאת מונה חסרה או לא תקינה" };
    if (odo != null && (!isFinite(odo) || odo < 0)) return { error: "קילומטראז׳ לא תקין" };
    if (price != null && (!isFinite(price) || price < 0 || price > MAX_PRICE_PER_KWH)) return { error: "מחיר לקוט״ש לא תקין" };
    return { value: { date, meter, odo, price: price > 0 ? price : null } };
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
    return {
      readings: [...readingsByDate.values()].sort(byDate),
      publicCharges: publicCharges.sort(byDate),
      settings,
    };
  }

  /** Converts application data to the JSON structure stored in zeekr-data.json, entries ordered by date. */
  function serializeData(data) {
    data = { ...data, readings: [...data.readings].sort(byDate), publicCharges: [...data.publicCharges].sort(byDate) };
    return {
      app: "zeekr-tracker",
      version: DATA_FORMAT_VERSION,
      updated: new Date().toISOString(),
      readings: data.readings.map((r) => ({ date: r.date, meter: r.meter, odo: r.odo ?? null, price: r.price ?? null })),
      public: data.publicCharges.map((p) => ({ id: p.id, date: p.date, kwh: p.kwh, cost: p.cost, provider: p.provider || "", note: p.note || "" })),
      settings: { ...data.settings },
    };
  }

  /** A detached copy of the current data, safe to modify before saving. */
  function snapshot() {
    return {
      readings: state.readings.map((r) => ({ ...r })),
      publicCharges: state.publicCharges.map((p) => ({ ...p })),
      settings: { ...state.settings },
    };
  }

  /** Replaces the in-memory data and refreshes the offline cache. */
  function applyData(data) {
    const clean = normalizeData(serializeData(data));
    state.readings = clean.readings;
    state.publicCharges = clean.publicCharges;
    state.settings = clean.settings;
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
   * @param {Reading[]} readings sorted by date
   * @param {PublicCharge[]} publicCharges
   */
  function buildModel(readings, publicCharges) {
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
          const span = {
            from: readings[anchor].date,
            to: readings[i].date,
            days: daysBetween(readings[anchor].date, readings[i].date),
            km, kwh, cost,
            efficiency: (kwh / km) * 100,
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
    const driven = { km: 0, kwh: 0, cost: 0, days: 0 };
    for (const s of spans) { driven.km += s.km; driven.kwh += s.kwh; driven.cost += s.cost; driven.days += s.days; }

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

  /* ======================================================================
   * 6. Charts
   * ==================================================================== */

  const CHART_HEIGHT = 210;

  function svgOpen(width, height) {
    return `<svg viewBox="0 0 ${width} ${height}" width="${width}" height="${height}" role="img" font-family="Heebo,sans-serif">`;
  }

  /** Horizontal grid line with its value label on the left axis. */
  function gridLine(x1, x2, y, label) {
    return `<line x1="${x1}" x2="${x2}" y1="${y}" y2="${y}" stroke="var(--line)" stroke-width="1"/>` +
      `<text x="${x1 - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${label}</text>`;
  }

  /**
   * Stacked bar chart: home cost at the bottom, public cost on top.
   * @param {HTMLElement} el
   * @param {{label:string, home:number, pub:number}[]} items
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
      svg += `<rect x="${x}" y="${baseY - hHome}" width="${barW}" height="${hHome}" rx="3" fill="var(--accent)"/>`;
      if (hPub > 0) svg += `<rect x="${x}" y="${baseY - hHome - hPub}" width="${barW}" height="${hPub}" rx="3" fill="var(--public)"/>`;
      if (step > 34) svg += `<text x="${x + barW / 2}" y="${baseY - hHome - hPub - 5}" text-anchor="middle" font-size="11" font-weight="500" fill="var(--fg)">${Math.round(it.home + it.pub)}</text>`;
      svg += `<text x="${x + barW / 2}" y="${CHART_HEIGHT - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${it.label}</text>`;
    });
    el.innerHTML = svg + "</svg>";
  }

  /**
   * Line chart of efficiency per span with a dashed line for the overall average.
   * @param {HTMLElement} el
   * @param {{label:string, value:number}[]} points
   * @param {number} average
   */
  function drawEfficiencyChart(el, points, average) {
    if (!points.length) { el.innerHTML = '<p class="empty">יופיע כאן אחרי שתי מדידות עם קילומטראז׳.</p>'; return; }
    points = points.slice(-CHART_MAX_POINTS);
    const width = Math.max(280, el.clientWidth || 600);
    const left = 38, bottom = 26, top = 22;
    const plotW = width - left - 24, plotH = CHART_HEIGHT - bottom - top;
    const values = points.map((p) => p.value).concat(average);
    const low = Math.max(0, Math.floor(Math.min(...values) - 2));
    const high = Math.ceil(Math.max(...values) + 2);
    const xAt = (i) => left + 12 + (points.length > 1 ? (plotW * i) / (points.length - 1) : plotW / 2);
    const yAt = (v) => top + plotH - (plotH * (v - low)) / (high - low);

    let svg = svgOpen(width, CHART_HEIGHT);
    for (let g = 0; g <= 3; g++) {
      const v = low + ((high - low) * g) / 3;
      svg += gridLine(left, width - 8, yAt(v), formatNumber(v, 0));
    }
    svg += `<line x1="${left}" x2="${width - 8}" y1="${yAt(average)}" y2="${yAt(average)}" stroke="var(--muted)" stroke-width="1.5" stroke-dasharray="5 4"/>`;
    if (points.length > 1) {
      const path = points.map((p, i) => (i ? "L" : "M") + xAt(i) + " " + yAt(p.value)).join(" ");
      svg += `<path d="${path} L${xAt(points.length - 1)} ${top + plotH} L${xAt(0)} ${top + plotH} Z" fill="var(--accent-soft)" stroke="none"/>`;
      svg += `<path d="${path}" fill="none" stroke="var(--accent)" stroke-width="2.5" stroke-linejoin="round"/>`;
    }
    points.forEach((p, i) => {
      const isLast = i === points.length - 1;
      svg += `<circle cx="${xAt(i)}" cy="${yAt(p.value)}" r="${isLast ? 5.5 : 4}" fill="var(--accent)" stroke="var(--surface)" stroke-width="2"/>`;
      svg += `<text x="${xAt(i)}" y="${yAt(p.value) - 10}" text-anchor="middle" font-size="11" font-weight="${isLast ? 700 : 500}" fill="var(--fg)">${formatNumber(p.value, 1)}</text>`;
      svg += `<text x="${xAt(i)}" y="${CHART_HEIGHT - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${p.label}</text>`;
    });
    el.innerHTML = svg + "</svg>";
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

  function renderOverview(model) {
    const { periods, spans, driven, spent, pending } = model;
    const efficiency = driven.km ? (driven.kwh / driven.km) * 100 : null;
    const costPerKm = driven.km ? driven.cost / driven.km : null;
    const fuelPerKm = (state.settings.fuelL100 / 100) * state.settings.fuelPrice;
    const saved = driven.km ? driven.km * fuelPerKm - driven.cost : null;

    $("k-eff").textContent = formatNumber(efficiency, 1);
    $("k-cpk").textContent = formatNumber(costPerKm == null ? null : costPerKm * 100, 1);
    $("k-month").textContent = formatNumber(spent.days ? (spent.total / spent.days) * 30 : null, 0);
    $("k-save").textContent = formatNumber(saved, 0);

    drawCostChart($("c-cost"), periods.map((p) => ({ label: formatDateShort(p.to), home: p.costHome, pub: p.costPublic })));
    const lastPeriod = periods[periods.length - 1];
    let splitText = lastPeriod
      ? "תקופה אחרונה (עד " + formatDate(lastPeriod.to) + "): " + describeSplit(lastPeriod.costHome, lastPeriod.costPublic) + "."
      : "";
    if (pending.count) {
      splitText += (splitText ? " " : "") + (pending.count === 1
        ? "מאז המדידה האחרונה נרשמה טעינה ציבורית אחת בסך " + formatNumber(pending.cost) + " ₪. היא תיכנס"
        : "מאז המדידה האחרונה נרשמו " + pending.count + " טעינות ציבוריות בסך " + formatNumber(pending.cost) + " ₪. הן ייכנסו") +
        " לתקופה הבאה כשתוזן מדידה.";
    }
    $("cost-split").textContent = splitText;

    drawEfficiencyChart($("c-eff"), spans.map((s) => ({ label: formatDateShort(s.to), value: s.efficiency })), efficiency || 0);
    const trend = $("trend");
    trend.textContent = "";
    trend.className = "";
    if (spans.length > 1) {
      const change = ((spans[spans.length - 1].efficiency - efficiency) / efficiency) * 100;
      trend.className = "pill " + (change <= 0 ? "good" : "warn");
      trend.textContent = "תקופה אחרונה: " + formatNumber(Math.abs(change), 0) + "% " + (change <= 0 ? "מתחת לממוצע" : "מעל הממוצע");
    }

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
        ["אותם ק״מ בבנזין", formatNumber(driven.km * fuelPerKm) + " ₪"],
        ["עלות לק״מ: חשמל מול בנזין", formatNumber(costPerKm * 100, 0) + " מול " + formatNumber(fuelPerKm * 100, 0) + " אג׳"],
        ["קצב נסיעה", formatNumber(kmPerDay * 30) + " ק״מ בחודש"],
        ["תחזית ק״מ לשנה", formatNumber(yearKm) + " ק״מ"],
        ["תחזית עלות חשמל לשנה", formatNumber(yearCost) + " ₪"],
        ["תחזית חיסכון לשנה", formatNumber(yearKm * fuelPerKm - yearKm * costPerKm) + " ₪"],
      ];
      facts.innerHTML = rows.map(([label, value]) => `<div><dt>${label}</dt><dd>${value}</dd></div>`).join("");
      $("facts-note").textContent = "ההשוואה מניחה רכב בנזין שצורך " + formatNumber(state.settings.fuelL100, 1) +
        " ליטר ל‑100 ק״מ ובנזין ב‑" + formatNumber(state.settings.fuelPrice, 2) +
        " ₪ לליטר. אפשר לשנות בהגדרות. הצריכה נמדדת במונה העמדה ולכן כוללת את הפסדי הטעינה.";
    } else {
      facts.innerHTML = '<p class="empty span-all">יופיע כאן אחרי שתי מדידות עם קילומטראז׳.</p>';
      $("facts-note").textContent = "";
    }

    // Efficiency by season, shown once at least two seasons have data.
    const seasons = { "קיץ": { kwh: 0, km: 0 }, "חורף": { kwh: 0, km: 0 }, "עונות מעבר": { kwh: 0, km: 0 } };
    for (const s of spans) {
      const month = new Date((toTime(s.from) + toTime(s.to)) / 2).getUTCMonth() + 1;
      const name = month >= 6 && month <= 9 ? "קיץ" : month === 12 || month <= 2 ? "חורף" : "עונות מעבר";
      seasons[name].kwh += s.kwh;
      seasons[name].km += s.km;
    }
    const withData = Object.keys(seasons).filter((name) => seasons[name].km > 0);
    $("season-panel").hidden = withData.length < 2;
    $("season").innerHTML = withData
      .map((name) => `<div><dt>${name}</dt><dd>${formatNumber((seasons[name].kwh / seasons[name].km) * 100, 1)} קוט״ש ל‑100 ק״מ</dd></div>`)
      .join("");
  }

  function renderHistory(model) {
    const { periods } = model;
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
        `<td>${formatNumber(priceForReading(state.readings, i), 4)}</td><td>${deleteButton("readings/" + r.date)}</td></tr>`);
      $("t-readings").innerHTML = "<tr><th>תאריך</th><th>מונה קוט״ש</th><th>קילומטראז׳</th><th>₪ לקוט״ש</th><th></th></tr>" + rows.reverse().join("");
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
        `<tr><td>${formatDate(p.date)}</td><td>${escapeHtml(p.provider) || "–"}</td><td>${formatNumber(p.kwh, 1)}</td><td>${formatNumber(p.cost, 2)}</td>` +
        `<td>${formatNumber(p.kwh ? p.cost / p.kwh : null, 2)}</td><td>${countedIn(p)}</td><td>${escapeHtml(p.note)}</td><td>${deleteButton("public/" + p.id)}</td></tr>`);
      $("t-public").innerHTML = "<tr><th>תאריך</th><th>חברה</th><th>קוט״ש</th><th>₪</th><th>₪ לקוט״ש</th><th>נספרת ב</th><th>הערה</th><th></th></tr>" + rows.join("");
    } else {
      $("t-public").innerHTML = '<tr><td class="empty">עוד לא נרשמו טעינות ציבוריות. מוסיפים בלשונית הזנה.</td></tr>';
    }
    $("export").hidden = !state.readings.length && !state.publicCharges.length;
  }

  /** Keeps form defaults in sync with the data without overwriting what the user is typing. */
  function renderFormDefaults() {
    const priceInput = $("r-price");
    if (!priceInput.dataset.touched && document.activeElement !== priceInput) priceInput.value = latestPrice();
    if (document.activeElement !== $("s-price")) $("s-price").value = state.settings.fuelPrice;
    if (document.activeElement !== $("s-l100")) $("s-l100").value = state.settings.fuelL100;
    // Suggest providers the user has typed before.
    const providers = [...new Set(state.publicCharges.map((p) => p.provider).filter(Boolean))].sort();
    $("providers").innerHTML = providers.map((name) => `<option value="${escapeHtml(name)}"></option>`).join("");
  }

  /** Redraws every screen from the current state. */
  function render() {
    const model = buildModel(state.readings, state.publicCharges);
    renderStatus();
    renderOverview(model);
    renderHistory(model);
    renderFormDefaults();
    renderReadingPreview();
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
      const kwhSinceAnchor = meter - anchor.meter +
        state.publicCharges.filter((p) => p.date > anchor.date && p.date <= date).reduce((sum, p) => sum + p.kwh, 0);
      text += " · " + formatNumber(km) + " ק״מ";
      if (km > 0) text += " · " + formatNumber((kwhSinceAnchor / km) * 100, 1) + " קוט״ש ל‑100 ק״מ";
      if (anchor !== previous) text += " (מחושב מאז " + formatDate(anchor.date) + ", המדידה האחרונה עם קילומטראז׳)";
    }
    el.textContent = text;
  }

  const NOT_CONNECTED = "צריך להתחבר ל‑OneDrive קודם.";

  async function onSubmitReading(event) {
    event.preventDefault();
    const date = $("r-date").value;
    const meter = inputNumber("r-meter");
    const odo = inputNumber("r-odo");
    const price = inputNumber("r-price");
    if (!state.connected) return showMessage("r-msg", NOT_CONNECTED);
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
        draft.readings.push({ date, meter, odo, price });
      });
      delete $("r-price").dataset.touched;
      showMessage("r-msg", "המדידה נשמרה." + (odo == null ? " בלי קילומטראז׳ הצריכה תחושב יחד עם התקופה הבאה." : ""), true);
      $("r-meter").value = "";
      $("r-odo").value = "";
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
    const provider = $("p-provider").value.trim();
    const note = $("p-note").value.trim();
    if (!state.connected) return showMessage("p-msg", NOT_CONNECTED);
    if (!date || !(kwh > 0) || cost == null || cost < 0) return showMessage("p-msg", "צריך תאריך, קוט״ש גדול מאפס וסכום.");
    try {
      // Always appended as a new entry: it never replaces a reading or another charge.
      await mutate((draft) => { draft.publicCharges.push({ id: newChargeId(), date, kwh, cost, provider, note }); });
      showMessage("p-msg", "הטעינה נשמרה.", true);
      for (const id of ["p-kwh", "p-cost", "p-provider", "p-note"]) $(id).value = "";
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
      await mutate((draft) => { draft.settings = { fuelPrice, fuelL100 }; });
      showMessage("s-msg", "נשמר.", true);
    } catch (err) {
      showMessage("s-msg", describeSaveError(err));
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
      await mutate((draft) => { draft.readings = []; draft.publicCharges = []; draft.settings = { ...DEFAULT_SETTINGS }; });
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

  const TABS = ["overview", "add", "history", "settings"];

  function showTab(tab) {
    document.querySelectorAll(".tabs button").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.tab === tab)));
    for (const name of TABS) $("v-" + name).hidden = name !== tab;
    try { localStorage.setItem(STORE.tab, tab); } catch (e) { /* storage unavailable */ }
    if (tab === "overview") render(); // charts measure their container, so draw once it is visible
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
  ];
  const TYPE_READING = "מדידה";
  const TYPE_PUBLIC = "טעינה ציבורית";

  /** Lower-cases and strips everything except Hebrew letters, Latin letters and digits. */
  const simplify = (text) => String(text ?? "").toLowerCase().replace(/[^a-z0-9א-ת]/g, "");

  /** Builds the CSV text for the given data (UTF-8 with BOM so Excel shows Hebrew correctly). */
  function toCsv(data) {
    const rows = [CSV_COLUMNS.map((c) => c.header)];
    for (const r of data.readings) rows.push([TYPE_READING, r.date, r.meter, r.odo ?? "", r.price ?? "", "", "", "", "", ""]);
    for (const p of data.publicCharges) rows.push([TYPE_PUBLIC, p.date, "", "", "", p.kwh, p.cost, p.provider || "", p.note || "", p.id]);
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
      if (json.settings && (json.settings.fuelPrice > 0 || json.settings.fuelL100 > 0)) result.settings = json.settings;
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
    for (const item of incoming.readings) {
      const existing = readings.get(item.date);
      if (!existing) counts.readings.added += 1;
      else if (same(existing, item, ["meter", "odo", "price"])) { counts.readings.unchanged += 1; continue; }
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
      for (const key of ["fuelPrice", "fuelL100"]) {
        const value = Number(incoming.settings[key]);
        if (value > 0 && value !== settings[key]) { settings[key] = value; counts.settingsChanged = true; }
      }
    }

    const data = { readings: [...readings.values()].sort(byDate), publicCharges: charges.sort(byDate), settings };

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
      readings: [{ date: "2026-10-01", meter: 4350.2, odo: 56200, price: 0.5451 }],
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
    if (plan.counts.settingsChanged) lines.push("הגדרות רכב הבנזין להשוואה יעודכנו לפי הקובץ.");
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
      setImportMessage("הייבוא הושלם: " + changed + " רשומות נוספו או עודכנו ונשמרו ב‑OneDrive.", "ok");
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

  function signOut() {
    storage.remove(STORE.token);
    storage.remove(STORE.cache);
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

    for (const id of ["r-date", "r-meter", "r-odo", "r-price"]) $(id).addEventListener("input", renderReadingPreview);
    $("r-price").addEventListener("input", () => { $("r-price").dataset.touched = "1"; });
    $("f-reading").addEventListener("submit", onSubmitReading);
    $("f-public").addEventListener("submit", onSubmitPublicCharge);
    $("f-fuel").addEventListener("submit", onSubmitFuelSettings);
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
    bindEvents();
    $("r-date").value = $("p-date").value = todayIso();
    $("app-version").textContent = APP_VERSION;
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
  const api = { parseDate, parseNumber, parseCsv, parseImportText, planMerge, normalizeData, serializeData, buildModel, toCsv };
  if (typeof window !== "undefined") window.ZeekrTracker = api;
  if (typeof document !== "undefined") start();
  else if (typeof globalThis !== "undefined") globalThis.ZeekrTracker = api;
})();
