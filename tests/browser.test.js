/**
 * Browser test for both copies of the app (beta/ and the main page): opens
 * each in a real Chromium with made-up data, goes through every tab and the
 * loan views, and fails on any script error or on a screen that did not fill in.
 *
 * Nothing leaves the machine: the page is served from this folder, every
 * request to another address (OneDrive, sign-in, fonts, weather) is blocked,
 * and the data is put straight into the page's offline cache.
 *
 * Needs Playwright with Chromium. Run: node --test tests/browser.test.js
 */
"use strict";
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const http = require("http");
const path = require("path");
const { chromium } = require("playwright");

const ROOT = path.join(__dirname, "..");
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".webmanifest": "application/manifest+json" };

/** Six months of made-up readings: 1,000 km and 160 kWh a month at 0.5 ILS, so 16.0 kWh/100 km and 8.0 agorot per km. */
function sampleData() {
  const readings = [];
  for (let i = 0; i <= 6; i++) {
    const date = new Date(Date.UTC(2026, 0 + i, 1)).toISOString().slice(0, 10);
    readings.push({ date, meter: 1000 + 160 * i, odo: 20000 + 1000 * i, price: 0.5, soc: null });
  }
  return {
    readings,
    public: [],
    settings: { fuelPrice: 7, fuelL100: 7, batteryKwh: 66 },
    loan: {
      buyDate: "2025-01-01", horizonMonths: 36, carPrice: 150000, carDepPct: 15, carMaint: 1000, carFee: 2000, carUpfront: 2000,
      altPrice: 60000, altKmL: 18, altDepPct: 10, altMaint: 2500, altFee: 1500, altUpfront: 1000,
      returnPct: 6, standAmount: 30000, standDropPct: 10, standFreeMonths: 42,
      schedules: [{ from: "2025-01-01", balance: 100000, ratePct: 5, months: 36, payment: 500 }],
      expenses: [{ date: "2025-07-01", amount: 900, note: "service" }],
    },
  };
}

/**
 * The three copies: where each is served, and the offline-cache key it reads first.
 * `current` marks the copies on the newest code (main gets it when beta is promoted);
 * `share` is the copy for other owners, without the loan tab.
 */
const COPIES = [
  { name: "trial copy", url: "/beta/", dir: "beta", cacheKey: "zk-data-test", trial: true, current: true },
  { name: "main copy", url: "/", dir: ".", cacheKey: "zk-data", trial: false },
  { name: "shared copy", url: "/share/", dir: "share", cacheKey: "zk-data-share", trial: false, current: true,
    share: true },
];

let server, browser, base;

test.before(async () => {
  server = http.createServer((req, res) => {
    let file = path.join(ROOT, decodeURIComponent(req.url.split("?")[0]));
    if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
    if (file.endsWith(path.sep)) file += "index.html";
    fs.readFile(file, (err, body) => {
      if (err) { res.writeHead(404); return res.end(); }
      res.writeHead(200, { "content-type": TYPES[path.extname(file)] || "application/octet-stream" });
      res.end(body);
    });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  base = `http://127.0.0.1:${server.address().port}`;
  browser = await chromium.launch();
});

test.after(async () => {
  if (browser) await browser.close();
  if (server) server.close();
});

/** Opens the trial page with the sample data; returns the page and the list of script errors. */
async function openApp(copy, viewport) {
  const context = await browser.newContext({ viewport, serviceWorkers: "block" });
  await context.addInitScript(([key, data]) => {
    if (sessionStorage.getItem("seeded")) return;
    localStorage.setItem(key, JSON.stringify(data));
    sessionStorage.setItem("seeded", "1");
  }, [copy.cacheKey, sampleData()]);
  await context.route((url) => !url.href.startsWith(base), (route) => route.abort());
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("console", (m) => { if (m.type() === "error" && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
  await page.goto(base + copy.url);
  await page.waitForSelector("#k-eff");
  return { page, errors };
}

/** Goes through the loan tab: every view fills the tiles, the chart and the sale table. */
async function checkLoanViews(page) {
  await page.click('.tabs button[data-tab="loan"]');
  assert.equal(await page.isHidden("#loan-missing"), true, "no missing inputs");
  for (const view of ["best", "24", "36", "84"]) {
    await page.selectOption("#lv-years", view);
    assert.match(await page.textContent("#tile-z"), /\d.*₪/, `Zeekr tile filled for view ${view}`);
    assert.match(await page.textContent("#tile-a"), /\d.*₪/, `alternative tile filled for view ${view}`);
  }
  assert.match(await page.textContent("#lv-best"), /^הכי משתלם, \d+ חודשים$/);
  await page.click("#lv-today");
  assert.ok(await page.locator("#c-loan svg").count(), "loan chart drawn");
  assert.ok((await page.locator("#sale-table tr").count()) > 2, "sale table filled");
}

for (const copy of COPIES) {
  test(`${copy.name}: version: header, Settings and the ?v= cache keys all match APP_VERSION`, async () => {
    const appJs = fs.readFileSync(path.join(ROOT, copy.dir, "app.js"), "utf8");
    const html = fs.readFileSync(path.join(ROOT, copy.dir, "index.html"), "utf8");
    const version = /const APP_VERSION = "([\d.]+)"/.exec(appJs)[1];
    assert.match(html, new RegExp(`app\\.js\\?v=${version.replace(".", "\\.")}"`), "index.html loads app.js with the current version");
    assert.match(html, new RegExp(`styles\\.css\\?v=${version.replace(".", "\\.")}"`), "index.html loads styles.css with the current version");
    const { page, errors } = await openApp(copy, { width: 400, height: 900 });
    assert.equal(await page.textContent("#head-version"), "v" + version);
    assert.equal(await page.textContent("#app-version"), version);
    assert.equal(await page.title(), copy.trial ? "ZEEKR TRACKER TEST" : "ZEEKR TRACKER", "each copy keeps its own name");
    if (copy.trial) assert.equal(await page.textContent("#head-name"), "ZEEKR TRACKER TEST", "trial header says TEST");
    assert.deepEqual(errors, []);
    await page.context().close();
  });

  test(`${copy.name}: every tab opens and fills in, on a phone and on a wide screen`, async () => {
    for (const viewport of [{ width: 400, height: 900 }, { width: 1280, height: 900 }]) {
      const { page, errors } = await openApp(copy, viewport);

      await page.click('.tabs button[data-tab="overview"]');
      assert.equal(await page.textContent("#k-eff"), "16.0", "average consumption");
      assert.equal(await page.textContent("#k-cpk"), "8.0", "cost per km in agorot");
      if (copy.current) {
        // 6,000 km from 1 Jan to 1 Jul 2026, 181 days: 6000 / 181 * 30 = 994.48, shown as 994.
        assert.equal(await page.textContent("#k-kpm"), "994", "average km per 30 days");
      }
      assert.ok(await page.locator("#c-cost svg").count(), "monthly cost chart drawn");

      assert.equal(await page.isVisible('.tabs button[data-tab="loan"]'), !copy.share,
        "loan tab hidden only on the shared copy");
      if (!copy.share) await checkLoanViews(page);

      await page.click('.tabs button[data-tab="add"]');
      await page.fill("#r-meter", "2200");
      await page.fill("#r-odo", "27000");
      assert.match(await page.textContent("#r-preview"), /ק״מ/, "live summary under the reading form");

      await page.click('.tabs button[data-tab="history"]');
      assert.equal(await page.locator("#t-readings tr").count(), 8, "header and seven readings");

      await page.click('.tabs button[data-tab="settings"]');
      assert.equal(await page.isVisible("#beta-panel"), copy.trial, "trial tools only on the trial copy");
      assert.equal(await page.isVisible("#beta-banner"), copy.trial, "trial notice only on the trial copy");
      if (copy.current) {
        assert.equal(await page.inputValue("#s-city"), "beit-yehoshua", "default weather city");
        assert.ok((await page.locator("#s-city option").count()) > 20, "city list filled");
      }
      assert.equal(await page.isVisible("#other-version"), !copy.share, "no link to other copies on the shared copy");

      assert.deepEqual(errors, [], `no script errors at ${viewport.width}px`);
      await page.context().close();
    }
  });

  if (!copy.share) test(`${copy.name}: loan base data is locked until Edit, and Cancel locks it again`, async () => {
    const { page, errors } = await openApp(copy, { width: 400, height: 900 });
    await page.click('.tabs button[data-tab="loan"]');
    assert.equal(await page.isDisabled("#l-car"), true);
    await page.click("#l-edit");
    assert.equal(await page.isDisabled("#l-car"), false);
    await page.fill("#l-car", "1");
    await page.click("#l-cancel");
    assert.equal(await page.isDisabled("#l-car"), true);
    assert.equal(await page.inputValue("#l-car"), "150000", "cancel restores the saved value");
    assert.deepEqual(errors, []);
    await page.context().close();
  });
}
