/**
 * Calculation tests, run against both copies of the app: the trial copy
 * (beta/app.js) and the main one (app.js). Every test runs once per copy.
 *
 * Every expected number below was worked out by hand (or with a separate
 * formula, noted next to it), not copied from the app's own output. All data
 * is made up: the repository is public, so no real readings or loan figures
 * belong here.
 *
 * Run: node --test tests/unit.test.js
 */
"use strict";
const { describe, test } = require("node:test");
const assert = require("node:assert/strict");
const { loadApp } = require("./load-app");

/** Values created inside the app's sandbox have their own prototypes; compare them as plain JSON. */
const plain = (value) => JSON.parse(JSON.stringify(value));
/** Equal to the agora (two decimals). */
const near = (actual, expected, message) => assert.ok(Math.abs(actual - expected) < 0.01, `${message || ""} expected ${expected}, got ${actual}`);

const reading = (date, meter, odo, price = 0.5, soc = null) => ({ date, meter, odo, price, soc });
const charge = (id, date, kwh, cost, provider = "", note = "") => ({ id, date, kwh, cost, provider, note });

for (const copy of ["beta/app.js", "app.js"]) {
  describe(copy, () => {
    const Z = loadApp(copy);

    /* ---------- Reading what people type or paste ---------- */

    test("dates: ISO and day-first forms, impossible dates rejected", () => {
      assert.equal(Z.parseDate("2026-3-5"), "2026-03-05");
      assert.equal(Z.parseDate("1/2/26"), "2026-02-01");
      assert.equal(Z.parseDate("15.10.2024"), "2024-10-15");
      assert.equal(Z.parseDate("31/2/2026"), null);
      assert.equal(Z.parseDate("hello"), null);
    });

    test("numbers: shekel sign, thousands separators and decimal comma", () => {
      assert.equal(Z.parseNumber("1,234.5"), 1234.5);
      assert.equal(Z.parseNumber("12,500"), 12500);
      assert.equal(Z.parseNumber("1,5"), 1.5);
      assert.equal(Z.parseNumber("₪ 12"), 12);
      assert.equal(Z.parseNumber(""), null);
      assert.ok(Number.isNaN(Z.parseNumber("abc")));
    });

    test("a reading needs a date and a meter; battery percent stays within 0-100", () => {
      assert.ok(Z.normalizeReading({ date: "2026-01-01", meter: "1000" }).value);
      assert.ok(Z.normalizeReading({ date: "", meter: 1000 }).error);
      assert.ok(Z.normalizeReading({ date: "2026-01-01", meter: -1 }).error);
      assert.ok(Z.normalizeReading({ date: "2026-01-01", meter: 1, soc: "120%" }).error);
      assert.equal(Z.normalizeReading({ date: "2026-01-01", meter: 1, soc: "80%" }).value.soc, 80);
    });

    test("CSV: quoted fields, doubled quotes and a semicolon delimiter", () => {
      const rows = Z.parseCsv('a;b;c\n"x;1";"say ""hi""";3\n');
      assert.deepEqual(plain(rows), [["a", "b", "c"], ["x;1", 'say "hi"', "3"]]);
    });

    test("import: Hebrew headers, typed rows, and a bad row reported by its line", () => {
      const csv = "סוג;תאריך;מונה קוט״ש;קילומטראז׳;מחיר לקוט״ש;קוט״ש;עלות\n" +
        "מדידה;1/1/2026;1000;10000;0.5;;\n" +
        "מדידה;31/2/2026;1100;11000;0.5;;\n" +
        "טעינה ציבורית;10/1/2026;;;;20;45\n";
      const result = Z.parseImportText(csv);
      assert.equal(result.readings.length, 1);
      assert.equal(result.publicCharges.length, 1);
      assert.equal(result.rejected.length, 1);
      assert.equal(result.rejected[0].where, "שורה 3");
    });

    test("export and import give back the same entries", () => {
      const data = {
        readings: [reading("2026-01-01", 1000, 10000, 0.5, 60), reading("2026-02-01", 1160, null, 0.55)],
        publicCharges: [charge("p1", "2026-01-10", 20, 45.5, "פז", 'note, with "quotes"')],
      };
      const back = Z.parseImportText(Z.toCsv(data));
      assert.equal(back.rejected.length, 0);
      assert.deepEqual(plain(back.readings), plain(data.readings));
      assert.deepEqual(plain(back.publicCharges), plain(data.publicCharges));
    });

    /* ---------- Merging an import into the data ---------- */

    test("merge: nothing is lost, same date updates, a repeated charge is recognised", () => {
      const current = Z.normalizeData({
        readings: [reading("2026-01-01", 1000, 10000, 0.5, 60)],
        public: [charge("a", "2026-01-10", 10, 20, "פז")],
      });
      const incoming = {
        readings: [reading("2026-01-01", 1000, 10000, 0.5, null), reading("2026-02-01", 1100, 10600)],
        publicCharges: [charge("", "2026-01-10", 10, 20, "פז", "updated note"), charge("", "2026-01-20", 5, 9)],
        settings: null,
      };
      const { data, counts, warnings } = Z.planMerge(current, incoming);
      assert.deepEqual(plain(counts.readings), { added: 1, updated: 0, unchanged: 1 });
      assert.deepEqual(plain(counts.publicCharges), { added: 1, updated: 1, unchanged: 0 });
      assert.equal(data.readings[0].soc, 60, "an empty battery cell keeps the stored value");
      assert.equal(data.publicCharges.find((p) => p.id === "a").note, "updated note");
      assert.equal(data.publicCharges.length, 2);
      assert.deepEqual(plain(data.loan), plain(current.loan), "a file without loan data keeps the loan inputs");
      assert.equal(warnings.length, 0);
    });

    test("merge: a meter going down is flagged", () => {
      const current = Z.normalizeData({ readings: [reading("2026-01-01", 1000, 10000)] });
      const { warnings } = Z.planMerge(current, { readings: [reading("2026-02-01", 900, 11000)], publicCharges: [], settings: null });
      assert.equal(warnings.length, 1);
    });

    test("the data file keeps the loan bills when saved and read again", () => {
      const loan = { buyDate: "2025-01-01", carPrice: 100000, expenses: [{ date: "2025-06-01", amount: 800, note: "service" }] };
      const again = Z.normalizeData(Z.serializeData(Z.normalizeData({ loan })));
      assert.deepEqual(plain(again.loan.expenses), [{ date: "2025-06-01", amount: 800, note: "service" }]);
    });

    /* ---------- Home charging: consumption and cost ---------- */

    test("one period: 160 kWh over 1,000 km at 0.5 ILS is 16 kWh/100 km and 80 ILS", () => {
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000), reading("2026-01-31", 1160, 11000)], [], 66);
      const p = m.periods[0];
      assert.equal(p.days, 30);
      near(p.kwhHome, 160);
      near(p.costHome, 80);
      near(p.efficiency, 16);
      near(m.spent.total, 80);
    });

    test("a public charge inside the period adds its energy and its cost", () => {
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000), reading("2026-01-31", 1160, 11000)],
        [charge("p", "2026-01-15", 20, 50)], 66);
      near(m.periods[0].efficiency, 18);   // (160 + 20) / 1000 km
      near(m.spent.public, 50);
      near(m.spent.total, 130);
    });

    test("a charge on a reading's date belongs to the period that reading closes", () => {
      const readings = [reading("2026-01-01", 1000, 10000), reading("2026-01-31", 1160, 11000)];
      const onClose = Z.buildModel(readings, [charge("p", "2026-01-31", 20, 50)], 66);
      near(onClose.periods[0].costPublic, 50);
      const onOpen = Z.buildModel(readings, [charge("p", "2026-01-01", 20, 50)], 66);
      near(onOpen.periods[0].costPublic, 0, "the opening day belongs to the period before");
    });

    test("a public charge after the last reading waits for the next one", () => {
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000), reading("2026-01-31", 1160, 11000)],
        [charge("p", "2026-02-05", 20, 50)], 66);
      assert.equal(m.pending.count, 1);
      near(m.spent.total, 80);
    });

    test("battery percent corrects the energy used, never the cost", () => {
      // 50% -> 70% of a 60 kWh battery: 12 kWh went into storage, so 148 kWh were used.
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000, 0.5, 50), reading("2026-01-31", 1160, 11000, 0.5, 70)], [], 60);
      near(m.spans[0].kwhUsed, 148);
      near(m.periods[0].efficiency, 14.8);
      near(m.spent.total, 80);
    });

    test("a reading without odometer joins its period to the next one", () => {
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000), reading("2026-01-16", 1080, null), reading("2026-01-31", 1160, 11000)], [], 66);
      assert.equal(m.spans.length, 1);
      assert.equal(m.periods[0].efficiency, null);
      near(m.periods[1].efficiency, 16);
      assert.equal(m.periods[1].mergedCount, 2);
      assert.equal(m.periods[1].kmMerged, 1000);
    });

    test("a reading without a price takes the previous price", () => {
      const m = Z.buildModel([reading("2026-01-01", 1000, 10000, 0.5), reading("2026-01-31", 1100, 10500, null)], [], 66);
      near(m.periods[0].costHome, 50);
    });

    test("a period is split between months by its days", () => {
      // 16 Jan to 15 Feb is 30 days: 16 in January, 14 in February, 10 kWh and 100 km a day.
      const readings = [reading("2026-01-16", 0, 0, 1), reading("2026-02-15", 300, 3000, 1)];
      const months = Z.buildMonths(readings, [], Z.buildModel(readings, [], 66));
      assert.deepEqual(plain(months.map((m) => m.key)), ["2026-01", "2026-02"]);
      near(months[0].kwhHome, 160);
      near(months[0].km, 1600);
      near(months[0].efficiency, 10);
      near(months[1].kwhHome, 140);
      assert.equal(months[0].partial, true);
      const years = Z.buildYears(months);
      assert.equal(years.length, 1);
      near(years[0].kwhHome, 300);
    });

    /* ---------- Loan arithmetic ---------- */

    test("Spitzer loan: 100,000 at 6% over 12 months", () => {
      near(Z.annuityPayment(100000, 6, 12), 8606.64);               // standard annuity formula
      near(Z.annuityBalance(100000, 6, 12, 6), 50748.08);
      assert.equal(Z.annuityBalance(100000, 6, 12, 12), 0);
      near(Z.annuityPayment(12000, 0, 12), 1000);
    });

    test("grace loan: a fixed payment leaves the rest owed at the end", () => {
      near(Z.scheduleBalance({ balance: 100000, ratePct: 6, months: 12, payment: 500 }, 12), 100000);   // interest only
      near(Z.scheduleBalance({ balance: 100000, ratePct: 6, months: 12, payment: 1000 }, 12), 93832.22);
      near(Z.scheduleBalance({ balance: 100000, ratePct: 6, months: 12, payment: 1000 }, 20), 93832.22, "no change after its end");
    });

    test("calendar helpers", () => {
      assert.equal(Z.addMonths("2024-01-31", 1), "2024-02-29");
      assert.equal(Z.addMonths("2024-10-07", 36), "2027-10-07");
      assert.equal(Z.monthsBetween("2024-01-15", "2024-03-14"), 1);
      assert.equal(Z.monthsBetween("2024-01-15", "2024-03-15"), 2);
    });

    test("loan inputs are cleaned: old litres per 100 km, bad schedules, order by date", () => {
      const loan = Z.normalizeLoan({
        altL100: 5,
        schedules: [
          { from: "2026-01-01", balance: 50000, ratePct: 5, months: 24 },
          { from: "2025-01-01", balance: 60000, ratePct: 5, months: 24 },
          { from: "2025-06-01", balance: 60000, ratePct: 80, months: 24 },
        ],
      });
      assert.equal(loan.altKmL, 20);
      assert.deepEqual(plain(loan.schedules.map((s) => s.from)), ["2025-01-01", "2026-01-01"]);
      assert.equal(loan.carIns, undefined, "insurance is no longer part of the inputs");
    });

    test("the comparison lists what it cannot run without", () => {
      const model = Z.buildLoanModel(Z.normalizeLoan({}), [], "2026-01-01");
      assert.ok(model.missing.includes("תאריך הקנייה"));
      assert.ok(model.missing.includes("לוח סילוקין אחד לפחות"));
    });

    /** A small made-up purchase: everything is zero unless a test sets it. */
    function loanOf(extra) {
      return Z.normalizeLoan({
        buyDate: "2025-01-01", carPrice: 100000, altPrice: 50000, standAmount: 0, standFreeMonths: null,
        horizonMonths: 12, schedules: [{ from: "2025-01-01", balance: 60000, ratePct: 0, months: 12 }],
        ...extra,
      });
    }

    test("loan model: depreciation and lost fund return, worked out by hand", () => {
      // Zeekr: 40,000 cash on day 0 plus 12 payments of 5,000, each grown at 1% a month to month 12:
      //   40,000 x 1.01^12 + 5,000 x (1.01^12 - 1) / 0.01 = 108,485.52, minus the car worth 90,000.
      // Alternative: 50,000 x 1.01^12 = 56,341.25, minus the car worth 45,000.
      const model = Z.buildLoanModel(loanOf({ carDepPct: 10, altDepPct: 10, returnPct: 12 }), [], "2026-01-01");
      const end = model.end;
      assert.equal(end.i, 12);
      near(end.costZ, 18485.52);
      near(end.costA, 11341.25);
      near(end.advantage, -7144.26);
      const sum = (parts) => Object.values(parts).reduce((a, b) => a + b, 0);
      near(sum(end.partsZ), end.costZ, "the Zeekr's parts add up to its total");
      near(sum(end.partsA), end.costA, "the alternative's parts add up to its total");
    });

    test("loan model: petrol for the same km at the price in Settings", () => {
      // 1,000 km a month at 20 km per litre is 50 litres; at 7.3 ILS that is 365 a month.
      const model = Z.buildLoanModel(loanOf({ kmMonth: 1000, altKmL: 20 }), [], "2026-01-01");
      near(model.end.partsA.fuel, 365 * 12);
    });

    test("loan model: a grace loan still open at its end goes on paying interest only", () => {
      const loan = loanOf({ horizonMonths: 18, schedules: [{ from: "2025-01-01", balance: 100000, ratePct: 6, months: 12, payment: 500 }] });
      const model = Z.buildLoanModel(loan, [], "2026-01-01");
      near(model.rows[15].payment, 500);
      near(model.rows[15].balance, 100000);
      near(model.end.partsZ.interest, 18 * 500);
    });

    test("loan model: the standing loan stops being owed when it rolls into the next car", () => {
      const loan = loanOf({ standAmount: 20000, standDropPct: 10, standFreeMonths: 6 });
      const rows = Z.buildLoanModel(loan, [], "2026-01-01").rows;
      near(rows[5].standOwed, 19140.99);   // 20,000 x 0.9^(5/12)
      assert.equal(rows[6].standOwed, 0);
      near(rows[6].advantage - rows[5].advantage, 19140.99, "the Zeekr gains exactly what is no longer owed");
    });

    test("best month: the highest advantage per month of ownership", () => {
      // The alternative pays a 1,200 yearly fee (100 a month), so the advantage grows 100 a month;
      // the 20,000 standing loan is released at month 24. Advantage per month: 100 before, 933 at 24, falling after.
      const loan = loanOf({ altFee: 1200, standAmount: 20000, standFreeMonths: 24 });
      assert.equal(Z.bestSaleMonth(loan, [], "2026-01-01"), 24);
    });
  });
}
