# Automatic tests

Two layers, both run by GitHub on every pull request (`.github/workflows/tests.yml`).

| File | What it checks | Time |
|---|---|---|
| `unit.test.js` | The calculations, run on every copy (`beta/app.js`, `app.js` and `share/app.js`): dates and numbers typed or imported, CSV import and merge, consumption and cost per period and per month, and the loan comparison (schedules, refinance, standing loan, depreciation, lost return, best sale month). | about 1 second |
| `browser.test.js` | Opens every copy (`beta/`, the main page and `share/`) in Chromium with sample data, goes through every tab and the loan views (not on the shared copy, which has no loan tab) on a phone and a wide screen, and fails on any script error. Also checks that the version in the header matches the `?v=` cache keys. | about 1.5 minutes on GitHub (Chromium is installed each run) |

All data in the tests is made up. The repository is public, so real readings or loan figures never go here.
Every expected number was worked out by hand and is explained next to the test.

Run locally (Node 22):

    node --test tests/unit.test.js
    node --test tests/browser.test.js   # needs Playwright with Chromium

When a feature is added or a calculation changes, add or update a test in the same change.
