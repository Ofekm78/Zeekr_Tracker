# ZEEKR TRACKER

A personal web app that tracks the electricity cost of a Zeekr X: home charger
readings, public charging, consumption, savings against petrol, and whether
financing the car was worth it. Hebrew, right to left, no build step and no
third-party libraries. Published by GitHub Pages from this repository.

- Main page: `https://ofekm78.github.io/Zeekr_Tracker/`
- Trial page: `https://ofekm78.github.io/Zeekr_Tracker/beta/`

All personal data lives in the user's own OneDrive (app folder), never here.
The repository is public: no data files, readings or loan figures are committed.

## What is where

| Path | What it is |
|---|---|
| `index.html`, `styles.css`, `app.js` | The main page: markup, look, and all logic (sections listed at the top of `app.js`). Works on `zeekr-data.json` in OneDrive. |
| `sw.js` | Keeps the page's own files on the phone so the installed app opens instantly. Never caches data. |
| `manifest.webmanifest`, `icons/` | Installing to the home screen: name, icons, opening screen, charging-company symbols. |
| `beta/` | The trial copy, same files plus its own name and cache. Works on a separate `zeekr-test.json`, so nothing there touches real data. Every change is made here first. |
| `data/fuel-prices.json` | Official petrol prices by effective date, shared by both copies. |
| `scripts/update_fuel_prices.py` | Reads the Ministry of Energy announcements and adds new prices to that file. |
| `.github/workflows/fuel-prices.yml` | Runs the script every morning; a red run means this month's price was not found. |
| `tests/` | Automatic tests: calculations on both copies, and a browser run through every tab. See `tests/README.md`. |
| `.github/workflows/tests.yml` | Runs the tests on every pull request and on `main`. |
| `logo/` | Master copies of the logo and symbol (SVG and PNG). Not loaded by the site. |

## How a change goes out

1. Change `beta/` and bump `APP_VERSION` in `beta/app.js` together with the
   `?v=` keys in `beta/index.html`.
2. Open a pull request; the tests must be green before it is merged.
3. After Ofek approves the trial version, copy `app.js`, `styles.css` and
   `index.html` to the root (keeping the main name in `index.html`).

Sign-in uses a Microsoft Entra app registration for personal accounts with the
`Files.ReadWrite.AppFolder` scope. Both page addresses (root and `/beta/`) must
stay registered as redirect URIs.
