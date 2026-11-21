#!/usr/bin/env python3
"""Keeps data/fuel-prices.json up to date with Israel's regulated petrol price.

The Ministry of Energy publishes each price change as a PDF on gov.il. This
script looks for the announcement of the current month (and of mid-month
updates), reads the maximum price of 95-octane petrol at self-service
including VAT, and records it with its effective date.

Run by .github/workflows/fuel-prices.yml every day. It can also be run by hand:
    python3 scripts/update_fuel_prices.py                 # look for announcements
    python3 scripts/update_fuel_prices.py 2026-11-01 7.35 # add a price manually

Exit codes: 0 = nothing new or file updated, 1 = no price found for this
month by day MISSING_AFTER_DAY (so the workflow fails and GitHub sends an
email), 2 = bad arguments.
"""
import datetime as dt
import json
import pathlib
import re
import subprocess
import sys
import tempfile
import urllib.request

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA_FILE = ROOT / "data" / "fuel-prices.json"
BASE = "https://www.gov.il/BlobFolder/news/"
MONTHS = ["january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"]
MIN_PRICE, MAX_PRICE = 4.0, 15.0
MISSING_AFTER_DAY = 5
USER_AGENT = "Mozilla/5.0 (fuel-price updater; github actions)"


def candidate_urls(year, month):
    """The address patterns the ministry has used for its announcements."""
    name, mm = MONTHS[month - 1], f"{month:02d}"
    folder = f"fuel-{name}-{year}"
    urls = [
        f"{BASE}{folder}/he/fuel-{name}-{year}.pdf",
        f"{BASE}{folder}/he/fuel_{name}{year}.pdf",
        f"{BASE}{folder}/he/fuel-{mm}-{year}.pdf",
        f"{BASE}{folder}/he/fuel-{month}-{year}.pdf",
        f"{BASE}fuel_{name[:3]}_{year}/he/{month}_{year}.pdf",
    ]
    # Mid-month updates, e.g. news/fuel-october-5-2026/he/fuel-5-october-2026.pdf
    for day in range(2, 29):
        urls.append(f"{BASE}fuel-{name}-{day}-{year}/he/fuel-{day}-{name}-{year}.pdf")
    return urls


def download(url):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read()
            return body if body.startswith(b"%PDF") else None
    except Exception:
        return None


def pdf_text(data):
    with tempfile.NamedTemporaryFile(suffix=".pdf") as f:
        f.write(data)
        f.flush()
        return subprocess.run(["pdftotext", "-layout", f.name, "-"], capture_output=True,
                              text=True, check=False).stdout


DATE = r"(\d{1,2})\.(\d{1,2})\.(20\d\d)"
PRICE = r"(\d{1,2}\.\d{2})"


def parse_announcement(text):
    """Returns (effective ISO date, price) or None.

    Extracted Hebrew may come out in logical or in visual (reversed) order, so
    both orders of the key phrase "לא יעלה על" (shall not exceed) are tried.
    The 95-octane sentence is the first one with that phrase.
    """
    flat = re.sub(r"\s+", " ", text)
    price = None
    for pattern in (r"לא יעלה על\s*" + PRICE, PRICE + r"\s*על יעלה אל"):
        match = re.search(pattern, flat)
        if match:
            price = float(match.group(1))
            break
    if price is None or not (MIN_PRICE <= price <= MAX_PRICE):
        return None
    # Effective date: the first date in the text; "4.10.2026-5.10.2026" means the 5th.
    match = re.search(DATE + r"(?:\s*-\s*" + DATE + r")?", flat)
    if not match:
        return None
    day, month, year = (match.group(4), match.group(5), match.group(6)) if match.group(4) else match.groups()[:3]
    try:
        date = dt.date(int(year), int(month), int(day)).isoformat()
    except ValueError:
        return None
    return date, price


def load():
    if DATA_FILE.exists():
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    return {"description": "", "prices": []}


def save(data):
    data["description"] = ("Maximum price of 95-octane petrol at self-service, incl. VAT, ILS per liter, "
                           "by effective date. Source: Israel Ministry of Energy announcements on gov.il.")
    data["updated"] = dt.date.today().isoformat()
    data["prices"].sort(key=lambda p: p["from"])
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def add(data, date, price, source):
    """Adds or corrects the price for a date. Returns True when something changed."""
    for entry in data["prices"]:
        if entry["from"] == date:
            if entry["price"] == price:
                return False
            entry.update(price=price, source=source)
            return True
    data["prices"].append({"from": date, "price": price, "source": source})
    return True


def main(argv):
    data = load()
    if len(argv) == 3:
        try:
            date = dt.date.fromisoformat(argv[1]).isoformat()
            price = float(argv[2])
        except ValueError:
            print("usage: update_fuel_prices.py [YYYY-MM-DD PRICE]")
            return 2
        if not MIN_PRICE <= price <= MAX_PRICE:
            print(f"price {price} is outside {MIN_PRICE}-{MAX_PRICE}")
            return 2
        if add(data, date, price, "manual"):
            save(data)
        print(f"{date}: {price}")
        return 0

    today = dt.date.today()
    found_this_month = any(p["from"][:7] == today.isoformat()[:7] for p in data["prices"])
    changed = False
    # Look at this month and, early in the month, at next month's announcement too.
    months = [(today.year, today.month)]
    if today.day >= 25:
        months.append((today.year + (today.month == 12), today.month % 12 + 1))
    for year, month in months:
        for url in candidate_urls(year, month):
            body = download(url)
            if not body:
                continue
            result = parse_announcement(pdf_text(body))
            if not result:
                text = re.sub(r"\s+", " ", pdf_text(body))
                print(f"could not read a price from {url}; text starts: {text[:1500]}")
                continue
            date, price = result
            print(f"{url}: {date} {price}")
            if date[:7] == today.isoformat()[:7]:
                found_this_month = True
            changed |= add(data, date, price, url)
    if changed:
        save(data)
        print("fuel-prices.json updated")
    if not found_this_month and today.day >= MISSING_AFTER_DAY:
        print("No announcement found for this month. Add it by hand: "
              "Actions > Fuel prices > Run workflow, with the date and price.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
