"""
FlatFinder Commercial — Toronto Full-Market Scraper
Platforms : Kijiji · Zumper · PadMapper · Craigslist · Rentals.ca · Apartments.ca
Coverage  : ALL unit types, ALL price ranges, ALL Toronto neighbourhoods
Output    : flatfinder_toronto.xlsx  (daily tab + All Listings tab)
            flatfinder_toronto_latest.csv  (Google Sheets import / API feed)
Run       : python flatfinder_scraper.py
Schedule  : GitHub Actions (.github/workflows/daily.yml) — runs 8 AM UTC daily
"""

import os, re, csv, time, random, logging, json, hashlib
import requests
from datetime import datetime, date
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── CONFIG ────────────────────────────────────────────────────────────────────
CITY        = "Toronto"
OUTPUT_DIR  = os.path.dirname(os.path.abspath(__file__))
XLSX_FILE   = os.path.join(OUTPUT_DIR, "flatfinder_toronto.xlsx")
CSV_FILE    = os.path.join(OUTPUT_DIR, "flatfinder_toronto_latest.csv")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-CA,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

COLS = [
    "ID", "Source", "Title", "Price", "Bedrooms", "Bathrooms",
    "Type", "Neighbourhood", "Address", "Utilities", "Pets",
    "TTC_Access", "Available", "URL", "Description", "Date_Scraped"
]

COL_WIDTHS = {
    "ID": 10, "Source": 12, "Title": 46, "Price": 10,
    "Bedrooms": 10, "Bathrooms": 10, "Type": 16,
    "Neighbourhood": 22, "Address": 30, "Utilities": 12,
    "Pets": 8, "TTC_Access": 12, "Available": 14,
    "URL": 14, "Description": 44, "Date_Scraped": 14
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger(__name__)

# ── STYLES ────────────────────────────────────────────────────────────────────
HDR_FILL   = PatternFill("solid", start_color="0D1B2A")
HDR_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=10)
EVEN_FILL  = PatternFill("solid", start_color="F4F6FB")
ODD_FILL   = PatternFill("solid", start_color="FFFFFF")
UTIL_FILL  = PatternFill("solid", start_color="D6F5E3")
PET_FILL   = PatternFill("solid", start_color="FFF3E0")
LINK_FONT  = Font(name="Arial", color="1155CC", underline="single", size=9)
BOLD9      = Font(name="Arial", size=9, bold=True)
REG9       = Font(name="Arial", size=9)
GREY8      = Font(name="Arial", size=8, color="666666")
CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT_W     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN       = Side(style="thin", color="D0D5E8")
BORDER     = Border(bottom=THIN)

# ── HELPERS ───────────────────────────────────────────────────────────────────
def fetch(url, retries=3, delay=2, extra_headers=None):
    h = dict(HEADERS)
    if extra_headers:
        h.update(extra_headers)
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=h, timeout=18)
            if r.status_code == 200:
                return r.text
            elif r.status_code == 403:
                log.warning(f"403 blocked: {url}")
                return None
            log.warning(f"HTTP {r.status_code} → {url}")
        except Exception as e:
            log.warning(f"Attempt {attempt+1} error: {e}")
        time.sleep(delay + random.uniform(0.5, 2.0))
    return None

def clean(t):
    return " ".join(t.strip().split()) if t else ""

def parse_price(text):
    if not text:
        return None
    m = re.search(r"[\d,]+", text.replace(",", ""))
    return int(m.group().replace(",", "")) if m else None

def detect_beds(text):
    t = text.lower()
    if any(w in t for w in ["bachelor", "studio", "bach", "0 bed"]):
        return "Bachelor/Studio"
    for n in ["5","6","7"]:
        if f"{n} bed" in t or f"{n}bed" in t or f"{n}br" in t:
            return f"{n}-Bed"
    for n, w in [("4","four"),("3","three"),("2","two"),("1","one")]:
        if f"{n} bed" in t or f"{n}bed" in t or f"{n}br" in t or f"{w} bed" in t:
            return f"{n}-Bed"
    return "Unknown"

def detect_baths(text):
    t = text.lower()
    for n in ["4","3","2"]:
        if f"{n} bath" in t or f"{n}bath" in t:
            return n
    return "1" if "bath" in t else "?"

def detect_utilities(text):
    t = text.lower()
    if any(w in t for w in ["all incl", "all-incl", "utilities incl", "all inclusive",
                              "heat incl", "hydro incl", "water incl", "bills incl",
                              "everything incl", "utilities included"]):
        return "Yes"
    if any(w in t for w in ["heat only", "water only", "hydro extra", "+ hydro",
                              "+ utilities", "utilities not", "hydro not incl"]):
        return "Partial"
    return "Check"

def detect_pets(text):
    t = text.lower()
    if any(w in t for w in ["pet friendly", "pets allowed", "pets ok", "dogs ok",
                              "cats ok", "pets welcome"]):
        return "Yes"
    if any(w in t for w in ["no pets", "pet free", "no dogs", "no cats"]):
        return "No"
    return "?"

def detect_ttc(text, address=""):
    t = (text + " " + address).lower()
    subway_keywords = [
        "subway", "ttc", "bloor-yonge", "spadina stn", "union stn",
        "osgoode", "st. patrick", "queen stn", "king stn", "dundas stn",
        "college stn", "wellesley", "sherbourne", "castle frank", "broadview",
        "chester", "pape", "donlands", "greenwood", "coxwell", "woodbine",
        "main street", "victoria park", "warden", "kennedy", "scarborough",
        "mccowan", "midland", "ellesmere", "lawrence east", "orion",
        "york mills", "sheppard", "wilson", "yorkdale", "lawrence",
        "eglinton", "davisville", "st. clair", "summerhill", "rosedale",
        "bloor-yonge", "bay", "museum", "queens park", "st. george",
        "dupont", "spadina", "bathurst", "ossington", "dufferin", "lansdowne",
        "dundas west", "runnymede", "jane", "runnymede", "old mill",
        "humber", "kipling", "islington", "royal york", "high park",
        "keele", "finch", "york", "pioneer village", "vaughan",
        "highway 407", "sheppard west", "downsview", "allen", "glencairn",
        "lawrence west", "yorkdale", "wilson", "finch west", "york university",
        "pioneer", "line 1", "line 2", "steps to subway", "walk to subway",
        "min to subway", "near subway", "close to subway"
    ]
    if any(kw in t for kw in subway_keywords):
        return "Subway"
    streetcar = ["streetcar", "504 ", "505 ", "506 ", "509 ", "510 ", "511 ",
                 "512 ", "queen st", "king st", "college st", "dundas st",
                 "bathurst st", "carlton"]
    if any(kw in t for kw in streetcar):
        return "Streetcar"
    if any(kw in t for kw in ["bus", "ttc bus", "transit"]):
        return "Bus"
    return "?"

def make_id(source, title, price):
    raw = f"{source}{title}{price}".encode()
    return hashlib.md5(raw).hexdigest()[:8].upper()

def detect_available(text):
    t = text.lower()
    patterns = [
        r"available\s+([\w\s,]+\d{4})",
        r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s.,-]+\d{1,2}[\s,]+\d{4}",
        r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
        r"(immediately|now|asap|right away)",
        r"march 1|april 1|may 1",
    ]
    for p in patterns:
        m = re.search(p, t)
        if m:
            return clean(m.group())[:20]
    return ""

# ── SCRAPERS ──────────────────────────────────────────────────────────────────

def scrape_kijiji():
    listings = []
    pages = [
        ("https://www.kijiji.ca/b-apartments-condos/city-of-toronto/c37l1700273", "All"),
        ("https://www.kijiji.ca/b-apartments-condos/city-of-toronto/bachelor+studio/c37l1700273a27949001", "Bachelor"),
        ("https://www.kijiji.ca/b-apartments-condos/city-of-toronto/1-bedroom/c37l1700273a29276001", "1-Bed"),
        ("https://www.kijiji.ca/b-apartments-condos/city-of-toronto/2-bedroom/c37l1700273a29277001", "2-Bed"),
        ("https://www.kijiji.ca/b-apartments-condos/city-of-toronto/3-bedroom/c37l1700273a29278001", "3-Bed"),
    ]
    for url, hint in pages:
        html = fetch(url)
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for item in soup.select("li[data-testid='listing-card-list-item'], div[data-testid='listing-card']"):
            try:
                title_el = item.select_one("[class*='title']")
                price_el = item.select_one("[class*='price']")
                loc_el   = item.select_one("[class*='location']")
                link_el  = item.select_one("a[href*='/v-']")
                desc_el  = item.select_one("[class*='description']")

                title    = clean(title_el.get_text()) if title_el else ""
                price_raw= clean(price_el.get_text()) if price_el else ""
                price    = parse_price(price_raw)
                loc      = clean(loc_el.get_text()) if loc_el else CITY
                href     = "https://www.kijiji.ca" + link_el["href"] if link_el and link_el.get("href") else ""
                desc     = clean(desc_el.get_text()) if desc_el else ""
                combined = title + " " + desc

                if not title:
                    continue
                listings.append({
                    "ID": make_id("Kijiji", title, price),
                    "Source": "Kijiji",
                    "Title": title,
                    "Price": price,
                    "Bedrooms": detect_beds(title) if hint == "All" else hint,
                    "Bathrooms": detect_baths(combined),
                    "Type": "Apartment",
                    "Neighbourhood": loc,
                    "Address": "",
                    "Utilities": detect_utilities(combined),
                    "Pets": detect_pets(combined),
                    "TTC_Access": detect_ttc(combined),
                    "Available": detect_available(combined),
                    "URL": href,
                    "Description": desc[:220],
                    "Date_Scraped": str(date.today()),
                })
            except Exception:
                continue
        time.sleep(random.uniform(2, 4))
    log.info(f"Kijiji: {len(listings)}")
    return listings


def scrape_zumper():
    listings = []
    urls = [
        "https://www.zumper.com/apartments-for-rent/toronto-on",
        "https://www.zumper.com/apartments-for-rent/toronto-on?beds=0",
        "https://www.zumper.com/apartments-for-rent/toronto-on?beds=1",
        "https://www.zumper.com/apartments-for-rent/toronto-on?beds=2",
        "https://www.zumper.com/apartments-for-rent/toronto-on?beds=3",
    ]
    for url in urls:
        html = fetch(url)
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for card in soup.select("article, [data-testid*='listing'], [class*='ListingCard'], [class*='listing-card']"):
            try:
                title_el = card.select_one("h2, h3, [class*='title'], [class*='address'], [class*='Address']")
                price_el = card.select_one("[class*='price'], [class*='Price']")
                link_el  = card.select_one("a[href]")
                bed_el   = card.select_one("[class*='bed'], [class*='Bed'], [class*='room']")
                bath_el  = card.select_one("[class*='bath'], [class*='Bath']")
                hood_el  = card.select_one("[class*='neighbour'], [class*='location'], [class*='Location']")

                title    = clean(title_el.get_text()) if title_el else ""
                price_raw= clean(price_el.get_text()) if price_el else ""
                price    = parse_price(price_raw)
                beds     = clean(bed_el.get_text()) if bed_el else ""
                baths    = clean(bath_el.get_text()) if bath_el else ""
                hood     = clean(hood_el.get_text()) if hood_el else CITY
                href     = link_el["href"] if link_el else ""
                if href and not href.startswith("http"):
                    href = "https://www.zumper.com" + href

                if not title:
                    continue
                listings.append({
                    "ID": make_id("Zumper", title, price),
                    "Source": "Zumper",
                    "Title": title,
                    "Price": price,
                    "Bedrooms": detect_beds(title + " " + beds),
                    "Bathrooms": detect_baths(baths) if baths else "?",
                    "Type": "Apartment",
                    "Neighbourhood": hood,
                    "Address": title if any(c.isdigit() for c in title[:5]) else "",
                    "Utilities": "Check",
                    "Pets": "?",
                    "TTC_Access": detect_ttc(title + " " + hood),
                    "Available": "",
                    "URL": href,
                    "Description": beds + (" | " + baths if baths else ""),
                    "Date_Scraped": str(date.today()),
                })
            except Exception:
                continue
        time.sleep(random.uniform(1.5, 3))
    log.info(f"Zumper: {len(listings)}")
    return listings


def scrape_padmapper():
    listings = []
    # Try public API first
    api_urls = [
        ("https://www.padmapper.com/api/t/1/listings?"
         "section=apartment&center_lat=43.6532&center_lng=-79.3832"
         "&zoom=11&beds_min=0&beds_max=6&limit=200", "api"),
    ]
    for url, mode in api_urls:
        try:
            r = requests.get(url, headers=HEADERS, timeout=18)
            if r.status_code == 200:
                data = r.json()
                items = data if isinstance(data, list) else data.get("data", data.get("listings", []))
                for item in (items or []):
                    price  = item.get("price") or item.get("min_price")
                    title  = item.get("address") or item.get("title") or item.get("street") or ""
                    slug   = item.get("slug") or item.get("id") or ""
                    href   = f"https://www.padmapper.com/rentals/{slug}" if slug else ""
                    beds   = str(item.get("bed_label") or item.get("beds") or "")
                    baths  = str(item.get("bath_label") or item.get("baths") or "?")
                    util   = "Yes" if item.get("utilities_included") else "Check"
                    pets   = "Yes" if item.get("pets_allowed") else "?"
                    hood   = item.get("neighbourhood") or item.get("neighborhood") or CITY
                    avail  = item.get("available_from") or ""
                    desc   = item.get("description") or ""
                    listings.append({
                        "ID": make_id("PadMapper", title, price),
                        "Source": "PadMapper",
                        "Title": clean(title),
                        "Price": price,
                        "Bedrooms": detect_beds(beds),
                        "Bathrooms": baths,
                        "Type": "Apartment",
                        "Neighbourhood": clean(hood),
                        "Address": clean(title),
                        "Utilities": util,
                        "Pets": pets,
                        "TTC_Access": detect_ttc(clean(hood) + " " + clean(desc)),
                        "Available": str(avail)[:20],
                        "URL": href,
                        "Description": clean(desc)[:220],
                        "Date_Scraped": str(date.today()),
                    })
                break
        except Exception as e:
            log.warning(f"PadMapper API: {e}")

    # Fallback: scrape HTML
    if not listings:
        html = fetch("https://www.padmapper.com/apartments/toronto-on")
        if html:
            soup = BeautifulSoup(html, "lxml")
            for card in soup.select("[class*='ListItem'], [class*='listing-card'], [class*='Card']"):
                try:
                    title_el = card.select_one("[class*='address'], h3, h2, [class*='title']")
                    price_el = card.select_one("[class*='price']")
                    link_el  = card.select_one("a[href]")
                    title    = clean(title_el.get_text()) if title_el else ""
                    price_raw= clean(price_el.get_text()) if price_el else ""
                    price    = parse_price(price_raw)
                    href     = link_el["href"] if link_el else ""
                    if href and not href.startswith("http"):
                        href = "https://www.padmapper.com" + href
                    if not title:
                        continue
                    listings.append({
                        "ID": make_id("PadMapper", title, price),
                        "Source": "PadMapper", "Title": title, "Price": price,
                        "Bedrooms": detect_beds(title), "Bathrooms": "?",
                        "Type": "Apartment", "Neighbourhood": CITY, "Address": "",
                        "Utilities": "Check", "Pets": "?",
                        "TTC_Access": detect_ttc(title), "Available": "",
                        "URL": href, "Description": "", "Date_Scraped": str(date.today()),
                    })
                except Exception:
                    continue
    log.info(f"PadMapper: {len(listings)}")
    return listings


def scrape_craigslist():
    """Craigslist — cautious: one request per category, long delays."""
    listings = []
    urls = [
        "https://toronto.craigslist.org/search/tor/apa",
        "https://toronto.craigslist.org/search/tor/roo",
    ]
    for url in urls:
        html = fetch(url, retries=2, delay=4)
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for item in soup.select("li.cl-search-result, .result-row"):
            try:
                title_el = item.select_one(".cl-app-anchor, .result-title, a.hdrlnk")
                price_el = item.select_one(".priceinfo, .result-price")
                hood_el  = item.select_one(".supertitle, .result-hood")
                link_el  = item.select_one("a[href]")
                meta_el  = item.select_one(".housing, .result-meta")

                title    = clean(title_el.get_text()) if title_el else ""
                price_raw= clean(price_el.get_text()) if price_el else ""
                price    = parse_price(price_raw)
                hood     = clean(hood_el.get_text()).strip("() ") if hood_el else CITY
                href     = link_el["href"] if link_el else ""
                if href and not href.startswith("http"):
                    href = "https://toronto.craigslist.org" + href
                meta     = clean(meta_el.get_text()) if meta_el else ""
                combined = title + " " + meta

                if not title:
                    continue
                listings.append({
                    "ID": make_id("Craigslist", title, price),
                    "Source": "Craigslist",
                    "Title": title,
                    "Price": price,
                    "Bedrooms": detect_beds(combined),
                    "Bathrooms": detect_baths(combined),
                    "Type": "Apartment",
                    "Neighbourhood": hood,
                    "Address": "",
                    "Utilities": detect_utilities(combined),
                    "Pets": detect_pets(combined),
                    "TTC_Access": detect_ttc(combined, hood),
                    "Available": detect_available(combined),
                    "URL": href,
                    "Description": meta[:220],
                    "Date_Scraped": str(date.today()),
                })
            except Exception:
                continue
        time.sleep(random.uniform(5, 9))  # extra cautious
    log.info(f"Craigslist: {len(listings)}")
    return listings


def scrape_rentals_ca():
    listings = []
    urls = [
        "https://rentals.ca/toronto",
        "https://rentals.ca/toronto?beds[]=bachelor-studio",
        "https://rentals.ca/toronto?beds[]=1-bedroom",
        "https://rentals.ca/toronto?beds[]=2-bedroom",
        "https://rentals.ca/toronto?beds[]=3-bedroom",
    ]
    for url in urls:
        html = fetch(url, extra_headers={"Referer": "https://rentals.ca/"})
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for card in soup.select(
            "[class*='listing-card'], [class*='ListingCard'], "
            "article[class*='listing'], [data-testid*='listing']"
        ):
            try:
                title_el = card.select_one("h2, h3, [class*='title'], [class*='address']")
                price_el = card.select_one("[class*='price'], [class*='Price'], [class*='rent']")
                hood_el  = card.select_one("[class*='location'], [class*='neighbourhood'], [class*='address']")
                link_el  = card.select_one("a[href]")
                bed_el   = card.select_one("[class*='bed'], [class*='Bed']")
                desc_el  = card.select_one("[class*='desc'], [class*='summary']")

                title    = clean(title_el.get_text()) if title_el else ""
                price_raw= clean(price_el.get_text()) if price_el else ""
                price    = parse_price(price_raw)
                hood     = clean(hood_el.get_text()) if hood_el else CITY
                href     = link_el.get("href","") if link_el else ""
                if href and not href.startswith("http"):
                    href = "https://rentals.ca" + href
                beds     = clean(bed_el.get_text()) if bed_el else ""
                desc     = clean(desc_el.get_text()) if desc_el else ""
                combined = title + " " + beds + " " + desc

                if not title:
                    continue
                listings.append({
                    "ID": make_id("Rentals.ca", title, price),
                    "Source": "Rentals.ca",
                    "Title": title,
                    "Price": price,
                    "Bedrooms": detect_beds(combined),
                    "Bathrooms": detect_baths(combined),
                    "Type": "Apartment",
                    "Neighbourhood": hood,
                    "Address": title if any(c.isdigit() for c in title[:5]) else "",
                    "Utilities": detect_utilities(combined),
                    "Pets": detect_pets(combined),
                    "TTC_Access": detect_ttc(combined, hood),
                    "Available": detect_available(combined),
                    "URL": href,
                    "Description": desc[:220],
                    "Date_Scraped": str(date.today()),
                })
            except Exception:
                continue
        time.sleep(random.uniform(2, 4))
    log.info(f"Rentals.ca: {len(listings)}")
    return listings


def scrape_apartments_ca():
    listings = []
    urls = [
        "https://www.apartments.ca/toronto/",
        "https://www.apartments.ca/toronto/bachelor-apartments/",
        "https://www.apartments.ca/toronto/1-bedroom-apartments/",
        "https://www.apartments.ca/toronto/2-bedroom-apartments/",
        "https://www.apartments.ca/toronto/3-bedroom-apartments/",
    ]
    for url in urls:
        html = fetch(url, extra_headers={"Referer": "https://www.apartments.ca/"})
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for card in soup.select(
            "[class*='listing'], [class*='Listing'], article, "
            "[class*='property-card'], [class*='PropertyCard']"
        ):
            try:
                title_el = card.select_one("h2, h3, [class*='title'], [class*='name'], [class*='address']")
                price_el = card.select_one("[class*='price'], [class*='rent'], [class*='Price']")
                hood_el  = card.select_one("[class*='location'], [class*='neighbourhood'], [class*='area']")
                link_el  = card.select_one("a[href]")
                bed_el   = card.select_one("[class*='bed'], [class*='Bed'], [class*='room']")
                desc_el  = card.select_one("[class*='desc'], [class*='summary'], [class*='amenity']")

                title    = clean(title_el.get_text()) if title_el else ""
                price_raw= clean(price_el.get_text()) if price_el else ""
                price    = parse_price(price_raw)
                hood     = clean(hood_el.get_text()) if hood_el else CITY
                href     = link_el.get("href","") if link_el else ""
                if href and not href.startswith("http"):
                    href = "https://www.apartments.ca" + href
                beds     = clean(bed_el.get_text()) if bed_el else ""
                desc     = clean(desc_el.get_text()) if desc_el else ""
                combined = title + " " + beds + " " + desc

                if not title or len(title) < 4:
                    continue
                listings.append({
                    "ID": make_id("Apartments.ca", title, price),
                    "Source": "Apartments.ca",
                    "Title": title,
                    "Price": price,
                    "Bedrooms": detect_beds(combined),
                    "Bathrooms": detect_baths(combined),
                    "Type": "Apartment",
                    "Neighbourhood": hood,
                    "Address": title if any(c.isdigit() for c in title[:5]) else "",
                    "Utilities": detect_utilities(combined),
                    "Pets": detect_pets(combined),
                    "TTC_Access": detect_ttc(combined, hood),
                    "Available": detect_available(combined),
                    "URL": href,
                    "Description": desc[:220],
                    "Date_Scraped": str(date.today()),
                })
            except Exception:
                continue
        time.sleep(random.uniform(2, 3.5))
    log.info(f"Apartments.ca: {len(listings)}")
    return listings


# ── DEDUP ─────────────────────────────────────────────────────────────────────
def deduplicate(listings):
    seen, out = set(), []
    for l in listings:
        key = l["ID"]
        if key not in seen:
            seen.add(key)
            out.append(l)
    return out


# ── XLSX WRITER ───────────────────────────────────────────────────────────────
def style_row(ws, ri, l, fill):
    ws.row_dimensions[ri].height = 18
    for ci, col in enumerate(COLS, 1):
        val  = l.get(col, "")
        cell = ws.cell(ri, ci)
        cell.border = BORDER

        if col == "Price":
            cell.value         = l.get("Price")
            cell.number_format = '"$"#,##0'
            cell.font          = Font(name="Arial", size=9, bold=True)
            cell.alignment     = CENTER
            cell.fill          = fill

        elif col == "URL" and val:
            cell.value     = "Open"
            cell.hyperlink = val
            cell.font      = LINK_FONT
            cell.alignment = CENTER
            cell.fill      = fill

        elif col == "Utilities":
            cell.value     = val
            cell.font      = Font(name="Arial", size=9, bold=(val=="Yes"),
                                  color="1A7A3C" if val=="Yes" else
                                        "E65100" if val=="Partial" else "555555")
            cell.fill      = UTIL_FILL if val=="Yes" else fill
            cell.alignment = CENTER

        elif col == "Pets":
            cell.value     = val
            cell.font      = Font(name="Arial", size=9,
                                  color="1A7A3C" if val=="Yes" else
                                        "C62828" if val=="No" else "555555")
            cell.fill      = PET_FILL if val=="Yes" else fill
            cell.alignment = CENTER

        elif col == "TTC_Access":
            color = {"Subway":"1155CC","Streetcar":"6A1B9A","Bus":"2E7D32"}.get(val,"555555")
            cell.value     = val
            cell.font      = Font(name="Arial", size=9, color=color,
                                  bold=(val in ("Subway","Streetcar")))
            cell.fill      = fill
            cell.alignment = CENTER

        elif col == "Title":
            cell.value     = val
            cell.font      = BOLD9
            cell.fill      = fill
            cell.alignment = LEFT_W

        elif col == "Description":
            cell.value     = val
            cell.font      = GREY8
            cell.fill      = fill
            cell.alignment = LEFT_W

        elif col in ("ID","Source","Bedrooms","Bathrooms","Type","Available","Date_Scraped"):
            cell.value     = val
            cell.font      = REG9
            cell.fill      = fill
            cell.alignment = CENTER

        else:
            cell.value     = val
            cell.font      = REG9
            cell.fill      = fill
            cell.alignment = LEFT_W


def write_sheet(ws, listings, sheet_label):
    ws.row_dimensions[1].height = 22
    for ci, col in enumerate(COLS, 1):
        cell = ws.cell(1, ci, col)
        cell.font      = HDR_FONT
        cell.fill      = HDR_FILL
        cell.alignment = CENTER
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = COL_WIDTHS[col]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}1"

    for ri, l in enumerate(listings, 2):
        fill = EVEN_FILL if ri % 2 == 0 else ODD_FILL
        style_row(ws, ri, l, fill)

    last = len(listings) + 2
    ws.cell(last, 1, f"Total: {len(listings)}").font = BOLD9
    ws.cell(last, 3, f'=COUNTA(C2:C{last-1})').font  = REG9
    ws.cell(last, 4, "Avg:").font                     = REG9
    ws.cell(last, 5, f'=IFERROR(AVERAGE(D2:D{last-1}),"-")').number_format = '"$"#,##0'
    ws.cell(last, 5).font = REG9


def write_xlsx(listings):
    today_str = str(date.today())

    if os.path.exists(XLSX_FILE):
        wb = load_workbook(XLSX_FILE)
    else:
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Today's tab
    if today_str in wb.sheetnames:
        del wb[today_str]
    ws_today = wb.create_sheet(title=today_str, index=0)
    write_sheet(ws_today, listings, today_str)

    # All Listings tab — append new unique rows
    ALL = "All Listings"
    if ALL not in wb.sheetnames:
        wa = wb.create_sheet(ALL)
        write_sheet(wa, [], ALL)
    else:
        wa = wb[ALL]

    existing_ids = set()
    for row in wa.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            existing_ids.add(row[0])

    new_rows = [l for l in listings if l["ID"] not in existing_ids]
    start_ri = wa.max_row + 1
    for ri, l in enumerate(new_rows, start_ri):
        fill = EVEN_FILL if ri % 2 == 0 else ODD_FILL
        wa.row_dimensions[ri].height = 18
        style_row(wa, ri, l, fill)

    # Stats sheet
    STATS = "📊 Stats"
    if STATS in wb.sheetnames:
        del wb[STATS]
    ws_stats = wb.create_sheet(STATS)
    write_stats_sheet(ws_stats, listings, today_str)

    wb.save(XLSX_FILE)
    log.info(f"XLSX saved → {XLSX_FILE}  |  today: {len(listings)}  |  new to All: {len(new_rows)}")


def write_stats_sheet(ws, listings, today_str):
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18

    title_cell = ws.cell(1, 1, f"FlatFinder Toronto — {today_str}")
    title_cell.font = Font(name="Arial", size=13, bold=True, color="0D1B2A")
    ws.merge_cells("A1:B1")

    rows = [
        ("Total listings today", len(listings)),
        ("Sources scraped", len(set(l["Source"] for l in listings))),
        ("Bachelor/Studio", sum(1 for l in listings if "Bach" in str(l.get("Bedrooms","")) or "Studio" in str(l.get("Bedrooms","")))),
        ("1-Bedroom",       sum(1 for l in listings if l.get("Bedrooms") == "1-Bed")),
        ("2-Bedroom",       sum(1 for l in listings if l.get("Bedrooms") == "2-Bed")),
        ("3-Bedroom+",      sum(1 for l in listings if l.get("Bedrooms") in ("3-Bed","4-Bed","5-Bed","6-Bed"))),
        ("Utilities Included", sum(1 for l in listings if l.get("Utilities") == "Yes")),
        ("Pet Friendly",    sum(1 for l in listings if l.get("Pets") == "Yes")),
        ("Subway Access",   sum(1 for l in listings if l.get("TTC_Access") == "Subway")),
        ("Avg Price (all)",  int(sum(l["Price"] for l in listings if l.get("Price")) / max(1, sum(1 for l in listings if l.get("Price"))))),
        ("Min Price",       min((l["Price"] for l in listings if l.get("Price")), default=0)),
        ("Max Price",       max((l["Price"] for l in listings if l.get("Price")), default=0)),
    ]

    for ri, (label, value) in enumerate(rows, 3):
        lc = ws.cell(ri, 1, label)
        vc = ws.cell(ri, 2, value)
        lc.font = REG9
        vc.font = BOLD9
        lc.fill = vc.fill = EVEN_FILL if ri % 2 == 0 else ODD_FILL
        lc.alignment = vc.alignment = Alignment(horizontal="left", vertical="center")
        vc.number_format = '"$"#,##0' if "Price" in label and "Avg" not in label else ('"$"#,##0' if "Price" in label else "General")
        lc.border = vc.border = BORDER

    # By source breakdown
    ws.cell(len(rows)+5, 1, "By Source").font = BOLD9
    by_source = {}
    for l in listings:
        by_source[l["Source"]] = by_source.get(l["Source"], 0) + 1
    for ri, (src, cnt) in enumerate(sorted(by_source.items(), key=lambda x:-x[1]), len(rows)+6):
        ws.cell(ri, 1, src).font  = REG9
        ws.cell(ri, 2, cnt).font  = BOLD9
        ws.cell(ri, 1).fill = ws.cell(ri, 2).fill = EVEN_FILL if ri%2==0 else ODD_FILL
        ws.cell(ri, 1).border = ws.cell(ri, 2).border = BORDER


# ── CSV WRITER ────────────────────────────────────────────────────────────────
def write_csv(listings):
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLS)
        writer.writeheader()
        for l in listings:
            writer.writerow({col: l.get(col, "") for col in COLS})
    log.info(f"CSV  saved → {CSV_FILE}")


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    log.info("=" * 55)
    log.info("  FlatFinder Commercial Scraper — START")
    log.info("=" * 55)

    all_listings = []
    all_listings += scrape_kijiji()
    all_listings += scrape_zumper()
    all_listings += scrape_padmapper()
    all_listings += scrape_craigslist()
    all_listings += scrape_rentals_ca()
    all_listings += scrape_apartments_ca()

    all_listings = deduplicate(all_listings)
    all_listings.sort(key=lambda x: (x.get("Price") or 999999))

    log.info(f"Total unique: {len(all_listings)}")
    write_xlsx(all_listings)
    write_csv(all_listings)
    log.info("  FlatFinder Commercial Scraper — DONE")
    log.info("=" * 55)

if __name__ == "__main__":
    main()
