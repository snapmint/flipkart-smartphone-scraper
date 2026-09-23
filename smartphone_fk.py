"""
Flipkart MOBILE Assortment Scraper — v1.0
=====================================================
Derived from the Trimmer scraper (v1.0), which itself derived from the
Air Fryer scraper (v1.0). Retains: leaf-card detection via data-id,
Playwright-only fetching, persistent tab per source, multi-source +
cross-source dedupe, consecutive-failure circuit breaker, checkpoint
saves, diagnostic mode.

v1.0 CHANGES (vs Trimmer v1.0):

  [A] BASE URL STRATEGY UNCHANGED — Mobiles scraped from a CATEGORY
      BROWSE page (".../mobiles/pr?sid=...&sort=popularity"). Pagination
      confirmed identical via "&page=N" on a live page-2 URL
      ("Showing 25 – 48 products of 7,917 products"). build_url()
      required ZERO changes.

      Ships with ONLY the single main "Popularity" category URL per
      user's plan — brand-filtered URLs to be appended to BASE_URLS
      later, same as the trimmer rollout. Multi-source loop already
      supports N sources with no further changes.

  [B] ⚠️ PRODUCTS PER PAGE = 24, NOT 40. Confirmed via live page text
      "Showing 1 – 24 products of 7,917 products". PRODUCTS_PER_PAGE_DEFAULT
      changed accordingly — this cascades correctly into
      determine_total_pages(), is_partial_page(), and the breaker logic
      with no other changes needed.

  [C] CARD ANCHOR STRUCTURE DIFFERS FROM TRIMMER TEMPLATE — mobile cards
      wrap the ENTIRE card in ONE anchor (a.k7wcnx, no title attribute),
      instead of two separate anchors (image-only + titled text anchor).
      find_product_link() still works unmodified (picks the single match).
      extract_title()'s fallback chain (div.RG5Slk) already handles this
      correctly since the anchor's own get_text() fails the "no ₹ in
      text" check (price text is baked into the same anchor) — ZERO
      changes needed to extract_title().

  [D] SPEC BLOCK IS A REAL <ul>/<li> LIST (div.CMXw7N > ul.HwRTzP >
      li.DTBslk), NOT a comma subtitle like trimmers' U_GKRr. Confirmed
      via inspect. extract_specs()'s EXISTING second-priority fallback
      (`ul.HwRTzP` selector) already targets this exact class — ZERO
      changes needed to extract_specs(). Only the DOWNSTREAM parsing of
      those spec strings (derive_attributes) is new, built around
      mobile-specific line formats:
        "6 GB RAM | 128 GB ROM | Expandable Upto 1 TB"
        "17.13 cm (6.745 inch) HD+ Display"
        "50MP Rear Camera | 8MP Front Camera"  (or "50MP + 2MP | 8MP Front Camera")
        "6000 mAh Battery" / "6000 mAh Lithium ion Battery"
        "T8200 Processor" / "Helio G99 Processor"
        "1 Year Manufacturer Warranty for Device and 6 Months ..."

  [E] ⚠️ RIBBON TEXT ("Bestseller") IS CSS ::before PSEUDO-CONTENT, NOT
      REAL DOM TEXT on this template — confirmed via inspect (div.o2uEoz
      has an empty node with a "::before" pseudo-element annotation
      showing the actual string "Bestseller"). BeautifulSoup/page.content()
      can NEVER see pseudo-element content. FIX: inject a small
      page.evaluate() JS snippet in SourceFetcher.fetch() BEFORE calling
      page.content() that reads getComputedStyle(el, '::before').content
      for every div.o2uEoz and writes it into a real `data-ribbon-text`
      attribute on that element. extract_badges() then reads that
      attribute first (falls back to get_text() for safety/future-proofing
      in case some ribbons DO render as real text).

  [F] PRICE BLOCK / EXCHANGE OFFER / BANK OFFER CONFIRMED STRUCTURALLY
      IDENTICAL to trimmer template via inspect: div.oFEPlD > div.QiMO5r
      > div.hZ3P6w (price), div.MaiFhH > div.hx1EGN > div.HZ0E6r.Rm9_cy
      (Exchange offer text + Bank Offer text). extract_price_block() and
      the badge-extraction offer logic needed NO changes.

  [G] RATING BLOCK CONFIRMED IDENTICAL: span[id^=productRating_] >
      div.MKiFS6 + span.PvbNMB with "X Ratings & Y Reviews" text.
      extract_rating_info() needed NO changes.

  [H] NEW DERIVED FIELDS for mobiles (replacing trimmer's Trimmer-Type/
      Category/Runtime/Length-Settings/Body-Material focus):
        - RAM / ROM (Storage) / Expandable Storage — parsed from the
          "X GB RAM | Y GB ROM | Expandable Upto Z TB" spec line
        - Display Size (inch) / Display Type — parsed from the
          "cm (X inch) TYPE Display" spec line
        - Rear Camera / Front Camera — parsed from the camera spec line,
          handling BOTH "50MP Rear Camera | 8MP Front Camera" and
          unlabeled dual-camera "50MP + 2MP | 8MP Front Camera" formats
        - Battery (mAh) — parsed from the "X mAh ... Battery" spec line
        - Processor — parsed from the "... Processor" spec line
        - Network Type (5G/4G/3G/2G) — derived from title+specs text
        - Color — mobiles have NO dedicated color subtitle div (unlike
          trimmers' U_GKRr); color lives INSIDE the title's trailing
          parentheses, e.g. "Ai+ Nova 2 5G (Blue, 128 GB)" → "Blue".
          Filtered so purely-numeric parenthetical content (e.g. just
          "(128 GB)" with no color) does NOT get misread as a color.
        - Warranty — raw text of whichever spec line mentions "Warranty"

  [I] KNOWN_BRANDS replaced with mobile-specific names (Samsung, Apple,
      Xiaomi, Redmi, POCO, Realme, Vivo, OPPO, Motorola, Nokia, Google,
      Infinix, iQOO, Micromax, Nothing, Honor, plus budget/local brands
      confirmed live: BOLTT, Ai+, Kechaoda, HMD).

  [J] Everything else (Playwright-only fetch, persistent-tab-per-source,
      circuit breaker, checkpointing, dual-source dedupe, leaf-card
      detection, Assured detection, Excel writer skeleton) retained
      verbatim — confirmed structurally identical via inspect.

SETUP:
    pip install beautifulsoup4 openpyxl playwright lxml
    playwright install chromium
"""

import re
import math
import time
import random
import logging
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Any

from bs4 import BeautifulSoup, Tag
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
#                  ⚙️ CONFIGURATION
# ============================================================

BASE_URLS = [
    {
        "label": "Popularity-Main",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
        ),
    },
    {
        "label": "Brand-Apple",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DApple"
        ),
    },
    {
        "label": "Brand-Google",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DGoogle"
        ),
    },
    {
        "label": "Brand-Motorola",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DMOTOROLA"
        ),
    },
    {
        "label": "Brand-Vivo",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3Dvivo"
        ),
    },
    {
        "label": "Brand-Oppo",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DOPPO"
        ),
    },
    {
        "label": "Brand-Nothing",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DNothing"
        ),
    },
    {
        "label": "Brand-Infinix",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DInfinix"
        ),
    },
    {
        "label": "Brand-Poco",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DPOCO"
        ),
    },
    {
        "label": "Brand-Realme",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3Drealme"
        ),
    },
    {
        "label": "Brand-Samsung",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DSamsung"
        ),
    },
    {
        "label": "Brand-Redmi",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DREDMI"
        ),
    },
    {
        "label": "Brand-OnePlus",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DOnePlus"
        ),
    },
    {
        "label": "Brand-iQOO",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DIQOO"
        ),
    },
    {
        "label": "Brand-AiPlus",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DAi%252B"
        ),
    },
    {
        "label": "Brand-Mi",
        "url": (
            "https://www.flipkart.com/mobiles/pr"
            "?sid=tyy%2C4io&marketplace=FLIPKART&sort=popularity"
            "&p%5B%5D=facets.brand%255B%255D%3DMi"
        ),
    },
]

OUTPUT_FILE = f"flipkart_mobile_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

MAX_PAGES = 150                         # Hard cap PER SOURCE (breaker usually ends sources earlier)
PRODUCTS_PER_PAGE_DEFAULT = 24          # [B] Confirmed: "Showing 1 – 24 products of 7,917 products"
MIN_ACCEPTABLE_PRODUCTS_RATIO = 0.70

CONSECUTIVE_FAILURE_LIMIT = 3
RECYCLE_CONTEXT_EVERY_N_PAGES = 25      # periodic tab refresh

PAGE_DELAY_MIN = 1.5                    # polite pacing between page loads (Playwright, not requests)
PAGE_DELAY_MAX = 3.5
SAVE_EVERY_N_PAGES = 5
MAX_RETRIES_PER_URL = 3                 # Playwright-level retries (fresh context each retry)

DIAGNOSTIC_MODE = False
DIAGNOSTIC_SOURCE_INDEX = 0
GLOBAL_DEDUPE = True

LOG_LEVEL = logging.INFO

# ============================================================
#                    LOGGING SETUP
# ============================================================
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# [C] Category-relevant keywords used by the anchor-climb fallback
# leaf-card detector (Strategy B) if the primary data-id strategy fails.
CATEGORY_KEYWORDS = [
    "smartphone", "mobile", "ram", "rom", "battery", "camera", "display",
    "processor", "5g", "4g", "storage", "mah", "inch", "warranty",
]

PATTERNS = {
    "discount": re.compile(r"(\d+)\s*%\s*off", re.IGNORECASE),
    "rating_count_worded": re.compile(r"([\d,]+)\s*Ratings?\b", re.IGNORECASE),
    "review_count": re.compile(r"([\d,]+)\s*Reviews?\b", re.IGNORECASE),
    "rating_count_paren": re.compile(r"\(\s*([\d,]+)\s*\)"),
    "rating_strict": re.compile(r"^([0-5](?:\.\d)?)$"),
    "total_products": re.compile(r"of\s+([\d,]+)\s+(?:results|products|items)", re.IGNORECASE),
    "junk_prefix": re.compile(r"^(add\s+to\s+compare|compare)\s*", re.IGNORECASE),
    "product_href": re.compile(r"/p/(itm[A-Za-z0-9]+|[A-Za-z0-9\-]+)"),
    "pid_query_param": re.compile(r"[?&]pid=([A-Z0-9]{10,20})", re.IGNORECASE),
    "itm_path_id": re.compile(r"/p/(itm[A-Za-z0-9]+)"),
    "or_pay": re.compile(r"Or\s*Pay\s*₹\s*([\d,]+)\s*\+\s*(\d+)", re.IGNORECASE),
    "exchange_discount": re.compile(
        r"Upto\s*₹\s*([\d,]+)\s*Off\s*on\s*Exchange", re.IGNORECASE
    ),
    # [H] NEW: mobile-specific spec-line parsers
    "ram_rom": re.compile(
        r"(\d+(?:\.\d+)?)\s*(GB|MB)\s*RAM\s*\|\s*(\d+(?:\.\d+)?)\s*(GB|MB)\s*ROM"
        r"(?:\s*\|\s*Expandable\s*Upto\s*(\d+(?:\.\d+)?)\s*(GB|TB))?",
        re.IGNORECASE
    ),
    "display": re.compile(
        r"([\d.]+)\s*cm\s*\(\s*([\d.]+)\s*inch\)\s*([A-Za-z0-9+\s]*?)\s*Display",
        re.IGNORECASE
    ),
    "battery_mah": re.compile(r"([\d,]+)\s*mAh", re.IGNORECASE),
    "processor_line": re.compile(r"^(.*?)\s*Processor\s*$", re.IGNORECASE),
    "color_from_title": re.compile(r"\(([^)]*)\)\s*$"),
}

# [I] Mobile-specific brands confirmed live (BOLTT, Ai+, Kechaoda, HMD)
# plus known category brands.
KNOWN_BRANDS = sorted([
    "Samsung", "Apple", "OnePlus", "Xiaomi", "Redmi", "POCO", "Realme",
    "Vivo", "OPPO", "Motorola", "Nokia", "Google", "Infinix", "Tecno",
    "iQOO", "Micromax", "Lava", "itel", "Nothing", "Asus", "Honor",
    "Huawei", "Sony", "LG", "HTC", "Panasonic", "BlackBerry", "Cat",
    "Gionee", "Karbonn", "Spice", "Videocon", "Zen", "Yu", "Meizu",
    "ZTE", "Alcatel", "HMD", "Jio", "iKall", "I Kall", "Intex",
    "Coolpad", "LYF", "BOLTT", "Ai+", "Kechaoda", "Cellecor",
    "Snexian", "Duke", "Ziox", "Swipe", "Sansui", "Renewed",
], key=len, reverse=True)

# [A] Safety-net ribbon labels — scanned against full card text as a
# backup in case a ribbon renders under yet another unseen class name,
# OR as real DOM text instead of a ::before pseudo-element.
RIBBON_KEYWORDS = [
    "bestseller", "trending", "flipkart's choice", "flipkart choice",
    "deal of the day", "hot deal", "lowest price", "top rated",
    "new launch", "limited deal", "most loved",
]


# ============================================================
#                  SMALL UTILITIES
# ============================================================

def clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return " ".join(text.replace("\xa0", " ").split()).strip()


def extract_number(raw: str) -> Optional[int]:
    if not raw:
        return None
    cleaned = str(raw).replace(",", "").replace("₹", "").strip()
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def extract_pid_from_url(href: str) -> str:
    m = PATTERNS["pid_query_param"].search(href)
    if m:
        return m.group(1).upper()
    m = PATTERNS["itm_path_id"].search(href)
    if m:
        return m.group(1)
    return href


def pid_from_card(card: Tag) -> Optional[str]:
    if card.has_attr("data-id") and card["data-id"].strip():
        return card["data-id"].strip().upper()
    parent = card.find_parent(attrs={"data-id": True})
    if parent and parent.get("data-id", "").strip():
        return parent["data-id"].strip().upper()
    return None


def has_category_keyword(text: str) -> bool:
    low = f" {text.lower()} "
    return any(kw in low for kw in CATEGORY_KEYWORDS)


# ============================================================
#                 LEAF-CARD DETECTION
# ============================================================

def get_leaf_cards(soup: BeautifulSoup) -> List[Tag]:
    leaves = []
    all_ids = soup.select("div[data-id]")
    logger.debug(f"Raw data-id elements found in DOM: {len(all_ids)}")
    for node in all_ids:
        if not node.find(attrs={"data-id": True}):
            leaves.append(node)

    if leaves:
        logger.debug(f"Identified {len(leaves)} leaf cards via strategy A (data-id)")
        return leaves

    logger.warning("No data-id leaves found! Using anchor-climb fallback.")
    anchors = soup.find_all("a", href=PATTERNS["product_href"])
    seen_pids, candidates = set(), []

    for a in anchors:
        href = a.get("href", "")
        pid = extract_pid_from_url(href)
        if not pid or pid == href or pid in seen_pids:
            continue
        seen_pids.add(pid)

        parent = a.parent
        for _ in range(6):
            if not parent:
                break
            txt = parent.get_text(" ")
            if "₹" in txt and has_category_keyword(txt):
                candidates.append(parent)
                break
            parent = parent.parent

    logger.debug(f"Identified {len(candidates)} leaf cards via strategy B (anchor-climb)")
    return candidates


def find_product_link(card: Tag) -> Optional[Tuple[str, str, Tag]]:
    """
    [C] Mobile cards wrap the ENTIRE card in ONE anchor (a.k7wcnx, no
    title attribute) — unlike trimmer/air-fryer's two-anchor split.
    Collecting all matching anchors and picking the "best" one still
    works fine since there's only one candidate here.
    """
    anchors = card.find_all("a", href=PATTERNS["product_href"])
    if anchors:
        best = max(
            anchors,
            key=lambda a: (bool(a.get("title")), len(a.get_text(strip=True)))
        )
        href = best.get("href", "")
        url = href if href.startswith("http") else f"https://www.flipkart.com{href}"
        return url, extract_pid_from_url(href), best

    for a in card.find_all("a", attrs={"target": "_blank"}):
        href = a.get("href", "")
        if "/p/" in href:
            url = href if href.startswith("http") else f"https://www.flipkart.com{href}"
            return url, extract_pid_from_url(href), a

    return None


# ============================================================
#                      FIELD EXTRACTORS
# ============================================================

def extract_title(card: Tag, link: Tag) -> Optional[str]:
    # 1. Anchor title attribute (usually absent on mobile's single-anchor card)
    if link is not None and hasattr(link, "get"):
        t = link.get("title")
        if t and len(t.strip()) > 10:
            return clean_text(t)

    # 1b. The anchor's own text — on mobiles this is the WHOLE card's
    # text (title + specs + price), so it correctly FAILS this check
    # ("₹" is present) and falls through to the div.RG5Slk selector below.
    if link is not None:
        t = clean_text(link.get_text(" "))
        if t and len(t) > 10 and "₹" not in t:
            return t

    # 2. Category-template class — [C] THIS is the primary path for mobiles
    div = card.select_one("div.RG5Slk")
    if div:
        t = clean_text(div.get_text(" "))
        if t and len(t) > 10:
            return t

    # 3. Structural heuristic fallback
    for col in card.find_all("div", class_=re.compile(r"col-\d+")):
        txt = clean_text(col.get_text(" "))
        if len(txt) > 30 and has_category_keyword(txt):
            return txt.split("  ")[0][:250]

    # 4. Last resort
    texts = [t for t in card.stripped_strings if len(t) > 20 and "http" not in t]
    if texts:
        return clean_text(PATTERNS["junk_prefix"].sub("", texts[0]))

    return None


def extract_rating_info(card: Tag) -> Dict[str, Any]:
    result = {"rating": None, "rating_count": None, "review_count": None}

    rating_span = card.select_one("span[id^='productRating_']") or card.select_one("div.MKiFS6")

    if rating_span:
        val = clean_text(rating_span.get_text(" "))
        m = PATTERNS["rating_strict"].match(val)
        if m:
            result["rating"] = m.group(1)
        else:
            m2 = re.search(r"\b([0-5]\.\d)\b", val)
            if m2:
                result["rating"] = m2.group(1)

    count_span = None
    if rating_span is not None:
        count_span = rating_span.find_next_sibling("span")
    if count_span is None:
        count_span = card.select_one("span.PvbNMB")

    if count_span is not None:
        txt = clean_text(count_span.get_text(" "))
        rc = PATTERNS["rating_count_worded"].search(txt)
        rv = PATTERNS["review_count"].search(txt)
        if rc:
            result["rating_count"] = extract_number(rc.group(1))
        if rv:
            result["review_count"] = extract_number(rv.group(1))

        if result["rating_count"] is None:
            pm = PATTERNS["rating_count_paren"].search(txt)
            if pm:
                result["rating_count"] = extract_number(pm.group(1))

    return result


def extract_price_block(card: Tag) -> Dict[str, Any]:
    result = {"price": None, "mrp": None, "discount_pct": None,
              "or_pay_price": None, "supercoins": None, "exchange_discount": None}

    full_card_text = clean_text(card.get_text(" "))

    price_container = (
        card.select_one("div[class*='QiMO5r']") or
        card.select_one("div[class*='oFEPlD']") or
        card
    )
    raw_text = clean_text(price_container.get_text(" ")) if price_container else ""
    if "₹" not in raw_text:
        raw_text = full_card_text

    cut_markers = [
        "Or Pay", "Off on Exchange", "Exchange",
        "Bank Offer", "No Cost EMI", "EMI starts", "Save extra with",
    ]
    for marker in cut_markers:
        idx = raw_text.find(marker)
        if idx != -1:
            upto_idx = raw_text.rfind("Upto", 0, idx)
            raw_text = raw_text[: (upto_idx if upto_idx != -1 else idx)]

    nums = [extract_number(p) for p in re.findall(r"₹\s?([\d,]+)", raw_text)]
    nums = [n for n in nums if n and n > 20]

    if nums:
        result["price"] = nums[0]
        if len(nums) > 1:
            if nums[1] >= nums[0]:
                result["mrp"] = nums[1]
            else:
                remaining = [n for n in nums[1:] if n >= nums[0]]
                if remaining:
                    result["mrp"] = max(remaining)

    dm = PATTERNS["discount"].search(raw_text)
    if dm:
        result["discount_pct"] = int(dm.group(1))

    if result["discount_pct"] and result["price"] and not result["mrp"]:
        try:
            result["mrp"] = int(result["price"] / (1 - result["discount_pct"] / 100))
        except ZeroDivisionError:
            pass

    om = PATTERNS["or_pay"].search(full_card_text)
    if om:
        result["or_pay_price"] = extract_number(om.group(1))
        result["supercoins"] = extract_number(om.group(2))

    ex = PATTERNS["exchange_discount"].search(full_card_text)
    if ex:
        result["exchange_discount"] = extract_number(ex.group(1))

    return result


def extract_specs(card: Tag) -> List[str]:
    """
    [D] div.U_GKRr (trimmer/air-fryer subtitle) does NOT exist on the
    mobile template — this selector correctly returns None and falls
    through. div.CMXw7N > ul.HwRTzP > li.DTBslk IS the mobile spec list
    (RAM/ROM, Display, Camera, Battery, Processor, Warranty) — and this
    is EXACTLY what the pre-existing `ul.HwRTzP` fallback already
    targets. ZERO changes needed to this function.
    """
    subtitle = card.select_one("div.U_GKRr")
    if subtitle:
        parts = [clean_text(p) for p in subtitle.get_text(" ").split(",")]
        parts = [p for p in parts if p]
        if parts:
            return parts

    ul = card.select_one("ul.HwRTzP")
    if ul:
        specs = [clean_text(li.get_text(" ")) for li in ul.find_all("li")]
        specs = [s for s in specs if s and len(s) > 2]
        if specs:
            return specs

    for ul in card.find_all("ul"):
        items = [clean_text(li.get_text(" ")) for li in ul.find_all("li")]
        items = [i for i in items if i and len(i) > 2]
        if len(items) >= 2:
            return items

    return []


def is_flipkart_assured(card: Tag) -> bool:
    """Confirmed unchanged for mobiles (still div.FIdGa1-style shield icon
    + fa_9e47c1.png)."""
    if card.select_one("div[class*='qYp2rh'] img") or card.select_one("div[class*='FIdGa1'] img"):
        return True
    for img in card.find_all("img"):
        src = (img.get("src") or "").lower()
        alt = (img.get("alt") or "").lower()
        if "/img/fa_" in src or "assured" in src or "assured" in alt:
            return True
    return "assured" in card.get_text(" ").lower()


def extract_badges(card: Tag) -> Dict[str, Any]:
    """
    [E] PRIMARY source for ribbon text is now the `data-ribbon-text`
    attribute injected by SourceFetcher.fetch() via page.evaluate()
    BEFORE the HTML snapshot is taken — because on this template the
    ribbon text ("Bestseller") is CSS ::before pseudo-element content,
    which is INVISIBLE to BeautifulSoup/page.content(). Falls back to
    real get_text() in case a future ribbon renders as real DOM text.

    SECONDARY: div.HZ0E6r pills (Bank Offer / Exchange offer chain) —
    same as trimmer template, confirmed unchanged via inspect.
    """
    text_low = card.get_text(" ").lower()
    badges = []

    assured = is_flipkart_assured(card)
    bank_offer = "bank offer" in text_low

    if assured:
        badges.append("Flipkart Assured")
    if bank_offer:
        badges.append("Bank Offer")

    # Primary: div.o2uEoz ribbon — read injected pseudo-element text first
    for ribbon in card.select("div.o2uEoz"):
        txt = clean_text(ribbon.get("data-ribbon-text", "") or ribbon.get_text(" "))
        if txt and txt not in badges:
            badges.append(txt)

    # Secondary: colored HZ0E6r pills — EXCLUDE exchange-offer fragments
    # that share the same class chain
    for pill in card.select("div.HZ0E6r"):
        txt = clean_text(pill.get_text(" "))
        low_txt = txt.lower()
        if (txt and 2 < len(txt) < 25 and "₹" not in txt
                and "exchange" not in low_txt and "upto" not in low_txt
                and txt not in badges):
            badges.append(txt)

    # Safety net: known ribbon labels via plain text match
    for kw in RIBBON_KEYWORDS:
        if kw in text_low:
            title_kw = kw.title()
            if title_kw not in badges:
                badges.append(title_kw)

    # Offer-type badges (EMI, exchange, delivery, etc.)
    for kw in ["no cost emi", "emi", "exchange", "free delivery", "special price",
               "big saving", "or pay"]:
        if kw in text_low:
            t = kw.title()
            if t not in badges:
                badges.append(t)

    return {
        "assured": "Yes" if assured else "No",
        "bank_offer": "Yes" if bank_offer else "No",
        "badges_str": ", ".join(dict.fromkeys(badges)),
    }


# ============================================================
#         [H] MOBILE-SPECIFIC SPEC-LINE PARSERS
# ============================================================

def parse_ram_rom(specs: List[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    combined = " | ".join(specs)
    m = PATTERNS["ram_rom"].search(combined)
    if not m:
        return None, None, None
    ram = f"{m.group(1)} {m.group(2).upper()}"
    rom = f"{m.group(3)} {m.group(4).upper()}"
    expandable = f"{m.group(5)} {m.group(6).upper()}" if m.group(5) else None
    return ram, rom, expandable


def parse_display(specs: List[str]) -> Tuple[Optional[str], Optional[str]]:
    for line in specs:
        m = PATTERNS["display"].search(line)
        if m:
            inch_val = m.group(2)
            dtype = clean_text(m.group(3))
            return inch_val, (dtype if dtype else None)
    return None, None


def parse_camera_specs(specs: List[str]) -> Tuple[Optional[str], Optional[str]]:
    """
    Handles BOTH observed formats:
      "50MP Rear Camera | 8MP Front Camera"   (explicitly labeled)
      "50MP + 2MP | 8MP Front Camera"         (unlabeled dual rear, before "|")
      "0.3MP Rear Camera"                     (rear only, no front line)
    """
    rear, front = None, None
    for line in specs:
        low = line.lower()
        if "camera" not in low and "mp" not in low:
            continue
        parts = [p.strip() for p in line.split("|")]
        for part in parts:
            plow = part.lower()
            if "front camera" in plow:
                front = clean_text(re.sub(r"front camera", "", part, flags=re.IGNORECASE))
            elif "rear camera" in plow:
                rear = clean_text(re.sub(r"rear camera", "", part, flags=re.IGNORECASE))
            elif "mp" in plow and rear is None and front is None:
                # Unlabeled first segment (e.g. "50MP + 2MP") — assume rear
                rear = clean_text(part)
        if rear or front:
            break
    return rear, front


def parse_battery(specs: List[str]) -> Optional[int]:
    for line in specs:
        if "mah" in line.lower():
            m = PATTERNS["battery_mah"].search(line)
            if m:
                return extract_number(m.group(1))
    return None


def parse_processor(specs: List[str]) -> Optional[str]:
    for line in specs:
        if "processor" in line.lower():
            m = PATTERNS["processor_line"].match(line.strip())
            if m:
                return clean_text(m.group(1))
            return clean_text(re.sub(r"processor", "", line, flags=re.IGNORECASE))
    return None


def parse_warranty(specs: List[str]) -> Optional[str]:
    for line in specs:
        if "warranty" in line.lower():
            return line
    return None


def derive_attributes(title: str, specs: List[str]) -> Dict[str, Any]:
    """[H] Mobile-specific: RAM/ROM/Expandable, Display Size/Type,
    Rear/Front Camera, Battery, Processor, Network Type, Color, Warranty
    (replaces trimmer's Trimmer-Type/Category/Runtime focus)."""
    out = {
        "brand": None, "ram": None, "rom": None, "expandable_storage": None,
        "display_size_inch": None, "display_type": None,
        "rear_camera": None, "front_camera": None,
        "battery_mah": None, "processor": None,
        "network_type": None, "color": None, "warranty": None,
    }
    spec_blob = " | ".join(specs)
    combined = f"{title} {spec_blob}"
    t_low = combined.lower()

    # Brand
    for b in KNOWN_BRANDS:
        if t_low.startswith(b.lower() + " ") or t_low == b.lower():
            out["brand"] = b
            break
    if not out["brand"]:
        for b in KNOWN_BRANDS:
            if re.search(rf"\b{re.escape(b.lower())}\b", t_low):
                out["brand"] = b
                break
    if not out["brand"]:
        first = title.strip().split()
        if first:
            out["brand"] = re.sub(r"[^A-Za-z0-9\-&+]", "", first[0]) or None

    # RAM / ROM / Expandable storage
    out["ram"], out["rom"], out["expandable_storage"] = parse_ram_rom(specs)

    # Display size / type
    out["display_size_inch"], out["display_type"] = parse_display(specs)

    # Rear / front camera
    out["rear_camera"], out["front_camera"] = parse_camera_specs(specs)

    # Battery
    out["battery_mah"] = parse_battery(specs)

    # Processor
    out["processor"] = parse_processor(specs)

    # Warranty (raw text)
    out["warranty"] = parse_warranty(specs)

    # Network type — priority 5G > 4G > 3G > 2G
    for kw in ["5g", "4g", "3g", "2g"]:
        if re.search(rf"\b{kw}\b", t_low):
            out["network_type"] = kw.upper()
            break

    # Color — from title's trailing parentheses, e.g. "(Blue, 128 GB)".
    # Filters out purely-numeric parenthetical content (storage-only,
    # no color) so it doesn't get misread as a color.
    m = PATTERNS["color_from_title"].search(title)
    if m:
        inner_parts = [clean_text(p) for p in m.group(1).split(",")]
        if inner_parts and inner_parts[0] and not re.search(r"\d", inner_parts[0]):
            out["color"] = inner_parts[0]

    return out


def validate_and_clean(title: Optional[str]) -> bool:
    if not title:
        return False
    t = title.strip().lower()
    if t in ("", "compare", "add to compare"):
        return False
    if t.startswith(("add to compare", "compare")):
        return False
    if len(title.strip()) < 5:
        return False
    if re.match(r"^[\d.,\s₹%]+$", title):
        return False
    return True


# ============================================================
#                      MAIN PARSER ENGINE
# ============================================================

def parse_single_card(card: Tag) -> Optional[Dict[str, Any]]:
    link_info = find_product_link(card)
    if not link_info:
        logger.debug("Skipping card: no product link found")
        return None

    url, url_pid, anchor_tag = link_info
    pid = pid_from_card(card) or url_pid

    raw_title = extract_title(card, anchor_tag)
    if not validate_and_clean(raw_title):
        logger.debug(f"Skipping card {pid}: invalid title '{raw_title}'")
        return None
    title = PATTERNS["junk_prefix"].sub("", raw_title).strip()

    rating_data = extract_rating_info(card)
    price_data = extract_price_block(card)
    specs = extract_specs(card)
    badge_data = extract_badges(card)
    derived = derive_attributes(title, specs)

    has_content = any([
        price_data["price"], rating_data["rating"],
        rating_data["rating_count"], specs,
    ])
    if not has_content:
        logger.warning(f"Card {pid} ('{title[:30]}...') had almost no data. Skipping.")
        return None

    return {
        "pid": pid, "brand": derived["brand"], "title": title, "url": url,
        "selling_price": price_data["price"], "mrp": price_data["mrp"],
        "discount_pct": price_data["discount_pct"],
        "or_pay_price": price_data["or_pay_price"], "supercoins": price_data["supercoins"],
        "exchange_discount": price_data["exchange_discount"],
        "avg_rating": rating_data["rating"], "total_ratings": rating_data["rating_count"],
        "total_reviews": rating_data["review_count"],
        "ram": derived["ram"], "rom": derived["rom"],
        "expandable_storage": derived["expandable_storage"],
        "display_size_inch": derived["display_size_inch"], "display_type": derived["display_type"],
        "rear_camera": derived["rear_camera"], "front_camera": derived["front_camera"],
        "battery_mah": derived["battery_mah"], "processor": derived["processor"],
        "network_type": derived["network_type"], "color": derived["color"],
        "warranty": derived["warranty"],
        "key_features": " | ".join(specs) if specs else "",
        "assured": badge_data["assured"], "bank_offer": badge_data["bank_offer"],
        "badges_offers": badge_data["badges_str"],
        "scraped_at": datetime.now().isoformat(timespec="seconds"),
    }


def parse_page(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "lxml")
    cards = get_leaf_cards(soup)

    if 0 < len(cards) < PRODUCTS_PER_PAGE_DEFAULT * MIN_ACCEPTABLE_PRODUCTS_RATIO:
        logger.debug(f"⚠️ Only {len(cards)} raw cards in DOM (expected ~{PRODUCTS_PER_PAGE_DEFAULT}).")

    results, pids_seen = [], set()
    for idx, card in enumerate(cards, 1):
        try:
            parsed = parse_single_card(card)
            if parsed and parsed["pid"] not in pids_seen:
                results.append(parsed)
                pids_seen.add(parsed["pid"])
        except Exception as e:
            logger.error(f"Error parsing card #{idx}: {e}", exc_info=True)
            continue

    logger.info(f"Successfully parsed {len(results)} products (raw cards: {len(cards)}).")
    return results


def is_partial_page(products: List[Dict], expected: int = PRODUCTS_PER_PAGE_DEFAULT) -> bool:
    if not products:
        return True
    return len(products) < (expected * MIN_ACCEPTABLE_PRODUCTS_RATIO)


# ============================================================
#     PLAYWRIGHT-ONLY NETWORK LAYER (no requests at all)
# ============================================================

_playwright_ctx = None
_browser_instance = None

# [E] JS injected before every page.content() snapshot — reads the
# CSS ::before pseudo-element content off ribbon divs (e.g. "Bestseller")
# and writes it into a real DOM attribute so BeautifulSoup can see it.
RIBBON_PSEUDO_ELEMENT_FIX_JS = """
() => {
    document.querySelectorAll('div.o2uEoz').forEach(el => {
        try {
            const style = window.getComputedStyle(el, '::before');
            let content = style && style.content;
            if (content && content !== 'none' && content !== 'normal') {
                content = content.replace(/^["']|["']$/g, '');
                if (content) el.setAttribute('data-ribbon-text', content);
            }
        } catch (e) {}
    });
}
"""


def get_or_create_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        return _browser_instance

    from playwright.sync_api import sync_playwright
    _playwright_ctx = sync_playwright().start()
    _browser_instance = _playwright_ctx.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
            "--disable-extensions", "--disable-background-networking",
            "--disable-default-apps", "--disable-sync",
            "--disable-blink-features=AutomationControlled",
            "--js-flags=--max-old-space-size=256",
        ]
    )
    logger.info("Playwright browser launched (reused across all sources/pages).")
    return _browser_instance


def close_shared_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        try:
            _browser_instance.close()
            logger.info("Playwright browser closed.")
        except Exception:
            pass
        _browser_instance = None
    if _playwright_ctx is not None:
        try:
            _playwright_ctx.stop()
        except Exception:
            pass
        _playwright_ctx = None


class SourceFetcher:
    """
    Manages ONE persistent browser tab (context+page) reused across ALL
    page fetches within a single source. Recycled periodically, and
    force-recycled if a fetch comes back with zero data-id cards.
    """
    def __init__(self):
        self.context = None
        self.page = None
        self.pages_served = 0

    def _open(self):
        browser = get_or_create_browser()
        self.context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1280, "height": 900},
            locale="en-IN",
        )
        self.page = self.context.new_page()

    def _close(self):
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        self.context = None
        self.page = None

    def _recycle_if_due(self):
        if self.pages_served > 0 and self.pages_served % RECYCLE_CONTEXT_EVERY_N_PAGES == 0:
            logger.info(f"  [tab] Recycling Playwright context after {self.pages_served} pages...")
            self._close()

    def fetch(self, url: str, target_card_count: int = PRODUCTS_PER_PAGE_DEFAULT) -> Optional[str]:
        if self.page is None:
            self._open()

        for attempt in range(MAX_RETRIES_PER_URL):
            try:
                self.page.goto(url, wait_until="domcontentloaded", timeout=30000)
                self.page.wait_for_timeout(1200)

                previous_count = 0
                for _ in range(14):
                    self.page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
                    self.page.wait_for_timeout(650)
                    current_count = self.page.evaluate(
                        "document.querySelectorAll('div[data-id]').length"
                    )
                    if current_count >= target_card_count:
                        break
                    if current_count == previous_count:
                        self.page.wait_for_timeout(900)
                    previous_count = current_count

                self.page.evaluate("window.scrollTo(0, 0)")
                self.page.wait_for_timeout(300)

                # [E] Ribbon pseudo-element fix — MUST run before page.content()
                try:
                    self.page.evaluate(RIBBON_PSEUDO_ELEMENT_FIX_JS)
                except Exception as ribbon_err:
                    logger.debug(f"  [tab] Ribbon pseudo-element fix failed (non-fatal): {ribbon_err}")

                html = self.page.content()
                self.pages_served += 1

                if "data-id=" in html:
                    self._recycle_if_due()
                    return html

                logger.warning(f"  [tab] Attempt {attempt+1}: zero data-id cards in HTML. "
                               f"Recycling context and retrying...")
                self._close()
                self._open()

            except Exception as e:
                logger.error(f"  [tab] Playwright error (attempt {attempt+1}): {e}. "
                             f"Recycling context and retrying...")
                self._close()
                self._open()
                time.sleep(random.uniform(2, 4))

        return None

    def close(self):
        self._close()


def fetch_and_parse_page(fetcher: SourceFetcher, url: str, page_label: str = "") -> List[Dict[str, Any]]:
    html = fetcher.fetch(url)
    products = parse_page(html) if html else []

    if is_partial_page(products):
        logger.warning(f"{page_label}: only {len(products)} products "
                       f"(expected ~{PRODUCTS_PER_PAGE_DEFAULT}) — forcing a fresh-context retry...")
        fetcher.close()
        html2 = fetcher.fetch(url)
        products2 = parse_page(html2) if html2 else []
        if len(products2) > len(products):
            logger.info(f"{page_label}: retry recovered {len(products2)} (vs {len(products)}).")
            products = products2

    return products


# ============================================================
#                  PAGINATION HELPERS
# ============================================================

def determine_total_pages(first_page_html: str, max_pages_cap: int) -> int:
    m = PATTERNS["total_products"].search(first_page_html)
    if m:
        total_prods = extract_number(m.group(1))
        if total_prods:
            calc_pages = math.ceil(total_prods / PRODUCTS_PER_PAGE_DEFAULT)
            final = min(calc_pages, max_pages_cap)
            logger.info(f"Detected {total_prods:,} total results ≈ {calc_pages} pages. "
                        f"Capping at {final}")
            return final
    logger.warning("Could not detect total results. Defaulting to max_pages_cap.")
    return max_pages_cap


def build_url(base_url: str, page_num: int) -> str:
    if page_num <= 1:
        return base_url
    return f"{base_url}&page={page_num}"


# ============================================================
#                      EXCEL WRITER
# ============================================================

COLUMN_HEADERS = [
    "Source", "Page No.", "PID", "Brand", "Product Name", "Product URL",
    "Selling Price (₹)", "MRP (₹)", "Discount (%)",
    "Or Pay Price (₹)", "SuperCoins", "Exchange Discount (₹)",
    "Avg Rating", "Total Ratings", "Total Reviews",
    "RAM", "ROM/Storage", "Expandable Storage",
    "Display Size (inch)", "Display Type",
    "Rear Camera", "Front Camera", "Battery (mAh)", "Processor",
    "Network Type", "Color", "Warranty",
    "Key Features", "Flipkart Assured", "Bank Offer",
    "Badges/Offers", "Scraped At",
]

COLUMN_WIDTHS = [16, 8, 20, 14, 58, 55, 14, 12, 11, 15, 12, 16, 10, 13, 13,
                 10, 12, 14, 14, 22, 16, 14, 12, 18, 12, 14, 40, 45, 16, 12, 32, 20]


def setup_excel(path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mobiles"

    header_fill = PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    ws.append(COLUMN_HEADERS)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, w in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    return wb, ws


def append_rows(ws, products: List[Dict], page_num: int, source_label: str):
    for p in products:
        ws.append([
            source_label, page_num,
            p["pid"], p["brand"], p["title"], p["url"],
            p["selling_price"], p["mrp"], p["discount_pct"],
            p["or_pay_price"], p["supercoins"], p["exchange_discount"],
            p["avg_rating"], p["total_ratings"], p["total_reviews"],
            p["ram"], p["rom"], p["expandable_storage"],
            p["display_size_inch"], p["display_type"],
            p["rear_camera"], p["front_camera"], p["battery_mah"], p["processor"],
            p["network_type"], p["color"], p["warranty"],
            p["key_features"], p["assured"], p["bank_offer"],
            p["badges_offers"], p["scraped_at"],
        ])


# ============================================================
#                  PER-SOURCE PROCESSING
# ============================================================

def process_source(wb, ws, source_label: str, source_url: str,
                    global_pids: set) -> Tuple[int, int, int]:
    logger.info("\n" + "=" * 70)
    logger.info(f"STARTING SOURCE: '{source_label}'")
    logger.info(f"URL: {source_url}")
    logger.info("=" * 70)

    fetcher = SourceFetcher()

    first_url = build_url(source_url, 1)
    products_batch_1 = fetch_and_parse_page(fetcher, first_url, page_label=f"[{source_label}] Page 1")

    first_html_raw = fetcher.page.content() if fetcher.page else ""

    if not products_batch_1 and not first_html_raw:
        logger.error(f"[{source_label}] Cannot reach page 1 at all. Abandoning this source.")
        fetcher.close()
        return 0, 0, 0

    total_pages = determine_total_pages(first_html_raw or "", MAX_PAGES)

    added_count = 0
    dupes_skipped = 0
    consecutive_failures = 0
    pages_attempted = 1

    if not products_batch_1:
        consecutive_failures += 1
        logger.error(f"[{source_label}] Page 1: COMPLETE FAILURE "
                     f"({consecutive_failures}/{CONSECUTIVE_FAILURE_LIMIT}).")
        ws.append([source_label, 1, "[ERROR]", "", "Failed to fetch page 1", first_url] +
                  [""] * (len(COLUMN_HEADERS) - 7) + [datetime.now().isoformat(timespec="seconds")])
    else:
        if GLOBAL_DEDUPE:
            fresh = [p for p in products_batch_1 if p["pid"] not in global_pids]
            dupes_skipped += len(products_batch_1) - len(fresh)
            global_pids.update(p["pid"] for p in fresh)
            products_batch_1 = fresh
        append_rows(ws, products_batch_1, 1, source_label)
        added_count += len(products_batch_1)
        logger.info(f"[{source_label}] Page 1: +{len(products_batch_1)} unique added.")

    try:
        for pg in range(2, total_pages + 1):
            if consecutive_failures >= CONSECUTIVE_FAILURE_LIMIT:
                logger.warning(
                    f"[{source_label}] {CONSECUTIVE_FAILURE_LIMIT} CONSECUTIVE failed pages "
                    f"hit at page {pg}. Abandoning this source and moving to the next one."
                )
                break

            url = build_url(source_url, pg)
            delay = random.uniform(PAGE_DELAY_MIN, PAGE_DELAY_MAX)
            logger.info(f"\n--- [{source_label}] Page {pg}/{total_pages} "
                        f"({delay:.1f}s delay | consec.fail {consecutive_failures}/"
                        f"{CONSECUTIVE_FAILURE_LIMIT}) ---")
            time.sleep(delay)
            pages_attempted += 1

            products = fetch_and_parse_page(fetcher, url, page_label=f"[{source_label}] Page {pg}")

            if not products:
                consecutive_failures += 1
                logger.error(f"[{source_label}] Page {pg}: COMPLETE FAILURE "
                             f"({consecutive_failures}/{CONSECUTIVE_FAILURE_LIMIT}).")
                ws.append([source_label, pg, "[ERROR]", "", f"Failed to fetch page {pg}", url] +
                          [""] * (len(COLUMN_HEADERS) - 7) + [datetime.now().isoformat(timespec="seconds")])
                continue

            consecutive_failures = 0

            if GLOBAL_DEDUPE:
                fresh = [p for p in products if p["pid"] not in global_pids]
                dupes_skipped += len(products) - len(fresh)
                global_pids.update(p["pid"] for p in fresh)
                products = fresh

            append_rows(ws, products, pg, source_label)
            added_count += len(products)
            logger.info(f"[{source_label}] Page {pg}: +{len(products)} unique added. "
                        f"Source running total: {added_count} (dupes skipped: {dupes_skipped})")

            if pg % SAVE_EVERY_N_PAGES == 0:
                wb.save(OUTPUT_FILE)
                logger.info(f"⚡️ Progress saved: {OUTPUT_FILE}")

    except KeyboardInterrupt:
        logger.warning(f"\n[{source_label}] Interrupted mid-source.")
        wb.save(OUTPUT_FILE)
        fetcher.close()
        raise

    fetcher.close()
    logger.info(f"\n[{source_label}] SOURCE COMPLETE: {added_count} unique products added, "
                f"{dupes_skipped} duplicates skipped, {pages_attempted} pages attempted.")
    return added_count, dupes_skipped, pages_attempted


# ============================================================
#                         MAIN
# ============================================================

def run_diagnostic_mode():
    src = BASE_URLS[DIAGNOSTIC_SOURCE_INDEX]
    logger.info("=" * 70)
    logger.info(f"DIAGNOSTIC MODE: Source '{src['label']}' — Page 1 only, no file written")
    logger.info("=" * 70)

    fetcher = SourceFetcher()
    products = fetch_and_parse_page(fetcher, build_url(src["url"], 1), page_label="Page 1 (diagnostic)")
    fetcher.close()

    if not products:
        logger.error("CRITICAL: Could not extract any products from page 1.")
        close_shared_browser()
        return

    logger.info(f"\nExtracted {len(products)} products (expected ~{PRODUCTS_PER_PAGE_DEFAULT}):\n")
    for i, p in enumerate(products, 1):
        print(f"{i:3d} | {str(p['pid']):>18} | {str(p['brand'] or '-'):<12} | "
              f"{str(p['avg_rating'] or '-'):>4}★ | ₹{str(p['selling_price'] or '-'):>7} "
              f"(MRP ₹{str(p['mrp'] or '-'):>7}, {str(p['discount_pct'] or '-')}%) | "
              f"{str(p['ram'] or '-'):>7} | {str(p['rom'] or '-'):>8} | "
              f"{str(p['battery_mah'] or '-'):>5}mAh | {str(p['network_type'] or '-'):<3} | "
              f"Assured:{p['assured']} | Badges:{p['badges_offers'][:30]}")
        print(f"      {p['title'][:95]}")
        if p["key_features"]:
            print(f"      Specs: {p['key_features'][:110]}")
        if p["rear_camera"] or p["front_camera"]:
            print(f"      Camera: Rear={p['rear_camera']} Front={p['front_camera']}")
        if p["exchange_discount"]:
            print(f"      Exchange Discount: ₹{p['exchange_discount']}")
        print("-" * 118)

    n = len(products)
    def pct(k): return f"{sum(1 for p in products if p[k] not in (None, '', 'No')) / n * 100:5.1f}%"
    print("\nFIELD FILL-RATE HEALTH CHECK")
    for key in ["brand", "selling_price", "mrp", "discount_pct", "avg_rating",
                "total_ratings", "total_reviews", "ram", "rom", "expandable_storage",
                "display_size_inch", "display_type", "rear_camera", "front_camera",
                "battery_mah", "processor", "network_type", "color", "warranty",
                "key_features", "badges_offers", "exchange_discount"]:
        print(f"   {key:<20} {pct(key)}")
    print(f"   {'assured=Yes':<20} {sum(1 for p in products if p['assured']=='Yes')/n*100:5.1f}%")

    close_shared_browser()


def main():
    if DIAGNOSTIC_MODE:
        run_diagnostic_mode()
        return

    logger.info("Starting Flipkart MOBILE Scraper (v1.0 — Playwright-only)")
    logger.info(f"Sources to run: {[s['label'] for s in BASE_URLS]}")
    logger.info(f"Output: {OUTPUT_FILE}")

    wb, ws = setup_excel(OUTPUT_FILE)
    global_pids = set()
    grand_total_added = 0
    grand_total_dupes = 0

    try:
        for source in BASE_URLS:
            added, dupes, pages = process_source(
                wb, ws, source["label"], source["url"], global_pids
            )
            grand_total_added += added
            grand_total_dupes += dupes
            wb.save(OUTPUT_FILE)

    except KeyboardInterrupt:
        logger.warning("\nInterrupted by user — saving progress before exit...")

    finally:
        wb.save(OUTPUT_FILE)
        close_shared_browser()
        logger.info("\n" + "=" * 70)
        logger.info(f"ALL SOURCES DONE. Grand total unique products: {grand_total_added}")
        logger.info(f"Grand total duplicates skipped: {grand_total_dupes}")
        logger.info(f"File saved: {OUTPUT_FILE}")
        logger.info("=" * 70)


if __name__ == "__main__":
    main()