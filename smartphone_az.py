"""
Amazon Mobiles Assortment Scraper — Robust v4.0
==========================================================
v3.0 recap (all Amazon-specific extraction logic — UNCHANGED in v4.0):
  • 3-tier badge/ribbon extraction (rio-badge-label → status-badge
    JSON props → whitelist regex scan of card text)
  • 3-tier ASIN resolution (data-asin → badge JSON props → /dp/ASIN)
  • Price fallback: precise selectors first, then offer-text-stripped
    regex scan, read in DOM order
  • Garbage-card filter (kills "Sponsored" video-ad ghost cards)
  • is_captcha_page() — detects Amazon's block interstitial, retries
  • extract_mobile_attributes(): RAM, storage, battery, charging watts,
    refresh rate, network, display inch, processor, color, tech tags
  • Live-DOM-count scroll polling (no fixed "scroll N times" guessing)

v4.0 CHANGES — Multi-Source Architecture (adopted from the Flipkart
scraper line), scaffolding-only. ZERO changes to any parsing/extraction
function — every selector, regex, and heuristic below is byte-identical
to v3.0.

  [A] BASE_URL + SECOND_BASE_URL (two hardcoded sources, one bespoke
      "supplemental pass" function) → BASE_URLS = [{"label","url"}, ...]
      — an open-ended list processed by ONE process_source() loop.
      Adding a 3rd/10th/20th source (brand search, another category
      node, a different sort order) is now just appending a dict —
      no new function needed, matching the Flipkart BASE_URLS pattern.

  [B] build_page_url() + build_second_page_url() → single build_url()
      — confirmed both source types paginate identically via "&page=N"
      with page 1 omitting the param, so one function covers both.

  [C] fetch_html_playwright() (standalone, one shared context for the
      ENTIRE run) → SourceFetcher class: persistent browser CONTEXT
      per source (opened at source start, closed at source end),
      periodic recycle (RECYCLE_CONTEXT_EVERY_N_PAGES) + recycle-on-
      captcha/error — same lifecycle pattern as the Flipkart scraper's
      SourceFetcher, adapted to Amazon's scroll-poll + captcha-retry
      internals (which are preserved verbatim inside the class method).

  [D] NEW: fetch_and_parse_page() wrapper — forces ONE fresh-context
      retry if a page returns under MIN_ACCEPTABLE_PRODUCTS_RATIO of
      expected cards, even when it's NOT a captcha (v3.0 only retried
      on captcha/exception, never on a silently-thin non-captcha page).
      Directly ported from the Flipkart scraper's robustness layer.

  [E] STOP_AFTER_EMPTY (single global counter across the whole v3.0
      run) → CONSECUTIVE_FAILURE_LIMIT, now scoped PER SOURCE inside
      process_source() — one dead/blocked source no longer kills
      sources queued after it.

  [F] Two Excel sheets doing the same job (PageSummary +
      PopularityPageSummary, one per hardcoded pass) → ONE PageSummary
      sheet with a Source column, ONE Products sheet with a Source
      column. Resumability (get_scraped_pages) and dedupe
      (get_existing_asins) now filter/aggregate by source label,
      generalized to N sources instead of 2.

  [G] Diagnostic mode: DIAGNOSTIC_SOURCE_INDEX lets you point at ANY
      source in BASE_URLS, not just the hardcoded keyword search.

  [H] OUTPUT_FILE renamed (schema changed — new Source column at
      position 1 — an old v3.0 workbook would load with mismatched
      headers if resumed against it).

SETUP:
    pip install beautifulsoup4 openpyxl playwright lxml
    playwright install chromium

USAGE:
    python3 -u amazon_mobiles_category_v4.py
"""

import re
import time
import random
import logging
import json
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from bs4 import BeautifulSoup, Tag
from playwright.sync_api import sync_playwright

# ============================================================
#                  ⚙️ CONFIGURATION
# ============================================================

# [A] Open-ended source list. Add more dicts here (brand-specific
# keyword searches, other category nodes, different sort orders, etc.)
# and process_source() will run each one with full resumability +
# cross-source ASIN dedupe — no new code needed.
BASE_URLS = [
    {
        "label": "Keyword-Mobiles",
        "url": (
            "https://www.amazon.in/s?k=mobiles"
            "&i=electronics"
            "&rh=n%3A1389401031"
        ),
    },
    {
        "label": "Category-PopularityRank",
        # Amazon's dedicated "Smartphones & Basic Mobiles" category
        # node, sorted explicitly by popularity-rank. Surfaces a
        # different (and ribbon-tag-dense — Bestseller/Amazon's
        # Choice/Trending) product mix than the keyword search.
        "url": (
            "https://www.amazon.in/s?i=electronics"
            "&rh=n%3A1389432031"
            "&s=popularity-rank"
            "&fs=true"
        ),
    },
]

# OUTPUT_FILE = "8th_az_mobiles_bestselling_assortment_v4.xlsx"
OUTPUT_FILE = f"amazon_mobile_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


END_PAGE = 400                         # hard cap regardless of detected total, PER SOURCE
PRODUCTS_PER_PAGE_DEFAULT = 24         # Amazon India search grid default
MIN_ACCEPTABLE_PRODUCTS_RATIO = 0.70   # below this → treat as partial render

CONSECUTIVE_FAILURE_LIMIT = 3          # [E] was STOP_AFTER_EMPTY — now per-source
RECYCLE_CONTEXT_EVERY_N_PAGES = 30     # [C] periodic tab/context refresh per source

DELAY_RANGE = (3.0, 6.0)
SAVE_EVERY_N_PAGES = 1
MAX_RETRIES = 3
NAV_TIMEOUT_MS = 30000

DIAGNOSTIC_MODE = False       # True = fetch+parse page 1 of one source only, print, exit
DIAGNOSTIC_SOURCE_INDEX = 0   # [G] which BASE_URLS entry diagnostic mode targets

LOG_LEVEL = logging.INFO

# ============================================================
#                    LOGGING SETUP
# ============================================================
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# ---- Precompiled regex patterns (class-name independent) — UNCHANGED ----
PATTERNS = {
    "discount_pct": re.compile(r"\(\s*(\d+)\s*%\s*off\s*\)", re.IGNORECASE),
    "total_results": re.compile(r"of\s+(?:over\s+)?([\d,]+)\s+results", re.IGNORECASE),
    "rating_from_alt": re.compile(r"([\d.]+)\s*out of\s*5", re.IGNORECASE),
    "asin_from_url": re.compile(r"/dp/([A-Z0-9]{10})"),
    "junk_title_prefix": re.compile(r"^\s*(sponsored)\s*", re.IGNORECASE),
    "price_rupee": re.compile(r"₹\s?([\d,]+)"),

    # ---- Mobile-specific title parsers ----
    "mobile_ram_storage_combo": re.compile(r"(\d+)\s*GB\s*\+\s*(\d+)\s*GB", re.IGNORECASE),
    "mobile_ram": re.compile(r"(\d+)\s*GB\s*RAM", re.IGNORECASE),
    "mobile_storage": re.compile(r"(\d+)\s*(GB|TB)\s*(?:Storage|ROM|Internal)", re.IGNORECASE),
    "mobile_battery": re.compile(r"(\d+)\s*mAh", re.IGNORECASE),
    "mobile_charging": re.compile(r"(\d+)\s*W\b(?:\s*(?:Fast\s*)?Charg\w*)?", re.IGNORECASE),
    "mobile_refresh_rate": re.compile(r"(\d+)\s*Hz", re.IGNORECASE),
    "mobile_network": re.compile(r"\b(5G|4G VoLTE|4G|3G)\b", re.IGNORECASE),
    "mobile_display_inch": re.compile(r"([\d.]+)\s*(?:inch(?:es)?|\")", re.IGNORECASE),
    "mobile_processor": re.compile(
        r"(Snapdragon\s*[\w\d\+]*|MediaTek\s*Dimensity\s*[\w\d]*|Dimensity\s*[\w\d]*|"
        r"Exynos\s*[\w\d]*|A\d{2}\s*Bionic|Tensor\s*G\d?|Helio\s*[\w\d]*|Bionic)",
        re.IGNORECASE,
    ),
    "mobile_first_parens": re.compile(r"\(([^)]+)\)"),
}

# Known merchandising ribbon labels — regex-scan fallback (Strategy 3)
KNOWN_SPECIAL_TAGS = [
    "Best seller", "Amazon's Choice", "Overall Pick", "Limited time deal",
    "Deal of the Day", "Lightning Deal", "Today's Deal",
    "Climate Pledge Friendly", "Small Business",
    "Trending", "Popular", "Most Wished For", "New Arrival",
]

# Technology keywords scanned out of titles into a single "Technology Tags" column
MOBILE_TECH_KEYWORDS = [
    "5G", "4G VoLTE", "AMOLED", "Super AMOLED", "OLED", "IPS LCD",
    "NFC", "Dual Sim", "Fast Charging", "Water Resistant", "IP68", "IP67",
    "Fingerprint Sensor", "Face Unlock", "In-Display Fingerprint",
    "Wireless Charging", "Stereo Speakers", "Corning Gorilla Glass",
]

CUT_MARKERS = [
    "with coupon", "Off on Select Bank Cards", "back with Amazon Pay",
    "M.R.P", "Flat INR", "Up to", "% back",
]


# ============================================================
#         ROBUST CARD PARSER (CSS-INDEPENDENT) — UNCHANGED
# ============================================================

def clean_text(text: Optional[str]) -> str:
    if not text:
        return ""
    return " ".join(text.split()).strip()


def extract_number(raw: Optional[str]) -> Optional[int]:
    if raw is None:
        return None
    s = str(raw).replace(",", "").replace("₹", "").replace("(", "").replace(")", "").strip()
    m = re.match(r"([\d.]+)\s*[kK]\+?", s)
    if m:
        try:
            return int(float(m.group(1)) * 1000)
        except ValueError:
            return None
    try:
        return int(float(s))
    except ValueError:
        return None


def extract_asin_from_url(url: Optional[str]) -> Optional[str]:
    if not url:
        return None
    m = PATTERNS["asin_from_url"].search(url)
    return m.group(1) if m else None


def get_leaf_cards(soup: BeautifulSoup) -> List[Tag]:
    """
    Strategy A (primary): standard search-result card container.
    Strategy B (fallback): anchor-climb — find /dp/ links, climb parents
    until a container with both ₹ and mobile keywords is found.
    """
    cards = soup.select("div.s-result-item[data-asin][data-component-type='s-search-result']")
    if cards:
        return cards

    logger.warning("Standard card selector found 0 results — using anchor-climb fallback (Strategy B).")
    mobile_keywords = ["ram", "rom", "storage", "battery", "mah", "display",
                       "smartphone", "camera", "processor"]
    anchors = soup.find_all("a", href=re.compile(r"/dp/[A-Z0-9]{10}"))
    seen_asins = set()
    candidates = []

    for a in anchors:
        asin = extract_asin_from_url(a.get("href", ""))
        if not asin or asin in seen_asins:
            continue
        seen_asins.add(asin)

        parent = a.parent
        for _ in range(6):
            if not parent:
                break
            txt = parent.get_text(" ").lower()
            if "₹" in txt and any(k in txt for k in mobile_keywords):
                candidates.append(parent)
                break
            parent = parent.parent

    logger.debug(f"Strategy B recovered {len(candidates)} candidate cards.")
    return candidates


def extract_badge_info(card: Tag) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    """3-tier ribbon badge extraction — UNCHANGED from v3.0."""
    ribbon_badge = None
    ribbon_supplementary = None
    badge_asin = None
    badge_type = None

    rio_main = card.select_one("span.rio-badge-label span[aria-label]")
    if rio_main:
        ribbon_badge = (rio_main.get("aria-label") or rio_main.get_text(strip=True)).strip()

    rio_sub = card.select_one("span.rio-badge-supplementary-text[aria-label]")
    if rio_sub:
        ribbon_supplementary = rio_sub.get("aria-label", "").strip()

    status_badge = card.select_one('span[data-component-type="s-status-badge-component"]')
    if status_badge:
        if not ribbon_badge:
            label_span = status_badge.select_one("span.a-badge-label")
            if label_span:
                pieces = [t.get_text(strip=True) for t in label_span.select("span.a-badge-text")]
                ribbon_badge = " ".join(pieces).strip() if pieces else label_span.get_text(" ", strip=True)

        props_raw = status_badge.get("data-component-props")
        if props_raw:
            try:
                props = json.loads(props_raw)
                badge_asin = props.get("asin")
                badge_type = props.get("badgeType")
            except (json.JSONDecodeError, TypeError):
                pass

    if not ribbon_badge:
        combined = card.get_text(" ", strip=True)
        for tag in KNOWN_SPECIAL_TAGS:
            if tag.lower() in combined.lower():
                ribbon_badge = tag
                break

    return ribbon_badge, ribbon_supplementary, badge_asin, badge_type


def extract_title_and_link(card: Tag) -> Tuple[Optional[str], Optional[str]]:
    """Cascading title + URL extraction. UNCHANGED from v3.0."""
    title_recipe = card.select_one('div[data-cy="title-recipe"]')
    h2_tag = (title_recipe.select_one("h2") if title_recipe else None) or card.select_one("h2")
    link_tag = (
        (title_recipe.select_one("a[href]") if title_recipe else None)
        or card.select_one("h2 a[href]")
        or card.select_one("a.a-link-normal[href]")
    )

    title = None
    if h2_tag:
        title = h2_tag.get("aria-label", "").strip() or h2_tag.get_text(strip=True)

    href = link_tag.get("href") if link_tag else None
    full_url = None
    if href:
        full_url = "https://www.amazon.in" + href.split("?")[0] if href.startswith("/") else href

    return title, full_url


def validate_and_clean(title: Optional[str], url: Optional[str]) -> Optional[str]:
    """Garbage-card filter — UNCHANGED from v3.0."""
    if not title:
        return None

    t = PATTERNS["junk_title_prefix"].sub("", title).strip()

    if not t or t.lower() == "sponsored":
        return None
    if len(t) < 5:
        return None
    if re.match(r"^[\d.,\s]+$", t):
        return None
    if url and url.startswith("javascript:"):
        return None
    if not url:
        return None

    return t


def extract_price_block(card: Tag) -> Dict[str, Any]:
    """Price extraction — UNCHANGED from v3.0."""
    result = {"price": None, "mrp": None, "discount_pct": None,
              "offer_text": None, "coupon_price": None}

    price_block = card.select_one('div[data-cy="price-recipe"]')
    scope = price_block or card

    price_tag = scope.select_one("span.a-price:not(.a-text-price) > span.a-offscreen")
    if price_tag:
        result["price"] = extract_number(price_tag.get_text(strip=True))

    mrp_tag = scope.select_one("span.a-price.a-text-price > span.a-offscreen")
    if mrp_tag:
        result["mrp"] = extract_number(mrp_tag.get_text(strip=True))

    if result["price"] is None:
        raw_text = scope.get_text(" ", strip=True)
        for marker in CUT_MARKERS:
            idx = raw_text.find(marker)
            if idx != -1:
                raw_text = raw_text[:idx]
        nums = [extract_number(m) for m in PATTERNS["price_rupee"].findall(raw_text)]
        nums = [n for n in nums if n and n > 100]
        if nums:
            result["price"] = nums[0]
            if len(nums) > 1 and result["mrp"] is None:
                result["mrp"] = max(nums[1:])

    disc_tag = scope.find(string=PATTERNS["discount_pct"])
    if disc_tag:
        m = PATTERNS["discount_pct"].search(disc_tag)
        if m:
            result["discount_pct"] = int(m.group(1))
    elif result["price"] and result["mrp"] and result["mrp"] > result["price"]:
        result["discount_pct"] = round((result["mrp"] - result["price"]) * 100 / result["mrp"])

    if price_block:
        offer_tag = price_block.select_one("span.a-truncate-full")
        if offer_tag:
            result["offer_text"] = offer_tag.get_text(strip=True)

    for span in card.find_all("span"):
        t = span.get_text(strip=True)
        if t.lower().startswith("you pay"):
            result["coupon_price"] = t
            break

    return result


def extract_rating_info(card: Tag) -> Dict[str, Any]:
    """UNCHANGED from v3.0."""
    result = {"rating": None, "rating_count": None}
    reviews_block = card.select_one('div[data-cy="reviews-block"]')
    scope = reviews_block or card

    rating_tag = scope.select_one("span.a-icon-alt")
    if rating_tag:
        m = PATTERNS["rating_from_alt"].search(rating_tag.get_text(strip=True))
        if m:
            result["rating"] = m.group(1)

    rc_tag = (
        scope.select_one("a[aria-label*='ratings'] span.a-size-mini")
        or scope.select_one("span.a-size-base.s-underline-text")
        or card.select_one("a[href*='#customerReviews'] span[aria-hidden='true']")
    )
    if rc_tag:
        result["rating_count"] = extract_number(rc_tag.get_text(strip=True))

    return result


def extract_bought_delivery_service(card: Tag) -> Dict[str, Any]:
    """UNCHANGED from v3.0."""
    result = {"bought_info": None, "delivery_info": None, "service": None}

    reviews_block = card.select_one('div[data-cy="reviews-block"]')
    scope = reviews_block or card
    for span in scope.find_all("span"):
        t = span.get_text(strip=True)
        if "bought in past month" in t.lower():
            result["bought_info"] = t
            break

    delivery_block = card.select_one('div[data-cy="delivery-recipe"]')
    if delivery_block:
        result["delivery_info"] = delivery_block.get_text(" ", strip=True)

    for div in card.select("div.a-row"):
        t = div.get_text(strip=True)
        if t.startswith("Service:"):
            result["service"] = t
            break

    return result


def extract_mobile_attributes(title: Optional[str]) -> Dict[str, Any]:
    """
    UNCHANGED from v3.0. Amazon search cards do NOT render a bullet
    spec list (unlike Flipkart), so title-text parsing is the ONLY
    reliable source here.
    """
    result = {
        "ram_gb": None, "storage_gb": None, "battery_mah": None,
        "charging_watts": None, "refresh_rate_hz": None, "network": None,
        "display_inch": None, "processor": None, "color": None,
        "technology_tags": "",
    }
    if not title:
        return result

    combo_m = PATTERNS["mobile_ram_storage_combo"].search(title)
    if combo_m:
        result["ram_gb"] = extract_number(combo_m.group(1))
        result["storage_gb"] = extract_number(combo_m.group(2))
    else:
        ram_m = PATTERNS["mobile_ram"].search(title)
        if ram_m:
            result["ram_gb"] = extract_number(ram_m.group(1))
        storage_m = PATTERNS["mobile_storage"].search(title)
        if storage_m:
            val = extract_number(storage_m.group(1))
            if val and storage_m.group(2).upper() == "TB":
                val *= 1024
            result["storage_gb"] = val

    batt_m = PATTERNS["mobile_battery"].search(title)
    if batt_m:
        result["battery_mah"] = extract_number(batt_m.group(1))

    charge_m = PATTERNS["mobile_charging"].search(title)
    if charge_m:
        result["charging_watts"] = extract_number(charge_m.group(1))

    hz_m = PATTERNS["mobile_refresh_rate"].search(title)
    if hz_m:
        result["refresh_rate_hz"] = extract_number(hz_m.group(1))

    net_m = PATTERNS["mobile_network"].search(title)
    if net_m:
        result["network"] = net_m.group(1).upper()

    disp_m = PATTERNS["mobile_display_inch"].search(title)
    if disp_m:
        try:
            result["display_inch"] = float(disp_m.group(1))
        except ValueError:
            pass

    proc_m = PATTERNS["mobile_processor"].search(title)
    if proc_m:
        result["processor"] = clean_text(proc_m.group(1))

    parens_m = PATTERNS["mobile_first_parens"].search(title)
    if parens_m:
        first_segment = parens_m.group(1).split(",")[0].strip()
        if first_segment and not re.search(r"\d", first_segment):
            result["color"] = first_segment

    found_tags = [kw for kw in MOBILE_TECH_KEYWORDS if kw.lower() in title.lower()]
    result["technology_tags"] = ", ".join(found_tags)

    return result


def has_content(price: Optional[int], rating: Optional[str], rating_count: Optional[int],
                 delivery_info: Optional[str], service: Optional[str]) -> bool:
    return any([price, rating, rating_count, delivery_info, service])


def parse_product_card(card: Tag) -> Optional[Dict[str, Any]]:
    """UNCHANGED from v3.0."""
    title_raw, url = extract_title_and_link(card)
    title = validate_and_clean(title_raw, url)
    if not title:
        logger.debug(f"Skipping card: invalid/garbage title '{title_raw}'")
        return None

    ribbon_badge, ribbon_supplementary, badge_asin, badge_type = extract_badge_info(card)

    asin = (card.get("data-asin") or "").strip()
    asin_source = "card_data-asin" if asin else None
    if not asin and badge_asin:
        asin, asin_source = badge_asin, "badge_data-component-props"
    if not asin:
        url_asin = extract_asin_from_url(url)
        if url_asin:
            asin, asin_source = url_asin, "product_url"
    if not asin:
        logger.debug(f"Skipping card '{title[:40]}': no ASIN found via any method")
        return None

    price_data = extract_price_block(card)
    rating_data = extract_rating_info(card)
    meta = extract_bought_delivery_service(card)
    mobile_attrs = extract_mobile_attributes(title)

    if not has_content(price_data["price"], rating_data["rating"],
                        rating_data["rating_count"], meta["delivery_info"], meta["service"]):
        logger.warning(f"Card {asin} ('{title[:30]}...') had almost no extracted data. Skipping.")
        return None

    sponsored = "Sponsored" in card.get_text(" ", strip=True)
    deal_tag = card.select_one("span.a-badge-text")
    deal_badge = deal_tag.get_text(strip=True) if deal_tag and deal_tag.get_text(strip=True) != ribbon_badge else None

    swatch_links = card.select("div.s-color-swatch-container a[aria-label]")
    colors_available = ", ".join(a.get("aria-label", "").strip() for a in swatch_links if a.get("aria-label")) or None

    img_tag = card.select_one("img.s-image")
    img_url = img_tag.get("src") if img_tag else None

    return {
        "asin": asin, "asin_source": asin_source,
        "title": title, "url": url,
        "price": price_data["price"], "mrp": price_data["mrp"],
        "discount_pct": price_data["discount_pct"],
        "offer_text": price_data["offer_text"], "coupon_price": price_data["coupon_price"],
        "rating": rating_data["rating"], "rating_count": rating_data["rating_count"],
        "bought_info": meta["bought_info"], "delivery_info": meta["delivery_info"],
        "service": meta["service"], "sponsored": sponsored,
        "deal_badge": deal_badge, "ribbon_badge": ribbon_badge,
        "ribbon_supplementary": ribbon_supplementary, "ribbon_badge_type": badge_type,
        "ram_gb": mobile_attrs["ram_gb"], "storage_gb": mobile_attrs["storage_gb"],
        "battery_mah": mobile_attrs["battery_mah"], "charging_watts": mobile_attrs["charging_watts"],
        "refresh_rate_hz": mobile_attrs["refresh_rate_hz"], "network": mobile_attrs["network"],
        "display_inch": mobile_attrs["display_inch"], "processor": mobile_attrs["processor"],
        "color": mobile_attrs["color"], "technology_tags": mobile_attrs["technology_tags"],
        "colors_available": colors_available, "image_url": img_url,
        "scraped_at": datetime.now().isoformat(),
    }


def parse_listing_page(html: str) -> List[Dict[str, Any]]:
    """UNCHANGED from v3.0."""
    soup = BeautifulSoup(html, "html.parser")
    cards = get_leaf_cards(soup)

    if 0 < len(cards) < PRODUCTS_PER_PAGE_DEFAULT * MIN_ACCEPTABLE_PRODUCTS_RATIO:
        logger.debug(f"⚠️ Only {len(cards)} raw cards in DOM (expected ~{PRODUCTS_PER_PAGE_DEFAULT}).")

    results = {}  # keyed by ASIN — automatic per-page dedup
    for card in cards:
        try:
            parsed = parse_product_card(card)
            if parsed:
                results[parsed["asin"]] = parsed
        except Exception as e:
            logger.error(f"Error parsing card: {e}", exc_info=True)

    logger.info(f"Parsed {len(results)} products from page (raw cards: {len(cards)}).")
    return list(results.values())


def is_partial_page(products: List[Dict], expected: int = PRODUCTS_PER_PAGE_DEFAULT) -> bool:
    if not products:
        return True
    return len(products) < (expected * MIN_ACCEPTABLE_PRODUCTS_RATIO)


def is_captcha_page(html: Optional[str]) -> bool:
    if not html:
        return True
    return (
        "api-services-support@amazon.com" in html
        or "Enter the characters you see below" in html
        or "validateCaptcha" in html
    )


def cross_check_total_results(html: str) -> None:
    m = PATTERNS["total_results"].search(html)
    if m:
        logger.info(f"Amazon claims 'over {m.group(1)} results' (informational only — "
                    f"real crawlable depth is usually far less; CONSECUTIVE_FAILURE_LIMIT handles the cutoff).")


def get_total_pages(html: str, default: int = END_PAGE) -> int:
    if not html:
        return default
    soup = BeautifulSoup(html, "html.parser")

    disabled = soup.select("span.s-pagination-item.s-pagination-disabled")
    for el in disabled:
        t = el.get_text(strip=True)
        if t.isdigit():
            return min(int(t), default)

    nums = [int(el.get_text(strip=True)) for el in soup.select("span.s-pagination-item, a.s-pagination-item")
            if el.get_text(strip=True).isdigit()]
    if nums:
        return min(max(nums), default)
    return default


def build_url(base_url: str, page_num: int) -> str:
    """
    [B] Replaces build_page_url()/build_second_page_url(). Confirmed
    both source types (keyword search AND category node) paginate
    identically via "&page=N", with page 1 omitting the param.
    """
    if page_num <= 1:
        return base_url
    return f"{base_url}&page={page_num}"


# ============================================================
#     [C] NETWORK LAYER — PERSISTENT CONTEXT PER SOURCE
# ============================================================
_playwright_ctx = None
_browser_instance = None


def get_or_create_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        return _browser_instance

    _playwright_ctx = sync_playwright().start()
    _browser_instance = _playwright_ctx.chromium.launch(
        headless=True,
        args=[
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-extensions",
            "--disable-background-networking",
            "--js-flags=--max-old-space-size=256",
        ],
    )
    logger.info("Playwright browser launched (shared across all sources).")
    return _browser_instance


def close_shared_browser():
    global _playwright_ctx, _browser_instance
    if _browser_instance is not None:
        try:
            _browser_instance.close()
        except Exception:
            pass
        _browser_instance = None
    if _playwright_ctx is not None:
        try:
            _playwright_ctx.stop()
        except Exception:
            pass
        _playwright_ctx = None
    logger.info("Playwright browser closed.")


def new_context(browser):
    return browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 900},
        locale="en-IN",
        extra_http_headers={"Accept-Language": "en-IN,en;q=0.9"},
    )


class SourceFetcher:
    """
    [C] Manages ONE persistent browser CONTEXT reused across all page
    fetches within a single source (individual tabs/pages within that
    context are still opened/closed per fetch, matching v3.0's original
    granularity). Recycled periodically, and force-recycled on
    captcha/error — mirrors the Flipkart scraper's SourceFetcher
    lifecycle, with Amazon's scroll-poll + captcha-retry internals
    preserved verbatim from v3.0's fetch_html_playwright().
    """
    def __init__(self, browser):
        self.browser = browser
        self.context = None
        self.pages_served = 0

    def _open(self):
        self.context = new_context(self.browser)

    def _close(self):
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        self.context = None

    def _recycle_if_due(self):
        if self.pages_served > 0 and self.pages_served % RECYCLE_CONTEXT_EVERY_N_PAGES == 0:
            logger.info(f"  [tab] Recycling Playwright context after {self.pages_served} pages...")
            self._close()

    def fetch(self, url: str, target_card_count: int = PRODUCTS_PER_PAGE_DEFAULT) -> Optional[str]:
        if self.context is None:
            self._open()

        for attempt in range(MAX_RETRIES):
            try:
                page = self.context.new_page()
                page.goto(url, timeout=NAV_TIMEOUT_MS, wait_until="domcontentloaded")
                page.wait_for_timeout(1500)

                previous_count = 0
                max_scroll_attempts = 8
                for i in range(max_scroll_attempts):
                    page.mouse.wheel(0, 2500)
                    page.wait_for_timeout(700)

                    current_count = page.evaluate(
                        "document.querySelectorAll("
                        "'div.s-result-item[data-asin][data-component-type=\"s-search-result\"]'"
                        ").length"
                    )
                    logger.debug(f"  [scroll {i+1}] {current_count} cards visible in DOM")

                    if current_count >= target_card_count:
                        break
                    if current_count == previous_count and i >= 2:
                        break
                    previous_count = current_count

                html = page.content()
                page.close()

                if is_captcha_page(html):
                    logger.warning(f"  (captcha/block detected, attempt {attempt+1}/{MAX_RETRIES}) — recycling & retrying...")
                    self._close()
                    self._open()
                    time.sleep(4 * (attempt + 1))
                    continue

                self.pages_served += 1
                self._recycle_if_due()
                return html

            except Exception as e:
                logger.error(f"  Playwright error (attempt {attempt+1}/{MAX_RETRIES}): {e}")
                self._close()
                self._open()
                time.sleep(3 * (attempt + 1))

        return None

    def close(self):
        self._close()


def fetch_and_parse_page(fetcher: SourceFetcher, url: str, page_label: str = "",
                          cached_html: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    [D] NEW — forces ONE fresh-context retry if a page returns under
    MIN_ACCEPTABLE_PRODUCTS_RATIO of expected cards, even when it's NOT
    a captcha (v3.0 only retried on captcha/exception). Ported from the
    Flipkart scraper's robustness layer.
    """
    html = cached_html if cached_html is not None else fetcher.fetch(url)
    products = parse_listing_page(html) if html else []

    if is_partial_page(products):
        logger.warning(f"{page_label}: only {len(products)} products "
                       f"(expected ~{PRODUCTS_PER_PAGE_DEFAULT}) — forcing a fresh-context retry...")
        fetcher.close()
        html2 = fetcher.fetch(url)
        products2 = parse_listing_page(html2) if html2 else []
        if len(products2) > len(products):
            logger.info(f"{page_label}: retry recovered {len(products2)} (vs {len(products)}).")
            products = products2

    return products


# ============================================================
#                      EXCEL WRITER
# ============================================================

COLUMN_HEADERS = [
    "Source", "Page", "ASIN", "ASIN Source", "Title", "URL",
    "Price", "MRP", "Discount %", "Offer Text", "Coupon Price",
    "Rating", "Rating Count", "Bought Info", "Delivery Info", "Service",
    "Sponsored", "Deal Badge", "Ribbon Badge", "Ribbon Supplementary Text",
    "Ribbon Badge Type", "RAM (GB)", "Storage (GB)", "Battery (mAh)",
    "Charging (W)", "Refresh Rate (Hz)", "Network", "Display (inch)",
    "Processor", "Color", "Technology Tags",
    "Colors Available", "Image URL", "Scraped At",
]


def init_workbook():
    """[F] Returns (wb, ws, summary_ws) — ONE PageSummary sheet, both
    sheets carry a Source column, replacing the two-sheet v3.0 design."""
    try:
        wb = openpyxl.load_workbook(OUTPUT_FILE)
        ws = wb["Products"]
        summary_ws = wb["PageSummary"]
    except (FileNotFoundError, KeyError):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Products"
        ws.append(COLUMN_HEADERS)

        header_fill = PatternFill(start_color="0066CC", end_color="0066CC", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        summary_ws = wb.create_sheet("PageSummary")
        summary_ws.append(["Source", "Page", "Product Count", "New ASINs Added"])

    return wb, ws, summary_ws


def get_scraped_pages(summary_ws, source_label: str) -> set:
    """[F] Filters PageSummary rows by Source, generalized from v3.0's
    single-source get_scraped_pages()."""
    pages = set()
    for r in summary_ws.iter_rows(min_row=2, values_only=True):
        if r and r[0] == source_label and r[1] is not None:
            try:
                pages.add(int(r[1]))
            except (ValueError, TypeError):
                continue
    return pages


def get_existing_asins(ws) -> set:
    """ASIN column shifted from col 2 → col 3 due to new Source column."""
    return {str(r[0]) for r in ws.iter_rows(min_row=2, min_col=3, max_col=3, values_only=True) if r and r[0]}


def append_product_row(ws, source_label: str, page_num: int, p: Dict[str, Any]):
    ws.append([
        source_label, page_num, p["asin"], p["asin_source"], p["title"], p["url"],
        p["price"], p["mrp"], p["discount_pct"], p["offer_text"], p["coupon_price"],
        p["rating"], p["rating_count"], p["bought_info"], p["delivery_info"], p["service"],
        p["sponsored"], p["deal_badge"], p["ribbon_badge"], p["ribbon_supplementary"],
        p["ribbon_badge_type"], p["ram_gb"], p["storage_gb"], p["battery_mah"],
        p["charging_watts"], p["refresh_rate_hz"], p["network"], p["display_inch"],
        p["processor"], p["color"], p["technology_tags"],
        p["colors_available"], p["image_url"], p["scraped_at"],
    ])


# ============================================================
#                     DIAGNOSTIC MODE
# ============================================================

def run_diagnostic_mode():
    src = BASE_URLS[DIAGNOSTIC_SOURCE_INDEX]
    logger.info("=" * 60)
    logger.info(f"DIAGNOSTIC MODE: Source '{src['label']}' — Page 1 only, no file written")
    logger.info("=" * 60)

    browser = get_or_create_browser()
    fetcher = SourceFetcher(browser)
    html = fetcher.fetch(build_url(src["url"], 1))
    fetcher.close()

    if not html:
        logger.error("CRITICAL: Could not fetch page 1.")
        close_shared_browser()
        return

    cross_check_total_results(html)
    products = parse_listing_page(html)

    if not products:
        logger.error("CRITICAL: Could not extract any products from page 1.")
        close_shared_browser()
        return

    logger.info(f"\nExtracted {len(products)} products:\n")
    for i, p in enumerate(products[:15], 1):
        tag_str = f"[{p['ribbon_badge']}] " if p["ribbon_badge"] else ""
        print(f"{i:3d} | ASIN:{p['asin']:>12} ({p['asin_source']}) | {str(p['rating']):>3}★ "
              f"({p['rating_count']}) | ₹{str(p['price']):>8} | {tag_str}{p['title'][:45]}...")
        print(f"     RAM:{p['ram_gb']}GB Storage:{p['storage_gb']}GB Battery:{p['battery_mah']}mAh "
              f"Charging:{p['charging_watts']}W Refresh:{p['refresh_rate_hz']}Hz")
        print(f"     Network:{p['network']} Display:{p['display_inch']}in Processor:{p['processor']} "
              f"Color:{p['color']}")
        print(f"     Tags: {p['technology_tags']}")
        print("-" * 100)

    n = len(products)
    def pct(k): return f"{sum(1 for p in products if p[k] not in (None, '', False)) / n * 100:5.1f}%"
    print("\nFIELD FILL-RATE HEALTH CHECK")
    for key in ["price", "mrp", "discount_pct", "rating", "rating_count", "ram_gb",
                "storage_gb", "battery_mah", "charging_watts", "refresh_rate_hz",
                "network", "display_inch", "processor", "color", "technology_tags",
                "ribbon_badge", "delivery_info"]:
        print(f"   {key:<16} {pct(key)}")

    close_shared_browser()


# ============================================================
#         [A] GENERIC PER-SOURCE PROCESSING
# ============================================================

def process_source(wb, ws, summary_ws, source_label: str, source_url: str,
                    browser, global_asins: set) -> Tuple[int, int, int]:
    """
    Replaces v3.0's two hardcoded functions (main loop + 
    run_popularity_supplemental_scrape). Runs ANY source the same way:
    resumable via PageSummary (filtered by source_label), cross-source
    deduped via global_asins, circuit-broken via CONSECUTIVE_FAILURE_LIMIT
    scoped to this source only.
    """
    logger.info("\n" + "=" * 70)
    logger.info(f"STARTING SOURCE: '{source_label}'")
    logger.info(f"URL: {source_url}")
    logger.info("=" * 70)

    fetcher = SourceFetcher(browser)
    scraped_pages = get_scraped_pages(summary_ws, source_label)

    first_url = build_url(source_url, 1)
    logger.info(f"[{source_label}] Fetching page 1 to detect total pages...")
    first_html = fetcher.fetch(first_url)

    if not first_html:
        logger.error(f"[{source_label}] Cannot reach page 1 at all. Abandoning this source.")
        fetcher.close()
        return 0, 0, 0

    cross_check_total_results(first_html)
    total_pages = get_total_pages(first_html, default=END_PAGE)
    total_pages = min(END_PAGE, total_pages) if END_PAGE else total_pages
    logger.info(f"[{source_label}] Will process pages 1 → {total_pages}")

    added_count = 0
    dupes_skipped = 0
    consecutive_failures = 0
    pages_attempted = 0
    pages_since_save = 0

    try:
        for page_num in range(1, total_pages + 1):
            if page_num in scraped_pages:
                logger.info(f"[{source_label} page {page_num}] already scraped, skipping")
                continue

            if consecutive_failures >= CONSECUTIVE_FAILURE_LIMIT:
                logger.warning(
                    f"[{source_label}] {CONSECUTIVE_FAILURE_LIMIT} CONSECUTIVE failed/empty "
                    f"pages hit at page {page_num}. Abandoning this source and moving to the next one."
                )
                break

            url = build_url(source_url, page_num)
            page_label = f"[{source_label}] Page {page_num}/{total_pages}"

            if page_num == 1:
                products = fetch_and_parse_page(fetcher, url, page_label, cached_html=first_html)
            else:
                delay = random.uniform(*DELAY_RANGE)
                logger.info(f"\n--- {page_label} ({delay:.1f}s delay | consec.fail "
                            f"{consecutive_failures}/{CONSECUTIVE_FAILURE_LIMIT}) ---")
                time.sleep(delay)
                products = fetch_and_parse_page(fetcher, url, page_label)

            pages_attempted += 1
            new_count = 0

            if not products:
                consecutive_failures += 1
                logger.error(f"{page_label}: COMPLETE FAILURE ({consecutive_failures}/{CONSECUTIVE_FAILURE_LIMIT}).")
                ws.append([source_label, page_num, "[ERROR]", "", f"Failed to fetch/parse page {page_num}", url]
                           + [""] * (len(COLUMN_HEADERS) - 7) + [datetime.now().isoformat()])
            else:
                consecutive_failures = 0
                for p in products:
                    if p["asin"] in global_asins:
                        dupes_skipped += 1
                        continue
                    global_asins.add(p["asin"])
                    new_count += 1
                    append_product_row(ws, source_label, page_num, p)
                added_count += new_count
                logger.info(f"{page_label}: {len(products)} cards parsed, +{new_count} new ASINs "
                            f"(source running total: {added_count}, dupes skipped: {dupes_skipped})")

            summary_ws.append([source_label, page_num, len(products), new_count])

            pages_since_save += 1
            if pages_since_save >= SAVE_EVERY_N_PAGES:
                wb.save(OUTPUT_FILE)
                pages_since_save = 0
                logger.info(f"⚡️ Progress saved: {OUTPUT_FILE}")

    except KeyboardInterrupt:
        logger.warning(f"\n[{source_label}] Interrupted mid-source.")
        wb.save(OUTPUT_FILE)
        fetcher.close()
        raise

    fetcher.close()
    logger.info(f"\n[{source_label}] SOURCE COMPLETE: {added_count} new unique ASINs added, "
                f"{dupes_skipped} duplicates skipped, {pages_attempted} pages attempted.")
    return added_count, dupes_skipped, pages_attempted


# ============================================================
#                         MAIN
# ============================================================

def main():
    if DIAGNOSTIC_MODE:
        run_diagnostic_mode()
        return

    logger.info(">>> SCRIPT STARTED (multi-source v4.0)")
    logger.info(f"Sources to run: {[s['label'] for s in BASE_URLS]}")
    logger.info(f"Output: {OUTPUT_FILE}")

    wb, ws, summary_ws = init_workbook()
    global_asins = get_existing_asins(ws)
    logger.info(f"Resuming with {len(global_asins)} ASINs already in workbook.")

    browser = get_or_create_browser()
    grand_total_added = 0
    grand_total_dupes = 0

    try:
        for source in BASE_URLS:
            added, dupes, pages = process_source(
                wb, ws, summary_ws, source["label"], source["url"], browser, global_asins
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
        logger.info(f"ALL SOURCES DONE. Grand total NEW unique ASINs added this run: {grand_total_added}")
        logger.info(f"Grand total cross-source duplicates skipped: {grand_total_dupes}")
        logger.info(f"Total unique ASINs in workbook: {len(global_asins)}")
        logger.info(f"File saved: {OUTPUT_FILE}")
        logger.info("=" * 70)


if __name__ == "__main__":
    main()