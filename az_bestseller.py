"""
Amazon Bestsellers — Smartphones — Simple Scraper
===================================================
Scrapes the Amazon.in "Bestsellers in Smartphones" list (100 products,
2 fixed pages). Standalone script — no dependency on the main
multi-source assortment scraper.

Fields captured: Rank, ASIN, Title, Price, Rating, Rating Count,
Product URL, Image URL.

SETUP:
    pip install beautifulsoup4 openpyxl playwright
    playwright install chromium

USAGE:
    python3 amazon_bestsellers_smartphones.py
"""

import re
import time
import logging
from datetime import datetime
from typing import Optional, Dict, List, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# ============================================================
#                      CONFIGURATION
# ============================================================

PAGE_URLS = [
    "https://www.amazon.in/gp/bestsellers/electronics/1805560031",
    "https://www.amazon.in/gp/bestsellers/electronics/1805560031/ref=zg_bs_pg_2_electronics?ie=UTF8&pg=2",
]

OUTPUT_FILE = f"amazon_bestseller_smartphones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

NAV_TIMEOUT_MS = 30000
MAX_RETRIES = 3
DELAY_BETWEEN_PAGES = 4.0

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
logger = logging.getLogger(__name__)

PRICE_RE = re.compile(r"₹\s?([\d,]+)")
RATING_RE = re.compile(r"([\d.]+)\s*out of\s*5", re.IGNORECASE)
RANK_RE = re.compile(r"#?\s*(\d+)")


def is_captcha_page(html: Optional[str]) -> bool:
    if not html:
        return True
    return (
        "api-services-support@amazon.com" in html
        or "Enter the characters you see below" in html
        or "validateCaptcha" in html
    )


def extract_number(raw: Optional[str]) -> Optional[int]:
    if raw is None:
        return None
    s = str(raw).replace(",", "").replace("₹", "").strip()
    try:
        return int(float(s))
    except ValueError:
        return None


# ============================================================
#                     NETWORK LAYER
# ============================================================

def fetch_html(url: str, target_count: int = 50) -> Optional[str]:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1366, "height": 1000},
            locale="en-IN",
            extra_http_headers={"Accept-Language": "en-IN,en;q=0.9"},
        )

        html = None
        for attempt in range(MAX_RETRIES):
            try:
                page = context.new_page()
                page.goto(url, timeout=NAV_TIMEOUT_MS, wait_until="domcontentloaded")
                page.wait_for_timeout(1500)

                # --- live-DOM-count scroll polling (ported from main scraper) ---
                previous_count = 0
                max_scroll_attempts = 25          # generous ceiling — bestseller grid is long
                stable_rounds = 0

                for i in range(max_scroll_attempts):
                    page.mouse.wheel(0, 2500)
                    page.wait_for_timeout(600)

                    current_count = page.evaluate(
                        "document.querySelectorAll('span.zg-bdg-text').length"
                    )
                    logger.debug(f"  [scroll {i+1}] {current_count} ranked cards visible")

                    if current_count >= target_count:
                        break

                    if current_count == previous_count:
                        stable_rounds += 1
                        if stable_rounds >= 3:      # plateaued for 3 checks in a row → stop
                            break
                    else:
                        stable_rounds = 0

                    previous_count = current_count

                # final nudge: scroll all the way to the bottom once more + settle
                page.keyboard.press("End")
                page.wait_for_timeout(1200)

                content = page.content()
                page.close()

                if is_captcha_page(content):
                    logger.warning(f"Captcha/block detected (attempt {attempt+1}/{MAX_RETRIES}), retrying...")
                    context.close()
                    context = browser.new_context(
                        user_agent=USER_AGENT,
                        viewport={"width": 1366, "height": 1000},
                        locale="en-IN",
                        extra_http_headers={"Accept-Language": "en-IN,en;q=0.9"},
                    )
                    time.sleep(4 * (attempt + 1))
                    continue

                final_count = BeautifulSoup(content, "html.parser").select("span.zg-bdg-text")
                logger.info(f"  Final rendered rank-badge count before parse: {len(final_count)}")

                html = content
                break

            except Exception as e:
                logger.error(f"Playwright error (attempt {attempt+1}/{MAX_RETRIES}): {e}")
                time.sleep(3 * (attempt + 1))

        context.close()
        browser.close()
        return html

# ============================================================
#                      PARSING LOGIC
# ============================================================

def parse_bestseller_card(card) -> Optional[Dict[str, Any]]:
    asin = card.get("data-asin", "").strip()
    if not asin:
        return None

    # --- Rank ---
    rank_tag = card.select_one("span.zg-bdg-text")
    rank = None
    if rank_tag:
        m = RANK_RE.search(rank_tag.get_text(strip=True))
        if m:
            rank = int(m.group(1))

    # Skip non-ranked blocks (ads/ghost cards without a rank badge)
    if rank is None:
        return None

    # --- Title + URL (the non-aria-hidden link) ---
    title = None
    url = None
    for a in card.select("a.a-link-normal.aok-block"):
        if a.get("aria-hidden") == "true":
            continue
        txt = a.get_text(strip=True)
        if txt:
            title = txt
            href = a.get("href")
            if href:
                url = "https://www.amazon.in" + href if href.startswith("/") else href
            break

    if not title:
        return None

    # --- Image ---
    img_tag = card.select_one("img.p13n-product-image")
    image_url = img_tag.get("src") if img_tag else None

    # --- Rating + Rating Count ---
    rating = None
    rating_count = None
    rating_anchor = card.select_one("div.a-icon-row a[aria-label]")
    if rating_anchor:
        alt_tag = rating_anchor.select_one("span.a-icon-alt")
        if alt_tag:
            m = RATING_RE.search(alt_tag.get_text(strip=True))
            if m:
                rating = m.group(1)
        count_tag = rating_anchor.select_one("span.a-size-small[aria-hidden='true']")
        if count_tag:
            rating_count = extract_number(count_tag.get_text(strip=True))

    # --- Price ---
    price = None
    price_tag = card.select_one("span.a-color-price")
    if price_tag:
        price_text = price_tag.get_text(" ", strip=True)
        m = PRICE_RE.search(price_text)
        if m:
            price = extract_number(m.group(1))

    return {
        "rank": rank,
        "asin": asin,
        "title": title,
        "price": price,
        "rating": rating,
        "rating_count": rating_count,
        "url": url,
        "image_url": image_url,
        "scraped_at": datetime.now().isoformat(),
    }


def parse_listing_page(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select("div[data-asin]")

    results = {}
    for card in cards:
        try:
            parsed = parse_bestseller_card(card)
            if parsed:
                results[parsed["asin"]] = parsed
        except Exception as e:
            logger.error(f"Error parsing card: {e}", exc_info=True)

    products = list(results.values())
    products.sort(key=lambda p: p["rank"])
    logger.info(f"Parsed {len(products)} ranked products from page (raw candidate divs: {len(cards)}).")
    return products


# ============================================================
#                      EXCEL WRITER
# ============================================================

COLUMN_HEADERS = ["Rank", "ASIN", "Title", "Price", "Rating", "Rating Count", "URL", "Image URL", "Scraped At"]


def init_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bestsellers"
    ws.append(COLUMN_HEADERS)

    header_fill = PatternFill(start_color="CC6600", end_color="CC6600", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    return wb, ws


def append_row(ws, p: Dict[str, Any]):
    ws.append([
        p["rank"], p["asin"], p["title"], p["price"], p["rating"],
        p["rating_count"], p["url"], p["image_url"], p["scraped_at"],
    ])


# ============================================================
#                         MAIN
# ============================================================

def main():
    logger.info(">>> Amazon Bestsellers Smartphones scraper started")
    wb, ws = init_workbook()

    seen_asins = set()
    total_added = 0

    for i, url in enumerate(PAGE_URLS, 1):
        logger.info(f"\n--- Fetching page {i}/{len(PAGE_URLS)}: {url} ---")
        html = fetch_html(url)

        if not html:
            logger.error(f"Page {i}: failed to fetch after retries. Skipping.")
            continue

        products = parse_listing_page(html)
        new_count = 0
        for p in products:
            if p["asin"] in seen_asins:
                continue
            seen_asins.add(p["asin"])
            append_row(ws, p)
            new_count += 1

        total_added += new_count
        logger.info(f"Page {i}: {len(products)} parsed, {new_count} new added.")

        wb.save(OUTPUT_FILE)
        logger.info(f"⚡️ Progress saved: {OUTPUT_FILE}")

        if i < len(PAGE_URLS):
            time.sleep(DELAY_BETWEEN_PAGES)

    wb.save(OUTPUT_FILE)
    logger.info("\n" + "=" * 60)
    logger.info(f"DONE. Total unique products scraped: {total_added}")
    logger.info(f"File saved: {OUTPUT_FILE}")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
