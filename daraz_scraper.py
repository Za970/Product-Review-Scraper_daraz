"""
Daraz Product Review Scraper
=============================
Scrapes: Product Name, Platform Name, Overall Rating, Seller Name,
         Reviewer Name, Reviewer Rating, Review Date, Comment, Scraped At
Saves results into an Excel (.xlsx) file.

HOW TO RUN:
    py daraz_scraper.py "https://www.daraz.pk/products/led-led-500-i428204760-s2266670210.html?c=&channelLpJumpArgs=&clickTrackInfo=query%253Asmart%252Btemperature%252Bwater%252Bbottle%253Bnid%253A428204760%253Bsrc%253ALazadaMainSrp%253Brn%253A5c23d733348f893201e11aea485f7808%253Bregion%253Apk%253Bsku%253A428204760_PK%253Bprice%253A870%253Bclient%253Adesktop%253Bsupplier_id%253A6005057619968%253Bsession_id%253A%253Bbiz_source%253Ahttps%253A%252F%252Fwww.daraz.pk%252F%253Bslot%253A10%253Butlog_bucket_id%253A470687%253Basc_category_id%253A4688%253Bitem_id%253A428204760%253Bsku_id%253A2266670210%253Bshop_id%253A1310928%253BtemplateInfo%253A-1_A3_C%2523115931_G%2523&configId=choice_PK_promotion&freeshipping=0&fs_ab=1&fuse_fs=&lang=en&location=Punjab&price=8.7E%202&priceCompare=skuId%3A2266670210%3Bsource%3Alazada-search-voucher%3Bsn%3A5c23d733348f893201e11aea485f7808%3BoriginPrice%3A87000%3BdisplayPrice%3A87000%3BisGray%3Afalse%3BsinglePromotionId%3A50000078939002%3BsingleToolCode%3ApromPrice%3BvoucherPricePlugin%3A0%3Btimestamp%3A1786447933017&ratingscore=4.72177573670111&request_id=5c23d733348f893201e11aea485f7808&review=2613&sale=12355&search=1&source=search&spm=a211g0.searchlist.list.10&stock=1&upItemIds=428204760"

IMPORTANT NOTES (read before running):
1. Daraz's page structure changes over time and can differ slightly between
   product pages. The CSS selectors below are a best-effort starting point.
   If a field comes back "N/A", you need to update that selector — see the
   "HOW TO FIND CORRECT SELECTORS" section at the bottom of this file.
2. Respect Daraz's robots.txt and Terms of Service. Don't hammer their
   servers — the delays in this script are intentional, don't remove them.
3. headless=False is used on purpose so you can WATCH the browser and see
   exactly what it's doing / where it might be getting stuck or blocked.
"""

import asyncio
import sys
from datetime import datetime

import pandas as pd
from playwright.async_api import async_playwright

PLATFORM_NAME = "Daraz"
OUTPUT_FILE = "daraz_reviews.xlsx"
MAX_REVIEW_PAGES = 500          # how many pages of reviews to click through
WAIT_BETWEEN_ACTIONS_MS = 1500 # be polite, don't rush the site


async def scrape_product(url: str, max_review_pages: int = MAX_REVIEW_PAGES):
    results = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            )
        )
        page = await context.new_page()

        print(f"Opening: {url}")
        await page.goto(url, timeout=60000)
        await page.wait_for_load_state("networkidle")

        # ---------- Product Name ----------
        try:
            product_name = await page.locator("h1.pdp-mod-product-badge-title").inner_text()
        except Exception:
            product_name = "N/A"

        # ---------- Seller Name ----------
        try:
            seller_name = await page.locator(".seller-name__detail a").first.inner_text()
        except Exception:
            seller_name = "N/A"

        # ---------- Overall Rating ----------
        try:
            overall_rating = await page.locator(".score-average").first.inner_text()
        except Exception:
            overall_rating = "N/A"

            # ---------- Price ----------
        try:
           price = await page.locator(".pdp-price_type_normal").first.inner_text()
        except Exception:
          price = "N/A"


        print(f"Product: {product_name} | Seller: {seller_name} | Rating: {overall_rating} | Price: {price}")

        # Scroll down so the reviews section lazy-loads
        await page.mouse.wheel(0, 3000)
        await page.wait_for_timeout(WAIT_BETWEEN_ACTIONS_MS)

        page_num = 1
        # Everything below is wrapped in try/except so that if ANYTHING goes
        # wrong mid-scrape (a selector fails, a button click errors, etc.),
        # whatever reviews we already collected in `results` are still kept
        # and returned — instead of being lost when the script crashes.
        try:
            while page_num <= max_review_pages:
                review_items = await page.locator(".mod-reviews .item").all()

                if not review_items:
                    print("No review items found on this page — check selectors.")
                    break

                print(f"Page {page_num}: found {len(review_items)} reviews")

                for item in review_items:
                    try:
                        reviewer_name = await item.locator(".middle .name").inner_text()
                    except Exception:
                        reviewer_name = "N/A"

                    try:
                        # Star rating is usually rendered as filled star icons —
                        # counting filled stars is a common way to get the number.
                        reviewer_rating = await item.locator(".star.fill").count()
                    except Exception:
                        reviewer_rating = "N/A"

                    try:
                        review_date = await item.locator(".top .title").inner_text()
                    except Exception:
                        review_date = "N/A"

                    try:
                        comment = await item.locator(".content").inner_text()
                    except Exception:
                        comment = "N/A"

                    results.append({
                        "Product Name": product_name,
                        "Platform Name": PLATFORM_NAME,
                        "Overall Rating": overall_rating,
                        "Seller Name": seller_name,
                        "Reviewer Name": reviewer_name,
                        "Reviewer Rating": reviewer_rating,
                        "Review Date": review_date,
                        "Comment": comment,
                        "Price": price,
                        "Scraped At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })

                # Try to go to next review page.
                # .first is important — Daraz renders more than one element
                # matching this selector (duplicate/hidden layout variants),
                # and Playwright refuses to act on an ambiguous match without it.
                next_btn = page.locator(".next-pagination-item.next").first
                if await next_btn.count() > 0:
                    is_disabled = await next_btn.get_attribute("class")
                    if is_disabled and "disabled" in is_disabled:
                        print("Reached last page of reviews.")
                        break
                    await next_btn.click()
                    await page.wait_for_timeout(WAIT_BETWEEN_ACTIONS_MS)
                    page_num += 1
                else:
                    print("No next-page button found — assuming last page.")
                    break
        except Exception as e:
            print(f"Stopped early due to an error while paging through reviews: {e}")
            print(f"Keeping the {len(results)} review(s) collected so far.")

        await browser.close()

    return results


def save_to_excel(data, filename=OUTPUT_FILE):
    if not data:
        print("Nothing to save.")
        return
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False, engine="openpyxl")
    print(f"Saved {len(data)} reviews to {filename}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print('Usage: py daraz_scraper.py "https://www.daraz.pk/products/led-led-500-i428204760-s2266670210.html?c=&channelLpJumpArgs=&clickTrackInfo=query%253Asmart%252Btemperature%252Bwater%252Bbottle%253Bnid%253A428204760%253Bsrc%253ALazadaMainSrp%253Brn%253A5c23d733348f893201e11aea485f7808%253Bregion%253Apk%253Bsku%253A428204760_PK%253Bprice%253A870%253Bclient%253Adesktop%253Bsupplier_id%253A6005057619968%253Bsession_id%253A%253Bbiz_source%253Ahttps%253A%252F%252Fwww.daraz.pk%252F%253Bslot%253A10%253Butlog_bucket_id%253A470687%253Basc_category_id%253A4688%253Bitem_id%253A428204760%253Bsku_id%253A2266670210%253Bshop_id%253A1310928%253BtemplateInfo%253A-1_A3_C%2523115931_G%2523&configId=choice_PK_promotion&freeshipping=0&fs_ab=1&fuse_fs=&lang=en&location=Punjab&price=8.7E%202&priceCompare=skuId%3A2266670210%3Bsource%3Alazada-search-voucher%3Bsn%3A5c23d733348f893201e11aea485f7808%3BoriginPrice%3A87000%3BdisplayPrice%3A87000%3BisGray%3Afalse%3BsinglePromotionId%3A50000078939002%3BsingleToolCode%3ApromPrice%3BvoucherPricePlugin%3A0%3Btimestamp%3A1786447933017&ratingscore=4.72177573670111&request_id=5c23d733348f893201e11aea485f7808&review=2613&sale=12355&search=1&source=search&spm=a211g0.searchlist.list.10&stock=1&upItemIds=428204760"')
        sys.exit(1)

    product_url = sys.argv[1]
    data = asyncio.run(scrape_product(product_url))
    save_to_excel(data)


# ============================================================
# HOW TO FIND CORRECT SELECTORS IF SOMETHING RETURNS "N/A"
# ============================================================
# 1. Open the Daraz product page in Chrome.
# 2. Right-click the element you want (e.g. a reviewer's name) -> "Inspect".
# 3. DevTools opens and highlights the HTML for that element.
# 4. Note its tag and class, e.g. <div class="middle"><div class="name">John</div></div>
# 5. Update the matching line in this script, e.g.:
#       item.locator(".middle .name")
# 6. Do this for each field: product name, seller name, overall rating,
#    reviewer name, reviewer rating (stars), review date, and comment text.
# 7. For the reviews section specifically, scroll to "Ratings & Reviews"
#    on the page first, THEN inspect — that section often loads in via
#    JavaScript after the page opens, so it won't be in the initial HTML.
