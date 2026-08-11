# Daraz Product Review Scraper

A Python script that opens a Daraz product page in a real browser (via
Playwright), scrolls to the reviews section, and collects product + review
data into an Excel file.

**Fields collected:** Product Name, Platform Name, Overall Rating, Seller
Name, Price, Reviewer Name, Reviewer Rating, Review Date, Comment, Scraped At.

## Disclaimer

This project is shared for **educational purposes** — to demonstrate browser
automation and web scraping techniques with Playwright. It is not affiliated
with or endorsed by Daraz.

Before using it:
- Review Daraz's [Terms of Service](https://www.daraz.pk/wow/i/en/terms) and `robots.txt` — automated data collection may not be permitted.
- Don't use this to overload their servers — keep the built-in delays.
- Don't scrape, store, or redistribute personal data (reviewer names, etc.) without a lawful basis to do so.
- You are responsible for how you use this script and any data it collects.

## Requirements

- Python 3.9+
- Google Chrome / Chromium (installed automatically by Playwright, see below)

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/<your-username>/daraz-scraper.git
cd daraz-scraper

# 2. Create a virtual environment (recommended)
python -m venv venv
source venv/bin/activate      # on Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Install the Playwright browser binaries
playwright install chromium
```

## Usage

```bash
python daraz_scraper.py "https://www.daraz.pk/products/<product-url>"
```

The script will:
1. Open the product page in a visible Chromium window (so you can see what's happening / where it might get stuck).
2. Scrape the product name, seller, overall rating, and price.
3. Scroll to the reviews section and page through reviews (up to `MAX_REVIEW_PAGES`, default 500).
4. Save everything to `daraz_reviews.xlsx` in the current directory.

## Configuration

Edit these constants near the top of `daraz_scraper.py`:

| Variable | Purpose | Default |
|---|---|---|
| `OUTPUT_FILE` | Output Excel filename | `daraz_reviews.xlsx` |
| `MAX_REVIEW_PAGES` | Max review pages to page through | `500` |
| `WAIT_BETWEEN_ACTIONS_MS` | Delay between actions (ms) — keep this reasonable to stay polite to the server | `1500` |

## Troubleshooting: fields returning "N/A"

Daraz occasionally changes its page structure, which can break CSS selectors.
If a field comes back as `N/A`:

1. Open the product page in Chrome.
2. Right-click the element in question → **Inspect**.
3. Note its class/tag in DevTools.
4. Update the corresponding `page.locator(...)` / `item.locator(...)` call in the script.
5. For review fields specifically, scroll to "Ratings & Reviews" *before* inspecting — that section loads in dynamically via JavaScript.

## License

MIT — see [LICENSE](LICENSE). Feel free to fork and adapt, but note the disclaimer above still applies to how you use it.
