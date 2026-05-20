"""
Unique Viewers Scraper — Standalone Extract
============================================
Reads YouTube Studio analytics URLs from the WS CSV file,
opens each page in parallel browser tabs, extracts the "Unique viewers"
value, and writes a clean output CSV containing:

    Channel | Channel_title | Unique_viewers

No Table_data.csv join required — this is a pure extract.

Usage
-----
    python scrape_unique_viewers.py
    python scrape_unique_viewers.py --tabs 9
    python scrape_unique_viewers.py --output my_results.csv
"""

import argparse
import logging
import os
import time
from pathlib import Path
from urllib.parse import urlparse

# CWE-532 guard: Playwright debug flags (DEBUG, PLAYWRIGHT_DEBUG) can dump full
# HTTP request headers — including session cookies — to stdout/log files.
# Unconditionally remove them so they cannot be set externally to leak credentials.
for _var in ("DEBUG", "PLAYWRIGHT_DEBUG", "PWDEBUG"):
    os.environ.pop(_var, None)

import pandas as pd
from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright

ALLOWED_HOSTS = {"studio.youtube.com"}

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# File paths  — update WK_DIR and WS_CSV to match your weekly folder
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
WK_DIR = BASE_DIR 
WS_CSV = WK_DIR / "YouTube-manual-WS-Original.csv"
DEFAULT_OUTPUT_CSV = WK_DIR / "unique_viewers_extract.csv"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_metric_text(text: str) -> int | None:
    """Convert a YouTube-style metric string to an integer.

    Examples:
        '1.23M' -> 1_230_000
        '45.6K' -> 45_600
        '9,812' -> 9_812
    """
    text = text.strip().replace(",", "").replace(" ", "")
    if not text:
        return None
    multiplier = 1
    if text.upper().endswith("M"):
        multiplier = 1_000_000
        text = text[:-1]
    elif text.upper().endswith("K"):
        multiplier = 1_000
        text = text[:-1]
    try:
        return int(float(text) * multiplier)
    except ValueError:
        return None


def load_ws_csv(path: Path) -> pd.DataFrame:
    """Load the WS CSV and return only rows that have a valid URL."""
    df = pd.read_csv(path, header=0, names=["Channel", "Channel_title", "URL"])
    df = df[
        df["Channel"].notna()
        & ~df["Channel"].isin(["Channel", "Studios", "Total"])
        & df["URL"].notna()
        & df["URL"].str.startswith("http", na=False)
    ].reset_index(drop=True)
    logger.info("Loaded %d channel URLs from %s", len(df), path.name)
    return df


# ---------------------------------------------------------------------------
# JavaScript injected into each page to locate the Unique Viewers cell
# ---------------------------------------------------------------------------
EXTRACT_JS = """() => {
    const headers = document.querySelectorAll(
        'yta-explore-table-header-cell .debug-metric-title'
    );
    let colIdx = -1;
    for (let i = 0; i < headers.length; i++) {
        if (headers[i].textContent.trim().toLowerCase().includes('unique viewer')) {
            colIdx = i;
            break;
        }
    }
    if (colIdx === -1) return null;

    const rows = document.querySelectorAll('yta-explore-table-row');
    if (!rows.length) return null;
    const cells = rows[0].querySelectorAll('.debug-metric-value');
    if (colIdx >= cells.length) return null;
    return cells[colIdx].textContent.trim();
}"""


def _validate_url(url: str) -> bool:
    """Return True only for https URLs on the allowed YouTube Studio host."""
    try:
        parsed = urlparse(url)
        return parsed.scheme == "https" and parsed.netloc in ALLOWED_HOSTS
    except Exception:
        return False


def scrape_unique_viewers(page: Page, url: str, channel_id: str, channel_title: str) -> int | None:
    """Navigate to a YouTube Studio analytics URL and return the Unique viewers integer."""
    if not _validate_url(url):
        logger.error("  Blocked unsafe URL for channel '%s' (must be https://studio.youtube.com)", channel_title)
        return None
    logger.info("-> %-30s  %s", channel_title, channel_id)
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    except Exception as exc:
        # Redact the full URL from logs to avoid leaking owner IDs
        logger.error("  Failed to load page for '%s': %s", channel_title, type(exc).__name__)
        return None

    try:
        page.wait_for_selector(".debug-metric-value", timeout=20_000)
    except PWTimeout:
        logger.warning("  Table did not render in time for '%s'", channel_title)
        return None

    # Brief settle for any remaining JS rendering
    page.wait_for_timeout(1_000)

    value = page.evaluate(EXTRACT_JS)
    if value:
        result = parse_metric_text(str(value))
        if result is not None:
            logger.info("  OK  %-30s  ->  %s", channel_title, f"{result:,}")
            return result

    logger.warning("  FAIL  Could not extract unique viewers for '%s'", channel_title)
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extract Unique Viewers from YouTube Studio and save a standalone CSV."
    )
    parser.add_argument(
        "--profile", default=None, metavar="DIR",
        help="Browser profile directory (default: ./browser_profile)",
    )
    parser.add_argument(
        "--output", default=str(DEFAULT_OUTPUT_CSV), metavar="CSV",
        help="Output CSV path (default: unique_viewers_extract.csv)",
    )
    parser.add_argument(
        "--tabs", type=int, default=9, metavar="N",
        help="Number of parallel browser tabs (default: 9)",
    )
    args = parser.parse_args()

    output_path = Path(args.output)
    profile_dir = str(Path(args.profile).resolve()) if args.profile else str(BASE_DIR / "browser_profile")
    num_tabs = max(1, args.tabs)

    # ---- Load the WS CSV (Channel | Channel_title | URL) ------------------
    ws_df = load_ws_csv(WS_CSV)
    if ws_df.empty:
        logger.error("No channel URLs found in %s — aborting.", WS_CSV.name)
        return

    tasks = [
        (str(r["Channel"]).strip(), str(r["Channel_title"]).strip(), str(r["URL"]).strip())
        for _, r in ws_df.iterrows()
    ]

    results: list[dict] = []

    # ---- Browser session --------------------------------------------------
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            profile_dir,
            headless=False,
            no_viewport=True,
            channel="chrome",
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
            ],
            ignore_default_args=["--enable-automation"],
        )

        # Login prompt (skipped automatically after the first run)
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://studio.youtube.com", wait_until="domcontentloaded")
        print("\n" + "=" * 60)
        print("  A Chrome window has opened.")
        print("  -> If already logged in, press ENTER immediately.")
        print("  -> Otherwise log in first, then press ENTER.")
        print("  (Login is saved — you won't be asked again.)")
        print("=" * 60)
        input("  Press ENTER when ready >  ")
        page.close()

        # Scrape across N parallel tabs using round-robin
        logger.info("Scraping %d channels using %d tab(s)...", len(tasks), num_tabs)
        start_time = time.time()

        chunks = [tasks[i::num_tabs] for i in range(num_tabs)]
        pages = [context.new_page() for _ in range(num_tabs)]

        try:
            max_len = max(len(c) for c in chunks)
            for step in range(max_len):
                for pg, chunk in zip(pages, chunks):
                    if step >= len(chunk):
                        continue
                    channel_id, channel_title, url = chunk[step]
                    value = scrape_unique_viewers(pg, url, channel_id, channel_title)
                    results.append({
                        "Channel": channel_id,
                        "Channel_title": channel_title,
                        "Unique_viewers": value,   # None if extraction failed
                    })
        finally:
            for pg in pages:
                pg.close()
            context.close()

    elapsed = time.time() - start_time
    extracted = sum(1 for r in results if r["Unique_viewers"] is not None)
    logger.info(
        "Scraping complete — %d / %d extracted in %.0fs",
        extracted, len(tasks), elapsed,
    )

    if not results:
        logger.error("No data extracted. Output file was NOT created.")
        return

    # ---- Build and save the output CSV ------------------------------------
    output_df = pd.DataFrame(results, columns=["Channel", "Channel_title", "Unique_viewers"])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_df.to_csv(output_path, index=False)

    logger.info(
        "Done! %d row(s) written to %s  (%d failed)",
        len(output_df), output_path, len(output_df) - extracted,
    )


if __name__ == "__main__":
    main()
