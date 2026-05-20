"""
Unique Viewers Scraper
======================
Reads YouTube Studio analytics URLs from YouTube-manual-WS-Original.csv,
finds all Table_data.csv files under Channel* subdirectories, collects only
the channel IDs that are actually needed, scrapes each channel ONCE, then
updates every matching Table_data.csv in place.

Usage
-----
    python main.py                        # search for Channel* folders in ./
    python main.py --search-dir /path/to  # search in a different root directory
    python main.py --tabs 8               # parallel browser tabs (default 5)
"""

import argparse
import logging
import os
import shutil
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

BASE_DIR = Path(__file__).parent
WS_CSV = BASE_DIR / "YouTube-manual-WS-Original.csv"

# ---------------------------------------------------------------------------
# CSV helpers
# ---------------------------------------------------------------------------

def parse_metric_text(text: str) -> int | None:
    """Convert a YouTube-style metric string to an integer (e.g. '1.23M' -> 1230000)."""
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
    """Load the master channel URL list."""
    df = pd.read_csv(path, header=0, names=["Channel", "Channel_title", "URL"])
    df = df[
        df["Channel"].notna()
        & ~df["Channel"].isin(["Channel", "Studios", "Total"])
        & df["URL"].notna()
        & df["URL"].str.startswith("http", na=False)
    ].reset_index(drop=True)
    logger.info("Loaded %d channel URLs from %s", len(df), path.name)
    return df


def find_table_files(search_dir: Path) -> list[Path]:
    """Find all 'Table data.csv' files in Channel* subdirectories."""
    tables = sorted(search_dir.glob("Channel*/Table data.csv"))
    if not tables:
        # Fallback: search one level deeper
        tables = sorted(search_dir.glob("*/Channel*/Table data.csv"))
    logger.info("Found %d 'Table data.csv' file(s) under %s", len(tables), search_dir)
    for t in tables:
        logger.info("  %s", t.relative_to(search_dir))
    return tables


# ---------------------------------------------------------------------------
# Extraction
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
    """Navigate to a YouTube Studio analytics URL and extract the Unique viewers value."""
    if not _validate_url(url):
        logger.error("  Blocked unsafe URL for channel '%s' (must be https://studio.youtube.com)", channel_title)
        return None
    logger.info("-> %-35s  %s", channel_title, channel_id)
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60_000)
    except Exception as exc:
        # Redact the full URL from logs to avoid leaking owner IDs
        logger.error("  Failed to load page for '%s': %s", channel_title, type(exc).__name__)
        return None

    try:
        page.wait_for_selector(".debug-metric-value", timeout=20_000)
    except PWTimeout:
        logger.warning("  Table did not render for '%s'", channel_title)
        return None

    page.wait_for_timeout(1_000)

    value = page.evaluate(EXTRACT_JS)
    if value:
        result = parse_metric_text(str(value))
        if result is not None:
            logger.info("  OK  ->  %s", f"{result:,}")
            return result

    logger.warning("  FAIL  Could not extract unique viewers for '%s'", channel_title)
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Scrape Unique Viewers once and update all Channel*/Table_data.csv files"
    )
    parser.add_argument(
        "--search-dir", default=str(BASE_DIR), metavar="DIR",
        help="Root directory to search for Channel*/Table_data.csv (default: script directory)"
    )
    parser.add_argument(
        "--profile", default=None, metavar="DIR",
        help="Browser profile directory (default: ./browser_profile)"
    )
    parser.add_argument(
        "--tabs", type=int, default=5, metavar="N",
        help="Number of parallel browser tabs (default: 5)"
    )
    args = parser.parse_args()

    search_dir = Path(args.search_dir).resolve()
    profile_dir = str(Path(args.profile).resolve()) if args.profile else str(BASE_DIR / "browser_profile")
    num_tabs = max(1, args.tabs)

    # 1. Load master URL list
    ws_df = load_ws_csv(WS_CSV)
    if ws_df.empty:
        logger.error("No channel URLs found in %s – aborting.", WS_CSV.name)
        return

    # 2. Find all Table_data.csv files
    table_files = find_table_files(search_dir)
    if not table_files:
        logger.error("No 'Channel*/Table data.csv' files found under %s – aborting.", search_dir)
        return

    # 3. Collect the union of channel IDs needed across ALL tables
    needed_ids: set[str] = set()
    table_dfs: dict[Path, pd.DataFrame] = {}
    for path in table_files:
        df = pd.read_csv(path, encoding="utf-8")
        table_dfs[path] = df
        # Support both "Channel" column names
        col = "Channel" if "Channel" in df.columns else None
        if col is None:
            logger.warning("  No 'Channel' column in %s – skipping", path)
            continue
        ids = df[col].dropna().astype(str).str.strip()
        ids = ids[~ids.isin(["Channel", "Total", ""])]
        needed_ids.update(ids.tolist())

    logger.info("Need unique viewers for %d distinct channel(s) across %d table(s)",
                len(needed_ids), len(table_files))

    # 4. Filter master URL list to only the channels that are actually needed
    tasks_df = ws_df[ws_df["Channel"].isin(needed_ids)].copy()
    not_in_ws = needed_ids - set(ws_df["Channel"])
    if not_in_ws:
        logger.warning("%d channel(s) in tables but missing from WS CSV:\n  %s",
                       len(not_in_ws), "\n  ".join(sorted(not_in_ws)))

    tasks = [
        (str(r["Channel"]).strip(), str(r["Channel_title"]).strip(), str(r["URL"]).strip())
        for _, r in tasks_df.iterrows()
    ]
    logger.info("Will scrape %d channel(s) (skipping %d not needed)",
                len(tasks), len(ws_df) - len(tasks))

    # 5. Scrape – each channel exactly once
    scraped: dict[str, int] = {}  # channel_id -> unique_viewers

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

        # Login prompt (first run only – session is saved to profile_dir)
        login_page = context.pages[0] if context.pages else context.new_page()
        login_page.goto("https://studio.youtube.com", wait_until="domcontentloaded")
        print("\n" + "=" * 60)
        print("  A Chrome window has opened.")
        print("  -> If already logged in, press ENTER immediately.")
        print("  -> Otherwise log in first, then press ENTER.")
        print("  (Login is saved – you won't be asked again.)")
        print("=" * 60)
        input("  Press ENTER when ready >  ")
        login_page.close()

        logger.info("Scraping %d channels using %d tab(s)...", len(tasks), num_tabs)
        start_time = time.time()

        # Round-robin tasks across tabs so all tabs stay busy
        chunks = [tasks[i::num_tabs] for i in range(num_tabs)]
        pages = [context.new_page() for _ in range(num_tabs)]

        try:
            max_len = max((len(c) for c in chunks), default=0)
            for step in range(max_len):
                for pg, chunk in zip(pages, chunks):
                    if step >= len(chunk):
                        continue
                    channel_id, channel_title, url = chunk[step]
                    value = scrape_unique_viewers(pg, url, channel_id, channel_title)
                    if value is not None:
                        scraped[channel_id] = value
        finally:
            for pg in pages:
                pg.close()
            context.close()

    elapsed = time.time() - start_time
    logger.info("Scraping complete – %d / %d channels extracted in %.0fs",
                len(scraped), len(tasks), elapsed)

    if not scraped:
        logger.error("No data extracted. No files were modified.")
        return

    # 6. Update every Table_data.csv with the scraped values
    total_updated = 0
    for path, df in table_dfs.items():
        col = "Channel" if "Channel" in df.columns else None
        if col is None:
            continue
        updated = 0
        for channel_id, unique_viewers in scraped.items():
            mask = df[col].astype(str).str.strip() == channel_id
            if mask.any():
                df.loc[mask, "Unique viewers"] = unique_viewers
                updated += 1
        if updated:
            # Backup original before overwriting
            backup = path.with_suffix(".csv.bak")
            shutil.copy2(path, backup)
            df.to_csv(path, index=False)
            logger.info("  Updated %d row(s) in %s  (backup: %s)", updated, path.relative_to(search_dir), backup.name)
            total_updated += updated
        else:
            logger.info("  No matching channels in %s", path.relative_to(search_dir))

    logger.info("Done! %d total row(s) updated across %d file(s)", total_updated, len(table_files))


if __name__ == "__main__":
    main()
