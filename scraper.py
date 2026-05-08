# scraper.py
# Downloads yearly Excel files from the Swedish Investment Fund Association
# website (fondbolagen.se) using Selenium stealth.
# The site is dynamic — we navigate to the quarterly statistics page,
# find download links, and download the needed year files.

import os
import sys
import time
import re
import logging
import subprocess
import random

import config

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Chrome version detection (Windows dev + Linux Docker)
# ─────────────────────────────────────────────────────────────────────────────

def get_chrome_version():
    """Detect Chrome major version -- works on Windows (dev) and Linux (Docker)."""
    if sys.platform == 'win32':
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r'Software\Google\Chrome\BLBeacon',
            )
            return winreg.QueryValueEx(key, 'version')[0].split('.')[0]
        except Exception:
            pass
    for cmd in ['google-chrome', 'google-chrome-stable',
                'chromium', 'chromium-browser']:
        try:
            out = subprocess.check_output(
                [cmd, '--version'], stderr=subprocess.DEVNULL
            ).decode()
            return out.strip().split()[-1].split('.')[0]
        except Exception:
            continue
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _human_delay(lo=0.4, hi=1.2):
    time.sleep(random.uniform(lo, hi))


# ─────────────────────────────────────────────────────────────────────────────
# Build driver
# ─────────────────────────────────────────────────────────────────────────────

def _build_driver(download_dir):
    """Create a Selenium stealth Chrome driver configured for Excel download."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    try:
        from selenium_stealth import stealth
    except ImportError:
        stealth = None

    abs_dl = os.path.abspath(download_dir)
    os.makedirs(abs_dl, exist_ok=True)

    opts = Options()
    if config.HEADLESS_MODE:
        opts.add_argument('--headless=new')
        opts.add_argument('--disable-gpu')

    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--window-size=1920,1080')
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_argument('--lang=en-US')
    opts.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/131.0.0.0 Safari/537.36'
    )

    prefs = {
        'download.default_directory': abs_dl,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': False,
        'profile.default_content_settings.popups': 0,
    }
    opts.add_experimental_option('prefs', prefs)
    opts.add_experimental_option('excludeSwitches', ['enable-automation'])
    opts.add_experimental_option('useAutomationExtension', False)

    driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(config.WAIT_TIMEOUT * 2)

    if stealth is not None:
        stealth(
            driver,
            languages=['en-US', 'en'],
            vendor='Google Inc.',
            platform='Win32',
            webgl_vendor='Intel Inc.',
            renderer='Intel Iris OpenGL Engine',
            fix_hairline=True,
        )
        logger.info('Selenium stealth applied')

    driver.execute_cdp_cmd(
        'Page.addScriptToEvaluateOnNewDocument',
        {'source': 'Object.defineProperty(navigator,"webdriver",'
                    '{get:()=>undefined})'},
    )

    driver.execute_cdp_cmd(
        'Page.setDownloadBehavior',
        {'behavior': 'allow', 'downloadPath': abs_dl},
    )

    logger.info(f'Chrome driver ready -- download dir: {abs_dl}')
    return driver


# ─────────────────────────────────────────────────────────────────────────────
# Determine which years to download
# ─────────────────────────────────────────────────────────────────────────────

def _get_years_needed():
    """
    Determine which year Excel files to download based on master CSV state.

    If config.TARGET_YEARS is set, use those.
    Otherwise, read the master CSV and find the last quarter with data,
    then return years from that year onward (to fill missing quarters).
    """
    import pandas as pd

    if config.TARGET_YEARS is not None:
        logger.info(f'Using configured TARGET_YEARS: {config.TARGET_YEARS}')
        return sorted(config.TARGET_YEARS)

    if not os.path.exists(config.MASTER_FILE):
        logger.info('No master file found — will download all available years')
        return None  # signal to download everything available

    try:
        import csv as _csv
        # Detect description sub-header row and skip it
        skiprows = []
        with open(config.MASTER_FILE, 'r', newline='', encoding='utf-8') as f:
            reader = _csv.reader(f)
            _header = next(reader, None)
            second = next(reader, None)
            if second and len(second) > 1 and ':' in str(second[1]):
                skiprows = [1]

        master_df = pd.read_csv(config.MASTER_FILE, index_col=0, skiprows=skiprows)
    except Exception as e:
        logger.warning(f'Could not read master: {e}')
        return None

    if master_df.empty:
        return None

    # Find quarters with actual data (not all-empty rows)
    quarters_with_data = []
    for idx in master_df.index:
        row = master_df.loc[idx]
        if row.notna().any() and (row != '').any():
            quarters_with_data.append(str(idx))

    if not quarters_with_data:
        return None

    # Parse the last quarter with data: "2024-Q3" -> year=2024, q=3
    last_q = quarters_with_data[-1]
    match = re.match(r'(\d{4})-Q(\d)', last_q)
    if not match:
        logger.warning(f'Could not parse last quarter: {last_q}')
        return None

    last_year = int(match.group(1))
    last_qnum = int(match.group(2))

    # Need to download from last_year (to get remaining quarters) onward
    current_year = time.localtime().tm_year
    years = list(range(last_year, current_year + 1))

    logger.info(f'Master last data: {last_q} — will download years: {years}')
    return years


# ─────────────────────────────────────────────────────────────────────────────
# Find download links on the page
# ─────────────────────────────────────────────────────────────────────────────

def _find_download_links(driver):
    """
    Parse the page to find all yearly Excel download links.
    Returns dict: {year: full_url, ...}
    """
    from selenium.webdriver.common.by import By

    logger.info('Scanning page for download links...')

    # Wait for the content area to load
    time.sleep(3)

    # Scroll down to the content section where links are
    driver.execute_script('window.scrollTo(0, document.body.scrollHeight / 2);')
    _human_delay(1.0, 2.0)

    links = driver.find_elements(By.CSS_SELECTOR, '.article-content a[href]')
    download_map = {}

    for link in links:
        href = link.get_attribute('href') or ''
        text = link.text.strip()

        # Match links like "Fund saving by category 2024" or "Fund savings by category 2023"
        # Extract year from link text
        year_match = re.search(r'\b(20\d{2})\b', text)
        if not year_match:
            continue

        # Skip aggregate files like "2000-2025"
        if re.search(r'\d{4}-\d{4}', text):
            continue

        year = int(year_match.group(1))

        # Ensure it's an Excel file link
        if not (href.endswith('.xlsx') or href.endswith('.xls')):
            continue

        # Build full URL if relative
        if href.startswith('/'):
            href = config.BASE_DOMAIN + href

        download_map[year] = href
        logger.debug(f'Found: {year} -> {href}')

    logger.info(f'Found {len(download_map)} download links: '
                f'{sorted(download_map.keys())}')
    return download_map


# ─────────────────────────────────────────────────────────────────────────────
# Download a single file
# ─────────────────────────────────────────────────────────────────────────────

def _download_file(driver, url, download_dir, year):
    """
    Download a file via Selenium (navigating to the direct link).
    Returns the path to the downloaded file.
    """
    import requests

    logger.info(f'Downloading {year} from: {url}')

    # Use requests with the same cookies from Selenium for reliable download
    cookies = {c['name']: c['value'] for c in driver.get_cookies()}
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                       'AppleWebKit/537.36 (KHTML, like Gecko) '
                       'Chrome/131.0.0.0 Safari/537.36',
        'Referer': config.BASE_URL,
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,'
                  'application/vnd.ms-excel,*/*',
    }

    year_dir = os.path.join(download_dir, str(year))
    os.makedirs(year_dir, exist_ok=True)

    # Determine filename from URL
    filename = url.split('/')[-1]
    filepath = os.path.join(year_dir, filename)

    for attempt in range(config.MAX_DOWNLOAD_RETRIES):
        try:
            resp = requests.get(url, cookies=cookies, headers=headers,
                                timeout=60, stream=True)
            resp.raise_for_status()

            with open(filepath, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)

            file_size = os.path.getsize(filepath)
            if file_size < 1000:
                logger.warning(f'File too small ({file_size} bytes), '
                               f'retry {attempt + 1}')
                time.sleep(config.RETRY_DELAY)
                continue

            logger.info(f'Downloaded: {filename} ({file_size:,} bytes)')
            return filepath

        except Exception as e:
            logger.warning(f'Download attempt {attempt + 1} failed: {e}')
            time.sleep(config.RETRY_DELAY)

    raise RuntimeError(f'Failed to download {url} after '
                       f'{config.MAX_DOWNLOAD_RETRIES} attempts')


# ─────────────────────────────────────────────────────────────────────────────
# Main download function
# ─────────────────────────────────────────────────────────────────────────────

def download():
    """
    Full scraping pipeline:
    1. Determine which years are needed
    2. Launch stealth Chrome
    3. Navigate to fondbolagen.se statistics page
    4. Find download links
    5. Download needed year files

    Returns list of (year, filepath) tuples, or empty list if nothing to do.
    """
    download_dir = config.DOWNLOAD_RUN_DIR
    os.makedirs(download_dir, exist_ok=True)

    years_needed = _get_years_needed()

    driver = None
    try:
        driver = _build_driver(download_dir)

        # Navigate to the page
        logger.info(f'Loading page: {config.BASE_URL}')
        driver.get(config.BASE_URL)
        _human_delay(3.0, 5.0)

        # Find all download links
        link_map = _find_download_links(driver)

        if not link_map:
            logger.error('No download links found on page')
            return []

        # Filter to needed years
        if years_needed is not None:
            to_download = {y: link_map[y] for y in years_needed if y in link_map}
            missing = [y for y in years_needed if y not in link_map]
            if missing:
                logger.warning(f'No download links found for years: {missing}')
        else:
            to_download = link_map

        if not to_download:
            logger.info('No files to download')
            return []

        logger.info(f'Downloading {len(to_download)} files: '
                    f'{sorted(to_download.keys())}')

        # Download each file
        results = []
        for year in sorted(to_download.keys()):
            url = to_download[year]
            try:
                filepath = _download_file(driver, url, download_dir, year)
                results.append((year, filepath))
            except Exception as e:
                logger.error(f'Failed to download {year}: {e}')

        logger.info(f'Downloaded {len(results)} files')
        return results

    finally:
        if driver:
            try:
                driver.quit()
                logger.info('Browser closed')
            except Exception:
                pass
