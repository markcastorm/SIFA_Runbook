# config.py
# SIFA -- Swedish Investment Fund Association: Fund Savings by Category (Quarterly)
# All constants, paths, column mappings, and settings

import os
from datetime import datetime

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR  = os.path.join(BASE_DIR, 'downloads')
OUTPUT_DIR    = os.path.join(BASE_DIR, 'output')
MASTER_DIR    = os.path.join(BASE_DIR, 'Master_Data')
MASTER_FILE   = os.path.join(MASTER_DIR, 'Master_SIFA_DATA.csv')

# ── Timestamped folders ──────────────────────────────────────────────────────
RUN_TIMESTAMP = datetime.now().strftime('%Y%m%d_%H%M%S')

DOWNLOAD_RUN_DIR  = os.path.join(DOWNLOAD_DIR, RUN_TIMESTAMP)
OUTPUT_RUN_DIR    = os.path.join(OUTPUT_DIR, RUN_TIMESTAMP)
LATEST_OUTPUT_DIR = os.path.join(OUTPUT_DIR, 'latest')

# ── Source ────────────────────────────────────────────────────────────────────
BASE_URL = 'https://www.fondbolagen.se/en/Facts_Indices/fund-savings-by-category-quarterly-statistics/'
BASE_DOMAIN = 'https://www.fondbolagen.se'

PROVIDER_NAME = 'SIFA'
DATASET_NAME  = 'SIFA'

# ── Browser ───────────────────────────────────────────────────────────────────
HEADLESS_MODE       = True
WAIT_TIMEOUT        = 60
PAGE_LOAD_DELAY     = 5
DOWNLOAD_WAIT_TIME  = 120

# ── Download settings ─────────────────────────────────────────────────────────
MAX_DOWNLOAD_RETRIES = 3
RETRY_DELAY          = 3.0

# ── Target years ──────────────────────────────────────────────────────────────
# Set to a list of years e.g. [2024, 2025] to download specific years,
# or None to auto-detect based on master CSV gaps.
TARGET_YEARS = None

# ── Column visibility gate ────────────────────────────────────────────────────
# When True, the extractor checks each source Excel column's hidden flag before
# extracting. The provider uses column hiding as a publication signal — a column
# is hidden until that data is officially released. Set to False to extract all
# columns regardless of visibility (legacy behaviour).
SKIP_HIDDEN_COLUMNS = True

# ── Output filenames ─────────────────────────────────────────────────────────
DATA_FILE_PATTERN = 'SIFA_DATA_{timestamp}.xlsx'
META_FILE_PATTERN = 'SIFA_META_{timestamp}.xlsx'
ZIP_FILE_PATTERN  = 'SIFA_{timestamp}.zip'

# ── Fund type sections (order matters — matches output column order) ─────────
# Each entry: (fund_type_code, display_name, category_codes_for_nominee, category_code_for_others)
# The nominee/others codes differ between ALLTYPES and other sections.

FUND_TYPES = [
    'ALLTYPES',
    'EQUITYFUND',
    'BALANCFUND',
    'LONGTERM',
    'SHORTTERM',
    'HEDGEFUND',
    'OTHERFUND',
]

FUND_TYPE_DISPLAY = {
    'ALLTYPES':   'All types of funds',
    'EQUITYFUND': 'Equity funds',
    'BALANCFUND': 'Balanced funds',
    'LONGTERM':   'Long term fixed income funds',
    'SHORTTERM':  'Short term fixed income funds',
    'HEDGEFUND':  'Hedge funds',
    'OTHERFUND':  'Other funds',
}

# ── Category labels in source Excel (order matches data rows) ────────────────
CATEGORY_LABELS = [
    'Swedish households, direct inv.',
    'ISK (Investment Savings Account)',
    'IPS (Individual Pension Saving)',
    'Unit linked',
    'Premium Pension Savings',
    # Nominee label varies: "Unallocated Nominee Accounts" or "Nominee Accounts"
    None,  # matched dynamically (contains "Nominee")
    'Non profit institutions serving households',
    'Swedish corporations',
    'Others',
    'TOTAL',
]

# ── Metrics per section (5 groups of 10 = 50 cols per fund type) ─────────────
METRICS = ['NETSAVING', 'NETSAVINGSUM', 'NETSAVINGPERC', 'NETASSET', 'NETASSETPERC']

# ── SIMBA codes: ABSOLUTE column ordering (350 columns) ──────────────────────
# These are the exact SIMBA codes from the master CSV header row 1.
# The ordering is immutable.

# --- ALL TYPES: uses legacy codes for Net Savings and Net Assets ---
_ALLTYPES_NETSAVING = [
    'SWEDISHHOUSEHOLDSDIRECTINV.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'ISK.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'IPS.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'UNITLINKED.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'PREMIUMPENSION.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'NOMINEEACCOUNTS.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'NPISH.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'SWEDISHCORPORATIONS.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'OTHERS.SAVINGS.FLOW.NONE.Q.1@SIFA',
    'SWEPENFND.ALLTYPES.TOTAL.NETSAVING.Q',
]

_ALLTYPES_CATEGORIES = [
    'SWEHOUSEHOLD', 'ISK', 'IPS', 'UNITLINK', 'PREMPENSSAVING',
    'NOMINEEACC', 'NONPROFINST', 'SWECORP', 'OTHERS', 'TOTAL',
]

_ALLTYPES_NETSAVINGSUM  = [f'SWEPENFND.ALLTYPES.{c}.NETSAVINGSUM.Q'  for c in _ALLTYPES_CATEGORIES]
_ALLTYPES_NETSAVINGPERC = [f'SWEPENFND.ALLTYPES.{c}.NETSAVINGPERC.Q' for c in _ALLTYPES_CATEGORIES]

_ALLTYPES_NETASSET = [
    'SWEDISHHOUSEHOLDSDIRECTINV.NETASSETS.FLOW.NONE.Q.1@SIFA',
    'ISK.TOTAL.FLOW.NONE.Q.1@SIFA',
    'IPS.TOTAL.FLOW.NONE.Q.1@SIFA',
    'SWEUNITLINKED.TOTAL.FLOW.NONE.Q.1@SIFA',
    'PREMIUMPENSION.TOTAL.FLOW.NONE.Q.1@SIFA',
    'NOMINEEACCOUNTS.TOTAL.FLOW.NONE.Q.1@SIFA',
    'SWENPISH.TOTAL.FLOW.NONE.Q.1@SIFA',
    'SWECORPORATIONS.TOTAL.FLOW.NONE.Q.1@SIFA',
    'SWEOTHERS.TOTAL.FLOW.NONE.Q.1@SIFA',
    'SWEPENFND.ALLTYPES.TOTAL.NETASSET.Q',
]

_ALLTYPES_NETASSETPERC = [f'SWEPENFND.ALLTYPES.{c}.NETASSETPERC.Q' for c in _ALLTYPES_CATEGORIES]

# --- Other fund types: consistent SWEPENFND pattern ---
# Note: category codes differ from ALLTYPES (NOMIEEACC not NOMINEEACC, OTHER not OTHERS)
_OTHER_CATEGORIES = [
    'SWEHOUSEHOLD', 'ISK', 'IPS', 'UNITLINK', 'PREMPENSSAVING',
    'NOMIEEACC', 'NONPROFINST', 'SWECORP', 'OTHER', 'TOTAL',
]


def _build_fund_type_codes(fund_type):
    """Build 50 SIMBA codes for a non-ALLTYPES fund type section."""
    cats = _OTHER_CATEGORIES
    codes = []
    codes += [f'SWEPENFND.{fund_type}.{c}.NETSAVING.Q'     for c in cats]
    codes += [f'SWEPENFND.{fund_type}.{c}.NETSAVINGSUM.Q'  for c in cats]
    codes += [f'SWEPENFND.{fund_type}.{c}.NETSAVINGPERC.Q' for c in cats]
    codes += [f'SWEPENFND.{fund_type}.{c}.NETASSET.Q'      for c in cats]
    codes += [f'SWEPENFND.{fund_type}.{c}.NETASSETPERC.Q'  for c in cats]
    return codes


# Build the complete ordered list of 350 SIMBA codes
DATA_COLUMNS = (
    _ALLTYPES_NETSAVING + _ALLTYPES_NETSAVINGSUM + _ALLTYPES_NETSAVINGPERC +
    _ALLTYPES_NETASSET + _ALLTYPES_NETASSETPERC +
    _build_fund_type_codes('EQUITYFUND') +
    _build_fund_type_codes('BALANCFUND') +
    _build_fund_type_codes('LONGTERM') +
    _build_fund_type_codes('SHORTTERM') +
    _build_fund_type_codes('HEDGEFUND') +
    _build_fund_type_codes('OTHERFUND')
)

# ── Human-readable descriptions (DATA row 2) ─────────────────────────────────
_CATEGORY_DESCRIPTIONS = [
    'Swedish households, direct inv.',
    'ISK (Investment Savings Account)',
    'IPS (Individual Pension Saving)',
    'Unit linked',
    'Premium Pension Savings',
    'Nominee Accounts',
    'Non profit institutions serving households',
    'Swedish corporations',
    'Others',
    'TOTAL',
]

_METRIC_DESCRIPTIONS = {
    'NETSAVING':     'Net savings',
    'NETSAVINGSUM':  'Net savings sum',
    'NETSAVINGPERC': 'Net savings %',
    'NETASSET':      'Net assets',
    'NETASSETPERC':  'Net assets %',
}


def _build_descriptions():
    """Build 350 human-readable descriptions matching DATA_COLUMNS order."""
    descs = []
    for fund_type in FUND_TYPES:
        fund_name = FUND_TYPE_DISPLAY[fund_type]
        for metric_key in METRICS:
            metric_name = _METRIC_DESCRIPTIONS[metric_key]
            for cat_name in _CATEGORY_DESCRIPTIONS:
                descs.append(f'{fund_name}: {cat_name}: {metric_name}')
    return descs


DATA_DESCRIPTIONS = _build_descriptions()

# ── Source Excel column mapping ──────────────────────────────────────────────
# Columns in the source Excel per section:
#   B=Q1, C=Q2, D=Q3, E=Q4 (net savings per quarter)
#   F=Net savings sum (cumulative), G=Net savings %
#   H=Net assets (period-end), I=Net assets %
EXCEL_QUARTER_COLS = {1: 2, 2: 3, 3: 4, 4: 5}  # quarter_num -> Excel col index (1-based)
EXCEL_SUM_COL  = 6
EXCEL_PCT_COL  = 7
EXCEL_ASSET_COL = 8
EXCEL_ASSET_PCT_COL = 9

# ── Metadata columns ─────────────────────────────────────────────────────────
METADATA_COLUMNS = [
    'CODE',
    'CODE_MNEMONIC',
    'DESCRIPTION',
    'FREQUENCY',
    'MULTIPLIER',
    'AGGREGATION_TYPE',
    'UNIT_TYPE',
    'DATA_TYPE',
    'DATA_UNIT',
    'SEASONALLY_ADJUSTED',
    'ANNUALIZED',
    'PROVIDER_MEASURE_URL',
    'PROVIDER',
    'SOURCE',
    'SOURCE_DESCRIPTION',
    'COUNTRY',
    'DATASET',
]

METADATA_DEFAULTS = {
    'FREQUENCY':            'Q',
    'MULTIPLIER':           0.0,
    'AGGREGATION_TYPE':     'END_OF_PERIOD',
    'UNIT_TYPE':            'LEVEL',
    'DATA_TYPE':            'AMOUNT',
    'DATA_UNIT':            'MSEK',
    'SEASONALLY_ADJUSTED':  'N',
    'ANNUALIZED':           'N',
    'PROVIDER_MEASURE_URL': BASE_URL,
    'PROVIDER':             PROVIDER_NAME,
    'SOURCE':               PROVIDER_NAME,
    'SOURCE_DESCRIPTION':   'Swedish Investment Fund Association - Fund Savings by Category (Quarterly)',
    'COUNTRY':              'SE',
    'DATASET':              DATASET_NAME,
}
