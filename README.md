# SIFA Runbook

Automated data pipeline for **Swedish Investment Fund Association (SIFA)** quarterly fund savings statistics.

Downloads Excel files from [fondbolagen.se](https://www.fondbolagen.se/en/Facts_Indices/fund-savings-by-category-quarterly-statistics/), extracts quarterly data across 7 fund types and 10 categories, and outputs SIMBA-standard DATA/META/ZIP files with a cumulative master CSV.

---

## Quick Start

```
# Install dependencies
pip install openpyxl pandas selenium selenium-stealth requests

# Run the pipeline
python main.py
# or double-click: run.bat
```

The pipeline will:
1. Check the master CSV for missing quarters
2. Download needed year files from fondbolagen.se
3. Extract and transform data (350 SIMBA columns)
4. Update the master CSV and generate output files

Output appears in `output/{timestamp}/` and `output/latest/`.

---

## Project Structure

```
SIFA_Runbook/
|-- main.py                 Entry point
|-- orchestrator.py         Pipeline controller (download -> extract -> generate)
|-- config.py               All constants, SIMBA codes, column mappings
|-- scraper.py              Selenium scraper for fondbolagen.se
|-- extractor.py            Excel parser and data transformer
|-- file_generator.py       Output file creator + master CSV manager
|-- run.bat                 Windows batch runner
|
|-- compare/                Comparison tool (standalone)
|   |-- compare.py          Compares pipeline output vs reference data
|   |-- config_compare.py   Tolerances and settings
|   |-- reference_input/    Drop reference files here
|   |-- run_compare.bat     Windows batch runner
|
|-- Master_Data/
|   |-- Master_SIFA_DATA.csv   Cumulative master (all quarters)
|
|-- output/
|   |-- latest/                Always-current copies
|   |   |-- SIFA_DATA.csv      Full data (all quarters)
|   |   |-- SIFA_DATA_latest.xlsx
|   |   |-- SIFA_META_latest.xlsx
|   |   |-- SIFA_latest.zip
|   |-- {timestamp}/           Timestamped run output
|
|-- downloads/
|   |-- {timestamp}/{year}/    Downloaded source Excel files
|
|-- Project_information/       Reference/sample files
|-- claude.md                  Full technical context for AI assistants
```

---

## Pipeline Steps

### Step 1: Download (scraper.py)

- Reads the master CSV to determine which years need new data
- Launches headless Chrome with selenium-stealth
- Navigates to the fondbolagen.se quarterly statistics page
- Locates yearly Excel download links
- Downloads via HTTP requests using Selenium session cookies
- Saves to `downloads/{timestamp}/{year}/`

Set `TARGET_YEARS` in config.py to override auto-detection:
```python
TARGET_YEARS = [2024, 2025]  # specific years
TARGET_YEARS = None           # auto-detect (default)
```

### Step 2: Extract (extractor.py)

- Opens each Excel with `data_only=True` (reads cached formula results)
- Uses the **copy method**: copies all values to a fresh sheet for clean reads
- Dynamically locates 7 fund type sections by scanning column A headers
- Detects which quarters have actual data (filters zero-filled placeholders)
- Formats values with `:.15g` (15 significant digits, matches Excel precision)
- Returns structured data: `{'2024-Q1': {simba_code: value, ...}, ...}`

### Step 3: Generate (file_generator.py)

- **Master CSV update**: Cumulative merge -- appends new quarters, fills gaps in existing ones, never overwrites existing data
- **DATA.xlsx**: New quarters from this run (SIMBA codes as headers)
- **META.xlsx**: Metadata for all 350 SIMBA codes
- **ZIP**: Bundle of DATA + META
- **SIFA_DATA.csv**: Full master snapshot (all quarters)
- Copies all files to `output/latest/`

---

## Data Model

**350 columns** = 7 fund types x 5 metrics x 10 categories

| Fund Types | Metrics | Categories |
|---|---|---|
| All types of funds | Net savings | Swedish households |
| Equity funds | Net savings sum | ISK |
| Balanced funds | Net savings % | IPS |
| Long term fixed income | Net assets | Unit linked |
| Short term fixed income | Net assets % | Premium Pension |
| Hedge funds | | Nominee Accounts |
| Other funds | | Non profit inst. |
| | | Swedish corporations |
| | | Others |
| | | TOTAL |

### Last-Quarter Rule

- Non-last quarters: only Net savings columns populated (70 of 350)
- Last quarter with data: all 5 metrics populated (350 of 350)

---

## Compare Tool

Post-run validation tool that compares pipeline output against a reference file.

```
# 1. Drop reference file
copy reference.xlsx compare\reference_input\

# 2. Run comparison
cd compare
python compare.py
# or double-click: run_compare.bat
```

Generates per-run output in `compare/run_{timestamp}/`:
- `report_mismatches.csv` -- value conflicts between pipeline and reference
- `report_missing.csv` -- values in reference but not in pipeline
- `report_new_values.csv` -- values in pipeline but not in reference
- `ANNOTATED_*.xlsx` -- color-coded Excel (red=mismatch, yellow=missing, green=new)
- `comparison_summary.txt` -- full text report
- `files/` -- copies of source Excel files for manual verification

For mismatches, the tool verifies against the original source Excel to determine which side is correct.

---

## Configuration

### config.py (Main Pipeline)

| Setting | Default | Description |
|---|---|---|
| `TARGET_YEARS` | `None` | Years to download (None = auto-detect) |
| `HEADLESS_MODE` | `True` | Run Chrome headless |
| `WAIT_TIMEOUT` | `60` | Selenium wait timeout (seconds) |
| `DOWNLOAD_WAIT_TIME` | `120` | Max download wait (seconds) |
| `MAX_DOWNLOAD_RETRIES` | `3` | Download retry attempts |
| `FLOAT_TOLERANCE` | N/A | Not used; values use :.15g formatting |

### compare/config_compare.py (Compare Tool)

| Setting | Default | Description |
|---|---|---|
| `FLOAT_TOLERANCE` | `0.001` | Tolerance for value comparison |
| `VERBOSE` | `True` | Show detailed progress |
| `MAX_CONSOLE_EXAMPLES` | `5` | Top mismatches to show in summary |

---

## Dependencies

| Package | Purpose |
|---|---|
| `openpyxl` | Excel read/write |
| `pandas` | DataFrame operations |
| `selenium` | Browser automation |
| `selenium-stealth` | Anti-detection |
| `requests` | HTTP file downloads |

Install all:
```
pip install openpyxl pandas selenium selenium-stealth requests
```

---

## Master CSV Format

```
Row 1: (empty) , SIMBA_CODE_1 , SIMBA_CODE_2 , ... (350 codes)
Row 2: (empty) , Description 1 , Description 2 , ... (human-readable)
Row 3: 2024-Q1 , -2846.84     , 10436.47     , ...
Row 4: 2024-Q2 , value        , value        , ...
...
```

The master is cumulative and append-only. Existing data is never modified.
