# SIFA Runbook - Claude Context

## What This Project Is

SIFA = **Swedish Investment Fund Association**. This pipeline automates the quarterly extraction of "Fund Savings by Category" data from [fondbolagen.se](https://www.fondbolagen.se/en/Facts_Indices/fund-savings-by-category-quarterly-statistics/), transforms it into SIMBA-standard format (350 columns), and maintains a cumulative master CSV.

The output format follows an internal SIMBA standard used across multiple runbooks (CPFX, SIFA, etc.). Each runbook is a standalone pipeline for a different data provider.

---

## Architecture Overview

```
main.py                  Entry point (calls orchestrator.main())
  orchestrator.py        Wires the 3-step pipeline
    scraper.py           Step 1: Selenium stealth Chrome -> download Excel files
    extractor.py         Step 2: Parse Excel -> extract quarterly data
    file_generator.py    Step 3: Generate DATA/META/ZIP + update master CSV
  config.py              All constants, paths, SIMBA codes, column mappings

compare/                 Post-run comparison tool (standalone)
  compare.py             Compares pipeline output vs reference file
  config_compare.py      Tolerances and display settings
  reference_input/       Drop reference .xlsx/.csv here
  run_compare.bat        Runner

run.bat                  Pipeline runner
```

---

## Pipeline Flow

```
1. SCRAPER (scraper.py)
   - Reads master CSV to find which years need data (auto-detect gaps)
   - Launches headless Chrome with selenium-stealth
   - Navigates to fondbolagen.se quarterly statistics page
   - Finds year-specific Excel download links
   - Downloads needed year files via requests (using Selenium cookies)
   - Output: downloads/{timestamp}/{year}/fund-saving-by-category-{year}.xlsx

2. EXTRACTOR (extractor.py)
   - Opens each Excel with openpyxl (data_only=True)
   - COPY METHOD: copies all cell values to a fresh sheet (_copy_values_sheet)
   - Dynamically finds 7 fund type sections by scanning column A
   - Detects which quarters have data (skips zero-filled future quarters)
   - Extracts values with :.15g formatting (15 significant digits, matches Excel)
   - Output: dict mapping 'YYYY-QN' -> {simba_code: value, ...}

3. FILE GENERATOR (file_generator.py)
   - Updates master CSV (cumulative: append new quarters, fill gaps, never overwrite)
   - Creates DATA.xlsx (new quarters only), META.xlsx, ZIP
   - Creates SIFA_DATA.csv (full master data with all quarters)
   - Copies everything to output/latest/
```

---

## Critical Data Structure: 350 Columns

The output has exactly **350 columns** = 7 fund types x 50 columns each.

Each fund type block has 50 columns = 5 metrics x 10 categories:

```
Metrics (5):
  NETSAVING      Net savings for the quarter (positions 0-9)
  NETSAVINGSUM   Net savings cumulative sum  (positions 10-19)
  NETSAVINGPERC  Net savings percentage      (positions 20-29)
  NETASSET       Net assets end-of-period    (positions 30-39)
  NETASSETPERC   Net assets percentage       (positions 40-49)

Categories (10):
  Swedish households, ISK, IPS, Unit linked, Premium Pension,
  Nominee Accounts, Non profit inst., Swedish corporations, Others, TOTAL

Fund Types (7):
  ALLTYPES, EQUITYFUND, BALANCFUND, LONGTERM, SHORTTERM, HEDGEFUND, OTHERFUND
```

### Last-Quarter Rule

- For **non-last quarters** (e.g., Q1, Q2 when Q4 exists): only NETSAVING (10 cols per section, 70 total) is populated
- For the **last quarter with data**: all 5 metrics are populated (50 cols per section, 350 total)
- This is because sum/pct/assets columns in the source Excel are cumulative and only meaningful at the last reported quarter

### SIMBA Code Naming Quirks

- **ALLTYPES section uses legacy codes** for NETSAVING and NETASSET (e.g., `SWEDISHHOUSEHOLDSDIRECTINV.SAVINGS.FLOW.NONE.Q.1@SIFA`)
- **All other sections** use consistent `SWEPENFND.{FUNDTYPE}.{CATEGORY}.{METRIC}.Q` pattern
- **Category code spelling differs**: ALLTYPES uses `NOMINEEACC` and `OTHERS`; other sections use `NOMIEEACC` and `OTHER` (yes, `NOMIEEACC` is a typo in the original data -- do NOT fix it)

---

## Source Excel Layout

Each downloaded Excel file has one sheet named by year (e.g., "2024"). Layout per fund type section:

```
Row N:   Section header    "All types of funds"
Row N+1: Column headers    Quarter 1 | Quarter 2 | Quarter 3 | Quarter 4 | Net savings | Net savings | Net assets | Net assets
Row N+2: Sub-headers       (empty)   | (empty)   | (empty)   | (empty)   | sum         | %           | <date>     | %
Row N+3: Category 1        B=Q1val   | C=Q2val   | D=Q3val   | E=Q4val   | F=Sum       | G=%         | H=Assets   | I=Assets%
...      (9 categories)
Row N+12: TOTAL

Excel column mapping:
  Col B (2) = Q1 net savings
  Col C (3) = Q2 net savings
  Col D (4) = Q3 net savings
  Col E (5) = Q4 net savings
  Col F (6) = Net savings sum (cumulative)
  Col G (7) = Net savings %
  Col H (8) = Net assets
  Col I (9) = Net assets %
```

---

## Key Technical Decisions

### Copy Method for Value Extraction

**Problem**: openpyxl's `data_only=True` reads cached formula results as raw IEEE 754 doubles with FP noise (e.g., `-1144.415502230001` instead of `-1144.41550223`).

**Previous approach (REMOVED)**: `_clean_float()` tried to round away FP noise using a 1e-9 relative error threshold. This **destroyed real precision** -- truncated values like `8442.24009065002` down to `8442.24009`.

**Current approach**:
1. Load with `data_only=True`
2. Copy all values to a fresh sheet (`_copy_values_sheet`)
3. Format with `f'{val:.15g}'` -- 15 significant digits, exactly matching Excel's formula bar display
4. This strips FP tail noise while preserving all meaningful digits

### Cumulative Master CSV

The master CSV (`Master_Data/Master_SIFA_DATA.csv`) is **append-only and cumulative**:
- New quarters are appended
- Existing quarters with empty cells get those cells filled in
- Existing data is **never overwritten or deleted**
- Each run only adds what's missing

The master CSV has a description sub-header row (row 2) with human-readable descriptions like "All types of funds: Swedish households, direct inv.: Net savings". Both `_read_master_csv` and `_write_master_csv` handle this row. The reader detects it by checking if row 2 contains ':' characters.

### Zero-Filled Quarter Detection

Years like 2026 (mid-year) have Q2-Q4 columns filled with **zeros** (not None) -- these are formula placeholders for future data. The `_detect_quarters_with_data()` function checks for non-zero values in non-TOTAL rows to correctly identify which quarters actually have data.

### Nominee Account Label Variation

The category label for nominee accounts varies between files: "Unallocated Nominee Accounts" vs "Nominee Accounts". The extractor doesn't rely on exact label matching -- it reads 10 consecutive rows starting from the first data row (found dynamically by skipping headers containing "Quarter" or "Net").

---

## Output File Structure

```
output/
  {timestamp}/
    SIFA_DATA_{timestamp}.xlsx    New quarters only (DATA format)
    SIFA_META_{timestamp}.xlsx    Metadata for all 350 SIMBA codes
    SIFA_{timestamp}.zip          Bundle of DATA + META
    SIFA_DATA.csv                 Full master data (ALL quarters, with description row)
  latest/
    SIFA_DATA_latest.xlsx         Copy of latest run's DATA xlsx
    SIFA_META_latest.xlsx         Copy of latest run's META xlsx
    SIFA_latest.zip               Copy of latest run's ZIP
    SIFA_DATA.csv                 Copy of latest run's full CSV

Master_Data/
  Master_SIFA_DATA.csv            Cumulative master (all quarters ever processed)

downloads/
  {timestamp}/
    {year}/
      fund-saving-by-category-{year}.xlsx
```

### DATA xlsx vs CSV distinction

- **DATA xlsx**: Contains only the NEW quarters from the current run
- **SIFA_DATA.csv**: Contains ALL quarters (full master snapshot)
- The compare tool uses the CSV for this reason

---

## Compare Tool (compare/)

Standalone comparison tool. Compares pipeline output against a manually-provided reference file.

### Usage
1. Drop reference `.xlsx` or `.csv` into `compare/reference_input/`
2. Run `compare/run_compare.bat` (or `python compare/compare.py`)
3. Results appear in `compare/run_{timestamp}/`

### What it does
- Auto-discovers pipeline output from `output/latest/SIFA_DATA.csv` (prefers CSV over xlsx)
- Auto-discovers source downloads for verification
- Compares overlapping quarters cell-by-cell
- For mismatches, verifies against source Excel to determine which side is correct
- Generates:
  - `report_mismatches.csv` -- value conflicts
  - `report_missing.csv` -- in reference but absent from pipeline
  - `report_new_values.csv` -- in pipeline but absent from reference
  - `ANNOTATED_*.xlsx` -- color-coded (red=mismatch, yellow=missing, green=new)
  - `comparison_summary.txt` -- full text report
  - `files/` subfolders with copied source Excel files

---

## Important Functions Reference

### config.py
- `DATA_COLUMNS` -- List of 350 SIMBA codes in exact output order (immutable)
- `DATA_DESCRIPTIONS` -- 350 human-readable descriptions matching DATA_COLUMNS
- `_build_fund_type_codes(fund_type)` -- Generates 50 SIMBA codes for a fund type
- `FUND_TYPES` -- 7 fund type codes in order
- `EXCEL_QUARTER_COLS` -- Maps quarter number to Excel column index {1:2, 2:3, 3:4, 4:5}
- `TARGET_YEARS` -- Set to list of years, or None for auto-detect

### scraper.py
- `download()` -- Full scraping pipeline. Returns list of (year, filepath) tuples
- `_get_years_needed()` -- Reads master CSV to determine which years to download
- `_find_download_links(driver)` -- Parses page for yearly Excel download links
- `_download_file(driver, url, download_dir, year)` -- Downloads via requests with Selenium cookies
- `_build_driver(download_dir)` -- Creates stealth Chrome WebDriver

### extractor.py
- `extract(excel_path, year)` -- Main extraction for one year file. Returns dict of quarter data
- `extract_multiple(file_list)` -- Extracts from multiple year files
- `_copy_values_sheet(wb)` -- Copy method: copies all cell values to fresh sheet
- `_get_cell_value(ws, row, col)` -- Reads cell, formats with :.15g, returns float or None
- `_find_sections(ws)` -- Dynamically locates 7 fund type sections in worksheet
- `_find_category_rows(ws, data_start_row)` -- Finds 10 category rows (9 + TOTAL)
- `_detect_quarters_with_data(ws, data_rows)` -- Checks which Q1-Q4 columns have non-zero data
- `_extract_section_data(ws, section, quarter_num, is_last_quarter)` -- Extracts 50 values for one section

### file_generator.py
- `FileGenerator.generate_files(new_data_dict, output_dir)` -- Orchestrates all output generation
- `FileGenerator.update_master(new_data_dict)` -- Cumulative merge into master CSV
- `FileGenerator.create_data_file(df, output_path)` -- Creates DATA xlsx
- `FileGenerator.create_meta_file(output_path)` -- Creates META xlsx
- `FileGenerator._read_master_csv(path)` -- Reads master CSV, auto-skips description row
- `FileGenerator._write_master_csv(df, path)` -- Writes master CSV with description row

### compare/compare.py
- `find_pipeline_output(project_root)` -- Auto-discovers CSV (preferred) or xlsx
- `find_reference_file(compare_dir)` -- Finds reference in reference_input/
- `ComparisonReport.perform_comparison()` -- Cell-by-cell comparison of overlapping quarters
- `ComparisonReport.verify_mismatches()` -- Checks mismatches against source Excel
- `ComparisonReport.create_annotated_excel()` -- Color-coded Excel output
- `ComparisonReport.run()` -- Full comparison pipeline

---

## Dependencies

```
openpyxl          Excel read/write (data_only=True for formula results)
pandas            DataFrame operations for master CSV
selenium          Web scraping (headless Chrome)
selenium-stealth  Anti-detection for Selenium
requests          Direct file downloads
```

---

## Known Constraints / Gotchas

1. **SIMBA code typo `NOMIEEACC`**: This is intentional and must match the existing master data. Do NOT correct it to `NOMINEEACC`.

2. **ALLTYPES uses different SIMBA code patterns** than the other 6 fund types. ALLTYPES has legacy `*.SAVINGS.FLOW.NONE.Q.1@SIFA` codes for NETSAVING and NETASSET, while others use `SWEPENFND.*.*.NETSAVING.Q`.

3. **The DATA xlsx only has new quarters**; the CSV has all quarters. The compare tool must use the CSV.

4. **Master CSV has a description sub-header row** (row 2). All readers must detect and skip it. Detection: check if row 2 first data cell contains ':'.

5. **Zero-filled future quarters** (e.g., 2026 Q2-Q4 with all zeros) must be excluded. The extractor checks for at least one non-zero value in non-TOTAL rows.

6. **Float precision**: All values use `:.15g` formatting (15 significant digits) to match Excel's formula bar display. Never round or clean beyond this.

7. **openpyxl has 350 columns** (exceeds .xls 256-col limit), so we use .xlsx (openpyxl), not .xls (xlwt).

8. **Selenium stealth**: The fondbolagen.se site requires stealth Chrome. The scraper uses random human delays and anti-detection measures.

---

## Verified Test Results (2024-2026)

- 2024 Q1-Q3 vs sample reference (SIFA_DATA_20250212.xlsx): **210 matches, 0 mismatches**
- 2024 extraction: 4 quarters (Q1-Q4), 70 cols for Q1-Q3, 350 for Q4
- 2025 extraction: 4 quarters (Q1-Q4), same pattern
- 2026 extraction: 1 quarter (Q1 only), 290 cols (last quarter with sum/assets)
- Total: 9 quarters, all verified accurate

Sample verified values (2026-Q1):
- `-1144.41550223` (Swedish households net saving)
- `8442.24009065002` (ISK net saving -- exact match)
- `-1728.83321435` (IPS net saving)
