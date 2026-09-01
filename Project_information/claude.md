# SIFA Runbook - Claude Context

## What This Project Is

SIFA = **Swedish Investment Fund Association**. This pipeline automates the quarterly extraction of "Fund Savings by Category" data from [fondbolagen.se](https://www.fondbolagen.se/en/Facts_Indices/fund-savings-by-category-quarterly-statistics/), transforms it into SIMBA-standard format (350 columns), and maintains a cumulative master CSV.

The output format follows an internal SIMBA standard used across multiple runbooks (CPFX, SIFA, etc.). Each runbook is a standalone pipeline for a different data provider.

**Pipeline purpose: pure extraction only.** The pipeline reads raw values from the source Excel. It does NOT compute, derive, or transform values. If a column is formula-driven in the source, it is treated according to the column visibility rule (see below).

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
   - Reads column visibility from the ORIGINAL sheet BEFORE copying
   - COPY METHOD: copies all cell values to a fresh sheet (_copy_values_sheet)
   - Dynamically finds 7 fund type sections by scanning column A
   - Checks which columns are hidden (provider publication signal)
   - Detects which quarters have data (hidden col check + zero-fill check)
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
  NETSAVING      Net savings for the quarter        (positions 0-9)
  NETSAVINGSUM   Net savings cumulative sum          (positions 10-19)
  NETSAVINGPERC  Net savings % (year-to-date)       (positions 20-29)
  NETASSET       Net assets end-of-period            (positions 30-39)
  NETASSETPERC   Net assets %                        (positions 40-49)

Categories (10):
  Swedish households, ISK, IPS, Unit linked, Premium Pension,
  Nominee Accounts, Non profit inst., Swedish corporations, Others, TOTAL

Fund Types (7):
  ALLTYPES, EQUITYFUND, BALANCFUND, LONGTERM, SHORTTERM, HEDGEFUND, OTHERFUND
```

### Last-Quarter Rule

- For **non-last quarters** (e.g., Q1, Q2 when Q2/Q4 exists): only NETSAVING is populated
  (10 cols per section x 7 sections = 70 total columns)
- For the **last quarter with data**: all 5 metrics are populated
  (50 cols per section x 7 sections = 350 total columns, minus any hidden summary cols)
- "Last quarter" = the highest-numbered visible, non-zero quarter in the source file
- Summary columns (F, G, H, I) are only extracted on the last quarter AND only if visible

### SIMBA Code Naming Quirks

- **ALLTYPES section uses legacy codes** for NETSAVING and NETASSET
  (e.g., `SWEDISHHOUSEHOLDSDIRECTINV.SAVINGS.FLOW.NONE.Q.1@SIFA`)
- **All other sections** use consistent `SWEPENFND.{FUNDTYPE}.{CATEGORY}.{METRIC}.Q` pattern
- **Category code spelling differs**: ALLTYPES uses `NOMINEEACC` and `OTHERS`;
  other sections use `NOMIEEACC` and `OTHER`
  (`NOMIEEACC` is a deliberate spelling in the original data — do NOT correct it)

---

## Source Excel Layout

Each downloaded Excel file has one sheet named by year (e.g., "2026"). Layout per fund type section:

```
Row N:    Section header    "All types of funds"
Row N+1:  Column headers    Quarter 1 | Quarter 2 | Quarter 3 | Quarter 4 | Net savings | Net savings | Net assets     | Net assets
Row N+2:  Sub-headers       (empty)   | (empty)   | (empty)   | (empty)   | sum         | %           | <date>         | %
Row N+3:  Category 1        B=Q1val   | C=Q2val   | D=Q3val   | E=Q4val   | F=Sum       | G=%         | H=Assets       | I=Assets%
...       (9 categories)
Row N+12: TOTAL

Excel column mapping (1-based):
  Col B (2) = Q1 net savings          <- raw input value
  Col C (3) = Q2 net savings          <- raw input value
  Col D (4) = Q3 net savings          <- raw input value
  Col E (5) = Q4 net savings          <- raw input value
  Col F (6) = Net savings sum         <- FORMULA: =SUM(B:E) for each row
  Col G (7) = Net savings %           <- FORMULA: =F/TOTAL_F*100
  Col H (8) = Net assets              <- raw for non-ALLTYPES; formula (sum of sections) for ALLTYPES
  Col I (9) = Net assets %            <- FORMULA: =H/TOTAL_H*100
```

### Important: Formula vs Raw Columns

The source Excel has both raw input values and formula-driven computed columns:

| Column | Type | Notes |
|--------|------|-------|
| B–E (Q1–Q4) | **Raw input** | Actual net savings values entered by provider |
| F (Net savings sum) | Formula `=SUM(B:E)` | Cumulative YTD sum |
| G (Net savings %) | Formula `=F/F_TOTAL*100` | YTD % share of each category |
| H (Net assets) | Raw for non-ALLTYPES; Formula for ALLTYPES | ALLTYPES sums the other 6 sections |
| I (Net assets %) | Formula `=H/H_TOTAL*100` | % share of total assets |

**The pipeline extracts these columns when they are VISIBLE** (see column visibility rule below).
When a column is hidden in the source, its data is not yet officially published and must not be extracted.

---

## Column Visibility — Publication Signal (CRITICAL)

**This is the most important rule in the extractor.**

The provider (fondbolagen.se) uses **column hiding** as a deliberate publication signal.
A column is hidden in the source Excel until the data it contains is officially released.
This applies to both quarterly net savings columns (B–E) and summary columns (F, G, H, I).

### How It Works in Practice

**Example: 2026 file downloaded when only Q1 was published (May 2026):**
```
Col B (Q1 net savings):     VISIBLE   — data published, extract
Col C (Q2 net savings):     HIDDEN    — not yet published, skip
Col D (Q3 net savings):     (no dim entry, values = None/0, skipped by zero-check)
Col E (Q4 net savings):     (no dim entry, values = None/0, skipped by zero-check)
Col F (Net savings sum):    HIDDEN    — not yet published, skip
Col G (Net savings %):      HIDDEN    — not yet published, skip
Col H (Net assets):         VISIBLE   — Q1 end-of-period balance (31/3/2026)
Col I (Net assets %):       VISIBLE   — extract
Result: 2026-Q1 gets net savings + net assets (31/3/2026) + net assets %
```

**Example: 2026 file downloaded when Q2 was published (August 2026):**
```
Col B (Q1 net savings):     VISIBLE   — extract
Col C (Q2 net savings):     VISIBLE   — extract
Col D (Q3 net savings):     HIDDEN    — not yet published, skip
Col E (Q4 net savings):     (no dim entry, values = 0 for ALLTYPES, skipped)
Col F (Net savings sum):    VISIBLE   — extract (Q1+Q2 cumulative sum)
Col G (Net savings %):      HIDDEN    — not yet published, skip
Col H (Net assets):         VISIBLE   — Q2 end-of-period balance (30/6/2026)
Col I (Net assets %):       VISIBLE   — extract
Result: 2026-Q1 gets net savings only; 2026-Q2 (last quarter) gets net savings + sum + assets
```

### openpyxl Detection Detail

- Column visibility is read from the **original source sheet** BEFORE the copy step
- The copy method (`_copy_values_sheet`) transfers cell values only — it does NOT copy
  column dimension metadata. Reading visibility from the copied sheet always shows all
  columns as visible (incorrect). Always read from `wb.active` before copying.
- `ws.column_dimensions[letter].hidden == True` → column is hidden
- Columns with **no dimension entry** default to visible. This is correct for most cases,
  but some columns (notably E/Q4 in mid-year files) may appear visible via openpyxl yet
  be visually hidden in Excel. This is handled by the zero-value detection as a fallback.
- A column that is hidden but has cached formula values in the file will still show values
  via `data_only=True`. The visibility check prevents those cached values from being extracted.

### Config Toggle

```python
# config.py
SKIP_HIDDEN_COLUMNS = True   # Set False to revert to legacy behaviour (extracts all columns)
```

Setting `SKIP_HIDDEN_COLUMNS = False` disables the visibility check entirely and extracts
all columns regardless of their hidden state. Use only for debugging or legacy comparison.

---

## Two-Step Pipeline Behaviour (Q1 → Q2 Publication Cycle)

The pipeline is designed to be run repeatedly as the provider publishes new data.
Each run picks up exactly what is new without disturbing what was already captured.

### Verified Test: 2026 Two-Step Simulation

**Step 1 — Run with Q1-only file (fund-saving-by-category-2026.xlsx, May 2026 version):**
```
Hidden cols: [3, 6, 7]  (C=Q2, F=Sum, G=%)
Quarters extracted: [2026-Q1]
2026-Q1 populated: 210 cols  (70 net savings + 70 net assets % + 62 net assets + 8 edge zeros)
Net assets date: 31/3/2026  (end of Q1)
Master: 2026-Q1 added as new row
```

**Step 2 — Run with Q2-published file (fund-saving-by-category-2026.xlsx, Aug 2026 version):**
```
Hidden cols: [4, 7]  (D=Q3, G=%)
Quarters extracted: [2026-Q1, 2026-Q2]
2026-Q1: already complete in master — SKIPPED (70 net savings not re-extracted)
2026-Q2 populated: 280 cols  (70 net savings + 70 net savings sum + 70 net assets % + 62 net assets)
Net assets date: 30/6/2026  (end of Q2)
Master: 2026-Q2 added as new row
```

**Final master comparison vs reference (SIFA_DATA_20250824.xlsx):**
```
2026-Q1: 209/210 matches, 0 missing, 0 extra, 1 precision difference only
         (IPS net savings: manual=-1728.83, output=-1728.83321435 — output is more precise)
2026-Q2: 280/280 matches, 0 differences, 0 missing, 0 extra — PERFECT
```

---

## Key Technical Decisions

### Copy Method for Value Extraction

**Problem**: openpyxl's `data_only=True` reads cached formula results as raw IEEE 754 doubles
with FP noise (e.g., `-1144.415502230001` instead of `-1144.41550223`).

**Previous approach (REMOVED)**: `_clean_float()` tried to round away FP noise using a 1e-9
relative error threshold. This **destroyed real precision** — truncated values like
`8442.24009065002` down to `8442.24009`.

**Current approach**:
1. Load with `data_only=True`
2. Copy all values to a fresh sheet (`_copy_values_sheet`)
3. Format with `f'{val:.15g}'` — 15 significant digits, exactly matching Excel's formula bar display
4. This strips FP tail noise while preserving all meaningful digits

**Critical note on copy + visibility**: The copy step must happen AFTER reading column
visibility. Column dimensions are not copied to the new sheet. Always call
`_get_visible_cols(wb.active)` on the source sheet before `_copy_values_sheet(wb)`.

### Cumulative Master CSV

The master CSV (`Master_Data/Master_SIFA_DATA.csv`) is **append-only and cumulative**:
- New quarters are appended
- Existing quarters with empty cells get those cells filled in
- Existing data is **never overwritten or deleted**
- Each run only adds what's missing

The master CSV has a description sub-header row (row 2) with human-readable descriptions
like "All types of funds: Swedish households, direct inv.: Net savings".
Both `_read_master_csv` and `_write_master_csv` handle this row.
The reader detects it by checking if row 2 contains ':' characters.

### Zero-Filled Quarter Detection

Years like 2026 (mid-year) have future quarter columns filled with **zeros** (not None) for
the ALLTYPES section — these are formula placeholders (`=0` or `=SUM(...)` of empty ranges).
Other fund type sections have `None` in future quarter cells.

`_detect_quarters_with_data()` handles both:
1. **Hidden column check** (primary): if `SKIP_HIDDEN_COLUMNS=True` and a quarter's column
   is hidden, that quarter is excluded regardless of cell values
2. **Zero-value check** (secondary / fallback): even if a column is not flagged as hidden,
   a quarter with all-zero values in non-TOTAL rows is excluded

### Why NETSAVINGPERC Is Often Empty

`NETSAVINGPERC` (col G, Net savings %) is a pure Excel formula: `=F_row / F_TOTAL * 100`.
The provider hides col G until a full publication cycle is complete.
In the 2026 file, col G remains hidden even when Q1 and Q2 are published.
The pipeline correctly produces NaN for NETSAVINGPERC whenever col G is hidden.

Do not be alarmed if NETSAVINGPERC columns are empty for recent quarters —
this is the correct behaviour matching the provider's publication state.

### Nominee Account Label Variation

The category label for nominee accounts varies between files:
- "Unallocated Nominee Accounts" (newer files)
- "Nominee Accounts" (older files)

The extractor doesn't rely on exact label matching. It reads 10 consecutive rows starting
from the first data row found dynamically (by skipping headers containing "Quarter" or "Net").
The Nominee row is always row 6 within the 10-row category block regardless of its label.

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

### DATA xlsx vs CSV Distinction

- **DATA xlsx**: Contains only the NEW quarters from the current run
- **SIFA_DATA.csv**: Contains ALL quarters (full master snapshot)
- The compare tool uses the CSV for this reason

---

## Important Functions Reference

### config.py
- `DATA_COLUMNS` — List of 350 SIMBA codes in exact output order (immutable)
- `DATA_DESCRIPTIONS` — 350 human-readable descriptions matching DATA_COLUMNS
- `_build_fund_type_codes(fund_type)` — Generates 50 SIMBA codes for a fund type
- `FUND_TYPES` — 7 fund type codes in order
- `EXCEL_QUARTER_COLS` — Maps quarter number to Excel column index `{1:2, 2:3, 3:4, 4:5}`
- `EXCEL_SUM_COL = 6` — Net savings sum column index
- `EXCEL_PCT_COL = 7` — Net savings % column index
- `EXCEL_ASSET_COL = 8` — Net assets column index
- `EXCEL_ASSET_PCT_COL = 9` — Net assets % column index
- `TARGET_YEARS` — Set to list of years, or None for auto-detect
- `SKIP_HIDDEN_COLUMNS = True` — Visibility gate; set False for legacy extraction

### scraper.py
- `download()` — Full scraping pipeline. Returns list of `(year, filepath)` tuples
- `_get_years_needed()` — Reads master CSV to determine which years to download
- `_find_download_links(driver)` — Parses page for yearly Excel download links
- `_download_file(driver, url, download_dir, year)` — Downloads via requests with Selenium cookies
- `_build_driver(download_dir)` — Creates stealth Chrome WebDriver

### extractor.py
- `extract(excel_path, year)` — Main extraction for one year file. Returns dict of quarter data
- `extract_multiple(file_list)` — Extracts from multiple year files
- `_get_visible_cols(ws)` — Returns set of visible (non-hidden) 1-based column indices.
  Must be called on the ORIGINAL source sheet before `_copy_values_sheet()`.
  Columns with no dimension entry are treated as visible (Excel default).
- `_copy_values_sheet(wb)` — Copy method: copies all cell values to fresh sheet.
  NOTE: does NOT copy column dimensions — visibility must be read before this call.
- `_get_cell_value(ws, row, col)` — Reads cell, formats with `:.15g`, returns float or None
- `_find_sections(ws)` — Dynamically locates 7 fund type sections in worksheet
- `_find_category_rows(ws, data_start_row)` — Finds 10 category rows (9 categories + TOTAL)
- `_detect_quarters_with_data(ws, data_rows, visible_cols=None)` — Checks which Q1–Q4
  columns have non-zero data AND are visible. Returns sorted list e.g. `[1, 2]`.
- `_extract_section_data(ws, section, quarter_num, is_last_quarter, visible_cols=None)` —
  Extracts up to 50 values for one fund type section. Summary columns (F, G, H, I) are
  only extracted when `is_last_quarter=True` AND the column is in `visible_cols`.

### file_generator.py
- `FileGenerator.generate_files(new_data_dict, output_dir)` — Orchestrates all output generation
- `FileGenerator.update_master(new_data_dict)` — Cumulative merge into master CSV.
  Appends new quarters; fills empty cells in existing quarters; never overwrites non-empty data.
- `FileGenerator.create_data_file(df, output_path)` — Creates DATA xlsx (new quarters only)
- `FileGenerator.create_meta_file(output_path)` — Creates META xlsx (all 350 codes)
- `FileGenerator._read_master_csv(path)` — Reads master CSV, auto-skips description row
- `FileGenerator._write_master_csv(df, path)` — Writes master CSV with description sub-header row

### compare/compare.py
- `find_pipeline_output(project_root)` — Auto-discovers CSV (preferred) or xlsx
- `find_reference_file(compare_dir)` — Finds reference in reference_input/
- `ComparisonReport.perform_comparison()` — Cell-by-cell comparison of overlapping quarters
- `ComparisonReport.verify_mismatches()` — Checks mismatches against source Excel
- `ComparisonReport.create_annotated_excel()` — Color-coded Excel output
- `ComparisonReport.run()` — Full comparison pipeline

---

## Compare Tool (compare/)

Standalone comparison tool. Compares pipeline output against a manually-provided reference file.

### Usage
1. Drop reference `.xlsx` or `.csv` into `compare/reference_input/`
2. Run `compare/run_compare.bat` (or `python compare/compare.py`)
3. Results appear in `compare/run_{timestamp}/`

### What It Does
- Auto-discovers pipeline output from `output/latest/SIFA_DATA.csv` (prefers CSV over xlsx)
- Auto-discovers source downloads for verification
- Compares overlapping quarters cell-by-cell
- For mismatches, verifies against source Excel to determine which side is correct
- Generates:
  - `report_mismatches.csv` — value conflicts
  - `report_missing.csv` — in reference but absent from pipeline
  - `report_new_values.csv` — in pipeline but absent from reference
  - `ANNOTATED_*.xlsx` — color-coded (red=mismatch, yellow=missing, green=new)
  - `comparison_summary.txt` — full text report
  - `files/` subfolders with copied source Excel files

---

## Dependencies

```
openpyxl          Excel read/write (data_only=True for cached formula values)
pandas            DataFrame operations for master CSV
selenium          Web scraping (headless Chrome)
selenium-stealth  Anti-detection for Selenium
requests          Direct file downloads
```

---

## Known Constraints / Gotchas

1. **SIMBA code typo `NOMIEEACC`**: Intentional — matches existing master data.
   Do NOT correct it to `NOMINEEACC`. ALLTYPES uses `NOMINEEACC`; all others use `NOMIEEACC`.

2. **ALLTYPES uses different SIMBA code patterns** than the other 6 fund types.
   ALLTYPES has legacy `*.SAVINGS.FLOW.NONE.Q.1@SIFA` codes for NETSAVING and NETASSET,
   while others use `SWEPENFND.*.*.NETSAVING.Q`.

3. **The DATA xlsx only has new quarters**; the CSV has all quarters.
   The compare tool must use the CSV.

4. **Master CSV has a description sub-header row** (row 2). All readers must detect and skip it.
   Detection: check if row 2 first data cell contains ':'.

5. **Column visibility must be read BEFORE the copy step.**
   `_copy_values_sheet()` creates a new sheet with no column dimensions.
   Reading visibility from the copied sheet returns all columns as visible (wrong).
   Always call `_get_visible_cols(wb.active)` on the source sheet first.

6. **Hidden columns with cached formula values**: openpyxl `data_only=True` reads cached
   formula results even from hidden columns. The visibility gate (`SKIP_HIDDEN_COLUMNS`)
   prevents these from entering the output. Without it, NETSAVINGPERC values appear in
   the output even when col G is hidden (not yet published).

7. **Zero-filled future quarters vs hidden quarters**: Two separate mechanisms:
   - Hidden check: col D (Q3) `hidden=True` → quarter excluded by visibility gate
   - Zero check: col E (Q4) has no dim entry but all zeros → excluded by zero-value detection
   Both are needed because the provider is not perfectly consistent in how it signals
   unpublished data.

8. **Float precision**: All values use `:.15g` formatting (15 significant digits) to match
   Excel's formula bar display. Never round or clean beyond this. The manual reference
   file (SIFA_DATA_20250824.xlsx) stores 2 decimal places in some cells — the pipeline
   output is more precise, which is correct.

9. **openpyxl 350-column limit**: Exceeds .xls 256-col limit, so we use .xlsx (openpyxl),
   not .xls (xlwt).

10. **Selenium stealth**: fondbolagen.se requires stealth Chrome.
    The scraper uses random human delays and anti-detection measures.

11. **Net assets date in header varies**: Col H header shows the end-of-period date for
    the most recently published quarter (e.g., "31/3 2026" for Q1-only, "30/6 2026" for Q2).
    The extractor does not use this date — it is informational only in the source layout.

12. **NETSAVINGSUM and NETSAVINGPERC are cumulative YTD metrics**, not per-quarter.
    They represent the total from Q1 through the last published quarter of the year.
    They only appear in the last-quarter row. When the last quarter advances (e.g., Q1 → Q2),
    the previous quarter's values are NOT removed from the master (append-only rule).
    This is correct: Q1 captured net assets as of 31/3; Q2 captures assets as of 30/6.
    They represent different measurement dates and coexist correctly in the master.

---

## Verified Test Results

### Column Visibility Detection (2026 files)

| File version | Hidden cols detected | Quarters extracted | NETSAVINGPERC |
|---|---|---|---|
| Q1-only (May 2026) | `[3, 6, 7]` — C, F, G | Q1 only | None (correct) |
| Q2-published (Aug 2026) | `[4, 7]` — D, G | Q1, Q2 | None (correct) |
| 2025 full year | `[7]` — G only | Q1, Q2, Q3, Q4 | None (correct) |

### Two-Step Simulation vs Manual Reference (SIFA_DATA_20250824.xlsx)

```
Step 1: Extract Q1-only file → 2026-Q1 written to master (210 populated cols)
Step 2: Extract Q2 file → 2026-Q1 skipped (already complete), 2026-Q2 added (280 cols)

Final comparison:
  2026-Q1: 209/210 matches | 0 missing | 0 extra | 1 precision diff (more precise)
  2026-Q2: 280/280 matches | 0 missing | 0 extra | 0 differences  ← PERFECT
```

### Historical Data (pre-visibility feature)

- 2024 Q1–Q4: extracted from full-year file; Q1–Q3 = 70 cols, Q4 = last quarter = 280+ cols
- 2025 Q1–Q4: same pattern; NETSAVINGPERC in master for Q4 came from old code (col G was
  already hidden in 2025 file — these values are retained as legacy data, not re-extracted
- 2026 Q1–Q2: extracted with new visibility feature; NETSAVINGPERC correctly empty

---

## Debugging Checklist

**NETSAVINGPERC values appearing unexpectedly:**
1. Check `config.SKIP_HIDDEN_COLUMNS` — must be `True`
2. Confirm visibility is read from `wb.active` BEFORE `_copy_values_sheet(wb)` in `extract()`
3. Run: `ws.column_dimensions.get('G')` on the source sheet — should show `hidden=True`
4. If values are in master from a previous run (old code), they came from a hidden col;
   remove the affected quarter rows from master and re-run

**Quarter appearing when it should be empty:**
1. Check if the quarter column is hidden: `ws.column_dimensions.get(letter).hidden`
2. Check if all values are zero (zero-fill placeholder) — `_detect_quarters_with_data` skips these
3. Verify `SKIP_HIDDEN_COLUMNS = True` in config

**Net assets in wrong quarter:**
- Net assets (col H) always reflects the end-of-period date for the LAST visible quarter
- Q1-only file: assets in 2026-Q1 row (31/3/2026)
- Q2-published file: assets in 2026-Q2 row (30/6/2026), Q1 has no assets
- This is correct behaviour — the master retains Q1 assets from the Q1-only run

**Master has stale data for a quarter:**
- The master is append-only; existing non-empty cells are never overwritten
- To re-extract a quarter cleanly: delete its row from `Master_SIFA_DATA.csv` and re-run
- Never modify the master manually except to delete entire rows for re-extraction

**Comparison tool shows differences vs manual reference:**
- Check if the manual was generated with old code (before visibility feature)
- Old code extracted col G regardless of hidden state → NETSAVINGPERC present in manual but not in new output
- This is expected and correct; the new output is more accurate
- The manual file `SIFA_DATA_20250824.xlsx` was generated before the visibility feature;
  use it only as a partial reference for net savings and net assets values
