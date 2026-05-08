"""
SIFA Comparison Tool
=====================
Dynamic path discovery - no manual path entry needed.

What it auto-discovers:
  - Pipeline output:    Latest run in ../output/latest/  (SIFA_DATA_latest.xlsx)
  - Source downloads:   Downloaded Excel files in ../downloads/  (for verification)
  - Reference file:     Any .xlsx/.csv dropped into ./reference_input/

Three separate CSV reports:
    report_mismatches.csv  -- value conflicts (both files have different data)
    report_missing.csv     -- in reference but absent from pipeline
    report_new_values.csv  -- in pipeline but absent from reference

  files/ subfolder with category sub-directories:
    files/mismatch/    -- source Excel files linked to mismatch rows
    files/missing/     -- source Excel files linked to missing rows
    files/new_values/  -- source Excel files linked to new-value rows

Usage:
  1. Drop your reference .xlsx/.csv into the  reference_input/  folder
  2. Run: python compare.py
  3. Find results in the new run_YYYYMMDD_HHMMSS/ folder

Settings (tolerances, verbosity) are in config_compare.py.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font
import csv
import os
import sys
import glob
import shutil
import re
from datetime import datetime
from collections import defaultdict

import pandas as pd

import config_compare as config


# ---------------------------------------------------------------------------
# Path Discovery
# ---------------------------------------------------------------------------

def find_pipeline_output(project_root):
    """
    Find the latest pipeline DATA file in output/latest/.

    Prefers CSV over xlsx because:
      - CSV contains ALL quarters (full master data)
      - xlsx only contains new quarters from the latest run
    """
    latest_dir = os.path.join(project_root, 'output', 'latest')

    if not os.path.isdir(latest_dir):
        raise FileNotFoundError(
            f'No output/latest/ directory found: {latest_dir}\n'
            'Run the pipeline first to generate output.'
        )

    # Prefer CSV (has full master data with all quarters)
    csv_files = glob.glob(os.path.join(latest_dir, 'SIFA_DATA.csv'))
    if csv_files:
        return csv_files[0]

    # Fall back to xlsx (only has new quarters from latest run)
    xlsx_files = glob.glob(os.path.join(latest_dir, 'SIFA_DATA_latest.xlsx'))
    if not xlsx_files:
        xlsx_files = glob.glob(os.path.join(latest_dir, 'SIFA_DATA_*.xlsx'))

    if not xlsx_files:
        raise FileNotFoundError(
            f'No SIFA_DATA.csv or SIFA_DATA_*.xlsx found in: {latest_dir}'
        )

    return xlsx_files[0]


def find_source_downloads(project_root):
    """
    Find downloaded source Excel files in ../downloads/.
    Returns dict: {year: filepath, ...}
    """
    download_root = os.path.join(project_root, 'downloads')
    if not os.path.isdir(download_root):
        return {}

    source_map = {}

    # Walk through all subdirectories to find Excel files
    for root, dirs, files in os.walk(download_root):
        for f in files:
            if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                # Try to extract year from path or filename
                year_match = re.search(r'\b(20\d{2})\b', os.path.join(root, f))
                if year_match:
                    year = int(year_match.group(1))
                    filepath = os.path.join(root, f)
                    # Keep the most recent file for each year
                    if year not in source_map or os.path.getmtime(filepath) > os.path.getmtime(source_map[year]):
                        source_map[year] = filepath

    return source_map


def find_reference_file(compare_dir):
    """Find the reference file in reference_input/."""
    ref_dir = os.path.join(compare_dir, 'reference_input')
    os.makedirs(ref_dir, exist_ok=True)

    # Accept both .xlsx and .csv
    ref_files = (
        glob.glob(os.path.join(ref_dir, '*.xlsx')) +
        glob.glob(os.path.join(ref_dir, '*.csv'))
    )
    ref_files = [f for f in ref_files if not os.path.basename(f).startswith('~$')]

    if not ref_files:
        raise FileNotFoundError(
            f'No reference .xlsx or .csv found in:\n  {ref_dir}\n\n'
            'Drop your reference file there and re-run.'
        )

    if len(ref_files) > 1:
        names = '\n  '.join(os.path.basename(f) for f in ref_files)
        raise ValueError(
            f'Multiple reference files found in reference_input/ - keep only one:\n  {names}'
        )

    return ref_files[0]


# ---------------------------------------------------------------------------
# Data Loaders
# ---------------------------------------------------------------------------

def _load_sifa_xlsx(path, label):
    """
    Load a SIFA DATA Excel file.

    Layout:
        Row 1: SIMBA codes (col A empty, B+ = codes)
        Row 2: descriptions
        Row 3+: data rows (col A = quarter label like '2024-Q1')

    Returns: (codes_list, cells_dict, descriptions_list)
        codes_list:  list of SIMBA codes
        cells_dict:  {(quarter_label, col_idx): value, ...}
        descriptions_list: list of description strings
    """
    if config.VERBOSE:
        print(f'Loading {label}: {os.path.basename(path)}')

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # Row 1: SIMBA codes starting from column B
    codes = []
    for c in range(2, ws.max_column + 1):
        code = ws.cell(row=1, column=c).value
        if code is None:
            break
        codes.append(str(code).strip())

    # Row 2: descriptions
    descriptions = []
    for c in range(2, 2 + len(codes)):
        desc = ws.cell(row=2, column=c).value
        descriptions.append(str(desc).strip() if desc else '')

    # Row 3+: data
    cells = {}
    quarters_found = []
    for r in range(3, ws.max_row + 1):
        quarter_val = ws.cell(row=r, column=1).value
        if quarter_val is None:
            continue
        ql = str(quarter_val).strip()
        if not ql:
            continue
        quarters_found.append(ql)

        for c_idx in range(len(codes)):
            v = ws.cell(row=r, column=c_idx + 2).value
            if v is not None:
                cells[(ql, c_idx)] = v

    wb.close()

    if config.VERBOSE:
        print(f'  Loaded: {len(codes)} codes, {len(cells)} cells, '
              f'{len(quarters_found)} quarters')

    return codes, cells, descriptions


def _load_sifa_csv(path, label):
    """
    Load a SIFA DATA CSV file (with description sub-header row).

    Layout:
        Row 1: empty + SIMBA codes
        Row 2: empty + descriptions
        Row 3+: quarter_label + values

    Returns: (codes_list, cells_dict, descriptions_list)
    """
    if config.VERBOSE:
        print(f'Loading {label}: {os.path.basename(path)}')

    with open(path, 'r', newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        header = next(reader, None)
        second = next(reader, None)

        # Row 1: codes (skip first empty column)
        codes = [c.strip() for c in header[1:]] if header else []

        # Row 2: check if it's a description row
        descriptions = []
        data_rows_raw = []
        if second and len(second) > 1 and ':' in str(second[1]):
            # It's a description row
            descriptions = [d.strip() for d in second[1:]]
            # Read remaining data rows
            for row in reader:
                if row and row[0].strip():
                    data_rows_raw.append(row)
        else:
            # No description row; second is a data row
            if second and second[0].strip():
                data_rows_raw.append(second)
            for row in reader:
                if row and row[0].strip():
                    data_rows_raw.append(row)

    cells = {}
    quarters_found = []
    for row in data_rows_raw:
        ql = row[0].strip()
        quarters_found.append(ql)
        for c_idx, val in enumerate(row[1:]):
            if c_idx < len(codes) and val.strip() != '':
                try:
                    cells[(ql, c_idx)] = float(val)
                except ValueError:
                    cells[(ql, c_idx)] = val.strip()

    if config.VERBOSE:
        print(f'  Loaded: {len(codes)} codes, {len(cells)} cells, '
              f'{len(quarters_found)} quarters')

    return codes, cells, descriptions


def load_data_file(path, label):
    """Load a SIFA data file (xlsx or csv)."""
    if path.lower().endswith('.csv'):
        return _load_sifa_csv(path, label)
    else:
        return _load_sifa_xlsx(path, label)


# ---------------------------------------------------------------------------
# Source Excel Verification
# ---------------------------------------------------------------------------

def _load_source_sections(ws):
    """
    Locate fund type sections in a source Excel worksheet.
    Returns list of {'fund_type_display': name, 'header_row': N, 'data_start': N}
    """
    # Known section names (lowercase)
    section_names = [
        'all types of funds',
        'equity funds',
        'balanced funds',
        'long term fixed income funds',
        'short term fixed income funds',
        'hedge funds',
        'other funds',
    ]

    sections = []
    for row in range(1, ws.max_row + 1):
        cell_val = ws.cell(row=row, column=1).value
        if cell_val is None:
            continue
        cell_text = str(cell_val).strip().lower()

        for sn in section_names:
            if cell_text == sn:
                # Find data start row
                data_start = None
                for r in range(row + 1, min(row + 10, ws.max_row + 1)):
                    v = ws.cell(row=r, column=1).value
                    if v is not None and str(v).strip():
                        if 'Quarter' not in str(v) and 'Net' not in str(v):
                            data_start = r
                            break
                if data_start is not None:
                    sections.append({
                        'fund_type_display': sn,
                        'header_row': row,
                        'data_start': data_start,
                    })
                break

    return sections


def verify_from_source(source_path, quarter_label, col_idx, pipeline_val, ref_val):
    """
    Open the source Excel file and check the actual value for a specific
    quarter/column to determine which side (pipeline or reference) is correct.

    Returns dict with verification info:
        source_value: the raw value from the source Excel
        pipeline_correct: True/False/None
        reference_correct: True/False/None
        note: human-readable explanation
    """
    # Parse quarter label: "2024-Q3" -> year=2024, q=3
    match = re.match(r'(\d{4})-Q(\d)', quarter_label)
    if not match:
        return {'source_value': None, 'note': f'Cannot parse quarter: {quarter_label}'}

    year = int(match.group(1))
    q_num = int(match.group(2))

    try:
        wb = openpyxl.load_workbook(source_path, data_only=True)
        ws = wb.active
    except Exception as e:
        return {'source_value': None, 'note': f'Cannot open source: {e}'}

    # Determine which section and metric this column maps to
    # 350 columns = 7 sections x 50 cols
    # Each section: 10 NETSAVING + 10 NETSAVINGSUM + 10 NETSAVINGPERC + 10 NETASSET + 10 NETASSETPERC
    section_idx = col_idx // 50
    position_in_section = col_idx % 50
    metric_idx = position_in_section // 10
    category_idx = position_in_section % 10

    # Find the section in the source worksheet
    sections = _load_source_sections(ws)
    if section_idx >= len(sections):
        wb.close()
        return {'source_value': None, 'note': f'Section {section_idx} not found in source'}

    section = sections[section_idx]

    # Find category rows starting from data_start
    data_rows = []
    row = section['data_start']
    while len(data_rows) < 10 and row <= ws.max_row:
        cell_val = ws.cell(row=row, column=1).value
        if cell_val is not None and str(cell_val).strip():
            data_rows.append(row)
            if str(cell_val).strip() == 'TOTAL':
                break
        row += 1

    if category_idx >= len(data_rows):
        wb.close()
        return {'source_value': None, 'note': f'Category row {category_idx} not found'}

    target_row = data_rows[category_idx]

    # Determine Excel column based on metric
    # metric 0 = NETSAVING -> quarter column (B=2, C=3, D=4, E=5)
    # metric 1 = NETSAVINGSUM -> col F=6
    # metric 2 = NETSAVINGPERC -> col G=7
    # metric 3 = NETASSET -> col H=8
    # metric 4 = NETASSETPERC -> col I=9
    quarter_col_map = {1: 2, 2: 3, 3: 4, 4: 5}

    if metric_idx == 0:
        target_col = quarter_col_map.get(q_num, 2)
    elif metric_idx == 1:
        target_col = 6
    elif metric_idx == 2:
        target_col = 7
    elif metric_idx == 3:
        target_col = 8
    elif metric_idx == 4:
        target_col = 9
    else:
        wb.close()
        return {'source_value': None, 'note': f'Unknown metric index: {metric_idx}'}

    source_value = ws.cell(row=target_row, column=target_col).value
    wb.close()

    # Clean the source value
    if isinstance(source_value, (int, float)):
        source_value = float(source_value)
    elif source_value is not None:
        try:
            source_value = float(str(source_value).replace(',', '').strip())
        except (ValueError, TypeError):
            pass

    # Compare with both sides
    result = {
        'source_value': source_value,
        'source_row': target_row,
        'source_col': target_col,
        'section': section['fund_type_display'],
        'pipeline_correct': None,
        'reference_correct': None,
        'note': '',
    }

    if source_value is None:
        result['note'] = 'Source cell is empty'
        return result

    def _values_match(a, b):
        if a is None or b is None:
            return a is None and b is None
        try:
            fa, fb = float(a), float(b)
            return abs(fa - fb) < config.FLOAT_TOLERANCE
        except (ValueError, TypeError):
            return str(a) == str(b)

    p_match = _values_match(source_value, pipeline_val)
    r_match = _values_match(source_value, ref_val)

    result['pipeline_correct'] = p_match
    result['reference_correct'] = r_match

    if p_match and r_match:
        result['note'] = 'Both match source (within tolerance)'
    elif p_match and not r_match:
        result['note'] = 'Pipeline matches source; reference differs'
    elif not p_match and r_match:
        result['note'] = 'Reference matches source; pipeline differs'
    else:
        result['note'] = 'Neither matches source exactly'

    return result


# ---------------------------------------------------------------------------
# Comparison Engine
# ---------------------------------------------------------------------------

class ComparisonReport:

    def __init__(self, pipeline_path, reference_path, source_downloads, pipeline_csv):
        self.pipeline_path  = pipeline_path
        self.reference_path = reference_path
        self.source_downloads = source_downloads  # {year: filepath}
        self.pipeline_csv   = pipeline_csv

        self.pipeline_codes  = []
        self.pipeline_cells  = {}
        self.pipeline_descs  = []
        self.reference_codes = []
        self.reference_cells = {}
        self.reference_descs = []

        self.matches    = []
        self.mismatches = []
        self.missing    = []
        self.extra      = []

        self.mismatch_coords = []
        self.missing_coords  = []
        self.extra_coords    = []

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def get_section_name(self, col_idx):
        """Return human-readable section name for a column index."""
        section_idx = col_idx // 50
        section_names = [
            'All types of funds', 'Equity funds', 'Balanced funds',
            'Long term fixed income', 'Short term fixed income',
            'Hedge funds', 'Other funds',
        ]
        if section_idx < len(section_names):
            return section_names[section_idx]
        return 'UNKNOWN'

    def get_metric_name(self, col_idx):
        """Return metric name for a column index."""
        position = col_idx % 50
        metric_idx = position // 10
        metric_names = [
            'Net savings', 'Net savings sum', 'Net savings %',
            'Net assets', 'Net assets %',
        ]
        if metric_idx < len(metric_names):
            return metric_names[metric_idx]
        return 'UNKNOWN'

    def get_category_name(self, col_idx):
        """Return category name for a column index."""
        position = col_idx % 50
        cat_idx = position % 10
        cat_names = [
            'Swedish households', 'ISK', 'IPS', 'Unit linked',
            'Premium Pension', 'Nominee Accounts',
            'Non profit inst.', 'Swedish corporations',
            'Others', 'TOTAL',
        ]
        if cat_idx < len(cat_names):
            return cat_names[cat_idx]
        return 'UNKNOWN'

    def compare_values(self, val1, val2):
        """Compare two values with float tolerance."""
        if val1 is None and val2 is None:
            return True, 0
        if val1 is None or val2 is None:
            return False, None
        try:
            f1, f2 = float(val1), float(val2)
            diff = f1 - f2
            if abs(diff) < config.FLOAT_TOLERANCE:
                return True, 0
            return False, diff
        except (ValueError, TypeError):
            return str(val1) == str(val2), None

    def build_row_dict(self, status, quarter, col_idx, code, out_val, ref_val, diff):
        section  = self.get_section_name(col_idx)
        metric   = self.get_metric_name(col_idx)
        category = self.get_category_name(col_idx)
        desc = ''
        if col_idx < len(self.pipeline_descs):
            desc = self.pipeline_descs[col_idx]

        # Find source file for the year
        year_match = re.match(r'(\d{4})', quarter)
        year = int(year_match.group(1)) if year_match else None
        source_file = ''
        if year and year in self.source_downloads:
            source_file = os.path.basename(self.source_downloads[year])

        return {
            'Pipeline_File':     os.path.basename(self.pipeline_path),
            'Reference_File':    os.path.basename(self.reference_path),
            'Status':            status,
            'Quarter':           quarter,
            'Excel_Column':      col_idx + 2,
            'Column_Code':       code or '',
            'Section':           section,
            'Category':          category,
            'Metric':            metric,
            'Column_Description': desc,
            'Pipeline_Value':    out_val if out_val is not None else '',
            'Reference_Value':   ref_val if ref_val is not None else '',
            'Difference':        round(diff, 6) if diff is not None else '',
            'Source_File':       source_file if source_file else 'NO SOURCE FILE',
            '_year':             year,
            '_col_idx':          col_idx,
        }

    # ------------------------------------------------------------------
    # Core comparison
    # ------------------------------------------------------------------

    def perform_comparison(self):
        if config.VERBOSE:
            print('\nPerforming comparison...')

        # Only compare quarters that exist in BOTH files
        pipeline_quarters = set(k[0] for k in self.pipeline_cells.keys())
        reference_quarters = set(k[0] for k in self.reference_cells.keys())
        common_quarters = sorted(pipeline_quarters & reference_quarters)

        if config.VERBOSE:
            print(f'  Pipeline quarters:  {sorted(pipeline_quarters)}')
            print(f'  Reference quarters: {sorted(reference_quarters)}')
            print(f'  Overlapping:        {common_quarters}')

        pipeline_only  = sorted(pipeline_quarters - reference_quarters)
        reference_only = sorted(reference_quarters - pipeline_quarters)

        if pipeline_only and config.VERBOSE:
            print(f'  Pipeline only (not compared): {pipeline_only}')
        if reference_only and config.VERBOSE:
            print(f'  Reference only (not compared): {reference_only}')

        # Compare overlapping quarters cell-by-cell
        num_codes = max(len(self.pipeline_codes), len(self.reference_codes))

        for quarter in common_quarters:
            for col_idx in range(num_codes):
                out_val = self.pipeline_cells.get((quarter, col_idx))
                ref_val = self.reference_cells.get((quarter, col_idx))
                code = self.pipeline_codes[col_idx] if col_idx < len(self.pipeline_codes) else ''

                if out_val is None and ref_val is None:
                    continue
                elif out_val is not None and ref_val is not None:
                    match, diff = self.compare_values(out_val, ref_val)
                    if match:
                        self.matches.append((quarter, col_idx))
                    else:
                        self.mismatches.append(
                            self.build_row_dict('MISMATCH', quarter, col_idx,
                                                code, out_val, ref_val, diff)
                        )
                        self.mismatch_coords.append((quarter, col_idx))
                elif out_val is None:
                    self.missing.append(
                        self.build_row_dict('MISSING', quarter, col_idx,
                                            code, None, ref_val, None)
                    )
                    self.missing_coords.append((quarter, col_idx))
                else:
                    self.extra.append(
                        self.build_row_dict('NEW_VALUE', quarter, col_idx,
                                            code, out_val, None, None)
                    )
                    self.extra_coords.append((quarter, col_idx))

        if config.VERBOSE:
            print(f'\n  Matches:    {len(self.matches)}')
            print(f'  Mismatches: {len(self.mismatches)}')
            print(f'  Missing:    {len(self.missing)}')
            print(f'  New values: {len(self.extra)}')

    # ------------------------------------------------------------------
    # Source verification for mismatches
    # ------------------------------------------------------------------

    def verify_mismatches(self):
        """Check each mismatch against the source Excel to determine who is correct."""
        if not self.mismatches or not self.source_downloads:
            return

        if config.VERBOSE:
            print('\nVerifying mismatches against source Excel files...')

        verified = 0
        for row in self.mismatches:
            year = row.get('_year')
            if year and year in self.source_downloads:
                result = verify_from_source(
                    self.source_downloads[year],
                    row['Quarter'],
                    row['_col_idx'],
                    row['Pipeline_Value'],
                    row['Reference_Value'],
                )
                row['Source_Value'] = result.get('source_value', '')
                row['Pipeline_Correct'] = result.get('pipeline_correct', '')
                row['Reference_Correct'] = result.get('reference_correct', '')
                row['Verification_Note'] = result.get('note', '')
                verified += 1
            else:
                row['Source_Value'] = ''
                row['Pipeline_Correct'] = ''
                row['Reference_Correct'] = ''
                row['Verification_Note'] = 'No source file available'

        if config.VERBOSE:
            print(f'  Verified {verified}/{len(self.mismatches)} mismatches')

            # Summary of verification
            p_correct = sum(1 for r in self.mismatches if r.get('Pipeline_Correct') is True)
            r_correct = sum(1 for r in self.mismatches if r.get('Reference_Correct') is True)
            both = sum(1 for r in self.mismatches
                       if r.get('Pipeline_Correct') is True and r.get('Reference_Correct') is True)
            neither = sum(1 for r in self.mismatches
                          if r.get('Pipeline_Correct') is False and r.get('Reference_Correct') is False)

            print(f'  Pipeline correct:  {p_correct}')
            print(f'  Reference correct: {r_correct}')
            if both:
                print(f'  Both correct (tolerance): {both}')
            if neither:
                print(f'  Neither correct: {neither}')

    # ------------------------------------------------------------------
    # Source file copier
    # ------------------------------------------------------------------

    def copy_source_files(self, rows, dest_dir):
        """Copy source Excel files referenced by rows into dest_dir."""
        os.makedirs(dest_dir, exist_ok=True)
        copied = {}

        for row in rows:
            year = row.get('_year')
            if year and year in self.source_downloads:
                src_path = self.source_downloads[year]
                fn = os.path.basename(src_path)
                if fn not in copied:
                    dest = os.path.join(dest_dir, fn)
                    if os.path.exists(src_path):
                        shutil.copy2(src_path, dest)
                        copied[fn] = dest
                    else:
                        copied[fn] = '(source not found)'

        return copied

    # ------------------------------------------------------------------
    # CSV writers
    # ------------------------------------------------------------------

    _CSV_HEADER = [
        'Pipeline_File', 'Reference_File',
        'Status', 'Quarter', 'Excel_Column', 'Column_Code',
        'Section', 'Category', 'Metric',
        'Column_Description',
        'Pipeline_Value', 'Reference_Value', 'Difference',
        'Source_File',
        'Source_Value', 'Pipeline_Correct', 'Reference_Correct',
        'Verification_Note', 'Copied_To',
    ]

    def _write_category_csv(self, rows, output_path, copied_map, label):
        if config.VERBOSE:
            print(f'  Writing {label} report: {os.path.basename(output_path)}  '
                  f'({len(rows)} rows)')

        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=self._CSV_HEADER,
                                    extrasaction='ignore')
            writer.writeheader()
            for row in rows:
                sf = row.get('Source_File', '')
                row['Copied_To'] = copied_map.get(sf, '') if sf and sf != 'NO SOURCE FILE' else ''
                writer.writerow(row)

    # ------------------------------------------------------------------
    # Summary report
    # ------------------------------------------------------------------

    def generate_summary_stats(self):
        total = len(self.matches) + len(self.mismatches) + len(self.missing) + len(self.extra)
        return {
            'timestamp':            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'pipeline_file':        self.pipeline_path,
            'reference_file':       self.reference_path,
            'total_cells_compared': total,
            'matches':              len(self.matches),
            'mismatches':           len(self.mismatches),
            'missing':              len(self.missing),
            'new_values':           len(self.extra),
            'match_rate':           (len(self.matches) / total * 100) if total > 0 else 0,
        }

    def analyze_by_category(self, rows, key_func):
        groups = defaultdict(list)
        for row in rows:
            groups[key_func(row)].append(row)
        return groups

    def write_summary_report(self, output_path, copied_mm, copied_ms, copied_nv):
        if config.VERBOSE:
            print(f'Writing summary report: {os.path.basename(output_path)}')

        stats = self.generate_summary_stats()

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('=' * 70 + '\n')
            f.write('SIFA COMPARISON REPORT\n')
            f.write('=' * 70 + '\n\n')
            f.write(f"Generated:  {stats['timestamp']}\n\n")

            f.write('FILES COMPARED:\n')
            f.write(f"  Pipeline Output : {os.path.basename(stats['pipeline_file'])}\n")
            f.write(f"  Reference       : {os.path.basename(stats['reference_file'])}\n")
            f.write(f"  Source files    : {len(self.source_downloads)} year(s)\n\n")

            f.write('OVERALL STATISTICS:\n')
            f.write(f"  Total cells compared : {stats['total_cells_compared']:,}\n")
            f.write(f"  Matches              : {stats['matches']:,} "
                    f"({stats['match_rate']:.2f}%)\n")
            f.write(f"  Mismatches           : {stats['mismatches']:,}\n")
            f.write(f"  Missing (in ref)     : {stats['missing']:,}\n")
            f.write(f"  New values (extra)   : {stats['new_values']:,}\n\n")

            f.write('OUTPUT FILES:\n')
            f.write(f"  {config.REPORT_MISMATCHES}  ({stats['mismatches']} rows)\n")
            f.write(f"  {config.REPORT_MISSING}     ({stats['missing']} rows)\n")
            f.write(f"  {config.REPORT_NEW_VALUES}  ({stats['new_values']} rows)\n\n")

            # --- Mismatches detail ---
            if self.mismatches:
                f.write('\u2500' * 70 + '\n')
                f.write('MISMATCHES\n')
                f.write('\u2500' * 70 + '\n')

                f.write('\n  By Section:\n')
                by_section = self.analyze_by_category(
                    self.mismatches, lambda r: r['Section'])
                for section, items in sorted(by_section.items()):
                    f.write(f"    {section}: {len(items)}\n")

                f.write('\n  By Metric:\n')
                by_metric = self.analyze_by_category(
                    self.mismatches, lambda r: r['Metric'])
                for metric, items in sorted(by_metric.items(),
                                            key=lambda x: -len(x[1])):
                    f.write(f"    {metric}: {len(items)}\n")

                f.write('\n  By Quarter:\n')
                by_quarter = self.analyze_by_category(
                    self.mismatches, lambda r: r['Quarter'])
                for quarter, items in sorted(by_quarter.items()):
                    f.write(f"    {quarter}: {len(items)}\n")

                # Source verification summary
                p_correct = sum(1 for r in self.mismatches
                                if r.get('Pipeline_Correct') is True)
                r_correct = sum(1 for r in self.mismatches
                                if r.get('Reference_Correct') is True)
                if p_correct or r_correct:
                    f.write('\n  Source Verification:\n')
                    f.write(f"    Pipeline matches source: {p_correct}\n")
                    f.write(f"    Reference matches source: {r_correct}\n")

                f.write(f'\n  Top mismatches (by absolute difference):\n')
                sorted_mm = sorted(
                    [r for r in self.mismatches if r['Difference'] != ''],
                    key=lambda r: abs(float(r['Difference']))
                    if r['Difference'] else 0,
                    reverse=True,
                )[:config.MAX_CONSOLE_EXAMPLES]
                for r in sorted_mm:
                    f.write(f"    Quarter: {r['Quarter']}  Col: {r['Excel_Column']}  "
                            f"{r['Section']} | {r['Category']} | {r['Metric']}\n")
                    f.write(f"      Pipeline : {r['Pipeline_Value']}\n")
                    f.write(f"      Reference: {r['Reference_Value']}\n")
                    f.write(f"      Diff     : {r['Difference']}\n")
                    if r.get('Source_Value') != '':
                        f.write(f"      Source   : {r.get('Source_Value', '')}\n")
                        f.write(f"      Verdict  : {r.get('Verification_Note', '')}\n")
                    f.write(f"      Source file: {r['Source_File']}\n")
                f.write('\n')

                if copied_mm:
                    f.write(f"  Source files copied to  "
                            f"files/{config.MISMATCH_FILES_DIR}/:\n")
                    for fn in sorted(copied_mm.keys()):
                        f.write(f"    {fn}\n")
                    f.write('\n')

            # --- Missing detail ---
            if self.missing:
                f.write('\u2500' * 70 + '\n')
                f.write('MISSING  (present in reference, absent in pipeline)\n')
                f.write('\u2500' * 70 + '\n')

                f.write('\n  By Section:\n')
                by_section = self.analyze_by_category(
                    self.missing, lambda r: r['Section'])
                for section, items in sorted(by_section.items()):
                    f.write(f"    {section}: {len(items)}\n")

                f.write('\n  By Quarter:\n')
                by_quarter = self.analyze_by_category(
                    self.missing, lambda r: r['Quarter'])
                for quarter, items in sorted(by_quarter.items()):
                    f.write(f"    {quarter}: {len(items)}\n")
                f.write('\n')

                if copied_ms:
                    f.write(f"  Source files copied to  "
                            f"files/{config.MISSING_FILES_DIR}/:\n")
                    for fn in sorted(copied_ms.keys()):
                        f.write(f"    {fn}\n")
                    f.write('\n')

            # --- New values detail ---
            if self.extra:
                f.write('\u2500' * 70 + '\n')
                f.write('NEW VALUES  (present in pipeline, absent in reference)\n')
                f.write('\u2500' * 70 + '\n')

                f.write('\n  By Section:\n')
                by_section = self.analyze_by_category(
                    self.extra, lambda r: r['Section'])
                for section, items in sorted(by_section.items()):
                    f.write(f"    {section}: {len(items)}\n")
                f.write('\n')

                if copied_nv:
                    f.write(f"  Source files copied to  "
                            f"files/{config.NEW_VALUES_FILES_DIR}/:\n")
                    for fn in sorted(copied_nv.keys()):
                        f.write(f"    {fn}\n")
                    f.write('\n')

            f.write('=' * 70 + '\n')
            f.write('Detailed CSV reports:\n')
            f.write(f"  {config.REPORT_MISMATCHES}\n")
            f.write(f"  {config.REPORT_MISSING}\n")
            f.write(f"  {config.REPORT_NEW_VALUES}\n")
            f.write('=' * 70 + '\n')

    # ------------------------------------------------------------------
    # Annotated Excel
    # ------------------------------------------------------------------

    def create_annotated_excel(self, source_path, output_path, coords_dict, label):
        """
        Create a copy of the DATA Excel with coloured highlights.
        Red = mismatch, Yellow = missing, Green = new value.
        """
        if config.VERBOSE:
            print(f'Creating annotated {label} file: '
                  f'{os.path.basename(output_path)}')

        shutil.copy2(source_path, output_path)

        # Only annotate .xlsx files
        if not output_path.lower().endswith('.xlsx'):
            if config.VERBOSE:
                print('  Skipped (not .xlsx)')
            return

        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        red_fill    = PatternFill(start_color='FFCCCC', end_color='FFCCCC',
                                  fill_type='solid')
        yellow_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC',
                                  fill_type='solid')
        green_fill  = PatternFill(start_color='CCFFCC', end_color='CCFFCC',
                                  fill_type='solid')
        bold_font   = Font(bold=True)

        # Map quarter labels to row numbers (Row 3+ in SIFA format)
        quarter_to_row = {}
        for r in range(3, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val:
                quarter_to_row[str(val).strip()] = r

        highlight_counts = {'MISMATCH': 0, 'MISSING': 0, 'NEW_VALUE': 0}
        fill_map = {
            'MISMATCH': red_fill,
            'MISSING': yellow_fill,
            'NEW_VALUE': green_fill,
        }

        for status, coords in coords_dict.items():
            fill = fill_map[status]
            for quarter, col_idx in coords:
                row_num = quarter_to_row.get(quarter)
                if row_num:
                    ws.cell(row=row_num, column=col_idx + 2).fill = fill
                    highlight_counts[status] += 1

        # Insert legend row at top
        ws.insert_rows(1)
        ws.cell(row=1, column=1, value='LEGEND:').font = bold_font
        ws.cell(row=1, column=2, value='RED = Mismatch').fill = red_fill
        ws.cell(row=1, column=3, value='YELLOW = Missing').fill = yellow_fill
        ws.cell(row=1, column=4, value='GREEN = New Value').fill = green_fill
        ws.cell(row=1, column=5,
                value=f'Annotated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')

        wb.save(output_path)
        wb.close()

        if config.VERBOSE:
            print(f'  Highlighted: {highlight_counts["MISMATCH"]} mismatches (red), '
                  f'{highlight_counts["MISSING"]} missing (yellow), '
                  f'{highlight_counts["NEW_VALUE"]} new values (green)')

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------

    def run(self):
        print('=' * 70)
        print('SIFA COMPARISON TOOL')
        print('=' * 70)
        print()
        print(f'Pipeline Output  : {self.pipeline_path}')
        print(f'Reference File   : {self.reference_path}')
        print(f'Source Downloads  : {len(self.source_downloads)} year file(s)')
        print()

        for label, path in [('Pipeline', self.pipeline_path),
                             ('Reference', self.reference_path)]:
            if not os.path.exists(path):
                print(f'ERROR: {label} file not found: {path}')
                return 1

        # Load data
        self.pipeline_codes, self.pipeline_cells, self.pipeline_descs = \
            load_data_file(self.pipeline_path, 'Pipeline Output')
        self.reference_codes, self.reference_cells, self.reference_descs = \
            load_data_file(self.reference_path, 'Reference')

        if self.pipeline_codes != self.reference_codes:
            # Check how many differ
            diffs = sum(1 for a, b in zip(self.pipeline_codes, self.reference_codes)
                        if a != b)
            print(f'\nWARNING: {diffs} column codes differ between pipeline '
                  f'and reference!')

        # Compare
        self.perform_comparison()

        # Verify mismatches against source Excel
        self.verify_mismatches()

        # Build run output directory
        timestamp    = datetime.now().strftime('%Y%m%d_%H%M%S')
        compare_dir  = os.path.dirname(os.path.abspath(__file__))
        run_dir      = os.path.join(compare_dir, f'run_{timestamp}')
        files_dir    = os.path.join(run_dir, config.FILES_DIR)
        mismatch_dir = os.path.join(files_dir, config.MISMATCH_FILES_DIR)
        missing_dir  = os.path.join(files_dir, config.MISSING_FILES_DIR)
        nv_dir       = os.path.join(files_dir, config.NEW_VALUES_FILES_DIR)

        os.makedirs(run_dir, exist_ok=True)

        if config.VERBOSE:
            print(f'\nOutput directory: {run_dir}')

        # ---------- Copy original pipeline & reference files ----------
        if config.VERBOSE:
            print('\nCopying original pipeline and reference files...')

        output_copy = os.path.join(
            run_dir, f'ORIGINAL_{os.path.basename(self.pipeline_path)}')
        ref_copy = os.path.join(
            run_dir, f'ORIGINAL_{os.path.basename(self.reference_path)}')
        shutil.copy2(self.pipeline_path, output_copy)
        shutil.copy2(self.reference_path, ref_copy)

        # ---------- Copy source files ----------
        if config.VERBOSE:
            print('\nCopying source files from downloads...')

        copied_mm = {}
        copied_ms = {}
        copied_nv = {}

        if self.mismatches:
            copied_mm = self.copy_source_files(self.mismatches, mismatch_dir)
            if config.VERBOSE:
                print(f'  Mismatch  : {len(copied_mm)} source file(s) -> '
                      f'files/{config.MISMATCH_FILES_DIR}/')

        if self.missing:
            copied_ms = self.copy_source_files(self.missing, missing_dir)
            if config.VERBOSE:
                print(f'  Missing   : {len(copied_ms)} source file(s) -> '
                      f'files/{config.MISSING_FILES_DIR}/')

        if self.extra:
            copied_nv = self.copy_source_files(self.extra, nv_dir)
            if config.VERBOSE:
                print(f'  New values: {len(copied_nv)} source file(s) -> '
                      f'files/{config.NEW_VALUES_FILES_DIR}/')

        # ---------- Annotated Excel ----------
        if config.VERBOSE:
            print()

        # Annotate pipeline output (only if xlsx)
        if self.pipeline_path.lower().endswith('.xlsx'):
            annotated_pipeline = os.path.join(
                run_dir,
                f'ANNOTATED_Pipeline_{os.path.basename(self.pipeline_path)}')
            self.create_annotated_excel(
                self.pipeline_path, annotated_pipeline,
                {'MISMATCH': self.mismatch_coords,
                 'NEW_VALUE': self.extra_coords},
                'Pipeline',
            )

        # Annotate reference (only if xlsx)
        if self.reference_path.lower().endswith('.xlsx'):
            annotated_ref = os.path.join(
                run_dir,
                f'ANNOTATED_Reference_{os.path.basename(self.reference_path)}')
            self.create_annotated_excel(
                self.reference_path, annotated_ref,
                {'MISMATCH': self.mismatch_coords,
                 'MISSING': self.missing_coords},
                'Reference',
            )

        # ---------- Separate CSV reports ----------
        if config.VERBOSE:
            print('\nWriting separate CSV reports...')

        self._write_category_csv(
            self.mismatches,
            os.path.join(run_dir, config.REPORT_MISMATCHES),
            copied_mm, 'Mismatches',
        )
        self._write_category_csv(
            self.missing,
            os.path.join(run_dir, config.REPORT_MISSING),
            copied_ms, 'Missing',
        )
        self._write_category_csv(
            self.extra,
            os.path.join(run_dir, config.REPORT_NEW_VALUES),
            copied_nv, 'New Values',
        )

        # ---------- Summary ----------
        summary_path = os.path.join(run_dir, config.SUMMARY_FILENAME)
        self.write_summary_report(summary_path, copied_mm, copied_ms,
                                  copied_nv)

        # ---------- Console summary ----------
        stats = self.generate_summary_stats()
        print('\n' + '=' * 70)
        print('COMPARISON COMPLETE')
        print('=' * 70)
        print(f"\nMatches    : {stats['matches']:,} ({stats['match_rate']:.2f}%)")
        print(f"Mismatches : {stats['mismatches']:,}")
        print(f"Missing    : {stats['missing']:,}")
        print(f"New values : {stats['new_values']:,}")
        print(f'\nOutput directory: {run_dir}')

        print('\nFiles created:')
        created = [
            os.path.basename(output_copy),
            os.path.basename(ref_copy),
        ]
        if self.pipeline_path.lower().endswith('.xlsx'):
            created.append(f'ANNOTATED_Pipeline_{os.path.basename(self.pipeline_path)}')
        if self.reference_path.lower().endswith('.xlsx'):
            created.append(f'ANNOTATED_Reference_{os.path.basename(self.reference_path)}')
        created += [
            config.REPORT_MISMATCHES,
            config.REPORT_MISSING,
            config.REPORT_NEW_VALUES,
            config.SUMMARY_FILENAME,
        ]
        for name in created:
            print(f'  {name}')

        print(f'\nfiles/ subfolders:')
        if copied_mm:
            print(f'  files/{config.MISMATCH_FILES_DIR}/   '
                  f'({len(copied_mm)} source file(s))')
        if copied_ms:
            print(f'  files/{config.MISSING_FILES_DIR}/    '
                  f'({len(copied_ms)} source file(s))')
        if copied_nv:
            print(f'  files/{config.NEW_VALUES_FILES_DIR}/ '
                  f'({len(copied_nv)} source file(s))')
        if not (copied_mm or copied_ms or copied_nv):
            print('  (no source files to copy)')

        print('=' * 70)
        return 0


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    print('=' * 60)
    print(' SIFA Compare  --  drop reference file and run me')
    print('=' * 60)
    print()

    try:
        compare_dir  = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(compare_dir)

        print('Discovering paths...')

        pipeline_file    = find_pipeline_output(project_root)
        source_downloads = find_source_downloads(project_root)
        reference_file   = find_reference_file(compare_dir)

        print(f'  Pipeline:   {os.path.basename(pipeline_file)}')
        print(f'  Reference:  {os.path.basename(reference_file)}')
        print(f'  Source DLs: {len(source_downloads)} year file(s)')

        report = ComparisonReport(
            pipeline_file, reference_file,
            source_downloads, None,
        )
        return report.run()

    except (FileNotFoundError, ValueError) as e:
        print(f'\nERROR: {e}')
        return 1

    except KeyboardInterrupt:
        print('\n\nInterrupted by user')
        return 130

    except Exception as e:
        print(f'\nERROR: {e}')
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
