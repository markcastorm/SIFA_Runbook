# test_extraction.py
# Test the data_only=True + copy-values approach on the 2024 source Excel.
# Compares raw extracted values against the sample SIFA_DATA_20250212.xlsx.
#
# This validates that our extractor reads values correctly by using
# the simplest possible approach: just read cell values directly.

import os
import sys
import openpyxl

# ── Paths ────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SOURCE_2024 = os.path.join(
    BASE_DIR, 'downloads', '20260508_164624', '2024',
    'fund-saving-by-category-2024.xlsx'
)

SAMPLE_FILE = os.path.join(
    BASE_DIR, 'Project_information', 'SIFA_DATA_20250212.xlsx'
)

OUTPUT_DIR = os.path.join(BASE_DIR, 'test_output')

# ── Source Excel layout ──────────────────────────────────────────────────
# Section headers in column A (same order as config.FUND_TYPES):
SECTION_HEADERS = [
    'All types of funds',
    'Equity funds',
    'Balanced funds',
    'Long term fixed income funds',
    'Short term fixed income funds',
    'Hedge funds',
    'Other funds',
]

# Quarter columns: B=Q1, C=Q2, D=Q3, E=Q4
# Summary columns: F=Sum, G=%, H=Assets, I=Assets%
QUARTER_COLS = {1: 2, 2: 3, 3: 4, 4: 5}
SUM_COL = 6
PCT_COL = 7
ASSET_COL = 8
ASSET_PCT_COL = 9


# ── Step 1: Copy values from source Excel ────────────────────────────────

def copy_values(input_path, output_path):
    """
    Load Excel with data_only=True (gets cached formula results),
    copy all cell values to a new sheet. This is the user's proven method.
    """
    print(f'Loading source: {os.path.basename(input_path)}')
    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb.active
    print(f'  Sheet: {ws.title}, {ws.max_row} rows x {ws.max_column} cols')

    # Create values-only sheet
    target = wb.create_sheet(title='Values_Only')
    for row in ws.iter_rows():
        for cell in row:
            target.cell(row=cell.row, column=cell.column, value=cell.value)

    wb.save(output_path)
    print(f'  Values-only copy saved: {output_path}')
    return output_path


# ── Step 2: Find sections and extract data ───────────────────────────────

def find_sections(ws):
    """Find all fund type section headers and their data start rows."""
    lookup = {name.lower(): name for name in SECTION_HEADERS}
    sections = []

    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val is None:
            continue
        text = str(val).strip().lower()

        if text in lookup:
            # Find where data rows start (skip column header + sub-header)
            data_start = None
            for r in range(row + 1, min(row + 10, ws.max_row + 1)):
                v = ws.cell(row=r, column=1).value
                if v is not None and str(v).strip():
                    label = str(v).strip()
                    if 'Quarter' not in label and 'Net' not in label:
                        data_start = r
                        break

            if data_start:
                sections.append({
                    'name': lookup[text],
                    'header_row': row,
                    'data_start': data_start,
                })

    return sections


def find_category_rows(ws, data_start):
    """Read 10 consecutive non-empty rows (9 categories + TOTAL)."""
    rows = []
    r = data_start
    while len(rows) < 10 and r <= ws.max_row:
        val = ws.cell(row=r, column=1).value
        if val is not None and str(val).strip():
            rows.append(r)
            if str(val).strip() == 'TOTAL':
                break
        r += 1
    return rows


def detect_quarters(ws, data_rows):
    """Detect which quarters have non-zero data (skip TOTAL row)."""
    check_rows = data_rows[:-1] if len(data_rows) > 1 else data_rows
    quarters = []
    for q_num, col_idx in QUARTER_COLS.items():
        for row_num in check_rows:
            val = ws.cell(row=row_num, column=col_idx).value
            if val is not None and isinstance(val, (int, float)) and val != 0:
                quarters.append(q_num)
                break
    return sorted(quarters)


def extract_section_values(ws, section, q_num, is_last_quarter):
    """
    Extract raw values for one section + one quarter.
    Returns list of 50 values (or fewer if not last quarter).
    Positions: 0-9=NETSAVING, 10-19=SUM, 20-29=PCT, 30-39=ASSET, 40-49=ASSETPCT
    """
    cat_rows = find_category_rows(ws, section['data_start'])

    # Pad to 10 if needed
    while len(cat_rows) < 10:
        cat_rows.append(None)

    q_col = QUARTER_COLS[q_num]
    values = []

    # NETSAVING (positions 0-9): always
    for row_num in cat_rows:
        if row_num is not None:
            val = ws.cell(row=row_num, column=q_col).value
            values.append(val if isinstance(val, (int, float)) else None)
        else:
            values.append(None)

    if is_last_quarter:
        # SUM (positions 10-19)
        for row_num in cat_rows:
            if row_num is not None:
                val = ws.cell(row=row_num, column=SUM_COL).value
                values.append(val if isinstance(val, (int, float)) else None)
            else:
                values.append(None)

        # PCT (positions 20-29)
        for row_num in cat_rows:
            if row_num is not None:
                val = ws.cell(row=row_num, column=PCT_COL).value
                values.append(val if isinstance(val, (int, float)) else None)
            else:
                values.append(None)

        # ASSET (positions 30-39)
        for row_num in cat_rows:
            if row_num is not None:
                val = ws.cell(row=row_num, column=ASSET_COL).value
                values.append(val if isinstance(val, (int, float)) else None)
            else:
                values.append(None)

        # ASSET PCT (positions 40-49)
        for row_num in cat_rows:
            if row_num is not None:
                val = ws.cell(row=row_num, column=ASSET_PCT_COL).value
                values.append(val if isinstance(val, (int, float)) else None)
            else:
                values.append(None)

    return values


# ── Step 3: Load sample reference ────────────────────────────────────────

def load_sample(path):
    """
    Load the sample SIFA_DATA_20250212.xlsx.
    Returns dict: {(quarter_label, col_idx): value}
    """
    print(f'Loading sample: {os.path.basename(path)}')
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # Row 1: SIMBA codes
    codes = []
    for c in range(2, ws.max_column + 1):
        code = ws.cell(row=1, column=c).value
        if code is None:
            break
        codes.append(str(code).strip())

    # Row 2: descriptions (skip)
    # Row 3+: data
    cells = {}
    quarters = []
    for r in range(3, ws.max_row + 1):
        ql = ws.cell(row=r, column=1).value
        if ql is None:
            continue
        ql = str(ql).strip()
        quarters.append(ql)
        for c_idx in range(len(codes)):
            v = ws.cell(row=r, column=c_idx + 2).value
            if v is not None:
                cells[(ql, c_idx)] = v

    wb.close()
    print(f'  Loaded: {len(codes)} codes, {len(cells)} cells, quarters: {quarters}')
    return codes, cells, quarters


# ── Step 4: Compare ──────────────────────────────────────────────────────

def clean_float(val):
    """Remove FP artifacts (same as extractor._clean_float)."""
    if val is None:
        return None
    if val == 0:
        return 0.0
    abs_val = abs(val)
    for dp in range(2, 16):
        candidate = round(val, dp)
        if candidate == 0 and abs_val > 1e-9:
            continue
        if abs(candidate - val) / max(abs_val, 1e-15) < 1e-9:
            return candidate
    return val


def main():
    print('=' * 60)
    print(' SIFA Extraction Test — values-only approach')
    print('=' * 60)
    print()

    # Verify files exist
    if not os.path.exists(SOURCE_2024):
        print(f'ERROR: Source file not found: {SOURCE_2024}')
        return 1
    if not os.path.exists(SAMPLE_FILE):
        print(f'ERROR: Sample file not found: {SAMPLE_FILE}')
        return 1

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Step 1: Copy values from source
    values_copy = os.path.join(OUTPUT_DIR, 'fund-saving-2024-values-only.xlsx')
    copy_values(SOURCE_2024, values_copy)
    print()

    # Step 2: Extract data from the values-only copy
    print('Extracting data from values-only copy...')
    wb = openpyxl.load_workbook(values_copy, data_only=True)
    ws = wb['Values_Only']

    sections = find_sections(ws)
    print(f'  Found {len(sections)} sections:')
    for s in sections:
        print(f'    {s["name"]} — header row {s["header_row"]}, data row {s["data_start"]}')

    # Detect quarters
    first_cat_rows = find_category_rows(ws, sections[0]['data_start'])
    quarters = detect_quarters(ws, first_cat_rows)
    print(f'  Quarters with data: {quarters}')
    last_q = max(quarters)

    # Extract all values
    extracted = {}  # {quarter_label: {col_idx: value}}
    for q_num in quarters:
        ql = f'2024-Q{q_num}'
        is_last = (q_num == last_q)
        row_values = {}
        col_offset = 0

        for section in sections:
            values = extract_section_values(ws, section, q_num, is_last)
            for pos, val in enumerate(values):
                global_idx = col_offset + pos
                row_values[global_idx] = val
            col_offset += 50

        populated = sum(1 for v in row_values.values() if v is not None)
        print(f'  {ql}: {populated} cells populated'
              f'{"  (last quarter — includes sum/assets)" if is_last else ""}')
        extracted[ql] = row_values

    wb.close()
    print()

    # Step 3: Load sample reference
    sample_codes, sample_cells, sample_quarters = load_sample(SAMPLE_FILE)
    print()

    # Step 4: Compare extracted vs sample for overlapping quarters
    print('=' * 60)
    print(' COMPARISON: Extracted (values-only) vs Sample')
    print('=' * 60)

    overlap_quarters = [q for q in extracted if q in sample_quarters]
    print(f'Overlapping quarters: {overlap_quarters}')
    print()

    total_compared = 0
    total_match = 0
    total_mismatch = 0
    total_missing_in_extracted = 0
    total_extra_in_extracted = 0
    mismatches = []

    for ql in overlap_quarters:
        ext_row = extracted[ql]
        q_compared = 0
        q_match = 0
        q_mismatch = 0

        for col_idx in range(350):
            ext_val = ext_row.get(col_idx)
            sample_val = sample_cells.get((ql, col_idx))

            if ext_val is None and sample_val is None:
                continue

            total_compared += 1
            q_compared += 1

            if ext_val is not None and sample_val is not None:
                # Both have values — compare
                ext_clean = clean_float(ext_val) if isinstance(ext_val, (int, float)) else ext_val
                try:
                    s_float = float(sample_val)
                    e_float = float(ext_clean)
                    if abs(s_float - e_float) < 0.001:
                        q_match += 1
                        total_match += 1
                    else:
                        q_mismatch += 1
                        total_mismatch += 1
                        section_name = SECTION_HEADERS[col_idx // 50] if col_idx // 50 < 7 else '?'
                        mismatches.append({
                            'quarter': ql,
                            'col_idx': col_idx,
                            'section': section_name,
                            'extracted': ext_clean,
                            'sample': sample_val,
                            'diff': e_float - s_float,
                        })
                except (ValueError, TypeError):
                    if str(ext_clean) == str(sample_val):
                        q_match += 1
                        total_match += 1
                    else:
                        q_mismatch += 1
                        total_mismatch += 1
            elif ext_val is None:
                total_missing_in_extracted += 1
            else:
                total_extra_in_extracted += 1

        print(f'{ql}:  {q_compared} compared,  {q_match} match,  {q_mismatch} mismatch')

    print()
    print('-' * 60)
    print(f'TOTAL:')
    print(f'  Compared:     {total_compared}')
    print(f'  Matches:      {total_match}')
    print(f'  Mismatches:   {total_mismatch}')
    print(f'  Missing:      {total_missing_in_extracted} (in sample but not extracted)')
    print(f'  Extra:        {total_extra_in_extracted} (in extracted but not sample)')

    if total_compared > 0:
        rate = total_match / total_compared * 100
        print(f'  Match rate:   {rate:.2f}%')

    if mismatches:
        print()
        print(f'MISMATCH DETAILS (first 20):')
        for m in mismatches[:20]:
            pos_in_section = m['col_idx'] % 50
            metric_idx = pos_in_section // 10
            cat_idx = pos_in_section % 10
            metrics = ['NETSAVING', 'SUM', 'PCT', 'ASSET', 'ASSETPCT']
            metric = metrics[metric_idx] if metric_idx < 5 else '?'
            print(f'  {m["quarter"]}  col {m["col_idx"]:3d}  '
                  f'{m["section"][:20]:20s}  {metric:10s}  cat{cat_idx}  '
                  f'extracted={m["extracted"]}  sample={m["sample"]}  '
                  f'diff={m["diff"]:.6f}')

    print()
    print('=' * 60)

    # Also run our pipeline extractor for comparison
    print()
    print('Now running pipeline extractor on same file for cross-check...')
    sys.path.insert(0, BASE_DIR)
    import extractor
    pipeline_data = extractor.extract(SOURCE_2024, 2024)

    print()
    print('=' * 60)
    print(' CROSS-CHECK: Values-only vs Pipeline extractor')
    print('=' * 60)

    cross_match = 0
    cross_mismatch = 0
    cross_mismatches = []

    # Import config to get SIMBA codes
    import config
    for ql in extracted:
        if ql not in pipeline_data:
            print(f'  {ql}: not in pipeline data')
            continue

        ext_row = extracted[ql]
        pipe_row = pipeline_data[ql]

        for col_idx in range(min(350, len(config.DATA_COLUMNS))):
            ext_val = ext_row.get(col_idx)
            code = config.DATA_COLUMNS[col_idx]
            pipe_val = pipe_row.get(code)

            if ext_val is None and pipe_val is None:
                continue

            if ext_val is not None and pipe_val is not None:
                try:
                    e = float(ext_val)
                    p = float(pipe_val)
                    if abs(e - p) < 0.001:
                        cross_match += 1
                    else:
                        cross_mismatch += 1
                        cross_mismatches.append({
                            'quarter': ql, 'col_idx': col_idx,
                            'values_only': ext_val, 'pipeline': pipe_val,
                            'diff': e - p,
                        })
                except (ValueError, TypeError):
                    if str(ext_val) == str(pipe_val):
                        cross_match += 1
                    else:
                        cross_mismatch += 1
            else:
                # One is None, the other isn't — likely FP cleaning (0 vs None)
                cross_mismatch += 1
                cross_mismatches.append({
                    'quarter': ql, 'col_idx': col_idx,
                    'values_only': ext_val, 'pipeline': pipe_val,
                    'diff': None,
                })

    print(f'  Match:    {cross_match}')
    print(f'  Mismatch: {cross_mismatch}')

    if cross_mismatches:
        print(f'\n  Differences (first 20):')
        for m in cross_mismatches[:20]:
            print(f'    {m["quarter"]}  col {m["col_idx"]:3d}  '
                  f'values_only={m["values_only"]}  pipeline={m["pipeline"]}  '
                  f'diff={m["diff"]}')

    print()
    print('=' * 60)
    print('TEST COMPLETE')
    print('=' * 60)
    return 0


if __name__ == '__main__':
    sys.exit(main())
