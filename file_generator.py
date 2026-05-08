# file_generator.py
# Generate DATA (.xlsx), META (.xlsx), ZIP output files for SIFA.
# Also updates the master CSV with new quarterly data.
# Uses openpyxl (not xlwt) because SIFA has 350 columns (exceeds .xls 256-col limit).

import os
import csv
import shutil
import zipfile
import logging

import openpyxl
import pandas as pd

import config

logger = logging.getLogger(__name__)


class FileGenerator:
    """
    Generates SIMBA-standard output files:
      - DATA file: SIMBA codes as headers, quarters as rows
      - META file: dataset metadata for each SIMBA code
      - ZIP file: contains both DATA and META
      - Master CSV: cumulative data with new quarters merged in
    Output goes to timestamped folder + 'latest' folder.
    """

    def __init__(self):
        self.logger = logger

    # ─────────────────────────────────────────────────────────────────────
    # DATA file
    # ─────────────────────────────────────────────────────────────────────

    def create_data_file(self, df, output_path):
        """
        Create the DATA Excel file (.xlsx).

        Layout:
            Row 1: SIMBA codes (column headers)
            Row 2: Human-readable descriptions
            Row 3+: data rows (one per quarter, col A = 'YYYY-QN')
        """
        self.logger.info('Creating DATA file...')

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'DATA'

        # Row 1: SIMBA codes (col A empty, cols B+ = codes)
        for col_idx, code in enumerate(config.DATA_COLUMNS):
            ws.cell(row=1, column=col_idx + 2, value=code)

        # Row 2: descriptions
        for col_idx, desc in enumerate(config.DATA_DESCRIPTIONS):
            ws.cell(row=2, column=col_idx + 2, value=desc)

        # Data rows
        for row_idx, (quarter_label, row) in enumerate(df.iterrows()):
            ws.cell(row=row_idx + 3, column=1, value=str(quarter_label))
            for col_idx, code in enumerate(config.DATA_COLUMNS):
                value = row.get(code)
                if pd.notna(value):
                    try:
                        ws.cell(row=row_idx + 3, column=col_idx + 2,
                                value=float(value))
                    except (ValueError, TypeError):
                        ws.cell(row=row_idx + 3, column=col_idx + 2,
                                value=str(value))

        wb.save(output_path)
        self.logger.info(f'DATA file saved: {output_path}  |  '
                         f'{len(df)} rows x {len(config.DATA_COLUMNS)} cols')
        return output_path

    # ─────────────────────────────────────────────────────────────────────
    # META file
    # ─────────────────────────────────────────────────────────────────────

    def create_meta_file(self, output_path):
        """
        Create the META Excel file with one row per SIMBA code.
        """
        self.logger.info('Creating META file...')

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'META'

        # Header row
        for col_idx, col_name in enumerate(config.METADATA_COLUMNS):
            ws.cell(row=1, column=col_idx + 1, value=col_name)

        # One row per SIMBA code
        for row_idx, (code, desc) in enumerate(
            zip(config.DATA_COLUMNS, config.DATA_DESCRIPTIONS)
        ):
            row_data = {
                'CODE': code,
                'CODE_MNEMONIC': code.split('.')[0] if '.' in code else code,
                'DESCRIPTION': desc,
            }
            for key, value in config.METADATA_DEFAULTS.items():
                if key not in row_data:
                    row_data[key] = value

            for col_idx, col_name in enumerate(config.METADATA_COLUMNS):
                value = row_data.get(col_name, '')
                ws.cell(row=row_idx + 2, column=col_idx + 1, value=value)

        wb.save(output_path)
        self.logger.info(f'META file saved: {output_path}  |  '
                         f'{len(config.DATA_COLUMNS)} codes')
        return output_path

    # ─────────────────────────────────────────────────────────────────────
    # ZIP file
    # ─────────────────────────────────────────────────────────────────────

    def create_zip_file(self, data_file, meta_file, zip_path):
        """Bundle DATA and META files into a single ZIP."""
        self.logger.info(f'Creating ZIP: {zip_path}')

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.write(data_file, os.path.basename(data_file))
            zf.write(meta_file, os.path.basename(meta_file))

        self.logger.info('ZIP created')
        return zip_path

    # ─────────────────────────────────────────────────────────────────────
    # Master CSV update
    # ─────────────────────────────────────────────────────────────────────

    def update_master(self, new_data_dict):
        """
        Merge new quarterly data into the master CSV.

        new_data_dict: dict mapping 'YYYY-QN' -> {simba_code: value, ...}

        Logic:
        - Load existing master (if it exists)
        - For each new quarter:
          - If quarter doesn't exist in master: add it
          - If quarter exists but has empty cells: fill them in
          - If quarter exists and has data: skip (don't overwrite)
        - Sort by quarter label
        - Save back
        """
        master_path = config.MASTER_FILE
        self.logger.info(f'Updating master: {master_path}')
        os.makedirs(os.path.dirname(master_path), exist_ok=True)

        # Load existing master (skip description row if present)
        if os.path.exists(master_path):
            master_df = self._read_master_csv(master_path)
            self.logger.info(f'Existing master: {len(master_df)} rows')
        else:
            master_df = pd.DataFrame(columns=config.DATA_COLUMNS)
            master_df.index.name = ''
            self.logger.info('No existing master — creating new')

        # Merge new data
        new_count = 0
        updated_count = 0

        for quarter_label, row_values in new_data_dict.items():
            if quarter_label in master_df.index:
                # Quarter exists — fill only empty cells
                existing_row = master_df.loc[quarter_label]
                filled = 0
                for col, val in row_values.items():
                    if val is not None and col in master_df.columns:
                        if pd.isna(existing_row.get(col)) or existing_row.get(col) == '':
                            master_df.at[quarter_label, col] = val
                            filled += 1

                if filled > 0:
                    updated_count += 1
                    self.logger.info(f'{quarter_label}: filled {filled} empty cells')
                else:
                    self.logger.info(f'{quarter_label}: already complete — skipped')
            else:
                # New quarter — add row
                new_row = pd.Series(index=config.DATA_COLUMNS, dtype=object)
                for col, val in row_values.items():
                    if col in new_row.index:
                        new_row[col] = val
                new_row.name = quarter_label
                master_df = pd.concat([master_df, new_row.to_frame().T])
                new_count += 1
                self.logger.info(f'{quarter_label}: added as new row')

        # Sort by quarter label (chronological)
        master_df.sort_index(inplace=True)

        # Save with description sub-header row
        self._write_master_csv(master_df, master_path)
        self.logger.info(f'Master saved: {new_count} new, {updated_count} updated, '
                         f'{len(master_df)} total rows')

        return master_df

    @staticmethod
    def _read_master_csv(path):
        """
        Read master CSV, skipping the description sub-header row (row 2)
        if it exists.
        """
        # Peek at second row to check if it's a description row
        with open(path, 'r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            header = next(reader, None)
            second = next(reader, None)

        skiprows = []
        if second and len(second) > 1:
            # If the second row's first data value contains ':'
            # (e.g. "All types of funds: ..."), it's a description row
            first_val = second[1] if len(second) > 1 else ''
            if ':' in str(first_val):
                skiprows = [1]  # skip the description row (0-indexed after header)

        df = pd.read_csv(path, index_col=0, skiprows=skiprows)
        # Drop any rows with NaN index
        df = df[df.index.notna()]
        return df

    @staticmethod
    def _write_master_csv(df, path):
        """
        Write master CSV with a description sub-header row (row 2)
        matching the sample DATA output format.
        """
        with open(path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Row 1: empty index col + SIMBA codes
            writer.writerow([''] + list(config.DATA_COLUMNS))
            # Row 2: empty index col + descriptions
            writer.writerow([''] + list(config.DATA_DESCRIPTIONS))
            # Data rows
            for idx in df.index:
                row = [str(idx)]
                for col in config.DATA_COLUMNS:
                    val = df.at[idx, col]
                    if pd.isna(val):
                        row.append('')
                    else:
                        row.append(val)
                writer.writerow(row)

    # ─────────────────────────────────────────────────────────────────────
    # Generate all outputs
    # ─────────────────────────────────────────────────────────────────────

    def generate_files(self, new_data_dict, output_dir):
        """
        Generate DATA, META, ZIP files. Update master CSV.

        Args:
            new_data_dict: dict mapping 'YYYY-QN' -> {simba_code: value}
            output_dir: timestamped output directory

        Returns:
            dict with paths to all created files
        """
        os.makedirs(output_dir, exist_ok=True)

        # Update master first (merges new data)
        master_df = self.update_master(new_data_dict)

        timestamp = config.RUN_TIMESTAMP

        data_filename = config.DATA_FILE_PATTERN.format(timestamp=timestamp)
        meta_filename = config.META_FILE_PATTERN.format(timestamp=timestamp)
        zip_filename  = config.ZIP_FILE_PATTERN.format(timestamp=timestamp)

        data_path = os.path.join(output_dir, data_filename)
        meta_path = os.path.join(output_dir, meta_filename)
        zip_path  = os.path.join(output_dir, zip_filename)

        # Build a DataFrame for the new quarters only (for the DATA xlsx)
        new_df = pd.DataFrame(columns=config.DATA_COLUMNS)
        for ql, rv in new_data_dict.items():
            row = pd.Series(index=config.DATA_COLUMNS, dtype=object)
            for col, val in rv.items():
                if col in row.index:
                    row[col] = val
            row.name = ql
            new_df = pd.concat([new_df, row.to_frame().T])
        new_df.sort_index(inplace=True)

        # Create output files
        self.create_data_file(new_df, data_path)
        self.create_meta_file(meta_path)
        self.create_zip_file(data_path, meta_path, zip_path)

        # Also save full combined CSV (with description row)
        csv_path = os.path.join(output_dir, f'{config.DATASET_NAME}_DATA.csv')
        self._write_master_csv(master_df, csv_path)
        self.logger.info(f'Full CSV saved: {csv_path}')

        # Copy to 'latest' folder
        latest_dir = config.LATEST_OUTPUT_DIR
        os.makedirs(latest_dir, exist_ok=True)

        latest_data = os.path.join(latest_dir, f'{config.DATASET_NAME}_DATA_latest.xlsx')
        latest_meta = os.path.join(latest_dir, f'{config.DATASET_NAME}_META_latest.xlsx')
        latest_zip  = os.path.join(latest_dir, f'{config.DATASET_NAME}_latest.zip')
        latest_csv  = os.path.join(latest_dir, f'{config.DATASET_NAME}_DATA.csv')

        shutil.copy2(data_path, latest_data)
        shutil.copy2(meta_path, latest_meta)
        shutil.copy2(zip_path, latest_zip)
        shutil.copy2(csv_path, latest_csv)

        self.logger.info(f'Files copied to latest: {latest_dir}')

        return {
            'data_file':   data_path,
            'meta_file':   meta_path,
            'zip_file':    zip_path,
            'csv_file':    csv_path,
            'latest_data': latest_data,
            'latest_meta': latest_meta,
            'latest_zip':  latest_zip,
            'latest_csv':  latest_csv,
        }
