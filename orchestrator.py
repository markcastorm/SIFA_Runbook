# orchestrator.py
# Wires the full SIFA pipeline: download → extract → generate.

import sys
import logging

import config
from scraper import download
from extractor import extract_multiple
from file_generator import FileGenerator

logger = logging.getLogger(__name__)


def main():
    """Run the full pipeline. Returns 0 on success, 1 on failure."""
    logging.basicConfig(
        stream=sys.stdout,
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    for noisy in ('selenium', 'selenium.webdriver', 'urllib3',
                   'urllib3.connectionpool'):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    try:
        logger.info('=== SIFA pipeline started ===')
        logger.info(f'Timestamp: {config.RUN_TIMESTAMP}')
        logger.info(f'Source:    {config.BASE_URL}')
        logger.info(f'Master:   {config.MASTER_FILE}')

        # ── Step 1: Download ─────────────────────────────────────────────
        logger.info('Step 1: Downloading Excel files from fondbolagen.se...')
        file_list = download()

        if not file_list:
            logger.info('No new files downloaded — pipeline finished')
            return 0

        logger.info(f'Downloaded {len(file_list)} files: '
                     f'{[f"{y}" for y, _ in file_list]}')

        # ── Step 2: Extract / Transform ──────────────────────────────────
        logger.info('Step 2: Extracting and transforming data...')
        new_data = extract_multiple(file_list)

        if not new_data:
            logger.warning('No data extracted — aborting')
            return 1

        logger.info(f'Extracted {len(new_data)} quarters: '
                     f'{list(new_data.keys())}')

        # ── Step 3: Generate output files ────────────────────────────────
        logger.info('Step 3: Generating output files...')
        generator = FileGenerator()
        output_files = generator.generate_files(new_data, config.OUTPUT_RUN_DIR)

        # ── Summary ──────────────────────────────────────────────────────
        logger.info('=== SIFA pipeline completed successfully ===')
        logger.info(f'Output dir:  {config.OUTPUT_RUN_DIR}')
        logger.info(f'Latest dir:  {config.LATEST_OUTPUT_DIR}')
        logger.info(f'DATA: {output_files["data_file"]}')
        logger.info(f'META: {output_files["meta_file"]}')
        logger.info(f'ZIP:  {output_files["zip_file"]}')
        logger.info(f'CSV:  {output_files["csv_file"]}')

        return 0

    except Exception as e:
        logger.exception(f'Pipeline failed: {e}')
        return 1
