"""
SIFA Compare Configuration
===========================
No file paths needed here - they are discovered automatically.

Edit only tolerances and display settings below.
"""

# =================================================================
# COMPARISON SETTINGS
# =================================================================

# Tolerance for floating point comparison
FLOAT_TOLERANCE = 0.001

# =================================================================
# DISPLAY SETTINGS
# =================================================================

# Show detailed progress during comparison
VERBOSE = True

# Maximum number of top-mismatch examples to show in summary report
MAX_CONSOLE_EXAMPLES = 5

# =================================================================
# REPORT FILENAMES
# =================================================================

REPORT_MISMATCHES  = 'report_mismatches.csv'
REPORT_MISSING     = 'report_missing.csv'
REPORT_NEW_VALUES  = 'report_new_values.csv'
SUMMARY_FILENAME   = 'comparison_summary.txt'

# =================================================================
# FILES SUBFOLDER NAMES
# =================================================================

FILES_DIR            = 'files'
MISMATCH_FILES_DIR   = 'mismatch'
MISSING_FILES_DIR    = 'missing'
NEW_VALUES_FILES_DIR = 'new_values'
