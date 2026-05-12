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

# Threshold to separate decimal mismatches from full mismatches
# abs(diff) <= threshold  ->  decimal mismatch (rounding)
# abs(diff) >  threshold  ->  full mismatch (wrong value)
DECIMAL_MISMATCH_THRESHOLD = 1.0

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

REPORT_MISMATCHES_FULL    = 'report_mismatches_full.csv'
REPORT_MISMATCHES_DECIMAL = 'report_mismatches_decimal.csv'
REPORT_MISSING            = 'report_missing.csv'
REPORT_NEW_VALUES         = 'report_new_values.csv'
SUMMARY_FILENAME          = 'comparison_summary.txt'

