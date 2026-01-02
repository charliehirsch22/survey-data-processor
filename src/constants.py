"""
Constants for Survey Data Processor.

This module contains all constant values used throughout the application.
"""

# Optional Windows-specific import
try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


# ============================================================================
# SHEET NAMES
# ============================================================================

SHEET_RAW_DATA = 'raw data'
SHEET_DATA_MAP = 'data map'
SHEET_COLUMN_QUESTION_MAP = 'column question map'
SHEET_LOOP_VARIABLES = 'loop variables'


# ============================================================================
# COLUMN WIDTHS
# ============================================================================

COL_WIDTH_NARROW = 3
COL_WIDTH_STANDARD = 13
COL_WIDTH_MEDIUM = 16
COL_WIDTH_WIDE = 20
COL_WIDTH_EXTRA_WIDE = 50


# ============================================================================
# COLORS
# ============================================================================

COLOR_PALE_BLUE = 'E6F3FF'


# ============================================================================
# COLUMN WIDTH CONFIGURATIONS FOR QUESTION TYPES
# ============================================================================

SINGLE_SELECT_WITH_OTHER_WIDTHS = {
    'A': 3, 'B': 3, 'C': 20, 'D': 16, 'E': 16,
    'F': 3, 'G': 16, 'H': 3, 'I': 16, 'J': 16,
    'K': 16, 'L': 16, 'M': 16, 'N': 16, 'O': 3,
    'P': 3, 'Q': 13
}

SINGLE_SELECT_WIDTHS = {
    'A': 3, 'B': 3, 'C': 20, 'D': 13, 'E': 13,
    'F': 3, 'G': 13, 'H': 3, 'I': 13, 'J': 13,
    'K': 13, 'L': 13, 'M': 13, 'N': 13
}
