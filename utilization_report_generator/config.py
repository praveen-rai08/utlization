"""
Configuration for the Utilization Report Generator
"""

# Working days per month (adjust for public holidays if needed)
WORKING_DAYS = {
    "Jan'26": 21, "Feb'26": 20, "Mar'26": 22,
    "Apr'26": 22, "May'26": 21, "Jun'26": 22,
    "Jul'26": 23, "Aug'26": 21, "Sep'26": 22,
    "Oct'26": 22, "Nov'26": 21, "Dec'26": 23,
}

# Sheet configuration:
# Each entry: (sheet_name, [(month_name, forecast_col_index, actual_col_index), ...])
# Column indices are 0-based (col A = 0)
SHEET_CONFIG = [
    ('2026-Jan-Feb-Mar', [
        ("Jan'26", 14, 15),
        ("Feb'26", 18, 19),
        ("Mar'26", 22, 23),
    ]),
    ('2026-Apr-May-Jun', [
        ("Apr'26", 14, 15),
        ("May'26", 18, 19),
        ("Jun'26", 22, 23),
    ]),
    ('2026-Jul-Sep', [
        ("Jul'26", 14, 15),
        ("Aug'26", 18, 19),
        ("Sep'26", 22, 23),
    ]),
    ('2026-Oct-Dec', [
        ("Oct'26", 14, 15),
        ("Nov'26", 18, 19),
        ("Dec'26", 22, 23),
    ]),
]

# Utilization thresholds
UTIL_HIGH = 90      # >= HIGH -> Green
UTIL_MEDIUM = 80    # >= MEDIUM -> Yellow, else Red

# Data row start index (0-based): row 0 = month header, row 1 = column header, row 2+ = data
DATA_START_ROW = 2

# Column indices for employee info (0-based)
COL_ASSOC_ID = 1
COL_ASSOC_NAME = 2
COL_GRADE = 3
COL_PROJECT = 5
COL_ACCOUNT = 7
COL_BILLABILITY = 10
COL_COUNTRY = 11
COL_OO = 12  # Onsite/Offshore
COL_CITY = 13

# Color palette
COLOR_DARK_BLUE = "1F3864"
COLOR_MID_BLUE = "2E75B6"
COLOR_LIGHT_BLUE = "BDD7EE"
COLOR_VERY_LIGHT = "DEEAF1"
COLOR_GREEN = "70AD47"
COLOR_LIGHT_GREEN = "E2EFDA"
COLOR_YELLOW = "FFD966"
COLOR_LIGHT_YEL = "FFF2CC"
COLOR_RED = "FF0000"
COLOR_LIGHT_RED = "FFE0E0"
COLOR_WHITE = "FFFFFF"
COLOR_GREY = "F2F2F2"
COLOR_DARK_GREY = "595959"
