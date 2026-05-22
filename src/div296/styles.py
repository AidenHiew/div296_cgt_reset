"""Colours, fills, fonts, number formats per spec §11."""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --- Colour palette (spec §11) ---
COLOUR_BAND_DARK_TEAL = "1D3B34"
COLOUR_INPUT_TEXT_BLUE = "0000FF"
COLOUR_CROSS_SHEET_GREEN = "008000"
COLOUR_INPUT_FILL_GREEN = "E1F5EE"
COLOUR_TRAP_FILL = "FBE9E9"
COLOUR_TRAP_TEXT = "A32D2D"
COLOUR_WIN_TEXT = "0F6E56"
COLOUR_COST_BASE_HEADER = "185FA5"
COLOUR_DIV296_HEADER = "0F6E56"

# --- Fonts ---
FONT_NAME = "Arial"
TITLE_FONT = Font(name=FONT_NAME, size=16, bold=True)
SECTION_BAND_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_NAME, size=10)
INPUT_FONT = Font(name=FONT_NAME, size=10, color=COLOUR_INPUT_TEXT_BLUE)
CROSS_SHEET_FONT = Font(name=FONT_NAME, size=10, color=COLOUR_CROSS_SHEET_GREEN)
TRAP_FONT = Font(name=FONT_NAME, size=10, color=COLOUR_TRAP_TEXT, bold=True)
WIN_FONT = Font(name=FONT_NAME, size=10, color=COLOUR_WIN_TEXT, bold=True)

# --- Fills ---
SECTION_BAND_FILL = PatternFill("solid", fgColor=COLOUR_BAND_DARK_TEAL)
INPUT_FILL = PatternFill("solid", fgColor=COLOUR_INPUT_FILL_GREEN)
TRAP_FILL = PatternFill("solid", fgColor=COLOUR_TRAP_FILL)

# --- Alignment ---
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")

# --- Borders ---
THIN = Side(style="thin", color="BFBFBF")
THIN_BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# --- Number formats ---
FMT_CURRENCY = '$#,##0;($#,##0);"-"'
FMT_PERCENT = "0.0%"
FMT_INT = "#,##0"
