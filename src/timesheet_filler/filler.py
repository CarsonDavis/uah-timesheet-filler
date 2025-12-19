"""Excel timesheet filling logic."""

from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from timesheet_filler.config import PersonConfig


# Cell mappings for the first sheet (2026-01)
# Other sheets reference these via formulas
CELL_MAP = {
    "name": "E6",
    "a_number": "J6",
    "position_number": "O6",
    "title": "R6",
    "fte": "R8",
    "employee_signature": "E30",
    "employee_signature_date": "R30",
}

# Fancy script font for signature
SIGNATURE_FONT = Font(
    name="Brush Script MT",
    size=14,
    italic=True,
)

# Labor distribution starts at row 15, columns C, D, E
LABOR_START_ROW = 15
LABOR_ORG_COL = "C"
LABOR_ACCOUNT_COL = "D"
LABOR_PERCENT_COL = "E"


def fill_timesheet(template_path: Path, config: PersonConfig, output_path: Path) -> Path:
    """Fill in the timesheet template with person config data.

    Args:
        template_path: Path to the blank Excel template
        config: Person configuration with data to fill
        output_path: Where to save the filled workbook

    Returns:
        Path to the saved workbook
    """
    wb = openpyxl.load_workbook(template_path)

    # Get the first sheet - all others reference it
    first_sheet = wb.worksheets[0]

    # Fill in personal info
    first_sheet[CELL_MAP["name"]] = config.name
    first_sheet[CELL_MAP["a_number"]] = config.a_number
    first_sheet[CELL_MAP["position_number"]] = config.position_number
    first_sheet[CELL_MAP["title"]] = config.title
    first_sheet[CELL_MAP["fte"]] = config.fte

    # Fill in labor distribution
    for i, labor in enumerate(config.labor):
        row = LABOR_START_ROW + i
        first_sheet[f"{LABOR_ORG_COL}{row}"] = labor.org_index
        first_sheet[f"{LABOR_ACCOUNT_COL}{row}"] = labor.account_code
        # Convert percent to number (e.g., "50" -> 50, ".5" -> 0.5)
        try:
            percent_val = float(labor.percent)
        except ValueError:
            percent_val = labor.percent
        first_sheet[f"{LABOR_PERCENT_COL}{row}"] = percent_val

    # Fill in signature and date on ALL sheets
    today = date.today().strftime("%Y.%m.%d")

    for sheet in wb.worksheets:
        # Add signature with fancy font
        sig_cell = sheet[CELL_MAP["employee_signature"]]
        sig_cell.value = config.name
        sig_cell.font = SIGNATURE_FONT

        # Add date
        sheet[CELL_MAP["employee_signature_date"]] = today

    # Save the filled workbook
    wb.save(output_path)
    return output_path
