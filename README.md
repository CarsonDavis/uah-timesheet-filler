# Timesheet Filler

Automates filling UAH ESSC Excel timesheets and exporting them as PDFs.

## Overview

This tool takes a blank Excel timesheet template and a details file containing personal information, fills in the first sheet with the provided data (which auto-populates the remaining sheets via formulas), and exports each sheet as a PDF with a specific naming convention.

## Requirements

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) for package management
- **Microsoft Excel** (macOS) or **LibreOffice** for PDF export

## Setup

1. Clone this repo and install dependencies:
   ```bash
   git clone git@github.com:CarsonDavis/uah-timesheet-filler.git
   cd uah-timesheet-filler
   uv sync
   ```

2. Copy the example details file:
   ```bash
   cp details.example.yaml details.yaml
   ```

3. Edit `details.yaml` with your information

## Usage

```bash
uv run timesheet-filler
```

This will:
1. Read your details from `details.yaml`
2. Fill in the blank timesheet template
3. Export all 26 pay periods as separate PDFs to `output/`

### Options

```bash
uv run timesheet-filler --help
uv run timesheet-filler --template path/to/template.xlsx
uv run timesheet-filler --details path/to/details.yaml
uv run timesheet-filler --output-dir path/to/output
uv run timesheet-filler --year 27  # For fiscal year 2027
```

## PDF Naming Convention

Files are named: `{LastName} {BWPeriod}-{Year}.pdf`

Examples:
- BW1: `Smith 1-26.pdf`
- BW26: `Smith 26-26.pdf`

## Upload

Completed timesheets are uploaded here:
https://docs.google.com/forms/d/e/1FAIpQLScCJ4kWFv0Qf-jHWDf_5NKYtlvG_5Oa46kC-NQ7utZvsubGVQ/viewform

## Fields

The following fields are filled from the details file:

### Header Information
| Field | Example |
|-------|---------|
| Name | Carson |
| A# | A12345678 |
| Position # | 342973 |
| Title | Research Scientist III Step 3 |
| FTE | 1.00 |

### Labor Distribution Table
| 6 digit Org/Index | Account Code | % Distribution |
|-------------------|--------------|----------------|
| 745A9M | 6150 | 0.5 (50%) |
| 745A9B | 6150 | 0.5 (50%) |

Note: Percentages are decimals (0.5 = 50%) and should sum to 1.0.

### Auto-filled
- **Employee Signature Date**: Today's date (on all sheets)
- **Pay Period / Check Date**: Calculated from template formulas
