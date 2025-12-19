# Timesheet Filler

Automates filling UAH ESSC Excel timesheets and exporting them as PDFs.

## Overview

This tool takes a blank Excel timesheet template and a details file containing personal information, fills in the first sheet with the provided data (which auto-populates the remaining sheets via formulas), and exports each sheet as a PDF with a specific naming convention.

## Project Structure

```
timesheet-filler/
├── src/timesheet_filler/   # Source code
├── details.toml            # Your personal information (copy from example)
├── details.example.toml    # Example details file
├── template.xlsx           # Blank timesheet template
├── pyproject.toml
└── README.md
```

## Setup

1. Copy the example details file:
   ```bash
   cp details.example.toml details.toml
   ```

2. Edit `details.toml` with your information

## Fields

The following fields are filled from the details file:

### Header Information
| Field | Example |
|-------|---------|
| Payroll ID | 2025-22 |
| Name | Carson |
| A# | A12345678 |
| Position # | 342973 |
| Title | Research Scientist III Step 3 |
| Department | Earth System Science Center |
| Pay Period | 10/01/25 to 10/14/25 |
| FTE | 1.00 |
| Home Labor | 740001 |
| Check Date | 10/24/25 |

### Labor Distribution Table
| 6 digit Org/Index | Account Code | % Distribution |
|-------------------|--------------|----------------|
| 745A9M | 6150 | 50% |
| 745A9B | 6150 | 50% |

### Signature
- **Employee Signature Date**: Auto-filled with today's date

## PDF Naming Convention

Files are named: `{LastName} {BWPeriod}-{Year}`

Examples:
- BW1: `Smith 1-26`
- BW26: `Smith 26-26`

## Upload

Completed timesheets are uploaded here:
https://docs.google.com/forms/d/e/1FAIpQLScCJ4kWFv0Qf-jHWDf_5NKYtlvG_5Oa46kC-NQ7utZvsubGVQ/viewform

## Installation

```bash
uv sync
```

## Usage

```bash
uv run timesheet-filler
```

## Development

This project uses:
- **uv** for package management
- **ruff** for linting
- **pyright** for type checking
