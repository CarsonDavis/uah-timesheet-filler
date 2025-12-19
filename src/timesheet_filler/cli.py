"""Command-line interface for timesheet filler."""

import argparse
from pathlib import Path

from timesheet_filler.config import load_config
from timesheet_filler.exporter import export_sheets_to_pdf
from timesheet_filler.filler import fill_timesheet


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Fill UAH ESSC timesheets and export to PDF"
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("ESSC Labor Report Timesheet - 2026 BLANK.xlsx"),
        help="Path to blank timesheet template",
    )
    parser.add_argument(
        "--details",
        type=Path,
        default=Path("details.toml"),
        help="Path to details TOML file",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory to save PDFs",
    )
    parser.add_argument(
        "--year",
        default="26",
        help="Year suffix for PDF naming (default: 26)",
    )

    args = parser.parse_args()

    # Validate inputs
    if not args.template.exists():
        print(f"Error: Template not found: {args.template}")
        return

    if not args.details.exists():
        print(f"Error: Details file not found: {args.details}")
        print("Hint: Copy details.example.toml to details.toml and fill in your info")
        return

    # Load config
    print(f"Loading config from {args.details}...")
    config = load_config(args.details)
    print(f"  Name: {config.name} {config.last_name}")
    print(f"  A#: {config.a_number}")
    print(f"  Labor entries: {len(config.labor)}")

    # Fill timesheet
    filled_path = args.output_dir / "filled_timesheet.xlsx"
    args.output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nFilling timesheet from {args.template}...")
    fill_timesheet(args.template, config, filled_path)
    print(f"  Saved filled workbook to {filled_path}")

    # Export PDFs
    print(f"\nExporting sheets to PDF...")
    try:
        pdf_paths = export_sheets_to_pdf(
            filled_path,
            args.output_dir,
            config.last_name,
            args.year,
        )
        print(f"  Exported {len(pdf_paths)} PDFs to {args.output_dir}/")
        for pdf in pdf_paths[:3]:
            print(f"    - {pdf.name}")
        if len(pdf_paths) > 3:
            print(f"    ... and {len(pdf_paths) - 3} more")
    except RuntimeError as e:
        print(f"\nError: {e}")
        return

    print("\nDone!")


if __name__ == "__main__":
    main()
