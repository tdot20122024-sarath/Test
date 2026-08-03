#!/usr/bin/env python3
"""
Convert a large CSV file into a multi-sheet .xlsx workbook.

Excel worksheets are capped at 1,048,576 rows. When a CSV has more data rows
than that, this script splits the data across consecutive sheets (Data_1,
Data_2, ...), repeating the header row on each sheet, and streams the file
in write-only mode so multi-million-row CSVs don't need to fit in memory.

Usage:
    python csv_to_multisheet_xlsx.py input.csv output.xlsx [--rows-per-sheet N] [--no-header]

Example:
    python csv_to_multisheet_xlsx.py big_data.csv big_data.xlsx
"""

import argparse
import csv
import sys

from openpyxl import Workbook

# Excel's hard limit is 1,048,576 rows per sheet. Reserve one for the header
# row that gets repeated on every sheet.
EXCEL_MAX_ROWS_PER_SHEET = 1_048_576
DEFAULT_ROWS_PER_SHEET = EXCEL_MAX_ROWS_PER_SHEET - 1


def convert(input_csv, output_xlsx, rows_per_sheet=DEFAULT_ROWS_PER_SHEET, has_header=True):
    if rows_per_sheet < 1 or rows_per_sheet > EXCEL_MAX_ROWS_PER_SHEET - 1:
        raise ValueError(
            f"rows_per_sheet must be between 1 and {EXCEL_MAX_ROWS_PER_SHEET - 1}"
        )

    wb = Workbook(write_only=True)

    with open(input_csv, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)

        header = next(reader) if has_header else None

        ws = None
        rows_in_sheet = 0
        sheet_index = 0
        total_rows = 0

        for row in reader:
            if ws is None or rows_in_sheet >= rows_per_sheet:
                sheet_index += 1
                ws = wb.create_sheet(title=f"Data_{sheet_index}")
                if header:
                    ws.append(header)
                rows_in_sheet = 0

            ws.append(row)
            rows_in_sheet += 1
            total_rows += 1

    if sheet_index == 0:
        # Empty CSV (header only or completely empty) - still produce a sheet.
        ws = wb.create_sheet(title="Data_1")
        if header:
            ws.append(header)

    wb.save(output_xlsx)
    return total_rows, sheet_index


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_csv", help="Path to the source CSV file")
    parser.add_argument("output_xlsx", help="Path to write the output .xlsx file")
    parser.add_argument(
        "--rows-per-sheet",
        type=int,
        default=DEFAULT_ROWS_PER_SHEET,
        help=f"Data rows per sheet, excluding header (default {DEFAULT_ROWS_PER_SHEET}, "
        f"Excel's max minus 1 for the header row)",
    )
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="Treat the CSV as having no header row",
    )
    args = parser.parse_args()

    total_rows, sheet_count = convert(
        args.input_csv,
        args.output_xlsx,
        rows_per_sheet=args.rows_per_sheet,
        has_header=not args.no_header,
    )

    print(f"Wrote {total_rows:,} data rows across {sheet_count} sheet(s) to {args.output_xlsx}")


if __name__ == "__main__":
    sys.exit(main())
