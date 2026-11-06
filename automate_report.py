#!/usr/bin/env python3
"""Create a cleaned sales workbook and regional revenue summary."""

import argparse
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import tempfile
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException

COLUMNS = ('Order ID', 'Date', 'Region', 'Product', 'Units', 'Unit Price')


def number(value, row, column):
    """Accept finite numeric values, including numeric text; reject missing values."""
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f'Row {row}: {column} must be numeric, got {value!r}') from None
    if not result.is_finite():
        raise ValueError(f'Row {row}: {column} must be finite')
    return result


def create_report(input_path, output_path):
    """Read the first worksheet and write a separate report. Return counts and totals."""
    source = Path(input_path).resolve()
    target = Path(output_path).resolve()
    if source == target or (target.exists() and os.path.samefile(source, target)):
        raise ValueError('Input and output must be different files')
    if target.suffix.lower() != '.xlsx':
        raise ValueError('Output path must end in .xlsx')

    report = Workbook()
    cleaned = report.active
    cleaned.title = 'Cleaned Data'
    cleaned.append((*COLUMNS, 'Revenue'))
    totals = defaultdict(Decimal)
    count = 0
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        rows = workbook.worksheets[0].iter_rows(values_only=True)
        headers = next(rows, ())
        for name in COLUMNS:
            if headers.count(name) != 1:
                raise ValueError(f'Expected exactly one {name!r} column in the first row')
        indices = [headers.index(name) for name in COLUMNS]
        for row_number, row in enumerate(rows, start=2):
            if all(value is None for value in row):
                continue
            values = [row[index] for index in indices]
            region = values[2]
            if not isinstance(region, str) or not region.strip():
                raise ValueError(f'Row {row_number}: Region must be nonempty text')
            values[2] = ' '.join(region.split()).title()
            units = number(values[4], row_number, 'Units')
            price = number(values[5], row_number, 'Unit Price')
            revenue = units * price
            values[4:6] = [units, price]
            cleaned.append((*values, revenue))
            # Preserve literal text instead of treating leading '=' as a formula.
            for cell in cleaned[cleaned.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
            cleaned.cell(cleaned.max_row, 2).number_format = 'yyyy-mm-dd'
            for column in (6, 7):
                cleaned.cell(cleaned.max_row, column).number_format = '#,##0.00'
            totals[values[2]] += revenue
            count += 1
    finally:
        workbook.close()

    summary = report.create_sheet('Regional Summary')
    summary.append(('Region', 'Total Revenue'))
    for region, total in sorted(totals.items()):
        summary.append((region, total))
        summary.cell(summary.max_row, 1).data_type = 's'
    overall = sum(totals.values(), Decimal(0))
    summary.append(('Overall Total', overall))
    for row in summary.iter_rows(min_row=2, min_col=2):
        row[0].number_format = '#,##0.00'
    for sheet in report:
        sheet.freeze_panes = 'A2'
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = 20
    cleaned.auto_filter.ref = cleaned.dimensions
    for cell in summary[summary.max_row]:
        cell.font = Font(bold=True)

    # Save atomically so a failed write does not leave a partial report.
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, suffix='.xlsx', delete=False) as file:
            temporary = Path(file.name)
        report.save(temporary)
        os.replace(temporary, target)
    finally:
        report.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    return count, dict(totals), overall


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path, help='Source Excel workbook (.xlsx)')
    parser.add_argument('output', type=Path, help='Destination report (.xlsx)')
    args = parser.parse_args()
    try:
        count, totals, overall = create_report(args.input, args.output)
    except (OSError, ValueError, BadZipFile, InvalidFileException) as error:
        parser.exit(1, f'Error: {error}\n')
    print(f'Created {args.output}: {count} transactions')
    for region, total in sorted(totals.items()):
        print(f'{region}: {total:,.2f}')
    print(f'Overall Total: {overall:,.2f}')


if __name__ == '__main__':
    main()
