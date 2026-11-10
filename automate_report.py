#!/usr/bin/env python3
"""Create a cleaned sales workbook and regional revenue summary."""

import argparse
import os
import tempfile
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils.exceptions import InvalidFileException

COLUMNS = ('Order ID', 'Date', 'Region', 'Product', 'Units', 'Unit Price')
CURRENCY_FORMAT = '#,##0.00'


def parse_finite_decimal(value, row, column):
    """Accept finite numeric values, including numeric text; reject missing values."""
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f'Row {row}: {column} must be numeric, got {value!r}') from None
    if not result.is_finite():
        raise ValueError(f'Row {row}: {column} must be finite')
    return result


def validate_report_paths(input_path, output_path):
    """Resolve paths and reject an unsafe or unsupported destination."""
    source = Path(input_path).resolve()
    target = Path(output_path).resolve()
    if source == target or (target.exists() and os.path.samefile(source, target)):
        raise ValueError('Input and output must be different files')
    if target.suffix.lower() != '.xlsx':
        raise ValueError('Output path must end in .xlsx')
    return source, target


def validate_headers(headers):
    """Return required column positions after checking for missing or duplicate names."""
    for name in COLUMNS:
        if headers.count(name) != 1:
            raise ValueError(f'Expected exactly one {name!r} column in the first row')
    return [headers.index(name) for name in COLUMNS]


def normalize_transaction(values, row_number):
    """Validate a transaction and return its cleaned values and revenue."""
    order_id, date, region, product, raw_units, raw_price = values
    if not isinstance(region, str) or not region.strip():
        raise ValueError(f'Row {row_number}: Region must be nonempty text')
    region = ' '.join(region.split()).title()
    units = parse_finite_decimal(raw_units, row_number, 'Units')
    if units < 0:
        raise ValueError(f'Row {row_number}: Units must be nonnegative, got {raw_units!r}')
    unit_price = parse_finite_decimal(raw_price, row_number, 'Unit Price')
    revenue = units * unit_price
    return (order_id, date, region, product, units, unit_price), revenue


def format_cleaned_row(sheet):
    """Preserve literal text and apply transaction date and currency formats."""
    row_number = sheet.max_row
    # Preserve literal text instead of treating leading '=' as a formula.
    for cell in sheet[row_number]:
        if isinstance(cell.value, str):
            cell.data_type = 's'
    sheet.cell(row_number, 2).number_format = 'yyyy-mm-dd'
    for column in (6, 7):
        sheet.cell(row_number, column).number_format = CURRENCY_FORMAT


def process_transactions(source, cleaned):
    """Read transactions into the cleaned worksheet and accumulate regional totals."""
    regional_totals = defaultdict(Decimal)
    transaction_count = 0
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        rows = workbook.worksheets[0].iter_rows(values_only=True)
        indices = validate_headers(next(rows, ()))
        for row_number, row in enumerate(rows, start=2):
            if all(value is None for value in row):
                continue
            values, revenue = normalize_transaction(
                [row[index] for index in indices], row_number,
            )
            cleaned.append((*values, revenue))
            format_cleaned_row(cleaned)
            regional_totals[values[2]] += revenue
            transaction_count += 1
    finally:
        workbook.close()
    return transaction_count, dict(regional_totals)


def write_regional_summary(report, regional_totals):
    """Add the sorted regional summary and return its sheet and overall total."""
    summary = report.create_sheet('Regional Summary')
    summary.append(('Region', 'Total Revenue'))
    for region, total in sorted(regional_totals.items()):
        summary.append((region, total))
        summary.cell(summary.max_row, 1).data_type = 's'
    overall = sum(regional_totals.values(), Decimal(0))
    summary.append(('Overall Total', overall))
    return summary, overall


def format_report(report, cleaned, summary):
    """Apply summary formatting and shared worksheet presentation."""
    for row in summary.iter_rows(min_row=2, min_col=2):
        row[0].number_format = CURRENCY_FORMAT
    for sheet in report:
        sheet.freeze_panes = 'A2'
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = 20
    cleaned.auto_filter.ref = cleaned.dimensions
    for cell in summary[summary.max_row]:
        cell.font = Font(bold=True)


def save_report_atomically(report, target):
    """Replace the destination only after a successful save and clean up resources."""
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


def create_report(input_path, output_path):
    """Read the first worksheet and write a separate report. Return counts and totals."""
    source, target = validate_report_paths(input_path, output_path)
    report = Workbook()
    cleaned = report.active
    cleaned.title = 'Cleaned Data'
    cleaned.append((*COLUMNS, 'Revenue'))
    transaction_count, regional_totals = process_transactions(source, cleaned)
    summary, overall = write_regional_summary(report, regional_totals)
    format_report(report, cleaned, summary)
    save_report_atomically(report, target)
    return transaction_count, regional_totals, overall


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
