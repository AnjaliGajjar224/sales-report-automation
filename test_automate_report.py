from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from automate_report import COLUMNS, create_report


SOURCE = Path(__file__).parent / 'sales_data_zero_units.xlsx'


@pytest.fixture
def spreadsheet(tmp_path):
    def write(rows, headers=COLUMNS):
        source = tmp_path / 'input.xlsx'
        workbook = Workbook()
        try:
            workbook.active.append(headers)
            for row in rows:
                workbook.active.append(row)
            workbook.save(source)
        finally:
            workbook.close()
        return source
    return write


def transaction(units=2, price=12.5, region='North'):
    return (1001, datetime(2026, 4, 2), region, 'Notebook', units, price)


@pytest.mark.parametrize('units, expected_revenue', [(2, 25), (0, 0)],
                         ids=['positive-units', 'zero-units'])
def test_valid_transaction(spreadsheet, tmp_path, units, expected_revenue):
    source = spreadsheet([transaction(units=units)])
    output = tmp_path / 'report.xlsx'
    assert create_report(source, output) == (
        1, {'North': expected_revenue}, expected_revenue,
    )
    report = load_workbook(output, data_only=True)
    try:
        assert list(report['Cleaned Data'].values)[1] == (
            *transaction(units=units), expected_revenue,
        )
        assert list(report['Regional Summary'].values) == [
            ('Region', 'Total Revenue'), ('North', expected_revenue),
            ('Overall Total', expected_revenue),
        ]
    finally:
        report.close()


@pytest.mark.parametrize('units, price, message', [
    (-1, 100, 'Units must be nonnegative'),
    (2, None, 'Unit Price must be numeric'),
    ('invalid', 100, 'Units must be numeric'),
    (None, 100, 'Units must be numeric'),
], ids=['negative-units', 'missing-price', 'nonnumeric-units', 'missing-units'])
def test_invalid_transaction(spreadsheet, tmp_path, units, price, message):
    source = spreadsheet([transaction(units=units, price=price)])
    output = tmp_path / 'report.xlsx'
    with pytest.raises(ValueError, match=f'Row 2: {message}'):
        create_report(source, output)
    assert not output.exists()


@pytest.mark.parametrize('missing_column', COLUMNS)
def test_missing_required_column(spreadsheet, tmp_path, missing_column):
    index = COLUMNS.index(missing_column)
    source = spreadsheet(
        [tuple(value for i, value in enumerate(transaction()) if i != index)],
        headers=tuple(name for name in COLUMNS if name != missing_column),
    )
    output = tmp_path / 'report.xlsx'
    with pytest.raises(ValueError, match=f'Expected exactly one {missing_column!r} column'):
        create_report(source, output)
    assert not output.exists()


def test_region_normalization_and_aggregation(spreadsheet, tmp_path):
    source = spreadsheet([
        transaction(region=region)
        for region in (' north ', 'NORTH', 'NoRtH', '  new   YORK  ', 'NEW YORK')
    ])
    output = tmp_path / 'report.xlsx'
    assert create_report(source, output) == (5, {'North': 75, 'New York': 50}, 125)
    report = load_workbook(output, data_only=True)
    try:
        assert [row[2] for row in list(report['Cleaned Data'].values)[1:]] == [
            'North', 'North', 'North', 'New York', 'New York',
        ]
        assert list(report['Regional Summary'].values) == [
            ('Region', 'Total Revenue'), ('New York', 50),
            ('North', 75), ('Overall Total', 125),
        ]
    finally:
        report.close()


def test_zero_units_workbook_totals_and_formatting(tmp_path):
    output = tmp_path / 'report.xlsx'
    assert create_report(SOURCE, output) == (
        7, {'North': 2500, 'South': 1700, 'West': 2300}, 6500,
    )
    report = load_workbook(output, data_only=True)
    try:
        cleaned = report['Cleaned Data']
        assert cleaned['A8'].value == 1007
        assert cleaned['E8'].value == 0
        assert cleaned['G8'].value == 0
        assert list(report['Regional Summary'].values) == [
            ('Region', 'Total Revenue'), ('North', 2500),
            ('South', 1700), ('West', 2300), ('Overall Total', 6500),
        ]
        assert cleaned['B8'].number_format == 'yyyy-mm-dd'
        for address in ('F8', 'G8'):
            assert cleaned[address].number_format == '#,##0.00'
        assert cleaned.freeze_panes == 'A2'
        assert cleaned['A1'].font.bold
    finally:
        report.close()
