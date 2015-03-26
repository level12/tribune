import operator
from io import BytesIO

from blazeutils.datastructures import BlankObject
from blazeutils.spreadsheets import xlsx_to_reader
import mock
import pytest
from xlsxwriter import Workbook

from tribune import SheetColumn, LabeledColumn, ReportSheet, ProgrammingError
from .reports import CarSheet, CarModelSection, CarDealerSheet
from .utils import find_sheet_col, find_sheet_row

from tribune_ta.model.entities import Person

car_data = lambda s: [
    {'year': 1998, 'make': 'Ford', 'model': 'Taurus', 'style': 'SE Wagon',
     'color': 'silver', 'book_value': 1500},
    {'year': 2004, 'make': 'Oldsmobile', 'model': 'Alero', 'style': '4D Sedan',
     'color': 'silver', 'book_value': 4500},
    {'year': 2003, 'make': 'Ford', 'model': 'F-150', 'style': 'XL Supercab',
     'color': 'blue', 'book_value': 8500},
]

dealer_data = lambda s: [
    {'sale_count': 5, 'sale_income': 75000, 'sale_expense': 25000,
     'rental_count': 46, 'rental_income': 8600, 'rental_expense': 2900,
     'lease_count': 27, 'lease_income': 10548, 'lease_expense': 3680,
     'total_net': 100000},
]


class TestSheetDecl(object):
    @mock.patch('tribune.tests.reports.CarSheet.fetch_records', car_data)
    def test_sheet_units(self):
        sheet = CarSheet(Workbook(BytesIO()))
        assert len(sheet.units) == 4

    @mock.patch('tribune.tests.reports.CarSheet.fetch_records', car_data)
    def test_section_units(self):
        sheet = CarSheet(Workbook(BytesIO()))
        section = sheet.units[1]
        assert len(section.units) == 3

    @mock.patch('tribune.tests.reports.CarSheet.fetch_records')
    def test_filter_args(self, m_fetch):
        sheet = CarSheet(Workbook(BytesIO()), filter_arg_a='foo', filter_arg_b='bar')
        m_fetch.assert_called_once_with(arg_a='foo', arg_b='bar')


class TestSheetColumn(object):
    def test_extract_data_no_key(self):
        col = SheetColumn()
        assert col.extract_data({}) == ''

    def test_extract_data_attr(self):
        col = SheetColumn('bar')
        assert col.extract_data(
            BlankObject(
                foo=1,
                bar='something',
                baz='else',
            )
        ) == 'something'

    def test_extract_data_dict(self):
        col = SheetColumn('bar')
        assert col.extract_data({'foo': 1, 'bar': 'something', 'baz': 'else'}) == 'something'

    def test_extract_data_string(self):
        col = SheetColumn('bar')
        assert col.extract_data({'foo': 1, 'baz': 'else'}) == 'bar'

    def test_extract_data_sacol(self):
        person = Person.testing_create()
        col = SheetColumn(Person.firstname)
        assert col.extract_data(person) == person.firstname

    def test_extract_data_sacol_altcolname(self):
        person = Person.testing_create(lastname=u'foo')
        col = SheetColumn(Person.lastname)
        assert col.extract_data(person) == 'foo'

    def test_extract_data_combined_sacol(self):
        person = Person.testing_create(intcol=8, floatcol=9)
        col = SheetColumn((Person.intcol, Person.floatcol, operator.add))
        assert col.extract_data(person) == 17

    def test_extract_data_combined_string(self):
        col = SheetColumn(('foo', 'bar', operator.add))
        assert col.extract_data({'foo': 9, 'bar': 8}) == 17

    def test_extract_data_combined_int(self):
        col = SheetColumn((9, 8, operator.add))
        assert col.extract_data({}) == 17

    def test_extract_data_combined_float(self):
        col = SheetColumn((9.5, 8, operator.add))
        assert col.extract_data({}) == 17.5

    def test_header_construction(self):
        class TestSheet(ReportSheet):
            pre_data_rows = 5
            fetch_records = lambda x: []
            SheetColumn('key', header_2='foo', header_4='bar')

        sheet = TestSheet(Workbook(BytesIO()))
        col = sheet.units[0]
        assert col.header_data == ['', '', 'foo', '', 'bar']

    def test_header_construction_overflows_sheet(self):
        class TestSheet(ReportSheet):
            pre_data_rows = 3
            fetch_records = lambda x: []
            SheetColumn('key', header_2='foo', header_4='bar')

        with pytest.raises(ProgrammingError) as exc_info:
            TestSheet(Workbook(BytesIO()))
        assert 'not enough pre-data rows on sheet' in str(exc_info)

    def test_width(self):
        col = SheetColumn('key')
        col.register_col_width('something')
        col.register_col_width('something else')
        assert col.xls_computed_width == 14


class TestLabeledColumn(object):
    def test_header_construction(self):
        col = LabeledColumn('foo\nbar', 'key')
        assert col._init_header == {0: 'foo', 1: 'bar'}

    def test_header_construction_offset(self):
        class _LabeledColumn(LabeledColumn):
            header_start_row = 4

        col = _LabeledColumn('foo\nbar', 'key')
        assert col._init_header == {4: 'foo', 5: 'bar'}


class TestOutput(object):
    @mock.patch('tribune.tests.reports.CarSheet.fetch_records', car_data)
    def generate_report(self):
        wb = Workbook(BytesIO())
        ws = wb.add_worksheet(CarSheet.sheet_name)
        CarSheet(wb, worksheet=ws)
        book = xlsx_to_reader(wb)
        return book.sheet_by_name('Car Info')

    def test_header(self):
        ws = self.generate_report()
        assert ws.cell_value(0, 0) == 'Cars'

    def test_column_headings(self):
        ws = self.generate_report()
        assert ws.cell_value(2, 0) == ''
        assert ws.cell_value(2, 1) == 'Year'
        assert ws.cell_value(2, 2) == 'Make'
        assert ws.cell_value(2, 3) == 'Model'
        assert ws.cell_value(2, 4) == 'Style'
        assert ws.cell_value(2, 5) == 'Color'
        assert ws.cell_value(2, 6) == 'Blue'
        assert ws.cell_value(3, 6) == 'Book'

    def test_column_data(self):
        ws = self.generate_report()
        for data_row, data in enumerate(car_data(None)):
            row = data_row + 4
            assert ws.cell_value(row, 0) == ''
            assert ws.cell_value(row, find_sheet_col(ws, 'Year', 2)) == data['year']
            assert ws.cell_value(row, find_sheet_col(ws, 'Make', 2)) == data['make']
            assert ws.cell_value(row, find_sheet_col(ws, 'Model', 2)) == data['model']
            assert ws.cell_value(row, find_sheet_col(ws, 'Style', 2)) == data['style']
            assert ws.cell_value(row, find_sheet_col(ws, 'Color', 2)) == data['color']
            assert ws.cell_value(row, find_sheet_col(ws, 'Blue', 2)) == data['book_value']


class TestPortraitOutput(object):
    @mock.patch('tribune.tests.reports.CarDealerSheet.fetch_records', dealer_data)
    def generate_report(self):
        wb = Workbook(BytesIO())
        ws = wb.add_worksheet(CarDealerSheet.sheet_name)
        CarDealerSheet(wb, worksheet=ws)
        book = xlsx_to_reader(wb)
        return book.sheet_by_name('Dealership Summary')

    def test_header(self):
        ws = self.generate_report()
        assert ws.cell_value(0, 0) == 'Dealership'

    def test_column_headings(self):
        ws = self.generate_report()
        assert ws.cell_value(3, 0) == ''
        assert ws.cell_value(3, 1) == 'Count'
        assert ws.cell_value(3, 2) == 'Income'
        assert ws.cell_value(3, 3) == 'Expense'
        assert ws.cell_value(3, 4) == 'Net'

    def test_row_data(self):
        ws = self.generate_report()
        row = find_sheet_row(ws, 'Sales', 0)
        assert ws.cell_value(row, 1) == 5
        assert ws.cell_value(row, 2) == 75000
        assert ws.cell_value(row, 3) == 25000
        assert ws.cell_value(row, 4) == 50000

        row = find_sheet_row(ws, 'Rentals', 0)
        assert ws.cell_value(row, 1) == 46
        assert ws.cell_value(row, 2) == 8600
        assert ws.cell_value(row, 3) == 2900
        assert ws.cell_value(row, 4) == 5700

        row = find_sheet_row(ws, 'Leases', 0)
        assert ws.cell_value(row, 1) == 27
        assert ws.cell_value(row, 2) == 10548
        assert ws.cell_value(row, 3) == 3680
        assert ws.cell_value(row, 4) == 6868

        row = find_sheet_row(ws, 'Totals', 0)
        assert ws.cell_value(row, 4) == 100000
