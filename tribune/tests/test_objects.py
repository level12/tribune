import operator
from io import BytesIO

from blazeutils.datastructures import BlankObject
from blazeutils.spreadsheets import xlsx_to_reader
import mock
import pytest
from xlsxwriter import Workbook
import wrapt

from tribune import (
    LabeledColumn,
    PortraitRow,
    ProgrammingError,
    ReportSheet,
    SheetColumn,
    SheetSection,
    SheetUnit,
)
from .entities import Person
from .reports import CarSheet, CarDealerSheet
from .utils import find_sheet_col, find_sheet_row


def car_data(self):
    return [
        {'year': 1998, 'make': 'Ford', 'model': 'Taurus', 'style': 'SE Wagon',
         'color': 'silver', 'book_value': 1500},
        {'year': 2004, 'make': 'Oldsmobile', 'model': 'Alero', 'style': '4D Sedan',
         'color': 'silver', 'book_value': 4500},
        {'year': 2003, 'make': 'Ford', 'model': 'F-150', 'style': 'XL Supercab',
         'color': 'blue', 'book_value': 8500},
    ]


def dealer_data(self):
    return [
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

    @mock.patch('tribune.tests.reports.CarSheet.fetch_records', car_data)
    def test_attribute_values_passed_to_section_instance(self):
        sheet = CarSheet(Workbook(BytesIO()))
        car_model_section = sheet.units[1]
        assert car_model_section.foo == 'bar'

    @mock.patch('tribune.tests.reports.CarSheet.fetch_records', car_data)
    def test_attribute_values_passed_to_column_instance(self):
        sheet = CarSheet(Workbook(BytesIO()))
        car_model_section = sheet.units[1]
        car_year_column = car_model_section.units[0]
        assert car_year_column.xls_width == 15

    @mock.patch('tribune.tests.reports.CarSheet.fetch_records')
    def test_filter_args(self, m_fetch):
        CarSheet(Workbook(BytesIO()), filter_arg_a='foo', filter_arg_b='bar')
        m_fetch.assert_called_once_with(arg_a='foo', arg_b='bar')

    def test_deep_copy_sqlalchemy(self):
        """SQLALchemy objects like InstrumentedAttribute do not like deep copies. In the cases used
        by tribune, anything inheriting QueryableAttribute is treated differently, simply assigning
        the value over on the copy (instead of copying and then copying all underlying objects)."""
        base_row = PortraitRow('foo', Person.firstname)
        base_row.new_instance(None)

    def test_deep_copy_sqlalchemy_tuple(self):
        """
        tribune supports tuple expressions as values (formulated as as a 3-tuple,
        `(value1, value2, operator)`). As the tuple-values could be QueryableAttributes,
        we don't necessarily want to deep-copy

        :return:
        """
        base_row = PortraitRow('foo', (Person.intcol, -1, operator.mul))
        base_row.new_instance(None)

        base_row = PortraitRow('foo', (-1, Person.intcol, operator.mul))
        base_row.new_instance(None)

        base_row = PortraitRow('foo', ((Person.intcol, -1, operator.mul), -1, operator.mul))
        base_row.new_instance(None)

    def test_deep_copy_unbound_method(self):
        """This test is kind of hard to set up, so it needs some explanation. A great deal of
        function objects are fine with the deep copy, but this particular case is not. What we
        have is an unbound method (status_render) wrapped in a decorator."""
        def status_render(row):
            pass

        @wrapt.decorator
        def some_decorator(wrapped, instance, args, kwargs):
            row, = args
            return wrapped(row)

        class DummyUnit(SheetUnit):
            def __init__(self):
                self.foo = some_decorator(status_render)

        base_unit = DummyUnit()
        base_unit.new_instance(None)

    def test_deep_copy_static_method(self):
        """This test is kind of hard to set up, so it needs some explanation. A great deal of
        function objects are fine with the deep copy, but this particular case is not. What we
        have is an static method (status_render) wrapped in a decorator."""
        @wrapt.decorator
        def some_decorator(wrapped, instance, args, kwargs):
            row, = args
            return wrapped(row)

        class DummyUnit(SheetUnit):
            @staticmethod
            def status_render(row):
                pass

            def __init__(self):
                self.foo = some_decorator(self.status_render)

        base_unit = DummyUnit()
        base_unit.new_instance(None)


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
            def fetch_records(self):
                return []

            pre_data_rows = 5
            SheetColumn('key', header_2='foo', header_4='bar')

        sheet = TestSheet(Workbook(BytesIO()))
        col = sheet.units[0]
        assert col.header_data == ['', '', 'foo', '', 'bar']

    def test_header_construction_overflows_sheet(self):
        class TestSheet(ReportSheet):
            def fetch_records(self):
                return []

            pre_data_rows = 3
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
    def generate_report(self, reportcls=CarSheet):
        wb = Workbook(BytesIO())
        ws = wb.add_worksheet(reportcls.sheet_name)
        reportcls(wb, worksheet=ws)
        book = xlsx_to_reader(wb)
        return book.sheet_by_name('Car Info')

    def test_header(self):
        ws = self.generate_report()
        assert ws.cell_value(0, 0) == 'Cars'

    def test_column_headings(self):
        def checkit(ws, row, column, value):
            assert ws.cell_value(row, column) == value

        ws = self.generate_report()
        for row, column, value in [
            (2, 0, ''),
            (2, 1, 'Year'),
            (2, 2, 'Make'),
            (2, 3, 'Model'),
            (2, 4, 'Style'),
            (2, 5, 'Color'),
            (2, 6, 'Blue'),
            (3, 6, 'Book'),
        ]:
            yield checkit, ws, row, column, value

    def test_column_data(self):
        def checkit(ws, row, column, value):
            assert ws.cell_value(row, column) == value

        ws = self.generate_report()
        for data_row, data in enumerate(car_data(None)):
            row = data_row + 4
            for column, value in [
                (0, ''),
                (find_sheet_col(ws, 'Year', 2), data['year']),
                (find_sheet_col(ws, 'Make', 2), data['make']),
                (find_sheet_col(ws, 'Model', 2), data['model']),
                (find_sheet_col(ws, 'Style', 2), data['style']),
                (find_sheet_col(ws, 'Color', 2), data['color']),
                (find_sheet_col(ws, 'Blue', 2), data['book_value']),
            ]:
                yield checkit, ws, row, column, value

    def test_merged_section_heading(self):
        class TestSheet(ReportSheet):
            class CarMergedHeaderSection(SheetSection):
                LabeledColumn('Style', 'style')
                LabeledColumn('Color', 'color')

                def write_header(self):
                    record_dict = {'foo': 'bar', 'baz': 'data'}
                    self.sheet.write_simple_merge(2, (record_dict, 'baz'))

            pre_data_rows = 1
            sheet_name = 'Car Info'
            CarMergedHeaderSection()

            def fetch_records(self):
                return []

        ws = self.generate_report(reportcls=TestSheet)
        assert ws.cell_value(0, 0) == 'data'

    def test_merged_section_heading_single_length(self):
        class TestSheet(ReportSheet):
            class CarMergedHeaderSection(SheetSection):
                LabeledColumn('Style', 'style')

                def write_header(self):
                    record_dict = {'foo': 'bar', 'baz': 'data'}
                    self.sheet.write_simple_merge(1, (record_dict, 'baz'))

            pre_data_rows = 1
            sheet_name = 'Car Info'
            CarMergedHeaderSection()

            def fetch_records(self):
                return []

        ws = self.generate_report(reportcls=TestSheet)
        assert ws.cell_value(0, 0) == 'data'


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
        def checkit(ws, row, column, value):
            assert ws.cell_value(row, column) == value

        ws = self.generate_report()
        for row, column, value in [
            (3, 0, ''),
            (3, 1, 'Count'),
            (3, 2, 'Income'),
            (3, 3, 'Expense'),
            (3, 4, 'Net'),
        ]:
            yield checkit, ws, row, column, value

    def test_row_data(self):
        def checkit(ws, row, column, value):
            assert ws.cell_value(row, column) == value

        ws = self.generate_report()
        row = find_sheet_row(ws, 'Sales', 0)
        for column, value in [
            (1, 5),
            (2, 75000),
            (3, 25000),
            (4, 50000),
        ]:
            yield checkit, ws, row, column, value

        row = find_sheet_row(ws, 'Rentals', 0)
        for column, value in [
            (1, 46),
            (2, 8600),
            (3, 2900),
            (4, 5700),
        ]:
            yield checkit, ws, row, column, value

        row = find_sheet_row(ws, 'Leases', 0)
        for column, value in [
            (1, 27),
            (2, 10548),
            (3, 3680),
            (4, 6868),
        ]:
            yield checkit, ws, row, column, value

        row = find_sheet_row(ws, 'Totals', 0)
        for column, value in [
            (4, 100000),
        ]:
            yield checkit, ws, row, column, value
