import tribune.sheet_import as si
import tribune.sheet_import.parsers as p
from .utils import assert_import_errors


def test_normalize_text():
    assert p.normalize_text('') == ''
    assert p.normalize_text('a') == 'a'
    assert p.normalize_text('a b') == 'a b'
    assert p.normalize_text('   ') == ''
    assert p.normalize_text(' a b') == 'a b'
    assert p.normalize_text('a b ') == 'a b'
    assert p.normalize_text(' a b ') == 'a b'
    assert p.normalize_text('A b') == 'a b'
    assert p.normalize_text(' A B ') == 'a b'


def parse_records_with(importer):
    return lambda data: list(
        importer.get_sheet_records(
            si.XlrdSheetData.from_nested_iters(data)))


class TestSheetImport(object):
    def test_sheet_import_no_fields(self):
        parse_records = parse_records_with(si.SheetImporter([]))

        assert parse_records([]) == []
        assert parse_records([[]]) == []
        assert parse_records([['data']]) == []

    def test_sheet_import_simple(self):
        parse_records = parse_records_with(si.SheetImporter([
            si.Field('Name', 'name', p.parse_text),
        ]))

        assert parse_records([['Name']]) == []  # no data
        assert parse_records([['name']]) == []  # different header format
        assert parse_records([[' name ']]) == []  # different header format
        assert parse_records([['Name'], ['Ed']]) == [{'name': 'Ed'}]  # one row, one column
        assert parse_records([['Name'], ['Ed', '']]) == [{'name': 'Ed'}]  # one row, extra column
        assert parse_records([['Name'], ['Ed'], ['Joe']]) \
            == [{'name': 'Ed'}, {'name': 'Joe'}]  # two rows

        # invalid header
        invalid_headers = (
            [],                     # No data
            [['Some other name']],  # Wrong header name
        )
        for case in invalid_headers:
            assert_import_errors({'Expected "Name" in header cell A1.'},
                                 lambda: parse_records(case))

    def test_sheet_import_data_validation(self):
        parse_records = parse_records_with(si.SheetImporter([
            si.Field('Name', 'name', p.chain(p.parse_text, p.validate_not_empty)),
            si.Field('Age',  'age',  p.parse_int),
        ]))

        assert parse_records([['Name', 'Age']]) == []
        assert parse_records([['Name', 'Age'], ['Ed', '27']]) == [{'name': 'Ed', 'age': 27}]

        ed_n_joe = [
            {'name': 'Ed',  'age': 27},
            {'name': 'Joe', 'age': 32},
        ]

        # Test with and without extra columns or formatting
        assert parse_records([['Name', 'Age'], ['Ed', '27'], ['Joe', '32']]) == ed_n_joe
        assert parse_records([['Name', 'Age'], ['Ed', '27.0'], ['Joe', '32.00']]) == ed_n_joe
        assert parse_records([['Name', 'Age'], ['Ed', ' 27 '], ['Joe', ' 32.0 ']]) == ed_n_joe
        assert parse_records([['Name', 'Age'], ['Ed', '27', '2'], ['Joe', '32', '3']]) == ed_n_joe

        # invalid header and/or data
        assert_import_errors(
            {'Expected "Age" in header cell B1.'},
            lambda: parse_records([['Name', 'SSN']])
        )
        assert_import_errors(
            {'Expected "Age" in header cell B1.'},  # header errors beat data errors
            lambda: parse_records([['Name', 'SSN'], ['', '2']])
        )
        assert_import_errors(
            {'Unexpected value in cell A2: must not be empty',
             'Unexpected value in cell B2: must be an integer'},
            lambda: parse_records([['Name', 'Age'], ['', '3.9']])
        )
        assert_import_errors(
            {'Unexpected value in cell A3: must not be empty',
             'Unexpected value in cell B3: must be an integer'},
            lambda: parse_records([['Name', 'Age'], ['Ed', '27'], ['', '3.9']])
        )

    def test_non_string_header(self):
        parse_records = parse_records_with(si.SheetImporter([
            si.Field('Age', 'age', p.parse_int)
        ]))
        assert_import_errors(
            {'Expected "Age" in header cell A1.'},
            lambda: parse_records([[29]])
        )

    def test_sheet_import_with_modify_record(self):
        class TestImport(si.SheetImporter):
            fields = (
                si.Field('Name', 'name', p.chain(p.parse_text, p.validate_not_empty)),
            )

            def modify_record(self, record):
                if record['name'] == 'Ed':  # Filter out Eds
                    return None

                record['age'] = 27
                return record

        parse_records = parse_records_with(TestImport())

        assert parse_records([['Name']]) == []
        assert parse_records([['Name'], ['Ed']]) == []  # Ed is filtered out

        joe = [
            {'name': 'Joe', 'age': 27},  # age is always reset to 27
        ]

        assert parse_records([['Name'], ['Ed', '27'], ['Joe', '32']]) == joe
        assert parse_records([['Name'], ['Joe', '32.00'], ['Ed', '27.0']]) == joe

        # invalid header and/or data
        assert_import_errors(
            {'Expected "Name" in header cell A1.'},
            lambda: parse_records([['Age', 'Name']])
        )
        assert_import_errors(
            {'Unexpected value in cell A2: must not be empty'},
            lambda: parse_records([['Name'], [''], ['data']])
        )

    def test_sheet_import_data_start_row(self):
        parse_records = parse_records_with(si.SheetImporter([
            si.Field('Is Cool', 'is_cool', p.parse_yes_no)
        ], data_start_row=3))

        parse_records([
            [''],
            [''],
            ['Is Cool'],
            ['yes'],
            ['no'],
        ]) == [{'is_cool': True}, {'is_cool': False}]

        assert_import_errors(
            {'Expected "Is Cool" in header cell A3.'},
            lambda: parse_records([['Is Cool']])
        )

    def test_sheet_import_as_report_sheet(self):
        importer = si.SheetImporter((
            si.Field('Name', 'name', p.chain(p.parse_text, p.validate_not_empty)),
            si.Field('Job',  'job',  p.parse_text),
        ))

        TestReportSheet = importer.as_report_sheet()
        assert TestReportSheet.pre_data_rows == importer.data_start_row
        # TODO More tests here.
