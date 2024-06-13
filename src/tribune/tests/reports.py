import operator

from blazeutils.datastructures import OrderedDict

from tribune import (
    BlankColumn,
    LabeledColumn,
    PortraitRow,
    ReportPortraitSheet,
    ReportSheet,
    SheetPortraitColumn,
    SheetSection,
    TotaledMixin,
)

"""
    Objects for testing normal landscape-layout reports
"""


class CarLabeledColumn(LabeledColumn):
    header_start_row = 2


class CarYearColumn(CarLabeledColumn):
    xls_width = 10


class CarTotaledColumn(TotaledMixin, CarLabeledColumn):
    pass


class CarModelSection(SheetSection):
    def __init__(self, *args, **kwargs):
        self.foo = kwargs.get('foo')

    CarYearColumn('Year', 'year', xls_width=15)
    CarLabeledColumn('Make', 'make')
    CarLabeledColumn('Model', 'model')


class CarStyleSection(SheetSection):
    CarLabeledColumn('Style', 'style')
    CarLabeledColumn('Color', 'color')


class CarSheet(ReportSheet):
    pre_data_rows = 4
    sheet_name = 'Car Info'

    BlankColumn()
    CarModelSection(foo='bar')
    CarStyleSection()
    CarTotaledColumn('Blue\nBook', 'book_value')

    def fetch_records(self, **kwargs):
        # gets mocked in tests to return correct sequences
        pass

    def write_sheet_header(self):
        self.write_header(0, 0, 'Cars', self.style_header)


"""
    Objects for testing portrait-layout reports
"""


class DealershipReportRow(PortraitRow):
    def __init__(self, label=None, comp=None, inc=None, exp=None, net=None,
                 comp_curr=False, **kwargs):
        super(DealershipReportRow, self).__init__()
        self.label = label
        self.comp = comp
        self.comp_curr = comp_curr
        self.inc = inc
        self.exp = exp
        self.net = net

        if self.net is None and (self.inc or self.exp):
            self.net = (self.inc, self.exp, operator.sub)

    def format_label(self):
        return None

    def format_component(self):
        if self.comp_curr:
            return self.sheet.style_currency
        return None

    def format_income(self):
        return self.sheet.style_currency

    def format_expense(self):
        return self.sheet.style_currency

    def format_net(self):
        return self.sheet.style_currency

    def render(self):
        self.sheet.columns['label'].write_data(self.label, self.format_label())
        self.sheet.columns['component'].write_data(self.comp, self.format_component())
        self.sheet.columns['income'].write_data(self.inc, self.format_income())
        self.sheet.columns['expense'].write_data(self.exp, self.format_expense())
        self.sheet.columns['net'].write_data(self.net, self.format_net())
        return True


class DealershipTotalRow(DealershipReportRow):
    def format_label(self):
        return self.sheet.combine_style_dicts(
            self.sheet.style_header,
            self.sheet.style_right,
            {'top': 6}
        )

    def format_component(self):
        return self.sheet.combine_style_dicts(
            self.sheet.style_currency,
            {'top': 6}
        )

    def format_income(self):
        return self.format_component()

    def format_expense(self):
        return self.format_component()

    def format_net(self):
        return self.format_component()


class CarDealerSheet(ReportPortraitSheet):
    pre_data_rows = 4
    sheet_name = 'Dealership Summary'

    DealershipReportRow('Sales', 'sale_count', 'sale_income', 'sale_expense')
    DealershipReportRow('Rentals', 'rental_count', 'rental_income', 'rental_expense')
    DealershipReportRow('Leases', 'lease_count', 'lease_income', 'lease_expense')
    DealershipTotalRow('Totals', None, None, None, 'total_net')

    def fetch_records(self, **kwargs):
        # gets mocked in tests to return correct sequences
        pass

    def init_columns(self):
        d = OrderedDict()
        d['label'] = SheetPortraitColumn(sheet=self, record=self.records[0])
        d['component'] = SheetPortraitColumn(sheet=self, record=self.records[0])
        d['income'] = SheetPortraitColumn(sheet=self, record=self.records[0])
        d['expense'] = SheetPortraitColumn(sheet=self, record=self.records[0])
        d['net'] = SheetPortraitColumn(sheet=self, record=self.records[0])
        return d

    def write_sheet_header(self):
        self.write_header(0, 0, 'Dealership', self.style_header)
        self.write_header(3, 1, 'Count', self.style_header)
        self.write_header(3, 2, 'Income', self.style_header)
        self.write_header(3, 3, 'Expense', self.style_header)
        self.write_header(3, 4, 'Net', self.style_header)
