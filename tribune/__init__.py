from decimal import Decimal
import inspect
import sys

from blazeutils.spreadsheets import WriterX
from xlsxwriter.format import Format

__all__ = [
    'SheetUnit',
    'SheetColumn',
    'SheetPortraitColumn',
    'SheetSection',
    'BlankColumn',
    'LabeledColumn',
    'TotaledMixin',
    'ReportSheet',
    'ReportPortraitSheet',
]


def column_letter(c):
    # c is the zero-based column index
    alphabet = 'abcdefghijklmnopqrstuvwxyz'.upper()
    if c < 26:
        return alphabet[c]
    return alphabet[(c / 26)-1]+alphabet[c % 26]


class ProgrammingError(Exception):
    pass


class _SheetDeclarativeMeta(type):
    """
        Metaclass to define/create units on a sheet section
        Borrowed in part from WebGrid's Grid setup
    """

    def __new__(cls, name, bases, class_dict):
        class_units = []

        # add columns from base classes
        for base in bases:
            base_units = getattr(base, '__cls_units__', ())
            class_units.extend(base_units)
        class_units.extend(class_dict.get('__cls_units__', ()))
        class_dict['__cls_units__'] = class_units

        return super(_SheetDeclarativeMeta, cls).__new__(cls, name, bases, class_dict)


class SheetUnit(object):
    # Basic sheet object, can be a row or a column

    def __new__(cls, *args, **kwargs):
        col_inst = super(SheetUnit, cls).__new__(cls)
        if '_dont_assign' not in kwargs:
            col_inst._assign_to_section()
        return col_inst

    def new_instance(self, sheet):
        cls = self.__class__
        addtl_args = {}

        # try to be smart about which attributes should get copied to the
        # new instance by looking for attributes on the class that have the
        # same name as arguments to the classes __init__ method
        for argname in inspect.getargspec(self.__init__).args:
            if argname != 'self' and hasattr(self, argname) and argname not in ('key', 'sheet',):
                addtl_args[argname] = getattr(self, argname)

        # call the unit's init, and pass the addtl args back in
        unit = cls(key=getattr(self, 'key', None), sheet=sheet, _dont_assign=True, **addtl_args)

        return unit

    def _assign_to_section(self):
        section_locals = sys._getframe(2).f_locals
        section_cls_units = section_locals.setdefault('__cls_units__', [])
        section_cls_units.append(self)


class SheetColumn(SheetUnit):
    """
        Basic column for a spreadsheet report. Holds the SQLAlchemy
        column and has methods for rendering various kinds of rows
    """
    def new_instance(self, sheet):
        column = SheetUnit.new_instance(self, sheet)
        column.xls_width = self.xls_width
        column.expr = self.expr
        column._construct_header_data(self._init_header)

        return column

    def __init__(self, key=None, sheet=None, **kwargs):
        """
            header_# in kwargs used to override default (blank) heading values
        """
        # key can be one of many types: str, Decimal, int, float, SA column,
        #   or a tuple of the above (with an operator)
        self.expr = None
        if key is not None and not isinstance(key, (basestring, tuple, Decimal, int, float)):
            self.expr = col = key
            # use column.key, column.name, or None in that order
            key = getattr(col, 'key', getattr(col, 'name', None))

        self.key = key
        self.xls_width = kwargs.get('xls_width')
        self.xls_computed_width = 0
        self.sheet = sheet

        # look for header values in kwargs, construct a header dict
        self._init_header = dict()
        for k, v in kwargs.iteritems():
            if k.startswith('header_'):
                self._init_header[int(k[7:])] = v

    def _construct_header_data(self, init_header):
        d = ['' for i in range(self.sheet.pre_data_rows)]
        for k, v in init_header.iteritems():
            try:
                d[k] = v
            except IndexError:
                raise ProgrammingError('not enough pre-data rows on sheet')
        self.header_data = d

    def xls_width_calc(self, value):
        if self.xls_width:
            return self.xls_width
        if isinstance(value, basestring) and len(value) and value[0] == '=':
            # formulas start with an equals
            return 0
        if isinstance(value, basestring):
            return len(value)
        return len(str(value))

    def register_col_width(self, value):
        # keep track of the width of column contents
        if value is None:
            return
        self.xls_computed_width = max(
            self.xls_computed_width,
            self.xls_width_calc(value)
        )

    def adjust_col_width(self):
        # set column width according to computed width by register_col_width
        if hasattr(self, 'col_idx'):
            final_width = max(min(self.xls_computed_width, 150), 5)

            # width calculation is 1/256th of width of zero character, using the
            # first font that occurs in the Excel file
            # the following calculation seems to get it done alright
            self.sheet.ws.set_column(self.col_idx, self.col_idx, int(final_width + 3))

    def fetch_val(self, record, key):
        if key is None:
            return ''
        elif isinstance(key, basestring) and hasattr(record, key):
            return getattr(record, key)
        elif isinstance(key, basestring) and (isinstance(record, dict) and key in record):
            return record[key]
        elif isinstance(key, tuple):
            # magic way to combine data fields in a cell. tuple is expected as
            #   (SA-col-1/string/int, SA-col-2/string/int, operator)
            val = key[2](
                self.fetch_val(record, key[0]),
                self.fetch_val(record, key[1])
            )
        elif not isinstance(key, (basestring, int, float, Decimal)):
            # must be a SA-col, find the key and hit the record
            datakey = getattr(key, 'key', getattr(key, 'name', None))
            val = getattr(record, datakey)
        else:
            # if none of the above, put the key in the cell
            val = key
        return val

    def extract_data(self, record):
        return self.fetch_val(record, self.key)

    def format_data(self):
        return None

    def format_header(self, header_row):
        # called for each pre_data_row row. header_row is zero-based.
        return None

    def format_total(self):
        return None

    def write_header(self, header_row):
        # set the column index to be used for adjustment of width
        # and write the stored header data from init
        self.col_idx = self.sheet.colnum
        val = self.header_data[header_row]
        self.register_col_width(val)
        self.sheet.awrite(val, self.format_header(header_row))

    def write_data(self, record):
        val = self.extract_data(record)
        self.register_col_width(val)
        self.sheet.awrite(val, self.format_data())

    def write_total(self):
        self.sheet.awrite('', self.format_total())


class SheetPortraitColumn(SheetColumn):
    """
        Column for ReportPortraitSheet. Not tied to a specific data field
        (as that will change from row to row), but tracks column width needed
    """
    def __init__(self, record=None, *args, **kwargs):
        SheetColumn.__init__(self, *args, **kwargs)
        self.xls_computed_width = 0
        self.record = record

    def write_data(self, data_or_field, fmt=None):
        val = self.fetch_val(self.record, data_or_field)
        if val is not None:
            self.register_col_width(val)
        self.sheet.awrite(val, fmt)


class BlankColumn(SheetColumn):
    def extract_data(self, record):
        return ''


class LabeledColumn(SheetColumn):
    header_start_row = 0

    def __init__(self, label=None, key=None, **kwargs):
        if label:
            label_rows = label.split('\n')
            for i_row, label_string in enumerate(label_rows):
                kwargs['header_{0}'.format(self.header_start_row + i_row)] = label_string
        SheetColumn.__init__(self, key=key, **kwargs)


class TotaledMixin(object):
    def write_total(self):
        sum_begin = self.sheet.pre_data_rows+1
        sum_end = self.sheet.rownum
        if sum_begin > sum_end:
            # no data rows
            self.sheet.awrite('', self.format_total())
            return
        self.sheet.awrite(
            '=SUM({0}{1}:{0}{2})'.format(
                column_letter(self.sheet.colnum),
                sum_begin,
                sum_end
            ),
            self.format_total()
        )


class SheetSection(SheetUnit):
    """
        Groups SheetColumns together. May override heading render for
        individual columns (e.g. need a merged cell to head the section)

        If a sheet references a section, it will not reference the individual
        columns directly but depend on the section to police its own
    """
    __metaclass__ = _SheetDeclarativeMeta
    __cls_units__ = ()

    def new_instance(self, sheet):
        cls = self.__class__
        section = cls(_dont_assign=True)
        section.sheet = sheet
        section.units = []
        for unit in self.__cls_units__:
            new_unit = unit.new_instance(section.sheet)
            section.units.append(new_unit)

        # try to be smart about which attributes should get copied to the
        # new instance by looking for attributes on the class that have the
        # same name as arguments to the classes __init__ method
        for argname in inspect.getargspec(self.__init__).args:
            if argname != 'self' and hasattr(self, argname):
                setattr(section, argname, getattr(self, argname))

        return section

    def __init__(self, *args, **kwargs):
        return super(SheetSection, self).__init__()

    def write_header(self, header_row):
        # track starting column and then delegate to subunits
        self.col_idx = self.sheet.colnum
        for u in self.units:
            u.write_header(header_row)

    def write_data(self, record):
        # delegate to subunits
        for u in self.units:
            u.write_data(record)

    def write_total(self):
        # delegate to subunits
        for u in self.units:
            u.write_total()

    def adjust_col_width(self):
        # delegate to subunits
        for u in self.units:
            u.adjust_col_width()


class ReportSheet(SheetSection, WriterX):
    freeze = None
    pre_data_rows = 0
    sheet_name = None

    def __init__(self, parent_book, worksheet=None, **kwargs):
        self.parent_book = parent_book
        if not worksheet:
            worksheet = parent_book.add_worksheet(self.sheet_name)
        WriterX.__init__(self, ws=worksheet)
        self.set_base_style_dicts()

        # fetch records first, as units may need some data for initialization
        filter_args = dict([
            (k[7:], v) for k, v in kwargs.iteritems() if k[:7] == 'filter_'
        ])
        self.records = self.fetch_records(**filter_args)

        # list of (row,col) tuples to guard when writing. Idea here is to have
        #   as short a list as possible, for performance (i.e. don't check
        #   xlwt's complete sheet or row list)
        self.locked_cells = []

        # generate list of units (whether basic units or sections) from declarative data
        self.units = []
        for unit in self.__cls_units__:
            new_unit = unit.new_instance(self)
            self.units.append(new_unit)

        if kwargs.get('auto_render', True):
            self.render()

    def colwidth(self, chars):
        return chars * 256

    def fontsize(self, pt):
        return pt * 20

    def add_format_to_workbook(self, fmt):
        # attach a Format object to the workbook, if it has not been already
        if fmt.xf_format_indices is None:
            fmt.xf_format_indices = self.parent_book.xf_format_indices
            fmt.dxf_format_indices = self.parent_book.dxf_format_indices
            self.parent_book.formats.append(fmt)

    def set_base_style_dicts(self):
        self.style_title = {'bold': True, 'font_size': 14}
        self.style_header = {'bold': True}
        self.style_row_header = {'bold': True, 'bottom': 1}
        self.style_row_pct = {'num_format': '0.00%'}
        self.style_row_pct_round = {'num_format': '0%'}
        self.style_row_totals = {'bold': True, 'top': 1}
        self.style_row_totals_pct = {'bold': True, 'top': 1, 'num_format': '0.00%'}
        self.style_italic = {'italic': True}
        self.style_date_mdy = {'num_format': 'M/D/YYYY'}
        self.style_numeric = {'num_format': '#,##0.00'}
        self.style_pct = {'num_format': '0.00%'}
        self.style_left = {'align': 'left'}
        self.style_center = {'align': 'center'}
        self.style_right = {'align': 'right'}
        self.style_currency = {'num_format': '$* #,##0.00;[Red]$* (#,##0.00);$* -  ;'}

    def combine_style_dicts(self, *args):
        # takes dicts such as those defined in set_base_styles, combines
        #   left-to-right, returns dict
        def add_dicts(a, b):
            return dict(a.items() + b.items() + [(k, a[k] + b[k]) for k in set(b) & set(a)])

        style = args[0]
        if len(args) == 1:
            return style
        iterargs = iter(args)
        next(iterargs)
        for s in iterargs:
            style = add_dicts(style, s)
        return style

    def fetch_records(self, **kwargs):
        raise NotImplementedError

    def freeze_pane(self, col, row):
        self.ws.panes_frozen = True
        self.ws.vert_split_pos = col
        self.ws.horz_split_pos = row

    def write_simple_merge(self, num_cols, data, style=None):
        # shorthand for WriterX.write_merge, for merge on single row
        self.merge_range(self.rownum, self.colnum, self.rownum,
                         self.colnum + num_cols - 1)
        self.write(self.rownum, self.colnum, data, style)
        self.colnum += num_cols

    def write_sheet_header(self):
        # use to write sheet header info which is not column-specific (e.g. title)
        pass

    def write_header(self, row, col, data, style=None):
        # calls write, and records row,col tuple to guard overwrites
        self.write(row, col, data, style)
        self.locked_cells.append((row, col))

    def lock_sheet_header(self):
        # should result in performance upgrade, constant time lookup on locked_cells
        self.locked_cells = set(self.locked_cells)

    def conform_style(self, style):
        if isinstance(style, dict):
            # styles can be either an xlsxwriter Format or a dict
            style = Format(style)
            self.add_format_to_workbook(style)
        return style

    def write(self, row, col, data, style=None):
        # wraps WriterX.write, checks row,col tuple against guarded cells
        if (row, col) not in self.locked_cells:
            self.ws.write(row, col, data, self.conform_style(style))

    def awrite(self, data, style=None, nextrow=False):
        super(ReportSheet, self).awrite(data, style=self.conform_style(style), nextrow=nextrow)

    def render_header(self):
        for i in range(self.pre_data_rows):
            for u in self.units:
                u.write_header(i)
            self.nextrow()

    def render_data(self):
        for r in self.records:
            for u in self.units:
                u.write_data(r)
            self.nextrow()

    def render_total(self):
        for u in self.units:
            u.write_total()
            u.adjust_col_width()

    def render(self):
        self.write_sheet_header()
        self.lock_sheet_header()
        self.render_header()
        self.render_data()
        self.render_total()
        if self.freeze:
            self.freeze_pane(*self.freeze)


class ReportPortraitSheet(ReportSheet):
    # generic reporting sheet in which data is laid out in rows, though perhaps
    #   with multiple columns providing depth

    def __init__(self, *args, **kwargs):
        kwargs['auto_render'] = False
        ReportSheet.__init__(self, *args, **kwargs)

        self.columns = self.init_columns()

        # assign the columns their indices. columns should be OrderedDict, so
        #   columns will come in the order designed
        for idx, c in enumerate(self.columns.values()):
            c.col_idx = idx

        self.render()

    def init_columns(self):
        # return an ordered dict of columns on the report
        raise NotImplementedError

    def render(self):
        self.write_sheet_header()
        self.lock_sheet_header()

        # reset location for data loop
        self.rownum = self.pre_data_rows
        self.colnum = 0

        # each unit is responsible for rendering all columns for its row
        #   if the unit's render returns True, this indicates it has rendered
        #   and the sheet can move to the next row
        for u in self.units:
            if u.render():
                self.nextrow()

        # columns have tracked the necessary widths, set here
        for c in self.columns.values():
            c.adjust_col_width()
