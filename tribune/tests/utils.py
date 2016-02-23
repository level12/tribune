import datetime as dt

import six
import sqlalchemy as sa
import sqlalchemy.orm as saorm
import sqlalchemy.sql as sasql
import wrapt

from tribune.sheet_import import SpreadsheetImportError


# utility functions to find a row or column for testing. Reduces need for
#   hard-coded row/col indices in tests.
#   call generally with the worksheet, desired cell value, the column/row to search
#       within, and starting/ending indices
#   returns -1 if the value is not found

def find_sheet_row(sheet, needle, column, start=0, end=100):
    try:
        for x in range(start, end+1):
            if sheet.cell_value(x, column) == needle:
                return x
    except IndexError:
        pass
    return -1


def find_sheet_col(sheet, needle, row, start=0, end=100):
    try:
        for x in range(start, end+1):
            if sheet.cell_value(row, x) == needle:
                return x
    except IndexError:
        pass
    return -1


def assert_import_errors(errors, func):
    try:
        func()
        assert False, 'Did not raise SpreadsheetImportError'
    except SpreadsheetImportError as e:
        assert set(e.errors) == set(errors)


class DefaultColsMixin(object):
    id = sa.Column(sa.Integer, primary_key=True)
    createdts = sa.Column(sa.DateTime, nullable=False, default=dt.datetime.now,
                          server_default=sasql.text('CURRENT_TIMESTAMP'))
    updatedts = sa.Column(sa.DateTime, nullable=False, default=dt.datetime.now,
                          server_default=sasql.text('CURRENT_TIMESTAMP'), onupdate=dt.datetime.now)


@wrapt.decorator
def transaction(f, decorated_obj, args, kwargs):
    """
        decorates a function so that a DB transaction is always committed after
        the wrapped function returns and also rolls back the transaction if
        an unhandled exception occurs.

        'ncm' = non class method (version)
    """
    from tribune.tests.entities import session as dbsess

    try:
        retval = f(*args, **kwargs)
        dbsess.commit()
        return retval
    except Exception:
        dbsess.rollback()
        raise


def transaction_classmethod(f):
    """
        like transaction() but makes the function a class method
    """
    return transaction(classmethod(f))


class MethodsMixin(object):
    @classmethod
    def _sa_sess(cls):
        from .entities import session
        return session

    @transaction_classmethod
    def add(cls, **kwargs):
        o = cls()
        o.from_dict(kwargs)
        cls._sa_sess().add(o)
        return o

    def from_dict(self, data):
        """
        Update a mapped class with data from a JSON-style nested dict/list
        structure.
        """
        # surrogate can be guessed from autoincrement/sequence but I guess
        # that's not 100% reliable, so we'll need an override

        mapper = saorm.object_mapper(self)

        for key, value in six.iteritems(data):
            if isinstance(value, dict):
                dbvalue = getattr(self, key)
                rel_class = mapper.get_property(key).mapper.class_
                pk_props = rel_class._descriptor.primary_key_properties

                # If the data doesn't contain any pk, and the relationship
                # already has a value, update that record.
                if not [1 for p in pk_props if p.key in data] and \
                   dbvalue is not None:
                    dbvalue.from_dict(value)
                else:
                    record = rel_class.update_or_create(value)
                    setattr(self, key, record)
            elif isinstance(value, list) and \
                    value and isinstance(value[0], dict):

                rel_class = mapper.get_property(key).mapper.class_
                new_attr_value = []
                for row in value:
                    if not isinstance(row, dict):
                        raise Exception(
                            'Cannot send mixed (dict/non dict) data '
                            'to list relationships in from_dict data.'
                        )
                    record = rel_class.update_or_create(row)
                    new_attr_value.append(record)
                setattr(self, key, new_attr_value)
            else:
                setattr(self, key, value)


class DefaultMixin(DefaultColsMixin, MethodsMixin):
    pass
