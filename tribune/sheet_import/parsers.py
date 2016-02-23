from six import text_type

from ..utils import (
    chain,
    raises,
)
from . import (
    SpreadsheetImportError,
    normalize_text,
)


def parse_text(value):
    """Parses text as a unicode string and strips leading/trailing whitespace."""
    return text_type(value).strip()


def parse_yes_no(value):
    """Parses boolean values represented as 'yes' or 'no' (case insensitive)."""
    result = {
        normalize_text('yes'): True,
        normalize_text('no'):  False
    }.get(normalize_text(value))
    if result is None:
        raise SpreadsheetImportError([u'must be "yes" or "no"'])
    return result


def parse_int(value):
    """Parses integers. It accepts floating-point representations as long as they can be
    safely transformed into integers without loss of precision."""
    try:
        as_float = float(value)
        as_int = int(as_float)
        # compare float and integer versions because, e.g. 3.0 == 3 but 3.1 != 3.
        return as_int if as_int == as_float else raises(ValueError())
    except (ValueError, TypeError):
        raise SpreadsheetImportError([u'must be an integer'])


def parse_int_as_text(value):
    """Parses integers but converts them to unicode strings in the result."""
    return text_type(parse_int(value))


def parse_number(value):
    """Parses any floating point number."""
    try:
        return float(value)
    except (ValueError, TypeError):
        raise SpreadsheetImportError([u'must be a number'])


def parse_int_or_text(value):
    """Tries to parse a value as an integer and turn it into text but if that parse fails, it
    reverts to simple text parsing."""
    try:
        return parse_int_as_text(value)
    except SpreadsheetImportError:
        return parse_text(value)


def default_to(default):
    """Returns a parser that gives the provided default if the value is empty."""
    return lambda value: default if value is None or text_type(value).strip() == '' else value


def validate_satisfies(pred, error_message):
    """Validates that a value satisfies the given predicate or issues the given error if it
    doesn't."""
    return lambda value: (value if pred(value)
                          else raises(SpreadsheetImportError([error_message])))


def validate_max_length(maximum):
    """Validates that the value has at most ``maximum`` digits."""
    return validate_satisfies(lambda v: len(v) <= maximum,
                              u'must be no more than {} characters long'.format(maximum))


validate_not_empty = chain(
    default_to(None),
    validate_satisfies(lambda v: v is not None, u'must not be empty')
)


def validate_min(min_):
    """Validates that a number is equal to or grater than the given minimum."""
    return validate_satisfies(lambda v: v >= min_, u'number must be no less than {}'.format(min_))


def validate_max(max_):
    """Validates that a number is equal to or less than the given maximum."""
    return validate_satisfies(lambda v: v <= max_,
                              u'number must be no greater than {}'.format(max_))


def validate_range(min_, max_):
    """Validates that a number is between a minimum and a maximum."""
    return chain(validate_max(max_), validate_min(min_))
