from six import text_type
from six.moves import (
    filter,
    map,
)

from ..utils import (
    chain,
    raises,
)
from . import (
    SpreadsheetImportError,
    normalize_text,
)


def map_(func):
    """A strict (non-iterator), curried version of map for convenience in parser chains."""
    return lambda values: list(map(func, values))


def filter_(func):
    """A strict (non-iterator), curried version of filter for convenience in parser chains."""
    return lambda values: list(filter(func, values))


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


def nullable(parser):
    """Returns a parser that wraps another parser and only applies it if the value is not empty.
    If the value is empty, this parser always converts it to `None`."""
    return lambda value: None if default_to(None)(value) is None else parser(value)


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


def validate_unique(others, thing='thing'):
    """Validates that a value is not a member of the given set.

    :param others: is the other things which already exist.
    :param thing: is the name of the "thing" which may or may not be taken.
    """
    others_set = frozenset(others)

    def parser(value):
        if value in others_set:
            raise SpreadsheetImportError([
                u'{thing} "{value}" already exists'.format(thing=thing, value=value)])
        return value
    return parser


def validate_one_of(choices, thing='thing', show_choices_in_error=False):
    """Validates that a value is a member of the given set of choices.

    :param choices: is the set of possible choices.
    :param thing: is the name of the "things" that make up the possible choices.
    :param show_choices_in_error: determines if all the choices should be rendered in the error
                                  message.
    """
    error_message = u'must be a valid {thing}{choices}'.format(
        thing=thing,
        choices=': ' + ', '.join(map(str, choices)) if show_choices_in_error else ''
    )
    choices_set = frozenset(choices)
    return validate_satisfies(lambda v: v in choices_set, error_message)


def parse_list(parser=lambda x: x, delim=' '):
    """Returns a parser that splits its input text by the given delimiter and runs the given parser
    on each element of the result.

    :param parser: is a parser to map onto each element in the parsed list. By default it is the
                   identity function.
    :param delim: is the delimiter to use for splitting the input into a list.
    """
    return chain(parse_text, lambda value: value.split(delim), map_(parser))


# Keeps only non-empty values in a list
filter_non_empty = chain(map_(default_to(None)), filter_(lambda x: x is not None))


def parse_lookup(mapping, thing='thing'):
    """Parses a value by looking it up in a mapping (usually a dictionary). If the value is missing,
    this raises a SpreadsheetImportError. If you want to supply a default instead, then provide a
    `collections.defaultdict` object as the mapping.

    :param mapping: is a mapping from raw value to desired value. It only needs to behave like a
                    `dict`.
    :param thing: is the name of the resulting type for error messages.
    """
    def parser(value):
        try:
            return mapping[value]
        except KeyError:
            raise SpreadsheetImportError([u'"{}" is not a valid {}'.format(value, thing)])

    return parser
