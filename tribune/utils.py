from itertools import (
    chain as _itertools_chain,
    count,
    islice,
    takewhile,
)
from functools import reduce
import string

from six.moves import map


def raises(e):
    """A functional form of `raise`."""
    raise e


def compose(*functions):
    """Function composition on a series of functions.

    Remember that function composition runs right-to-left: `f . g . h = f(g(h(x)))`. As a unix
    pipeline, it would be written: `h | g | f`.

    From https://mathieularose.com/function-composition-in-python/.
    """
    return reduce(lambda f, g: lambda x: f(g(x)), functions, lambda x: x)


def chain(*functions):
    """Returns a new function based on a series of given functions by chaining them together via
    function composition. This is exactly equal to function composition, but represents the chaining
    from left to right instead of from right to left.

    Use this to combine parsers and validators.
    """
    return compose(*reversed(functions))


def unzip(iterable):
    """Unzips/transposes an iterable of tuples into a tuple of lists.

    WARNING: When given an empty iterable, this returns an empty list instead of a tuple. If you
    need a consistent interface then do something like this:

        left, right = list(unzip(two_columned_list)) or ([], [])

    :returns: an iterable of the unzipped input.
    """
    return map(list, zip(*iterable))


def flatten(iterable):
    """Takes an iterable of iterables and flattens it by one layer (e.g. [[1],[2]] becomes [1,2]).
    """
    return list(_itertools_chain(*iterable))


def split_every(n, iterable):
    """Returns an iterable that spits an iterable into n-sized chunks. The last chunk may have less
    than n elements.

    See http://stackoverflow.com/a/22919323/503377."""
    items = iter(iterable)
    return takewhile(bool, (list(islice(items, n)) for _ in count(0)))


def column_letter(n):
    """Returns the spreadsheet column name for a given 0-based column index.

    Adapted from http://stackoverflow.com/a/4532562/503377. Spreadsheet column names are a strange
    base-26 number (for 26 letters in the alphabet) where each place value is 1-based instead of
    0-based.

    :param n: is a 0-based column index.
    """
    alphabet = string.ascii_uppercase

    n = n + 1  # Turn our 0-based number into a 1-based number.
    name = ''  # initial column name
    while n > 0:
        n = n - 1  # Turn this place value into a 0-based number.
        name = alphabet[n % len(alphabet)] + name
        n //= len(alphabet)
    return name
