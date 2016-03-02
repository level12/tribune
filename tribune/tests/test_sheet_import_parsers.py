from collections import defaultdict

import pytest
import tribune.sheet_import.parsers as p
from .utils import assert_import_errors


def assert_really_equal(left, right):
    assert type(left) == type(right)
    assert left == right


def test_chain():
    assert p.chain(lambda _: 'Python')(0.99) == 'Python'
    assert p.chain(lambda x: x + 1, lambda x: x - 1)(0) == 0
    assert p.chain(lambda x: x * 2, lambda x: x - 1)(1) == 1
    assert p.chain(lambda x: x - 1, lambda x: x * 2)(1) == 0


def test_parse_text():
    assert p.parse_text('') == u''
    assert p.parse_text('  ') == u''
    assert p.parse_text(' A ') == u'A'


def test_parse_yes_no():
    assert p.parse_yes_no('yes') is True
    assert p.parse_yes_no('YES') is True
    assert p.parse_yes_no('  YES  ') is True
    assert p.parse_yes_no('no') is False
    assert p.parse_yes_no('NO') is False
    assert p.parse_yes_no('  NO  ') is False

    for case in {'yen', 'YEN', ''}:
        assert_import_errors({'must be "yes" or "no"'}, lambda: p.parse_yes_no(case))


def test_parse_int():
    assert_really_equal(p.parse_int('1'), 1)
    assert_really_equal(p.parse_int('-1'), -1)
    assert_really_equal(p.parse_int('123'), 123)
    assert_really_equal(p.parse_int('-123'), -123)
    assert_really_equal(p.parse_int('39.000'), 39)
    assert_really_equal(p.parse_int('-39.000'), -39)
    assert_really_equal(p.parse_int('1e+3'), 1000)
    assert_really_equal(p.parse_int(1), 1)

    for case in {'', 'a', '1ob', 'ob1', '3.2', 'one'}:
        assert_import_errors({'must be an integer'}, lambda: p.parse_int(case))


def test_parse_int_as_text():
    assert_really_equal(p.parse_int_as_text('1'), u'1')


def test_parse_number():
    assert_really_equal(p.parse_number('1'), 1.0)
    assert_really_equal(p.parse_number('-1'), -1.0)
    assert_really_equal(p.parse_number('3.9'), 3.9)
    assert_really_equal(p.parse_number('-3.9'), -3.9)
    assert_really_equal(p.parse_number('1e-3'), 0.001)
    assert_really_equal(p.parse_number(1.0), 1.0)

    for case in {'', 'a1', '1a', 'one'}:
        assert_import_errors({'must be a number'}, lambda: p.parse_number(case))


def test_parse_int_or_text():
    assert_really_equal(p.parse_int_or_text(''), u'')
    assert_really_equal(p.parse_int_or_text('   '), u'')
    assert_really_equal(p.parse_int_or_text('1'), u'1')
    assert_really_equal(p.parse_int_or_text('1.0'), u'1')
    assert_really_equal(p.parse_int_or_text(' 937'), u'937')
    assert_really_equal(p.parse_int_or_text('1e3'), u'1000')
    assert_really_equal(p.parse_int_or_text('a b c'), u'a b c')
    assert_really_equal(p.parse_int_or_text('  a b c  '), u'a b c')


def test_default_to():
    assert p.default_to('1')('') == '1'
    assert p.default_to('1')('  ') == '1'
    assert p.default_to('1')('1') == '1'
    assert p.default_to('1')('2') == '2'
    assert p.default_to(1)(None) == 1


def test_nullable():
    assert p.nullable(lambda _: 10)('') is None
    assert p.nullable(lambda _: 10)(None) is None
    assert p.nullable(lambda _: 10)('  ') is None
    assert p.nullable(lambda _: 10)('a') == 10
    assert p.nullable(lambda _: 10)(False) == 10
    assert p.nullable(lambda _: 10)(0) == 10


def test_validate_satisfies():
    def always(_):
        return True

    def never(_):
        return False

    assert p.validate_satisfies(always, '')('some-value') == 'some-value'
    assert p.validate_satisfies(always, '')(9) == 9

    assert_import_errors({'nope'}, lambda: p.validate_satisfies(never, 'nope')('some-value'))
    assert_import_errors({'nope'}, lambda: p.validate_satisfies(never, 'nope')(9))

    is_two = p.validate_satisfies(lambda x: x == 2, 'must be 2')
    assert is_two(2) == 2
    assert_import_errors({'must be 2'}, lambda: is_two(3))


def test_validate_not_empty():
    assert p.validate_not_empty('something') == 'something'
    assert_import_errors({'must not be empty'}, lambda: p.validate_not_empty(''))
    assert_import_errors({'must not be empty'}, lambda: p.validate_not_empty('  '))


def test_validate_max_length():
    assert p.validate_max_length(1)('a') == 'a'
    assert p.validate_max_length(2)('a') == 'a'
    assert p.validate_max_length(200)('a') == 'a'
    assert p.validate_max_length(200)('a'*200) == 'a'*200
    assert_import_errors({'must be no more than 1 characters long'},
                         lambda: p.validate_max_length(1)('ab'))
    assert_import_errors({'must be no more than 200 characters long'},
                         lambda: p.validate_max_length(200)('a'*201))


def test_validate_min():
    assert p.validate_min(-30)(-29) == -29
    assert p.validate_min(0)(0) == 0
    assert p.validate_min(1)(1) == 1
    assert p.validate_min(1)(2) == 2
    assert p.validate_min(1)(3.9) == 3.9
    assert p.validate_min(3.8)(3.9) == 3.9
    assert p.validate_min(float('-inf'))(1) == 1
    assert p.validate_min(100)(100) == 100

    assert_import_errors({'number must be no less than 12'},
                         lambda: p.validate_min(12)(11))
    assert_import_errors({'number must be no less than -30'},
                         lambda: p.validate_min(-30)(-31))
    assert_import_errors({'number must be no less than 3.9'},
                         lambda: p.validate_min(3.9)(3.8))


def test_validate_max():
    assert p.validate_max(-30)(-31) == -31
    assert p.validate_max(1)(1) == 1
    assert p.validate_max(3.9)(3.9) == 3.9
    assert p.validate_max(float('inf'))(1) == 1
    assert p.validate_max(100)(100) == 100

    assert_import_errors({'number must be no greater than 12'},
                         lambda: p.validate_max(12)(13))
    assert_import_errors({'number must be no greater than -30'},
                         lambda: p.validate_max(-30)(-29))
    assert_import_errors({'number must be no greater than 3.9'},
                         lambda: p.validate_max(3.9)(4))


def test_validate_range():
    assert p.validate_range(-30, 30)(0) == 0

    assert_import_errors({'number must be no greater than 30'},
                         lambda: p.validate_range(-30, 30)(31))
    assert_import_errors({'number must be no less than -30'},
                         lambda: p.validate_range(-30, 30)(-31))


def test_validate_unique():
    assert p.validate_unique({})(1) == 1
    assert p.validate_unique({})('abc') == 'abc'
    assert p.validate_unique({})(' abc ') == ' abc '
    assert p.validate_unique({1})(2) == 2
    assert p.validate_unique({2})(1) == 1
    assert p.validate_unique({1, 2})(3) == 3
    assert p.validate_unique(set(range(100)))(101) == 101

    fail_cases = [({1}, 1), ({1, 2}, 2), (set(range(100)), 99)]

    for others, value in fail_cases:
        with pytest.raises(p.SpreadsheetImportError):
            p.validate_unique(others)(value)


def test_validate_one_of():
    assert p.validate_one_of({1})(1) == 1
    assert p.validate_one_of({2})(2) == 2
    assert p.validate_one_of({1, 2})(1) == 1
    assert p.validate_one_of({1, 2})(2) == 2

    fail_cases = [({1}, 2), ({1, 2}, 3), (set(range(100)), 101)]

    for others, value in fail_cases:
        with pytest.raises(p.SpreadsheetImportError):
            p.validate_one_of(others)(value)

    assert_import_errors({u'must be a valid thing'}, lambda: p.validate_one_of({1})(2))
    assert_import_errors({u'must be a valid chicken'},
                         lambda: p.validate_one_of({1}, thing='chicken')(2))
    assert_import_errors(
        {u'must be a valid chicken: 1, 2, 3'},
        lambda: p.validate_one_of([1, 2, 3], thing='chicken', show_choices_in_error=True)(4)
    )


def test_parse_list():
    assert p.parse_list(delim=' ')('') == [u'']
    assert p.parse_list(delim=' ')('a') == [u'a']
    assert p.parse_list(delim=' ')('  a   ') == [u'a']
    assert p.parse_list(delim=',')('  a   ') == [u'a']
    assert p.parse_list(delim=' ')('a b') == [u'a', u'b']
    assert p.parse_list(delim=',')('a b') == [u'a b']
    assert p.parse_list(delim=',')('a, b, c') == [u'a', u' b', u' c']
    assert p.parse_list(p.parse_text, delim=',')('a, b, c') == [u'a', u'b', u'c']
    assert p.parse_list(lambda x: x.upper(), delim=',')('a,b,c') == [u'A', u'B', u'C']


def test_parse_lookup():
    assert p.parse_lookup({'a': 1})('a') == 1
    assert p.parse_lookup({'a': 1, 'b': 2})('b') == 2
    assert p.parse_lookup(defaultdict(lambda: None, {'a': 1}))('a') == 1
    assert p.parse_lookup(defaultdict(lambda: None, {'a': 1}))('b') is None


def test_filter_non_empty():
    assert p.filter_non_empty([]) == []
    assert p.filter_non_empty([1]) == [1]
    assert p.filter_non_empty([0, 1]) == [0, 1]
    assert p.filter_non_empty([False, True]) == [False, True]
    assert p.filter_non_empty([None]) == []
    assert p.filter_non_empty(['']) == []
    assert p.filter_non_empty(['', 1, None]) == [1]
    assert p.filter_non_empty((1, '', None, 2)) == [1, 2]
