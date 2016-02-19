from tribune.utils import (
    column_letter,
    split_every,
)


def test_column_letter():
    assert column_letter(0) == 'A'
    assert column_letter(2) == 'C'
    assert column_letter(25) == 'Z'
    assert column_letter(26) == 'AA'
    assert column_letter(27) == 'AB'
    assert column_letter(26*2) == 'BA'
    assert column_letter(26*3) == 'CA'
    assert column_letter(26**2+25) == 'ZZ'
    assert column_letter(26**2+26) == 'AAA'


def test_split_every():
    assert list(split_every(1, [])) == []
    assert list(split_every(1, [1])) == [[1]]
    assert list(split_every(1, [1, 2])) == [[1], [2]]
    assert list(split_every(2, [1, 2])) == [[1, 2]]
    assert list(split_every(3, [1, 2])) == [[1, 2]]
    assert list(split_every(2, [1, 2, 3])) == [[1, 2], [3]]
