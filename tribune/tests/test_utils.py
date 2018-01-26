import pytest

from tribune.utils import (
    column_letter,
    deepcopy_tuple_except,
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


class TestDeepCopyTupleExcept(object):
    class DoDeepCopy:
        def __init__(self, value):
            self.value = value

        def __eq__(self, other):
            return self.value == other.value

    class DoNotDeepCopy(DoDeepCopy):
        def __deepcopy__(self, memodict={}):
            raise Exception('Deep Copy!')

    def test_simple(self):
        original = (1, 2, 3)
        clone = deepcopy_tuple_except(original)
        assert original == clone

    def test_deep(self):
        original = (1, (2, self.DoNotDeepCopy([3, 4, 5]), 6), 7)
        clone = deepcopy_tuple_except(original, self.DoNotDeepCopy)
        assert original == clone

    def test_clone_is_clone(self):
        original = (1, [2, 3, 4], 5)
        clone = deepcopy_tuple_except(original)
        original[1].append(99)
        assert original != clone
        assert clone[0] == 1
        assert clone[1] == [2, 3, 4]
        assert clone[2] == 5

    def test_objects_are_deeply_copied(self):
        original = (1, self.DoDeepCopy([2, 3, 4]), 5)
        clone = deepcopy_tuple_except(original)
        assert original == clone

        assert clone[1].value == [2, 3, 4]
        original[1].value.append(99)
        assert clone[1].value == [2, 3, 4]

    def test_exceptions_are_not_deeply_copied(self):
        original = (1, self.DoNotDeepCopy([2, 3, 4]), 5)
        clone = deepcopy_tuple_except(original, self.DoNotDeepCopy)
        assert original == clone

        assert clone[1].value == [2, 3, 4]
        original[1].value.append(99)
        assert clone[1].value == [2, 3, 4, 99]

        assert original == clone

    def test_exceptions_are_not_copied_at_all(self):
        original = (1, self.DoNotDeepCopy([2, 3, 4]), 5)
        clone = deepcopy_tuple_except(original, self.DoNotDeepCopy)

        assert original[1] is clone[1]

    def test_exceptions_are_used(self):
        original = (1, self.DoNotDeepCopy([2, 3, 4]), 5)

        with pytest.raises(Exception):
            deepcopy_tuple_except(original)


def test_split_every():
    assert list(split_every(1, [])) == []
    assert list(split_every(1, [1])) == [[1]]
    assert list(split_every(1, [1, 2])) == [[1], [2]]
    assert list(split_every(2, [1, 2])) == [[1, 2]]
    assert list(split_every(3, [1, 2])) == [[1, 2]]
    assert list(split_every(2, [1, 2, 3])) == [[1, 2], [3]]
