"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Unit tests for validation helpers (libs.validation).
                 No PyQt/QuantLib required; runs with numpy only.
"""

import unittest
import numpy as np
import sys
import os

_here = os.path.dirname(os.path.abspath(__file__))
_xlbricks = os.path.dirname(_here)
_root = os.path.dirname(_xlbricks)
if _root not in sys.path:
    sys.path.insert(0, _root)

from xlbricks.libs.validation import (
    _is_missing,
    _check_required,
    _check_array_2d,
    _ERROR_PREFIX,
)


class TestIsMissing(unittest.TestCase):
    def test_none_is_missing(self):
        self.assertTrue(_is_missing(None))

    def test_nan_float_is_missing(self):
        self.assertTrue(_is_missing(float('nan')))
        self.assertTrue(_is_missing(np.nan))

    def test_empty_string_is_missing(self):
        self.assertTrue(_is_missing(''))
        self.assertTrue(_is_missing('   '))
        self.assertTrue(_is_missing('\t'))

    def test_nan_string_is_missing(self):
        self.assertTrue(_is_missing('nan'))
        self.assertTrue(_is_missing('NaN'))

    def test_empty_array_is_missing(self):
        self.assertTrue(_is_missing(np.array([])))
        self.assertTrue(_is_missing(np.array([[]])))

    def test_all_nan_array_is_missing(self):
        self.assertTrue(_is_missing(np.array([[np.nan, np.nan]])))

    def test_valid_scalars_not_missing(self):
        self.assertFalse(_is_missing(0))
        self.assertFalse(_is_missing(1.0))
        self.assertFalse(_is_missing('a'))
        self.assertFalse(_is_missing(True))

    def test_valid_array_not_missing(self):
        self.assertFalse(_is_missing(np.array([[1, 2], [3, 4]])))


class TestCheckRequired(unittest.TestCase):
    def test_missing_returns_error_string(self):
        out = _check_required('key', None)
        self.assertIsNotNone(out)
        self.assertTrue(out.startswith(_ERROR_PREFIX))
        self.assertIn('key', out)
        self.assertIn('required', out)

    def test_valid_returns_none(self):
        self.assertIsNone(_check_required('key', 'x'))
        self.assertIsNone(_check_required('key', 1))

    def test_allow_none_with_none_returns_none(self):
        self.assertIsNone(_check_required('opt', None, allow_none=True))

    def test_allow_none_with_missing_string_still_errors(self):
        out = _check_required('opt', '', allow_none=True)
        self.assertIsNotNone(out)


class TestCheckArray2d(unittest.TestCase):
    def test_none_required_returns_error(self):
        out = _check_array_2d('data', None)
        self.assertIsNotNone(out)
        self.assertIn('data', out)
        self.assertIn('required', out)

    def test_none_optional_returns_none(self):
        self.assertIsNone(_check_array_2d('data', None, required=False))

    def test_not_array_returns_error(self):
        out = _check_array_2d('data', 'hello')
        self.assertIsNotNone(out)
        self.assertIn('2D range', out)

    def test_1d_array_returns_error(self):
        out = _check_array_2d('data', np.array([1, 2, 3]))
        self.assertIsNotNone(out)

    def test_empty_2d_returns_error(self):
        out = _check_array_2d('data', np.array([[]]))
        self.assertIsNotNone(out)
        self.assertIn('empty', out)

    def test_valid_2d_returns_none(self):
        self.assertIsNone(_check_array_2d('data', np.array([[1, 2], [3, 4]])))
        self.assertIsNone(_check_array_2d('data', np.array([['a']])))


if __name__ == '__main__':
    unittest.main()
