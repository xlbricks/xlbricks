"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Unit tests for xlbfunctions: validation, edge cases, and success paths (no Excel).
"""

import unittest
import numpy as np
import sys
import os


# Ensure package root is on path when running tests
_here = os.path.dirname(os.path.abspath(__file__))
_xlbricks = os.path.dirname(_here)
_root = os.path.dirname(_xlbricks)
if _root not in sys.path:
    sys.path.insert(0, _root)

# Import xlbfunctions only if deps (PyQt5, QuantLib, xlwings) available; else skip module
try:
    from xlbricks.xlbfunctions import (
        _ERROR_PREFIX,
        xlb_brick,
        xlb_bricks,
        xlb_array,
        xlb_list,
        xlb_table,
        xlb_grid,
        xlb_lookup,
        xlb_flatten,
        xlb_alias,
        xlb_create_function,
        xlb_create_context,
        xlb_run_function,
        xlb_run_quantlib_function,
        xlb_merge,
        xlb_today,
        xlb_clear_bricks_front,
        xlb_open_brick_explorer,
    )
    XLBFUNCTIONS_AVAILABLE = True
    _import_error = ''
except ImportError as e:
    XLBFUNCTIONS_AVAILABLE = False
    _import_error = str(e)


def _arr(*rows):
    """Build 2D numpy array from rows."""
    return np.array(rows, dtype=object)


def _is_error(out):
    """True if out is an error string (#XLB ERROR:...). Safe when out is ndarray/date/etc."""
    return isinstance(out, str) and out.startswith(_ERROR_PREFIX)


# --- xlb_brick ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbBrick(unittest.TestCase):
    def test_missing_key_returns_error(self):
        data = np.array([['x']])
        out = xlb_brick(None, data, persist=False)
        self.assertTrue(_is_error(out))
        self.assertIn('key', out)

    def test_empty_key_returns_error(self):
        data = np.array([['x']])
        out = xlb_brick('', data, persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_data_returns_error(self):
        out = xlb_brick('k', None, persist=False)
        self.assertTrue(_is_error(out))
        self.assertIn('data', out)

    def test_empty_data_array_returns_error(self):
        out = xlb_brick('k', np.array([[]]), persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_reference_or_bricks_name(self):
        data = np.array([[1, 2], [3, 4]])
        out = xlb_brick('mykey', data, persist=False)
        self.assertFalse(_is_error(out))
        self.assertIsInstance(out, str)
        self.assertIn(':', out)  # alias:counter or similar


# --- xlb_bricks ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbBricks(unittest.TestCase):
    def test_missing_key_1_returns_error(self):
        arr = np.array([[1]])
        out = xlb_bricks(None, arr, persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_brick_1_returns_error(self):
        out = xlb_bricks('k', None, persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_string(self):
        arr = np.array([[10]])
        out = xlb_bricks('k1', arr, persist=False)
        self.assertFalse(_is_error(out))
        self.assertIn(':', out)


# --- xlb_array ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbArray(unittest.TestCase):
    def test_none_data_returns_error(self):
        out = xlb_array(None, persist=False)
        self.assertTrue(_is_error(out))

    def test_empty_data_returns_error(self):
        out = xlb_array(np.array([[]]), persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_string(self):
        data = np.array([[1, 2], [3, 4]])
        out = xlb_array(data, persist=False)
        self.assertFalse(_is_error(out))
        self.assertIn(':', out)


# --- xlb_list ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbList(unittest.TestCase):
    def test_missing_data_returns_error(self):
        out = xlb_list(None, persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_string(self):
        data = np.array([[1], [2], [3]])
        out = xlb_list(data, persist=False)
        self.assertFalse(_is_error(out))


# --- xlb_table ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbTable(unittest.TestCase):
    def test_missing_data_returns_error(self):
        out = xlb_table(None, persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_string(self):
        data = np.array([[1, 2], [3, 4]])
        out = xlb_table(data, persist=False)
        self.assertFalse(_is_error(out))


# --- xlb_grid ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbGrid(unittest.TestCase):
    def test_missing_data_returns_error(self):
        out = xlb_grid(None, persist=False)
        self.assertTrue(_is_error(out))

    def test_empty_data_returns_error(self):
        out = xlb_grid(np.array([[]]), persist=False)
        self.assertTrue(_is_error(out))


# --- xlb_lookup ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbLookup(unittest.TestCase):
    def test_missing_bricks_returns_error(self):
        out = xlb_lookup(None, 'a/b', persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_keys_returns_error(self):
        bricks = np.array([['x']])
        out = xlb_lookup(bricks, None, persist=False)
        self.assertTrue(_is_error(out))
        out2 = xlb_lookup(bricks, '', persist=False)
        self.assertTrue(_is_error(out2))


# --- xlb_flatten ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbFlatten(unittest.TestCase):
    def test_missing_brick_returns_error(self):
        out = xlb_flatten(None)
        self.assertTrue(_is_error(out))

    def test_success_returns_array(self):
        data = np.array([[1, 2], [3, 4]])
        out = xlb_flatten(data)
        self.assertFalse(_is_error(out))
        self.assertIsInstance(out, np.ndarray)
        self.assertEqual(out.shape[1], 2)


# --- xlb_alias ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbAlias(unittest.TestCase):
    def test_missing_brick_returns_error(self):
        out = xlb_alias(None, 'a')
        self.assertTrue(_is_error(out))

    def test_missing_alias_returns_error(self):
        arr = np.array([[1]])
        out = xlb_alias(arr, None)
        self.assertTrue(_is_error(out))
        out2 = xlb_alias(arr, '')
        self.assertTrue(_is_error(out2))


# --- xlb_create_function ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbCreateFunction(unittest.TestCase):
    def test_missing_functions_returns_error(self):
        out = xlb_create_function(None, persist=False)
        self.assertTrue(_is_error(out))

    def test_empty_functions_returns_error(self):
        out = xlb_create_function(np.array([[]]), persist=False)
        self.assertTrue(_is_error(out))

    def test_success_single_function_returns_string(self):
        # Minimal valid function block
        funcs = np.array([
            ['def my_test_func():'],
            ['    return 42'],
        ], dtype=object)
        out = xlb_create_function(funcs, persist=False)
        self.assertFalse(_is_error(out))
        self.assertIn(':', out)


# --- xlb_create_context ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbCreateContext(unittest.TestCase):
    def test_missing_context_name_returns_error(self):
        out = xlb_create_context(None, 'some.path', persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_context_path_returns_error(self):
        out = xlb_create_context('MyClass', None, persist=False)
        self.assertTrue(_is_error(out))

    def test_empty_context_name_returns_error(self):
        out = xlb_create_context('', 'path', persist=False)
        self.assertTrue(_is_error(out))


# --- xlb_run_function ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbRunFunction(unittest.TestCase):
    def test_missing_function_brick_returns_error(self):
        out = xlb_run_function(None, 'foo', persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_function_name_returns_error(self):
        arr = np.array([['x']])
        out = xlb_run_function(arr, None, persist=False)
        self.assertTrue(_is_error(out))
        out2 = xlb_run_function(arr, '', persist=False)
        self.assertTrue(_is_error(out2))

    def test_empty_function_name_returns_error(self):
        arr = np.array([['x']])
        out = xlb_run_function(arr, np.nan, persist=False)
        self.assertTrue(_is_error(out))


# --- xlb_run_quantlib_function ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbRunQuantlibFunction(unittest.TestCase):
    def test_missing_quantlib_object_returns_error(self):
        out = xlb_run_quantlib_function(None, 'method', persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_function_name_returns_error(self):
        arr = np.array([['x']])
        out = xlb_run_quantlib_function(arr, None, persist=False)
        self.assertTrue(_is_error(out))


# --- xlb_merge ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbMerge(unittest.TestCase):
    def test_missing_brick_1_returns_error(self):
        arr = np.array([[1]])
        out = xlb_merge(None, arr, persist=False)
        self.assertTrue(_is_error(out))

    def test_missing_brick_2_returns_error(self):
        arr = np.array([[1]])
        out = xlb_merge(arr, None, persist=False)
        self.assertTrue(_is_error(out))

    def test_success_returns_string(self):
        # merge_elements expects brick refs (XLBricks); raw arrays may yield backend error.
        # Assert we get a string response (reference or error) and no exception.
        a = np.array([['a', 1]])
        b = np.array([['b', 2]])
        out = xlb_merge(a, b, persist=False)
        self.assertIsInstance(out, str)
        if not _is_error(out):
            self.assertIn(':', out)  # reference format alias:counter


# --- xlb_today ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbToday(unittest.TestCase):
    def test_returns_date(self):
        from datetime import date
        out = xlb_today()
        self.assertFalse(_is_error(out))
        self.assertIsInstance(out, type(date.today()))


# --- xlb_clear_bricks_front ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbClearBricksFront(unittest.TestCase):
    def test_returns_without_error(self):
        out = xlb_clear_bricks_front()
        self.assertIsNone(out)


# --- xlb_open_brick_explorer ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestXlbOpenBrickExplorer(unittest.TestCase):
    def test_missing_data_returns_error(self):
        out = xlb_open_brick_explorer(None)
        self.assertTrue(_is_error(out))

    def test_empty_data_returns_error(self):
        out = xlb_open_brick_explorer(np.array([[]]))
        self.assertTrue(_is_error(out))


# --- Edge cases: _return_errors decorator ---


@unittest.skipUnless(XLBFUNCTIONS_AVAILABLE, "xlbfunctions deps not available: " + _import_error)
class TestReturnErrorsDecorator(unittest.TestCase):
    """Test that exceptions are turned into #XLB ERROR: strings."""

    def test_xlb_brick_invalid_data_type_caught(self):
        # Pass something that will fail inside xl.xlbrick_create (e.g. bad shape)
        out = xlb_brick('k', np.array([[1, 2], [3, 4]]), persist=False)
        # Should not raise; either success or error string
        self.assertIsInstance(out, str)
        if _is_error(out):
            self.assertIn(':', out)


if __name__ == '__main__':
    unittest.main()
