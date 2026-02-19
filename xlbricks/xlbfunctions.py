"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Excel UDF (User Defined Function) entry points for XLBricks;
                 exposes xlb_* functions to Excel via xlwings and delegates to xlfunctions.
"""

import sys
import inspect
import os.path as osp
import numpy as np
import xlwings as xw
from datetime import datetime
from functools import wraps
import xlbricks.libs.xlfunctions as xl
from xlbricks.libs.utility_functions import XLUtils

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
from xlbricks.ui.explorer import Explorer
from xlbricks.ui.tree_model import DictionaryTreeModel, node_structure_from_dict
from xlbricks.ui.config_editor import show_config_editor
from xlbricks.libs.xlbricks_frontstack import XLBricksFrontStack
from xlbricks.libs.validation import (
    _ERROR_PREFIX,
    _is_missing,
    _check_required,
    _check_array_2d,
)


def _return_errors(f):
    """Decorator: catch exceptions and return a #ERROR: message string for Excel.
    Preserves the wrapped function's signature so xlwings UDF inspection still works."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        try:
            return f(*args, **kwargs)
        except Exception as e:
            return _ERROR_PREFIX + '%s: %s' % (type(e).__name__, str(e))
    wrapper.__signature__ = inspect.signature(f)
    return wrapper


@xw.func
@xw.arg('key')
@xw.arg('data', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_brick(key, data, persist=True, xlapp=None):
    """Create a single named brick from Excel data.
    
    Returns a reference string (e.g., 'mykey:1') that can be used in other functions.
    """
    err = _check_required('key', key) or _check_array_2d('data', data)
    if err:
        return err
    return xl.xlbrick_create(key, data, persist, xlapp)


@xw.func
@xw.arg('key_1')
@xw.arg('brick_1', np.array, ndim=2)
@xw.arg('key_2')
@xw.arg('brick_2', np.array, ndim=2)
@xw.arg('key_3')
@xw.arg('brick_3', np.array, ndim=2)
@xw.arg('key_4')
@xw.arg('brick_4', np.array, ndim=2)
@xw.arg('key_5')
@xw.arg('brick_5', np.array, ndim=2)
@xw.arg('key_6')
@xw.arg('brick_6', np.array, ndim=2)
@xw.arg('key_7')
@xw.arg('brick_7', np.array, ndim=2)
@xw.arg('key_8')
@xw.arg('brick_8', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_bricks(key_1, brick_1, key_2=None, brick_2=None, key_3=None, brick_3=None,
               key_4=None, brick_4=None, key_5=None, brick_5=None, key_6=None, brick_6=None,
               key_7=None, brick_7=None, key_8=None, brick_8=None, persist=True, xlapp=None):
    """Create multiple named bricks at once (up to 8 key-value pairs).
    
    Returns a reference to the collection of bricks.
    """
    err = _check_required('key_1', key_1) or _check_array_2d('brick_1', brick_1)
    if err:
        return err
    return xl.xlbricks_create(key_1, brick_1, key_2, brick_2, key_3, brick_3, key_4, brick_4,
                              key_5, brick_5, key_6, brick_6, key_7, brick_7, key_8, brick_8, persist, xlapp)


@xw.func
@xw.arg('data', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_array(data, persist=True, xlapp=None):
    """Store an Excel range as a brick array.
    
    Preserves the original structure and data types of the range.
    """
    err = _check_array_2d('data', data)
    if err:
        return err
    return xl.array_create(data, persist, xlapp)


@xw.func
@xw.arg('data', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_list(data, persist=True, xlapp=None):
    """Convert an Excel range into a flattened list.
    
    Useful for creating one-dimensional sequences from multi-cell ranges.
    """
    err = _check_array_2d('data', data)
    if err:
        return err
    return xl.list_create(data, persist, xlapp)


@xw.func
@xw.arg('data', np.array, ndim=2)
@xw.arg('index', np.array, ndim=2)
@xw.arg('columns', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_table(data, columns=None, index=None, persist=True, xlapp=None):
    """Create a pandas DataFrame from Excel data with optional column names and row index.
    
    Ideal for structured tabular data with headers.
    """
    err = _check_array_2d('data', data)
    if err:
        return err
    return xl.table_create(data, columns, index, persist, xlapp)


@xw.func
@xw.arg('data', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_grid(data, persist=True, xlapp=None):
    """Parse a grid layout where the first column contains keys and remaining columns contain data.
    
    Each row with a key in column 1 becomes a separate brick.
    """
    err = _check_array_2d('data', data)
    if err:
        return err
    return xl.grid_create(data, persist, xlapp)


@xw.func
@xw.arg('bricks', np.array, ndim=2)
@xw.arg('keys')
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_lookup(bricks, keys=None, persist=True, xlapp=None):
    """Retrieve a nested brick using a path like 'parent/child/grandchild'.
    
    Use forward slashes to navigate through nested brick structures.
    """
    err = _check_array_2d('bricks', bricks) or _check_required('keys', keys)
    if err:
        return err
    return xl.lookup_element(bricks, keys, persist, xlapp)


@xw.func
@xw.arg('brick', np.array, ndim=2)
@_return_errors
def xlb_flatten(brick):
    """Extract the raw data from a brick reference.
    
    Returns the underlying array or value without the brick wrapper.
    """
    err = _check_array_2d('brick', brick)
    if err:
        return err
    return xl.flatten_element(brick)


@xw.func
@xw.arg('brick', np.array, ndim=2)
@_return_errors
def xlb_alias(brick, alias):
    """Assign a custom name (alias) to an existing brick.
    
    Makes bricks easier to reference with memorable names instead of cell addresses.
    """
    err = _check_array_2d('brick', brick) or _check_required('alias', alias)
    if err:
        return err
    return xl.assign_alias(brick, alias)


@xw.func
@xw.arg('functions', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_create_function(functions, persist=True, xlapp=None):
    """Define Python functions directly in Excel cells.
    
    Write function code in a range and execute it to create callable functions.
    """
    err = _check_array_2d('functions', functions)
    if err:
        return err
    return xl.create_function_objects(functions, persist, xlapp)


@xw.func
@xw.arg('context_name')
@xw.arg('context_path')
@xw.arg('args', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_create_context(context_name, context_path, args=None, persist=True, xlapp=None):
    """Create an instance of a Python class from a module path.
    
    Useful for instantiating objects like QuantLib contexts with optional arguments.
    """
    err = _check_required('context_name', context_name) or _check_required('context_path', context_path)
    if err:
        return err
    return xl.create_context_object(context_name, context_path, args, persist, xlapp)


@xw.func
@xw.arg('function_brick', np.array, ndim=2)
@xw.arg('function_name')
@xw.arg('args', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_run_function(function_brick, function_name, args=None, persist=True, xlapp=None):
    """Execute a function stored in a brick with optional arguments.
    
    Pass arguments as key-value pairs in a range.
    """
    err = _check_array_2d('function_brick', function_brick) or _check_required('function_name', function_name)
    if err:
        return err
    return xl.run_function(function_brick, function_name, args, persist, xlapp)


@xw.func
@xw.arg('quantlib_object', np.array, ndim=2)
@xw.arg('function_name')
@xw.arg('args', np.array, ndim=2)
@xw.arg('xl_app', vba='Application')
@_return_errors
def xlb_run_quantlib_function(quantlib_object, function_name, args=None, persist=True, xl_app=None):
    """Call a method on a QuantLib object stored in a brick.
    
    Enables financial calculations using QuantLib directly from Excel.
    """
    err = _check_array_2d('quantlib_object', quantlib_object) or _check_required('function_name', function_name)
    if err:
        return err
    return xl.run_quantlib_function(quantlib_object, function_name, args, persist, xl_app)


@xw.func
@xw.arg('brick_1', np.array, ndim=2)
@xw.arg('brick_2', np.array, ndim=2)
@xw.arg('brick_3', np.array, ndim=2)
@xw.arg('brick_4', np.array, ndim=2)
@xw.arg('brick_5', np.array, ndim=2)
@xw.arg('xlapp', vba='Application')
@_return_errors
def xlb_merge(brick_1, brick_2, brick_3=None, brick_4=None, brick_5=None, persist=True, xlapp=None):
    """Combine multiple bricks into a single collection (up to 5 bricks).
    
    All keys from the input bricks are merged into one unified brick.
    """
    err = _check_array_2d('brick_1', brick_1) or _check_array_2d('brick_2', brick_2)
    if err:
        return err
    return xl.merge_elements(brick_1, brick_2, brick_3, brick_4, brick_5, persist, xlapp)


@xw.func
@_return_errors
def xlb_today():
    """Return today's date.
    
    Simple utility function for getting the current date in Excel.
    """
    return datetime.today()


@xw.func
@_return_errors
def xlb_clear_bricks_front():
    """Clear all bricks from memory.
    
    Use this to reset the brick storage and free up memory.
    """
    return xl.clear_bricks_front()


@xw.func
@_return_errors
def xlb_open_bricks_explorer():
    """Open a visual explorer window showing all bricks currently in memory.
    
    Browse the brick hierarchy and view data in a tree and table interface.
    """
    explorer_app = QApplication(sys.argv)
    img_path = _get_image_path('stars.png')
    explorer_app.setWindowIcon(QIcon(img_path))
    model = DictionaryTreeModel(node_structure_from_dict(XLBricksFrontStack().to_dict()))
    wizard = Explorer(model)
    wizard.display()
    sys.exit(explorer_app.exec_())


@xw.func
@xw.arg('data', np.array, ndim=2)
@_return_errors
def xlb_open_brick_explorer(data):
    """Open a visual explorer window for a specific brick.
    
    View the structure and contents of a single brick in detail.
    """
    err = _check_array_2d('data', data)
    if err:
        return err
    if XLUtils.is_bricks_front_name(data):
        key = ''.join(data[0, 0].split(':')[:-1])
        element = XLUtils.get_bricks(data)
        explorer_app = QApplication(sys.argv)
        img_path = _get_image_path('stars.png')
        explorer_app.setWindowIcon(QIcon(img_path))
        model = DictionaryTreeModel(node_structure_from_dict({key: element.to_dict()}))
        wizard = Explorer(model)
        wizard.display_one_element()
        sys.exit(explorer_app.exec_())


@xw.func
@_return_errors
def xlb_open_config_editor():
    """Open the XLBricks config editor UI (same as XLBricks Wizard, from Excel)."""
    config_app = QApplication(sys.argv)
    img_path = _get_image_path('settings.png')
    config_app.setWindowIcon(QIcon(img_path))
    show_config_editor()
    sys.exit(0)


def _get_package_dir():
    """Return the absolute path to the xlbricks package directory (where xlbfunctions.py lives)."""
    this_file = getattr(sys.modules[__name__], '__file__', None)
    if not this_file:
        return ''
    return osp.abspath(osp.dirname(this_file))


def _get_image_path(name: str):
    """Return path to an icon from the imgs folder. Works when run from Excel (uses absolute path)."""
    pkg_dir = _get_package_dir()
    if not pkg_dir:
        return ''
    imgs_dir = osp.join(pkg_dir, 'imgs')
    path = osp.join(imgs_dir, name)
    if osp.isfile(path):
        return path
    return ''


if __name__ == '__main__':
    xw.serve()
