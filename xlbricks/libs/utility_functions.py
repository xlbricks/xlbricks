"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: XLUtils (get_bricks, crop_range, etc.), XLBricksFunction decorator,
                 and XLBricksUtils; bridges Excel ranges to brick structures.
"""

import numpy as np
import pandas as pd
from xlbricks.libs.xlbricks import XLBrick, XLBricks
from xlbricks.libs.xlbricks_front import XLBricksFront
from xlbricks.libs.xlbricks_frontstack import XLBricksFrontStack, add_bricks_to_front_stack, delete_bricks_from_front_stack


class XLBricksFunction(object):
    """Decorator that wraps functions to automatically manage brick storage and references.
    
    Registers results in the front stack and returns brick references for Excel.
    """

    def __init__(self, is_dynamic: bool = False):
        """Initialize the decorator.
        
        If is_dynamic is True, returns raw output for non-brick results instead of wrapping them.
        """
        self.is_dynamic = is_dynamic

    def __call__(self, f):
        """Apply the decorator to a function.
        
        Wraps the function to add results to the front stack and return references.
        """
        def wrap(*args, **kwargs):
            xl_output = f(*args, **kwargs)
            if self.is_dynamic and not isinstance(xl_output, XLBricksFront):
                return xl_output

            add_bricks_to_front_stack(xl_output)
            return xl_output.bricks_full_name
        return wrap


class XLUtils(object):
    """Utilities for converting Excel ranges to bricks and managing brick references.
    
    Handles data cropping, type conversion, and front stack lookups.
    """

    @staticmethod
    def get_bricks_front(data):
        """Look up a brick by its reference string.
        
        Expects a 1x1 cell containing 'alias:counter' format.
        """
        if XLUtils.is_bricks_front_name(data):
            key = ''.join(data[0, 0].split(':')[:-1])
            return XLBricksFrontStack()[key]
        else:
            return None

    @staticmethod
    def get_bricks(data):
        """Convert Excel range data into brick objects.
        
        Automatically crops empty cells, resolves brick references, and converts data types.
        """
        data = XLUtils.crop_range(data)
        bricks_front = XLUtils.get_bricks_front(data)
        if bricks_front is None:
            if data.dtype.type is np.str_:
                data = pd.DataFrame(data).apply(pd.to_numeric, errors='ignore').values
            return XLBrick(None, data)
        else:
            delete_bricks_from_front_stack(bricks_front)
            return bricks_front.xlbricks

    @staticmethod
    def delete_bricks(data):
        """Delete a brick from memory using its reference.
        
        Expects a 1x1 cell containing the brick reference to remove.
        """
        if XLUtils.is_bricks_front_name(data):
            key = ''.join(data[0, 0].split(':')[:-1])
            del XLBricksFrontStack()[key]

    @staticmethod
    def is_bricks_front_name(data):
        """Check if data contains a brick reference string.
        
        Returns True if it's a 1x1 cell with a colon-separated reference.
        """
        if data.shape == (1, 1) and isinstance(data[0, 0], str) and ':' in data[0, 0]:
            return True
        else:
            return False

    @staticmethod
    def crop_range(data):
        """Trim empty rows and columns from all edges of a data range.
        
        Automatically detects and removes surrounding empty cells.
        """
        if data.dtype.type is np.str_:
            return XLUtils._crop_range(data, lambda x: x == 'nan')
        elif data.dtype.type is np.object_:
            return XLUtils._crop_range(data, lambda x: pd.isnull(x))
        else:
            return XLUtils._crop_range(data, lambda x: np.isnan(x))

    @staticmethod
    def _crop_range(data, func_isnan):
        """Internal helper to crop range using a custom empty-cell detection function.
        
        Removes empty rows and columns from all four edges.
        """
        while func_isnan(data[0, :]).all():
            data = np.delete(data, 0, axis=0)

        while func_isnan(data[:, 0]).all():
            data = np.delete(data, 0, axis=1)

        rows, cols = data.shape
        row_idx = rows - 1
        while row_idx > 0 and func_isnan(data[row_idx, :]).all():
            data = np.delete(data, row_idx, axis=0)
            row_idx -= 1

        rows, cols = data.shape
        col_idx = cols - 1
        while col_idx > 0 and func_isnan(data[:, col_idx]).all():
            data = np.delete(data, col_idx, axis=1)
            col_idx -= 1

        return data

    @staticmethod
    def active_cell_address(xl_app):
        """Get the full address of the cell calling the function.
        
        Returns format: '[Workbook]Sheet!Address' for persistence tracking.
        """
        active_cell = xl_app.Caller
        worksheet = active_cell.Parent
        workbook = worksheet.Parent
        address = '[%s]%s!%s' % (workbook.Name, worksheet.Name, active_cell.Address)
        return address


class XLBricksUtils(object):
    """Utilities for converting Python data structures into brick objects.
    
    Handles dictionaries and lists, creating appropriate brick hierarchies.
    """

    @staticmethod
    def element_from_dictionary(input_data):
        """Convert a nested dictionary into a brick hierarchy.
        
        Each key-value pair becomes a brick, with nested dicts creating sub-bricks.
        """

        qd_element = XLBricks()
        for key, data in input_data.items():
            if isinstance(data, dict):
                qd_element[key] = XLBricksUtils.element_from_dictionary(data)
            else:
                qd_element[key] = XLBricks(None, data)

        return qd_element

    @staticmethod
    def element_from_list(input_data, key_prefix):
        """Convert a list into numbered bricks.
        
        Creates bricks with keys like 'prefix_1', 'prefix_2', etc.
        """

        qd_element = XLBricks()
        for idx, res in enumerate(input_data, 1):
            qd_element['%s_%s' % (key_prefix, idx)] = XLBricks(None, res)

        return qd_element

