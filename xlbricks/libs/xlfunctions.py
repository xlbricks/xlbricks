"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Core XLBricks operations: create bricks/arrays/tables/grids,
                 run user-defined functions, lookup, merge; used by xlbfunctions.
"""
import numpy as np
import pandas as pd
import importlib
from copy import deepcopy
from xlbricks.libs.xlbricks import XLBrick, XLBricks
from xlbricks.libs.xlbricks_front import XLBricksFront
from xlbricks.libs.xlbricks_frontstack import XLBricksFrontStack
from xlbricks.libs.utility_functions import XLUtils, XLBricksUtils, XLBricksFunction


@XLBricksFunction(False)
def xlbrick_create(key, data, persist=True, xlapp=None):
    """Create a single named brick from data.
    
    Internal function called by xlb_brick Excel function.
    """
    xlbricks = XLBricks()
    xlbricks[key] = XLUtils.get_bricks(data)

    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def xlbricks_create(key_1, brick_1, key_2=None, brick_2=None, key_3=None, brick_3=None,
                    key_4=None, brick_4=None, key_5=None, brick_5=None, key_6=None, brick_6=None,
                    key_7=None, brick_7=None, key_8=None, brick_8=None, persist=True, xlapp=None):
    """Create multiple named bricks in one operation.
    
    Internal function called by xlb_bricks Excel function.
    """

    xlbricks = XLBricks()
    for idx in range(1, 9):
        key = locals().get('key_%s' % idx, None)
        xlbrick = locals().get('brick_%s' % idx, None)
        if key is not None and xlbrick is not None:
            xlbricks[key] = XLUtils.get_bricks(xlbrick)

    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def array_create(data, persist=True, xlapp=None):
    """Create a brick containing array data.
    
    Stores the data as-is without modification.
    """
    xlbricks = XLUtils.get_bricks(data)
    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def list_create(data, persist=True, xlapp=None):
    """Create a brick containing a flattened list.
    
    Converts 2D ranges into 1D Python lists.
    """
    xlbricks = XLUtils.get_bricks(data)
    if isinstance(xlbricks.value, np.ndarray):
        xlbricks.value = xlbricks.value.flatten().tolist()
    else:
        xlbricks.value = [xlbricks.value]

    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def table_create(data, columns=None, index=None, persist=True, xlapp=None):
    """Create a pandas DataFrame brick with optional headers and index.
    
    Enables structured table operations on Excel data.
    """

    data_brick = XLUtils.get_bricks(data)

    if index is not None:
        index_brick = XLUtils.get_bricks(index)
        index = index_brick.value.flatten()

    if columns is not None:
        columns_brick = XLUtils.get_bricks(columns)
        columns = columns_brick.value.flatten()

    xlbrick = XLBrick(None, pd.DataFrame(data_brick.value, index, columns))
    xlbrick_front = create_bricks_front(xlbrick, xlapp, persist)
    return xlbrick_front


@XLBricksFunction(False)
def grid_create(data, persist=True, xlapp=None):
    """Parse a grid where column 1 has keys and remaining columns have values.
    
    Creates multiple bricks from a single structured range.
    """
    xlbricks = XLBricks()

    if data.dtype.type is np.str_:
        keys_idx = np.argwhere(np.array(data[:, 0]) != 'nan').flatten()
    elif data.dtype.type is np.object_:
        keys_idx = np.argwhere(np.array(~pd.isnull(data[:, 0]))).flatten()
    else:
        keys_idx = np.argwhere(np.array(~np.isnan(data[:, 0]))).flatten()

    for f_idx, s_idx in zip(keys_idx, keys_idx[1:]):
        xlbricks[data[f_idx, 0]] = XLUtils.get_bricks(data[f_idx:s_idx, 1:])

    xlbricks[data[s_idx, 0]] = XLUtils.get_bricks(data[s_idx:, 1:])

    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def lookup_element(bricks, keys=None, persist=True, xlapp=None):
    """Navigate to a nested brick using a slash-separated path.
    
    Example: 'parent/child' retrieves the child brick from parent.
    """
    xlbricks = XLUtils.get_bricks(bricks)
    keys = [key.strip(' \t\n\r') for key in keys.split('/')]
    xlbricks_front = create_bricks_front(xlbricks[keys], xlapp, persist)
    return xlbricks_front


@XLBricksFunction(True)
def flatten_bricks(bricks):
    """Extract raw data from a brick as a numpy array.
    
    Returns the underlying value without brick metadata.
    """
    xlbricks = XLUtils.get_bricks(bricks)
    xlbricks_value = xlbricks.value

    if not isinstance(xlbricks_value, np.ndarray):
        xlbricks_value = np.array(xlbricks_value)

    if xlbricks_value.ndim == 1:
        xlbricks_value = xlbricks_value.reshape(-1, 1)

    return xlbricks_value


def _func_line_sanitize(cell):
    """Clean a cell value for safe Python code execution.
    
    Converts empty/NaN cells to empty strings while preserving indentation.
    """
    if cell is None:
        return ''
    if isinstance(cell, (float, np.floating)) and np.isnan(cell):
        return ''
    try:
        if pd.isnull(cell):
            return ''
    except (TypeError, ValueError):
        pass
    s = str(cell) if isinstance(cell, str) else str(cell)
    if not s.strip():
        return ''
    if s.strip().lower() == 'nan':
        return ''
    return s


@XLBricksFunction(False)
def create_function_objects(funcs, persist=True, xlapp=None):
    """Parse and execute Python function definitions from Excel cells.
    
    Allows defining custom functions directly in Excel ranges.
    """

    funcs = XLUtils.get_bricks(funcs)
    funcs = np.array([[funcs.value]] if isinstance(funcs.value, str) else funcs.value)

    funcs_mask = list()
    for row in funcs:
        line = _func_line_sanitize(row[0])
        first_char = line.split(' ')[0] if line else ''
        funcs_mask.append(first_char == 'from' or first_char == 'import' or first_char == 'def')

    funcs_split_t = np.array_split(funcs[:, 0], np.argwhere(funcs_mask).flatten())
    funcs_split = ['\n'.join(_func_line_sanitize(line) for line in func) for func in funcs_split_t[1:]]

    # Ensure blocks ending with ':' (def/class/if/for etc.) have at least one indented line
    def _ensure_block_has_body(code):
        stripped = code.rstrip()
        if stripped.endswith(':'):
            return code + '\n    pass'
        return code

    funcs_split = [_ensure_block_has_body(block) for block in funcs_split]

    res = dict()
    [exec(func, res) for func in funcs_split]
    xlbricks = {k: XLBrick(None, v) for k, v in res.items()}
    xlbricks_front = create_bricks_front(XLBricks(None, xlbricks), xlapp, persist)
    return xlbricks_front


@XLBricksFunction(True)
def create_context_object(context_name, context_path, args=None, persist=True, xlapp=None):
    """Instantiate a Python class from a module path.
    
    Creates objects like QuantLib contexts with optional constructor arguments.
    """

    context = get_context_object(context_name, context_path)
    if args is None:
        context_instance = context()
    else:
        args = XLUtils.get_bricks(args)
        context_instance = context(**args.to_dict())

    xlbricks_front = create_bricks_front(XLBrick(None, context_instance), xlapp, persist)
    return xlbricks_front


@XLBricksFunction(True)
def run_function(function_objects, function_name, args=None, persist=True, xlapp=None):
    """Execute a stored function by name with optional arguments.
    
    Runs functions created with xlb_create_function or stored in bricks.
    """
    function_objects = XLUtils.get_bricks(function_objects)
    function_object = function_objects[[function_name]].value

    if function_object is None:
        func = getattr(function_object.value, function_name)
    else:
        func = function_objects[[function_name]].value

    if args is None:
        func_results = func()
    else:
        args = XLUtils.get_bricks(args)
        func_results = func(**args.to_dict())

    if isinstance(func_results, dict):

        xlbrick = XLBricksUtils.element_from_dictionary(func_results)
    elif isinstance(func_results, (list, tuple)):
        xlbrick = XLBricksUtils.element_from_list(func_results, '%s_res' % function_name)
    else:
        xlbrick = XLBrick(None, func_results)

    xlbricks_front = create_bricks_front(xlbrick, xlapp, persist)
    return xlbricks_front

@XLBricksFunction(True)
def run_quantlib_function(quantlib_objects, function_name, args=None, persist=True, xl_app=None):
    """Call a method on a QuantLib object with optional arguments.
    
    Enables financial calculations using QuantLib library methods.
    """

    quantlib_objects = XLUtils.get_bricks(quantlib_objects)
    func = getattr(quantlib_objects.value, function_name)

    if args is None:
        func_results = func()
    else:
        args = XLUtils.get_bricks(args)
        ##func_results = _run_quantlib_function(func, args.to_quantlib_dict())
        func_results = _run_quantlib_function(func, args.to_dict())

    if isinstance(func_results, dict):
        xlbrick = XLBricksUtils.element_from_dictionary(func_results)
    elif isinstance(func_results, (list, tuple)):
        xlbrick = XLBricksUtils.element_from_list(func_results, '%s_res' % function_name)
    else:
        xlbrick = XLBrick(None, func_results)

    xlbricks_front = create_bricks_front(xlbrick, xl_app, persist)
    return xlbricks_front


@XLBricksFunction(True)
def flatten_element(brick):
    xlbrick = XLUtils.get_bricks(brick)
    return xlbrick.value


@XLBricksFunction(False)
def merge_elements(brick_1=None, brick_2=None, brick_3=None, brick_4=None,
                   brick_5=None, persist=True, xlapp=None):
    """Combine multiple bricks into a single collection.
    
    Merges all key-value pairs from input bricks.
    """
    xlbricks = XLBricks()
    for idx in range(1, 6):
        brick = locals().get('brick_%s' % idx, None)
        if brick is not None:
            xlbricks.bricks.update(XLUtils.get_bricks(brick).bricks)
    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def replace_elements(bricks, key_1, brick_1, key_2=None, brick_2=None, key_3=None, brick_3=None,
                     key_4=None, brick_4=None, key_5=None, brick_5=None, persist=False, xlapp=None):
    """Update specific values in a brick collection.
    
    Replaces bricks at specified paths while keeping the rest unchanged.
    """

    xlbricks = deepcopy(XLUtils.get_bricks(bricks))
    xlbricks.alias = None
    xlbricks.persist = persist

    for idx in range(1, 6):
        keys = locals().get('key_%s' % idx, None)
        brick = locals().get('brick_%s' % idx, None)
        if keys is not None and brick is not None:
            keys = [key.strip(' \t\n\r') for key in keys.split('/')]
            xlbricks.replace(keys, XLUtils.get_bricks(bricks))

    xlbricks_front = create_bricks_front(xlbricks, xlapp, persist)
    return xlbricks_front


@XLBricksFunction(False)
def assign_alias(bricks, alias):
    """Give a custom name to a brick for easier referencing.
    
    Creates a memorable alias instead of using cell addresses.
    """
    xlbricks = XLUtils.get_bricks(bricks)
    xlbricks_front = XLBricksFront(alias, xlbricks, True)
    return xlbricks_front


@XLBricksFunction(True)
def delete_bricks(bricks):
    """Remove a brick from memory.
    
    Frees up the storage used by the specified brick.
    """
    XLUtils.delete_element(bricks)
    return 'DELETED FROM MEMORY'


def clear_bricks_front():
    """Clear all bricks from the front stack.
    
    Removes all stored bricks from memory.
    """
    XLBricksFrontStack().clear()


def create_bricks_front(xlbricks, xlapp, persist):
    """Wrap bricks in a front object with persistence tracking.
    
    Associates bricks with their Excel cell location if persist is True.
    """
    if persist:
        cell_address = XLUtils.active_cell_address(xlapp)
    else:
        cell_address = None
    xlbricks_front = XLBricksFront(cell_address, xlbricks, persist)
    return xlbricks_front


def get_context_object(class_name: str, class_path: str) -> object:
    """Import and return a class from a module path.
    
    Dynamically loads Python classes for instantiation.
    """
    module = importlib.import_module(class_path)
    class_object = getattr(module, class_name)
    return class_object

def _run_quantlib_function(func, args_dict):
    """Execute a QuantLib function with arguments as kwargs or positional args.
    
    Tries keyword arguments first, falls back to positional if that fails.
    """
    try:
        return func(**args_dict)
    except TypeError:
        return func(*[arg for arg in args_dict.values()])
