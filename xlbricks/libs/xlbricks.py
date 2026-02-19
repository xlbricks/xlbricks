"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: XLBrick and XLBricks data structures; in-memory representation
                 of Excel ranges and nested key-value bricks (incl. QuantLib helpers).
"""

import abc
import numpy as np
import QuantLib as ql
from datetime import datetime
from collections import OrderedDict


class XLBrickAbstract(metaclass=abc.ABCMeta):
    """Base class for brick data structures.
    
    Defines the interface for accessing and serializing brick data.
    """

    @abc.abstractmethod
    def __getitem__(self):
        pass

    @abc.abstractmethod
    def to_dict(self):
        pass


class XLBricks(XLBrickAbstract):
    """A collection of nested bricks organized as key-value pairs.
    
    Supports hierarchical structures where bricks can contain other bricks.
    """
    
    def __init__(self, key=None, bricks=None):
        self.key = key
        self.bricks = bricks or OrderedDict()

    def replace(self, keys, brick):
        """Replace a brick at a specific path in the hierarchy.
        
        Keys is a list representing the path to the brick to replace.
        """
        xl_brick = self.__getitem__(keys)
        if xl_brick is None:
            raise ValueError('brick not found at %s' % keys)

        if isinstance(brick, XLBricks):
            xl_brick = self.__getitem__(keys[:-1])
            xl_brick[keys[-1]] = brick
        else:
            xl_brick.value = brick.value

    def __setitem__(self, key, brick):
        self.bricks[key] = brick

    def __getitem__(self, keys):
        if len(keys) == 0:
            return self

        node = self
        for key in keys:
            node = node.bricks[key]
        return node

    def to_dict(self):
        """Convert the brick collection to a nested dictionary.
        
        Recursively serializes all child bricks.
        """
        child_dict = OrderedDict()
        for key, qd_item in self.bricks.items():
            child_dict[key] = qd_item.to_dict()

        if self.key is None:
            return child_dict
        else:
            return OrderedDict([(self.key, child_dict)])

    def to_quantlib_dict(self):
        """Convert to dictionary with QuantLib-compatible types.
        
        Transforms dates, strings, and numeric values into QuantLib objects.
        """
        f = np.vectorize(_cast_quantlib_variable)
        child_dict = OrderedDict()
        for key, qd_item in self.bricks.items():
            child_dict[f(key).item()] = qd_item.to_quantlib_dict()

        if self.key is None:
            return child_dict
        else:
            return OrderedDict([(f(self.key).item(), child_dict)])


class XLBrick(XLBrickAbstract):
    """A single brick containing a value (array, scalar, or object).
    
    The fundamental data unit in XLBricks, wrapping any type of data.
    """

    def __init__(self, key=None, value=None):
        self.key = key
        self.value = value

    def __getitem__(self, keys):
        pass

    def to_dict(self):
        """Convert the brick to a dictionary or return its raw value.
        
        Returns the value directly if no key is set, otherwise returns {key: value}.
        """
        if self.key is None:
            return self.value
        return OrderedDict([(self.key, self.value)])

    def to_quantlib_dict(self):
        """Convert brick value to QuantLib-compatible format.
        
        Handles date conversion, type casting, and QuantLib object preservation.
        """
        if 'QuantLib' in str(type(self.value)):
            return self.value

        f = np.vectorize(_cast_quantlib_variable)
        if isinstance(self.value, np.ndarray):
            xlb_array = f(self.value)
        else:
            xlb_array = np.array(f(self.value))

        if xlb_array.shape == (1, 1):
            if isinstance(xlb_array[0, 0], (np.float, np.int32, np.str, np.bool_)):
                xlb_value = xlb_array[0, 0].item()
            else:
                xlb_value = xlb_array[0, 0]
        else:
            xlb_value = xlb_array.tolist()

        if self.key is None:
            return xlb_value

        return OrderedDict([(f(self.key).item(), xlb_value)])


def _cast_quantlib_variable(x):
    """Convert Python values to QuantLib types.
    
    Handles datetime to ql.Date conversion, string evaluation, and type coercion.
    """
    if isinstance(x, datetime):
        return ql.Date(x.day, x.month, x.year)
    elif isinstance(x, float) and x.is_integer():
        return int(x)
    elif isinstance(x, str) and x[:3] == 'ql.':
        return eval(x)
    elif isinstance(x, str) and x.lower() in ['true', 'false']:
        return bool(x)
    else:
        return x