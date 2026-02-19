"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: XLBricks package: Excel-integrated brick structures and UDFs.
"""

__version__ = "0.1.0"
__author__ = "julij.jegorov"
__license__ = "MIT"

# Core brick data structures
from xlbricks.libs.xlbricks import (
    XLBrick,
    XLBricks,
    XLBrickAbstract,
)

# Front and stack management
from xlbricks.libs.xlbricks_front import XLBricksFront
from xlbricks.libs.xlbricks_frontstack import XLBricksFrontStack

# Utility classes
from xlbricks.libs.utility_functions import (
    XLUtils,
    XLBricksUtils,
    XLBricksFunction,
)

# Validation helpers
from xlbricks.libs.validation import (
    _is_missing,
    _check_required,
    _check_array_2d,
    _ERROR_PREFIX,
)

# Excel UDF functions (main entry point for xlwings)
from xlbricks import xlbfunctions

__all__ = [
    # Version info
    "__version__",
    "__author__",
    "__license__",
    
    # Core classes
    "XLBrick",
    "XLBricks",
    "XLBrickAbstract",
    "XLBricksFront",
    "XLBricksFrontStack",
    
    # Utilities
    "XLUtils",
    "XLBricksUtils",
    "XLBricksFunction",
    
    # Validation
    "_is_missing",
    "_check_required",
    "_check_array_2d",
    "_ERROR_PREFIX",
    
    # UDF module
    "xlbfunctions",
]