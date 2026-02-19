"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Pytest configuration; filters known third-party and environment warnings.
"""

import warnings

# Apply filters early so they apply to test collection and imports
warnings.filterwarnings(
    "ignore",
    category=DeprecationWarning,
    message=r".*imp module is deprecated.*",
)
warnings.filterwarnings(
    "ignore",
    category=DeprecationWarning,
    module="pywintypes",
)
warnings.filterwarnings(
    "ignore",
    category=ResourceWarning,
    message=r".*unclosed.*event loop.*",
)
warnings.filterwarnings(
    "ignore",
    category=ResourceWarning,
    message=r".*unclosed.*socket.*",
)
