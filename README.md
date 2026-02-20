# xlbricks

Excel-integrated brick structures and User Defined Functions (UDFs) for Python.

## Overview

xlbricks is a Python package that provides a powerful framework for working with Excel through xlwings, offering:

- **Brick Data Structures**: In-memory representation of Excel ranges as hierarchical key-value structures
- **Excel UDFs**: Custom functions callable directly from Excel spreadsheets
- **QuantLib Integration**: Financial calculations and analytics with QuantLib support
- **PyQt5 UI**: Interactive explorers and editors for managing brick structures
- **Validation Framework**: Robust input validation for Excel data

## Features

- **XLBrick & XLBricks**: Core data structures for organizing Excel data hierarchically
- **Excel Integration**: Seamless bidirectional communication with Excel via xlwings
- **Front Stack Management**: Manage multiple brick collections with undo/redo capabilities
- **Configuration Editor**: GUI for managing xlbricks settings
- **Data Explorer**: Visual interface for inspecting and navigating brick structures
- **Utility Functions**: Helper functions for common Excel operations

## Installation

Install xlbricks using pip:

```bash
pip install xlbricks
```

### Requirements

- Python 3.11 or higher
- Microsoft Excel (for xlwings integration)
- Windows operating system (required for xlwings Excel integration)

### Dependencies

xlbricks automatically installs the following dependencies:

- `numpy` - Numerical computing
- `pandas` - Data manipulation and analysis
- `xlwings` - Excel integration
- `PyQt5` - GUI components
- `QuantLib-Python` - Financial calculations

**Note**: QuantLib-Python can be challenging to install on some systems. If you encounter issues, please refer to the [QuantLib installation guide](https://www.quantlib.org/install.shtml).

## Quick Start

### Using xlbricks in Excel

1. Create an Excel workbook and set up xlwings:

```python
import xlwings as xw
from xlbricks import xlbfunctions
```

2. Use xlbricks UDFs in your Excel formulas:

```excel
=xlb_brick("mydata", A1:B10)
=xlb_get("mydata", "key1")
```

### Using xlbricks in Python

```python
from xlbricks.libs.xlbricks import XLBrick, XLBricks
import numpy as np

# Create a brick from data
data = np.array([[1, 2], [3, 4]])
brick = XLBrick(key="mydata", data=data)

# Create a collection of bricks
bricks = XLBricks(key="root")
bricks.bricks["mydata"] = brick

# Access brick data
print(brick.to_dict())
```

### Configuration

After installation, configure xlbricks by editing the `xlbricks.json` file:

```json
{
  "APPS_PATH": "C:\\path\\to\\your\\xlbricks\\applications",
  "INTERPRETER": "C:\\path\\to\\your\\pythonw.exe",
  "PYTHONPATH": "C:\\path\\to\\your\\python\\models",
  "CONTEXT": {
    "PythonFunctions": "myapp.python_function.context"
  }
}
```

You can also use the built-in configuration editor:

```python
from xlbricks.ui.config_editor import show_config_editor
show_config_editor()
```

## Excel UDF Functions

xlbricks provides several User Defined Functions for Excel:

- `xlb_brick(key, data, persist=True)` - Store Excel range as a brick
- `xlb_get(key, *path)` - Retrieve data from a brick
- `xlb_delete(key)` - Delete a brick
- `xlb_explorer()` - Open the brick explorer GUI
- `xlb_config()` - Open the configuration editor

## Development

### Running Tests

```bash
pytest xlbricks/tests/
```

### Installing from Source

```bash
git clone <repository-url>
cd xlbricks
pip install -e .
```

### Development Dependencies

```bash
pip install -e .[dev]
```

## Project Structure

```
xlbricks/
├── libs/           # Core functionality
│   ├── xlbricks.py           # Brick data structures
│   ├── xlfunctions.py        # Function implementations
│   ├── validation.py         # Input validation
│   └── utility_functions.py  # Helper utilities
├── ui/             # PyQt5 GUI components
│   ├── explorer.py           # Brick explorer
│   ├── config_editor.py      # Configuration editor
│   └── tree_model.py         # Tree view models
├── tests/          # Unit tests
└── xlbfunctions.py # Excel UDF entry points
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

julij.jegorov

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## Support

For issues, questions, or contributions, please visit the project repository.
