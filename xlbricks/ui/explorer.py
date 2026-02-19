"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: PyQt Explorer window to browse the XLBricks front stack
                 (tree view and table); opened from Excel.
"""

import numpy as np
import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtWidgets import QWidget, QLineEdit, QTreeView, QTableView, QPushButton, QSplitter, QVBoxLayout, QMessageBox
from xlbricks.libs.xlbricks_frontstack import XLBricksFrontStack

from xlbricks.ui.pandas_model import PandasModel
from xlbricks.ui.tree_model import DictionaryTreeModel, node_structure_from_dict


class Singleton(object):
    """Ensures only one instance of a class exists.
    
    Base class for singleton pattern implementation.
    """
    def __new__(cls):
        if not hasattr(cls, 'instance'):
            cls.instance = super(Singleton, cls).__new__(cls)
        return cls.instance


class ExplorerTreeView(QTreeView):
    """Tree view widget for displaying brick hierarchies.
    
    Supports keyboard navigation and F5 refresh.
    """

    keyPressedNavigation = QtCore.pyqtSignal()
    keyPressedRefresh = QtCore.pyqtSignal()

    def __init__(self, model):
        super(QTreeView, self).__init__()
        self.setHeaderHidden(True)
        self.setModel(model)

    def keyPressEvent(self, event):
            super().keyPressEvent(event)
            if event.key() == QtCore.Qt.Key_Return \
                                                or event.key() == QtCore.Qt.Key_Down \
                                    or event.key() == QtCore.Qt.Key_Up \
                                    or event.key() == QtCore.Qt.Key_Left \
                                    or event.key() == QtCore.Qt.Key_Right:
                self.keyPressedNavigation.emit()

            elif event.key() == QtCore.Qt.Key_F5:
                self.keyPressedRefresh.emit()

            else:
                super().keyPressEvent(event)

    def refresh(self):
        """Reload the tree view with current brick data from memory."""
        model = DictionaryTreeModel(node_structure_from_dict(XLBricksFrontStack().to_dict()))
        self.setModel(model)


class ExplorerTableView(QTableView):
    """Table view widget for displaying brick data values.
    
    Shows the contents of selected bricks in tabular format.
    """
    
    def __init__(self):
        super(QTableView, self).__init__()

    def refresh(self, data=None):
        """Update the table view with new data.
        
        Accepts DataFrames, arrays, or scalar values.
        """
        if data is None:
            df = pd.DataFrame()
        elif isinstance(data, pd.DataFrame):
            df = data
        elif isinstance(data, np.ndarray):
            df = pd.DataFrame(data)
        else:
            df = pd.DataFrame(np.array([data]))
        self.setModel(PandasModel(df))


class Explorer(QWidget):
    """Main explorer window for browsing bricks.
    
    Split view with tree navigation on left and data table on right.
    """

    def __init__(self, model):
        """Initialize the explorer with a tree model.
        
        Sets up the UI with tree and table views.
        """
        super(Explorer, self).__init__()
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        self._entry = QLineEdit()
        self._tree_view = ExplorerTreeView(model)
        self._table_view = ExplorerTableView()
        self._button = QPushButton()

    def refresh(self):
        """Reload both tree and table views with current data."""
        self._tree_view.refresh()
        self._table_view.refresh()

    def load_data_frame(self):
        """Load the selected brick's data into the table view.
        
        Triggered when user clicks or navigates in the tree.
        """
        try:
            index = self._tree_view.currentIndex()
            if not index.isValid():
                return
            node = index.model().get_node(index)
            if not hasattr(node, 'value') or node.value is None:
                self._table_view.refresh(None)
                return
            data_frame = node.value
            # Make a defensive copy to isolate from Excel COM thread
            if isinstance(data_frame, pd.DataFrame):
                data_frame = data_frame.copy(deep=True)
            elif isinstance(data_frame, np.ndarray):
                data_frame = np.copy(data_frame)
            self._table_view.refresh(data_frame)
        except Exception as e:
            print(f"Error loading data frame: {e}")
            import traceback
            traceback.print_exc()
            self._table_view.refresh(None)

    def display(self):
        """Show the explorer window for viewing all bricks.
        
        Displays the full brick collection in a split-pane interface.
        """
        self.setWindowTitle('Object Viewer')
        self.setMinimumSize(600, 400)
        self._tree_view.clicked.connect(self.load_data_frame)
        self._tree_view.keyPressedNavigation.connect(self.load_data_frame)
        vertical_box = QVBoxLayout()
        splitter = QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(self._tree_view)
        splitter.addWidget(self._table_view)
        splitter.setSizes([100, 200])
        vertical_box.addWidget(splitter)
        self.setLayout(vertical_box)
        self.show()

    def display_one_element(self):
        """Show the explorer window for a single brick.
        
        Displays one brick's structure and contents in detail.
        """
        self.setWindowTitle('Object Viewer')
        self.setMinimumSize(200, 300)
        self._tree_view.clicked.connect(self.load_data_frame)
        self._tree_view.keyPressedNavigation.connect(self.load_data_frame)
        vertical_box = QVBoxLayout()
        splitter = QSplitter(QtCore.Qt.Horizontal)
        splitter.addWidget(self._tree_view)
        splitter.addWidget(self._table_view)
        splitter.setSizes([100, 200])
        vertical_box.addWidget(splitter)
        self.setLayout(vertical_box)
        self.show()

