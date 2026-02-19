"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Qt table model for pandas DataFrames; used by Explorer detail view.
"""

from PyQt5 import QtCore
import pandas as pd
import numpy as np


class PandasModel(QtCore.QAbstractTableModel):
    """Qt model adapter for displaying pandas DataFrames in table views.
    
    Handles data formatting, headers, and display properties.
    """

    def __init__(self, df=pd.DataFrame(), parent=None):
        """Initialize the model with a DataFrame.
        
        Defaults to an empty DataFrame if none provided.
        """
        QtCore.QAbstractTableModel.__init__(self, parent=parent)
        self._df = df

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        """Provide header labels for rows and columns.
        
        Returns DataFrame column names and index values.
        """
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                try:
                    return self._df.columns.tolist()[section]
                except (IndexError, ):
                    return QtCore.QVariant()
            elif orientation == QtCore.Qt.Vertical:
                try:
                    return str(self._df.index.tolist()[section])
                except (IndexError, ):
                    return QtCore.QVariant()
        elif role == QtCore.Qt.TextAlignmentRole:
            return QtCore.QVariant(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)
        elif role == QtCore.Qt.SizeHintRole:
            pass
        return QtCore.QVariant()

    def data(self, index, role=QtCore.Qt.DisplayRole):
        """Provide cell data for display.
        
        Formats numeric values to 5 decimal places for readability.
        """
        if role == QtCore.Qt.DisplayRole:
            cell_value = self._df.iloc[index.row(), index.column()]
            # Check for numeric types (numpy and python floats)
            if isinstance(cell_value, (float, np.floating)):
                cell_value = np.around(cell_value, 5)
            return QtCore.QVariant(str(cell_value))

        elif role == QtCore.Qt.TextAlignmentRole:
            return QtCore.QVariant(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)

        return QtCore.QVariant()

    def rowCount(self, parent=QtCore.QModelIndex()):
        """Return the number of rows in the DataFrame."""
        return len(self._df.index)

    def columnCount(self, parent=QtCore.QModelIndex()):
        """Return the number of columns in the DataFrame."""
        return len(self._df.columns)

