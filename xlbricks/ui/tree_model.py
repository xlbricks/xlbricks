"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Qt tree model for an arbitrary dictionary; used by Explorer.
"""

import pandas as pd
from xlbricks.ui.node import Node
from PyQt5 import QtCore

class DictionaryTreeModel(QtCore.QAbstractItemModel):
    """Qt model for displaying nested dictionaries as a tree structure.
    
    Converts hierarchical data into a format suitable for tree views.
    """

    def __init__(self, root, parent=None):
        """Initialize the tree model with a root node.
        
        The root node should be created using node_structure_from_dict.
        """
        super(DictionaryTreeModel, self).__init__(parent)
        self._rootNode = root

    def rowCount(self, parent):
        """Return the number of children for a given parent node."""
        if not parent.isValid():
            parent_node = self._rootNode
        else:
            parent_node = parent.internalPointer()

        return parent_node.child_count()

    def columnCount(self, parent):
        """Return the number of columns (always 1 for tree display)."""
        return 1

    def data(self, index, role):
        """Provide data for display in the tree view.
        
        Returns node names for rendering.
        """
        if not index.isValid():
            return None

        node = index.internalPointer()
        if role == QtCore.Qt.DisplayRole:
            return node.data(index.column())

    def parent(self, index):
        """Get the parent index for a given node index."""
        node = self.get_node(index)
        parent_node = node.parent()
        if parent_node == self._rootNode:
            return QtCore.QModelIndex()

        return self.createIndex(parent_node.row(), 0, parent_node)

    def index(self, row, column, parent):
        """Create an index for accessing a specific node in the tree."""
        parent_node = self.get_node(parent)
        child_item = parent_node.child(row)

        if child_item:
            return self.createIndex(row, column, child_item)
        else:
            return QtCore.QModelIndex()

    def get_node(self, index):
        """Retrieve the Node object from a model index.
        
        Returns the root node if index is invalid.
        """
        if index.isValid():
            node = index.internalPointer()
            if node:
                return node
        return self._rootNode


def node_structure_from_dict(datadict, parent=None, root_node=None):
    """Convert a nested dictionary into a tree of Node objects.
    
    Recursively builds the node hierarchy for tree model display.
    """
    if not parent:
        root_node = Node('Root')
        parent = root_node

    for name, data in datadict.items():
        node = Node(name, parent)
        node.name = str(name)
        if isinstance(data, dict):
            node_structure_from_dict(data, node, root_node)
        elif isinstance(data, pd.DataFrame):
            node.value = data.head(min(25, len(data)))
        else:
            node.value = data

    return root_node

