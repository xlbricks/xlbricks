"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Tree node for key-value hierarchies; used by DictionaryTreeModel.
"""


class Node(object):
    """A tree node for hierarchical data structures.
    
    Represents a single item in a tree with name, value, parent, and children.
    """

    def __init__(self, name, parent=None):
        """Create a new tree node.
        
        Automatically adds itself as a child of the parent if provided.
        """
        self._name = name
        self._parent = parent
        self._children = []
        self._value = None
        if parent is not None:
            parent.add_child(self)

    def add_child(self, child):
        """Add a child node to this node."""
        self._children.append(child)

    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        self._value = value

    def child(self, row):
        """Get the child node at the specified index."""
        return self._children[row]

    def child_count(self):
        """Return the number of children this node has."""
        return len(self._children)

    def parent(self):
        """Get the parent node of this node."""
        return self._parent

    def row(self):
        """Get this node's position in its parent's children list."""
        if self._parent is not None:
            return self._parent._children.index(self)

    def data(self, column):
        """Get data for display in a specific column.
        
        Column 0 returns name, column 1 returns value.
        """
        if column is 0:
            return self.name
        elif column is 1:
            return self.value

    def set_data(self, column, value):
        """Set data for a specific column.
        
        Column 0 sets name, column 1 sets value.
        """
        if column is 0:
            self.name = value
        if column is 1:
            self.value = value

