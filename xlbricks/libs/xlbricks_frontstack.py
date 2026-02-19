"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: Singleton stack of XLBricksFront instances; tracks bricks by alias/UUID.
"""

from xlbricks.libs.xlbricks_front import XLBricksFront

class Singleton(object):
    """Ensures only one instance of a class exists.
    
    Base class for singleton pattern implementation.
    """

    def __new__(cls):
        if not hasattr(cls, 'instance'):
            cls.instance = super(Singleton, cls).__new__(cls)
        return cls.instance


class XLBricksFrontStack(Singleton):
    """Global storage for all active bricks in memory.
    
    Singleton that maintains a dictionary of brick references accessible from Excel.
    """

    front_stack = dict()

    def __contains__(self, item):
        """Check if a brick reference exists in the stack."""
        return item in self.front_stack

    def __setitem__(self, key, value):
        """Store or update a brick in the stack."""
        self.front_stack[key] = value

    def __getitem__(self, item):
        """Retrieve a brick from the stack by its reference."""
        return self.front_stack.get(item, None)

    def __delitem__(self, item):
        """Remove a brick from the stack."""
        del self.front_stack[item]

    def clear(self):
        """Remove all bricks from the stack."""
        self.front_stack.clear()

    def to_dict(self):
        """Export all bricks as a nested dictionary.
        
        Useful for viewing or serializing the entire brick collection.
        """
        res_dict = dict()
        for key, bricks_front in self.front_stack.items():
            res_dict[key] = bricks_front.xlbricks.to_dict()
        return res_dict


def add_bricks_to_front_stack(bricks: XLBricksFront):
    """Add a brick to the global stack with automatic version incrementing.
    
    Updates the counter if a brick with the same name already exists.
    """
    container_name = bricks.bricks_name
    if container_name in XLBricksFrontStack():
        bricks.counter = XLBricksFrontStack()[container_name].counter + 1
    XLBricksFrontStack()[container_name] = bricks


def delete_bricks_from_front_stack(bricks: XLBricksFront):
    """Remove non-persistent bricks from the stack.
    
    Only deletes bricks that were created with persist=False.
    """
    if not bricks.persist and bricks.bricks_name in XLBricksFrontStack():
        del XLBricksFrontStack()[bricks.bricks_name]
