"""
    Author: julij.jegorov
    Date: 15/02/2026
    Description: XLBricksFront: wrapper around xlbricks with alias, persist flag, and counter.
"""

import uuid


class XLBricksFront(object):
    """Wrapper for bricks with versioning and persistence tracking.
    
    Manages brick references, counters, and whether they persist across sessions.
    """

    def __init__(self, alias, xlbricks, persist=True):
        """Initialize a front wrapper for bricks.
        
        Alias can be a cell address or custom name. Persist controls memory retention.
        """
        self.counter = 0
        self.alias = alias
        self.xlbricks = xlbricks
        self.persist = persist
        self.uuid = str(uuid.uuid1())

    @property
    def bricks_name(self):
        """Get the identifier for this brick (alias if persistent, UUID otherwise)."""
        return self.alias if self.persist else self.uuid

    @property
    def bricks_full_name(self):
        """Get the full reference string including version counter.
        
        Format: 'name:counter' (e.g., 'myBrick:3').
        """
        return '%s:%s' % (self.bricks_name, self.counter)


