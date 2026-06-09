"""
AdvancedExcel package.

This __init__ makes the main class importable as:

    from AdvancedExcel import AdvancedExcel
    # or
    import AdvancedExcel
    ae = AdvancedExcel.AdvancedExcel(...)

The old style `from . import AdvancedExcel` (submodule) is still supported
for any code that relied on it.
"""

from .AdvancedExcel import AdvancedExcel

# Constants are available as AdvancedExcel.ROW_INDEX etc. (class attributes)
# and also as module-level names via the aliases in AdvancedExcel.py.
