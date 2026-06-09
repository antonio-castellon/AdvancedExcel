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

# Re-export for convenience and to satisfy linters
__all__ = ["AdvancedExcel"]

# Constants and version are available as AdvancedExcel.XXX
# (class attributes) and also as module-level names.
from .AdvancedExcel import ROW_INDEX, COL_INDEX, DATA, __version__  # noqa: F401


