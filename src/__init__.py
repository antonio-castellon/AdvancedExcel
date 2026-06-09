"""
Legacy compatibility for direct 'src' on PYTHONPATH (rare after pip install -e .).

Recommended after `pip install -e .` or `pip install advanced-excel`:

    from advanced_excel import AdvancedExcel
"""

from advanced_excel import AdvancedExcel
from advanced_excel import ROW_INDEX, COL_INDEX, DATA, __version__

__all__ = ["AdvancedExcel", "ROW_INDEX", "COL_INDEX", "DATA", "__version__"]




