"""
advanced_excel package (standard Python package name).

This makes the library fully importable and compatible after `pip install -e .`
or `pip install advanced-excel`.

Recommended import:
    from advanced_excel import AdvancedExcel

The main class is AdvancedExcel (CamelCase is standard for classes).
Implementation is in core.py (the thin class + mixin imports) and the various
*Mixin modules.
"""

from .core import COL_INDEX, DATA, ROW_INDEX, AdvancedExcel, __version__

__all__ = ["AdvancedExcel", "ROW_INDEX", "COL_INDEX", "DATA", "__version__"]
