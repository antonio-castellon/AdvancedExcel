"""
AdvancedExcel - refactored modular version.

The public API (class name, all method names and signatures, constants,
sspipe-friendly usage, etc.) is 100% preserved.

Implementation has been split into focused mixins for rows, columns,
table/block/entity extraction, lookup, I/O, cleaning, and reshape.
All cross-calls (self.other_method) resolve via the final MRO.
"""

# Import all the behavior mixins
from .blocks import BlockTableMixin
from .cleaning import CleaningMixin
from .columns import ColumnMixin
from .io import IOMixin
from .lookup import LookupMixin
from .reshape import ReshapeMixin
from .rows import RowMixin


class AdvancedExcel(
    IOMixin,
    LookupMixin,
    BlockTableMixin,
    RowMixin,
    ColumnMixin,
    CleaningMixin,
    ReshapeMixin,
):
    """
    A Python wrapper for Excel files, providing easy access to their contents.

    This class simplifies working with Excel files by providing a collection of
    methods for common tasks such as reading data, manipulating DataFrames, and
    formatting data. It aims to make Excel data accessible and usable within
    Python.

    The implementation is split across focused mixins (io, lookup, blocks,
    rows, columns, cleaning, reshape) while preserving the exact public API
    and documented usage patterns (including sspipe chaining via p(instance.method)).

    **Constants (still available as AdvancedExcel.ROW_INDEX etc. and on instances):**
    * `ROW_INDEX`:  A string constant used as a key for row indices.
    * `COL_INDEX`:  A string constant used as a key for column indices.
    * `DATA`:  A string constant used as a key for data values.

    **Attributes:**
    * `name`: The name of the Excel file (can be set).
    * `sheets`: A dictionary to store DataFrames representing sheets (keys are sheet names).
    """

    # Public constants (preserved for backward compatibility with user code
    # that does row[AdvancedExcel.DATA] etc. after calling getAllRowsFromKey)
    ROW_INDEX = "row_index"
    COL_INDEX = "col_index"
    DATA = "data"

    __version__ = "1.1.0"


# Module-level aliases for the constants (so that both
#   AdvancedExcel.ROW_INDEX
# and
#   from .AdvancedExcel import ROW_INDEX
# continue to work, preserving any existing user code / docstring examples).
ROW_INDEX = AdvancedExcel.ROW_INDEX
COL_INDEX = AdvancedExcel.COL_INDEX
DATA = AdvancedExcel.DATA

__version__ = AdvancedExcel.__version__


# The heavy lifting (including the openpyxl hyperlink monkey-patch) now lives in
# the individual mixin modules. Importing this module pulls in all behavior.
#
# Example usage (unchanged from before):
#
#   from AdvancedExcel import AdvancedExcel
#   from sspipe import p
#
#   data = AdvancedExcel("studies_excel.xlsx")
#   results = data.getSheet("Raw data") | p(data.getDataFrameBlocksByKeyName, "Batch Number")
#   ...
#
# All methods that used to be top-level functions taking (self, df, ...) are now
# real methods on the class thanks to the mixin composition.
