import ntpath
import re

import numpy as np
import pandas as pd
import xlrd
from openpyxl import load_workbook
from pandas.io.excel._openpyxl import OpenpyxlReader
from pyxlsb import open_workbook


class IOMixin:
    """
    Mixin providing Excel file loading, sheet access, and related I/O helpers.

    Handles multi-engine reading (openpyxl / xlrd / pyxlsb), populates self.sheets,
    and provides the basic getSheet / getSheetNames API plus private path and
    sheetname helpers. Also applies the openpyxl hyperlink monkey-patch.
    """

    def __init__(self, file, filetype=None):
        """
        Initializes an ExcelFile object by loading the specified file and creating
        an internal distribution of all sheets as pandas DataFrames.

        This constructor determines the appropriate engine (openpyxl, xlrd, or pyxlsb)
        based on the file extension (.xlsx, .xlsm, .xls, or other assumed to be .xlsb).
        It then iterates through each sheet in the file, reading the data into a
        pandas DataFrame and storing it in the `self.sheets` dictionary, where
        the keys are the sheet names and the values are the corresponding DataFrames.

        Args:
            file (str): The path to the Excel file.
            filetype (dtype, optional): Data type for columns. Defaults to None.  If given, it is passed to pandas `read_excel` function.

        Raises:
            FileNotFoundError: If the specified file does not exist.  (Implicitly raised by underlying libraries)
            Exception: If there is an error during file processing or sheet reading. (Implicitly raised by underlying libraries)

        """
        self.name = self._path_leaf(file)
        _engine = None
        self.sheets = {}

        if re.search(r"(?i)\.xlsx|\.xlsm", self.name):
            AllSheets = load_workbook(filename=file).sheetnames
            _engine = "openpyxl"
        elif re.search(r"(?i)\.xls$", self.name):
            AllSheets = xlrd.open_workbook(filename=file).sheet_names()
            _engine = "xlrd"
        else:
            AllSheets = self._get_sheetnames(file)
            _engine = "pyxlsb"

        for x in AllSheets:
            # print(x)
            dfraw = pd.read_excel(
                file,
                engine=_engine,
                sheet_name=x,
                dtype=filetype,
                skiprows=0,
                engine_kwargs={"keep_vba": False},
            )
            self.sheets[x] = dfraw

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        # Currently a no-op. Future versions may release workbook resources here.
        self.sheets.clear()
        return False

    def get_sheet(self, sheetname):
        """
        Retrieves a specific sheet from the loaded Excel file as a pandas DataFrame.

        Args:
            sheetname (str): The name of the sheet to retrieve.

        Returns:
            pandas.DataFrame: The DataFrame representing the requested sheet, or None if the sheet does not exist.
        """
        if sheetname in self.sheets.keys():
            return self.sheets[sheetname]
        else:
            return None

    def get_sheet_names(self):
        """
        Returns a list of all sheet names present in the loaded Excel file.

        Returns:
            list: A list of strings, where each string is the name of a sheet in the Excel file.
        """
        return self.sheets.keys()

    def _path_leaf(self, path):
        """
        Extracts the filename from a path.

        This method extracts the filename (leaf) from a given path using `ntpath.split`.

        Args:
            path (str): The file path.

        Returns:
            str: The filename (leaf) of the path.
        """
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)

    def _get_sheetnames(self, filepath):
        """
        Gets sheet names from an Excel file without fully opening it.

        This method uses the `open_workbook` function (presumably from a library like
        `openpyxl` or `xlrd`) to efficiently retrieve the sheet names from an Excel
        file *without* fully loading the file's contents.

        Args:
            filepath (str): The path to the Excel file.

        Returns:
            list: A list of sheet names.
        """
        with open_workbook(filepath) as wb:
            return [sheet.name for sheet in wb.sheets]  # Return a list of sheet names.


# --- openpyxl hyperlink monkey-patch (applied on import of this module) ---


def _convert_cell(self, cell):
    """
    Converts a cell value from openpyxl to a standard Python type.

    This is a monkey-patched version of the `_convert_cell` method from `openpyxl`.
    It adds support for handling hyperlinks by returning the hyperlink target.

    Args:
        cell (openpyxl.cell.cell.Cell): The cell to convert.

    Returns:
        Scalar: The converted cell value (can be a hyperlink target, a standard
        Python type, or np.nan for errors).
    """
    from openpyxl.cell.cell import TYPE_BOOL, TYPE_ERROR, TYPE_NUMERIC

    if hasattr(cell, "hyperlink") and hasattr(cell.hyperlink, "target"):
        return cell.hyperlink.target  # Return hyperlink target
    elif cell.data_type == TYPE_ERROR:
        return np.nan
    elif cell.data_type == TYPE_BOOL:
        return bool(cell.value)
    elif cell.value is None:
        return ""  # For compatibility with xlrd
    elif cell.data_type == TYPE_NUMERIC:
        val = int(cell.value)
        if val == cell.value:
            return val
        return float(cell.value)
    return cell.value


OpenpyxlReader._convert_cell = _convert_cell  # Apply the monkey patch when io is imported
