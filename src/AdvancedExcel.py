import sys
import os
import ntpath
import xlrd
from pandas._typing import Scalar

from pandas.io.excel._openpyxl import OpenpyxlReader
from openpyxl import load_workbook
from pyxlsb import open_workbook
import pandas as pd
import numpy as np
import re

sys.path.append('../')
sys.path.append(os.path.dirname(os.path.realpath(__file__)))


class AdvancedExcel(object):
    """
A Python wrapper for Excel files, providing easy access to their contents.

This class simplifies working with Excel files by providing a collection of
methods for common tasks such as reading data, manipulating DataFrames, and
formatting data. It aims to make Excel data accessible and usable within
Python.

**Key Features:**

* **Data Extraction:**  Methods for reading data from sheets, including
  handling merged cells and various data types.
* **DataFrame Manipulation:** A wide range of methods for cleaning, transforming,
  and manipulating pandas DataFrames, including operations like removing rows/columns,
  filling missing values, merging/splitting columns, and changing data types.
* **Header and Index Management:** Functions for setting rows as headers, adding
  headers as rows, and resetting/dropping indices.
* **String and Value Manipulation:** Methods for stripping whitespace, replacing
  values, and changing the case of strings.
* **Convenience Functions:** Helper functions for tasks like transposing DataFrames,
  merging multiple DataFrames, and extracting filenames.

**Usage Tip:**  This class is designed to be used effectively with the `sspipe`
library (https://github.com/sspipe/sspipe).  `sspipe` allows for chaining
method calls in a readable and functional style, making complex data
transformations easier to manage.  For example:

```python
from sspipe import Pipe, p
from your_module import AdvancedExcel  # Replace your_module

excel_file = AdvancedExcel()
df = (
    excel_file.read_excel("my_excel_file.xlsx", "Sheet1")
    | p(excel_file.removeEmptyRows)
    | p(excel_file.fillForwardRowValues, columnNames=["ColumnA", "ColumnB"])
    | p(excel_file.setColumn2Numeric, columnName="ColumnC")
    | p(excel_file.stripAll)
    | p(excel_file.removeDuplicates)
)

# Now 'df' contains the transformed DataFrame.
```

**Constants:**

* `ROW_INDEX`:  A string constant used as a key for row indices.
* `COL_INDEX`:  A string constant used as a key for column indices.
* `DATA`:  A string constant used as a key for data values.

**Attributes:**

* `name`: The name of the Excel file (can be set).
* `sheets`: A dictionary to store DataFrames representing sheets (keys are sheet names).

"""

    ROW_INDEX = "row_index"
    COL_INDEX = "col_index"
    DATA = "data"

    name = None
    sheets = {}
    """ it contains all sheets as dataframes """

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

        if re.search(r'(?i)\.xlsx|\.xlsm', self.name):
            AllSheets = load_workbook(filename=file).sheetnames
            _engine = 'openpyxl'
        elif re.search(r'(?i)\.xls$', self.name):
            AllSheets = xlrd.open_workbook(filename=file).sheet_names()
            _engine = 'xlrd'
        else:
            AllSheets = self._get_sheetnames(file)
            _engine = 'pyxlsb'

        for x in AllSheets:
            # print(x)
            dfraw = pd.read_excel(file,
                                  engine=_engine,
                                  sheet_name=x,
                                  dtype=filetype,
                                  skiprows=0,
                                  engine_kwargs={"keep_vba": False})
            self.sheets[x] = dfraw

    def getSheet(self, sheetname):
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

    def getSheetNames(self):
        """
        Returns a list of all sheet names present in the loaded Excel file.

        Returns:
            list: A list of strings, where each string is the name of a sheet in the Excel file.
        """
        return self.sheets.keys()

    def getNextValue(self, _sheet, keyname, limitoffset=0):
        """
        Finds all occurrences of a keyname within a sheet and retrieves the values
        immediately to the right of each occurrence, within a specified offset limit.

        This method searches for the `keyname` within the provided sheet (`_sheet`).
        Because the data might have typos or inconsistent capitalization, it uses a
        non-optimized approach to search for the key (avoiding converting the entire
        DataFrame to lowercase).  It retrieves all rows containing the key using
        `getAllRowsFromKey`, then iterates through these rows, calling the internal
        `_getNextValue` method to get the value to the right of the key, respecting
        the `limitoffset`.

        Args:
            _sheet (pandas.DataFrame): The sheet to search within.
            keyname (str): The keyname to search for.
            limitoffset (int, optional): The maximum offset to the right to consider
                for retrieving the value. Defaults to 0.

        Returns:
            list: A list of values found to the right of each occurrence of the
                keyname, within the offset limit.  Returns an empty list if the keyname
                is not found, or if no values are found within the offset limit.
        """
        #
        # row = _sheet.loc[_sheet.isin([keyname]).any(axis=1)]
        # asList = row.index.tolist()
        # the problem of this approach is that the original/source data from the dataframe could have
        # different typos or to be case-sensitive , that's why we use a non optimized way to search inside
        # (avoiding generating another dataframe with all data in lowercase
        #
        data_rows = self.getAllRowsFromKey(_sheet, keyname)
        allResults = []
        for row in data_rows:
            allResults.append(
                self._getNextValue(row[AdvancedExcel.DATA], row[AdvancedExcel.COL_INDEX] + 1, limitoffset))

        return allResults

    def getNextValueFromKeyList(self, _sheet, keynames=[], limitoffset=0):
        """
         Searches a sheet for the first occurrence of any keyname from a given list
         and returns the value immediately to its right, within a specified offset limit.

         This method iterates through the provided list of `keynames`.  For each `keyname`,
         it searches the `_sheet` using `getAllRowsFromKey`. If a `keyname` is found
         (i.e., `getAllRowsFromKey` returns a non-empty list), it retrieves the first
         row containing that key and extracts the value to the right of the key using
         `_getNextValue`, respecting the `limitoffset`.  The search stops after the
         first keyname is found.

         Args:
             _sheet (pandas.DataFrame): The sheet to search within.
             keynames (list, optional): A list of keynames to search for. Defaults to None.
                                        If None, an empty list is effectively used, and an empty list will be returned.
             limitoffset (int, optional): The maximum offset to the right to consider
                 for retrieving the value. Defaults to 0.

         Returns:
             list: A list containing the single value found to the right of the first
                 occurring keyname, within the offset limit.  Returns an empty list
                 if none of the keynames are found, or if no value is found within
                 the offset limit for the first found key.
         """
        allResults = []
        for keyname in keynames:
            data_rows = self.getAllRowsFromKey(_sheet, keyname)
            if len(data_rows) > 0:
                row = data_rows[0]
                allResults.append(
                    self._getNextValue(row[AdvancedExcel.DATA], row[AdvancedExcel.COL_INDEX] + 1, limitoffset))
                break

        return allResults

    def getAllRowsFromKey(self, _sheet, keyname):
        """
        Finds all rows in a sheet that contain a specific keyname (case-insensitive)
        and returns information about their location.

        This method iterates through each cell in the provided sheet (`_sheet`).
        It checks if the cell value is a string, is not NaN, and, after converting
        both the cell value and the `keyname` to lowercase and stripping whitespace,
        if they are equal. If a match is found, it stores the entire row as a
        pandas Series, the row index, and the column index in a dictionary
        and appends it to the `dataRows` list.

        Args:
            _sheet (pandas.DataFrame): The sheet to search within.
            keyname (str): The keyname to search for (case-insensitive).

        Returns:
            list: A list of dictionaries. Each dictionary contains the following keys:
                - `AdvancedExcel.DATA` (pandas.Series): The entire row containing the key.
                - `AdvancedExcel.ROW_INDEX` (int): The row index of the key.
                - `AdvancedExcel.COL_INDEX` (int): The column index of the key.
                Returns an empty list if the `keyname` is not found.
        """

        dataRows = []

        for row in range(_sheet.shape[0]):  # df is the DataFrame
            for col in range(_sheet.shape[1]):
                if not pd.isna(_sheet.iloc[row, col]) \
                        and isinstance(_sheet.iloc[row, col], str) \
                        and _sheet.iloc[row, col].lower().strip() == keyname.lower().strip():
                    dataRows.append({AdvancedExcel.DATA: _sheet.loc[row],
                                     AdvancedExcel.ROW_INDEX: row,
                                     AdvancedExcel.COL_INDEX: col
                                     })
                    break

        return dataRows

    def getValueFromPosition(self, _sheet, col, row, as_string=True):
        """
         Retrieves a value from a specific cell in a sheet using its column and row indices.

         **Important:** This method is provided for exceptional cases where other data
         retrieval methods are not feasible.  Directly accessing data using fixed
         coordinates is generally discouraged due to its inflexibility.

         **Row Indexing:**  Because the first row of the DataFrame corresponds to the
         header (column names), you must adjust the Excel row number when using this
         method.  Specifically:
             - Excel row 1 (header) corresponds to `row = -1`
             - Excel row 2 corresponds to `row = 0`
             - Excel row 3 corresponds to `row = 1`
             and so on.

         Args:
             _sheet (pandas.DataFrame): The sheet to retrieve the value from.
             col (int): The column index (0-based).
             row (int): The adjusted row index (see explanation above).
             as_string (bool, optional): If True (default), converts the retrieved
                 value to a string and strips whitespace. If False, returns the
                 value as its original type.

         Returns:
             str or object: The value from the specified cell.  Returns a string
                 if `as_string` is True, otherwise returns the value with its
                 original data type.  If the cell is part of the header (row = -1),
                 the corresponding column name is returned.
         """

        if row == -1:
            oValue = _sheet.columns[col]
        else:
            oValue = _sheet.iloc[row, col]
        if isinstance(oValue, str):
            return oValue.strip()
        else:
            if as_string:
                return str(oValue).strip()
            else:
                return oValue

    def getAllPairValues(self, _sheet):
        """
        Extracts key-value pairs from a sheet, assuming a maximum of one pair per row.

        This method iterates through each row of the provided sheet (`_sheet`).
        For each row, it drops any NaN values and checks if exactly two values
        remain. If so, it treats the first remaining value as the key and the
        second as the value, storing them in the `pairKeyValues` dictionary.

        Args:
            _sheet (pandas.DataFrame): The sheet to extract key-value pairs from.

        Returns:
            dict: A dictionary where keys are the first value in a two-value row
                and values are the second value in that row.  Rows with fewer
                or more than two non-NaN values are ignored.
        """
        pairKeyValues = {}

        for row in range(_sheet.shape[0]):
            the_row = _sheet.loc[row].dropna()
            if len(the_row) == 2:
                pairKeyValues[the_row[0]] = the_row[1]

        return pairKeyValues

    def getAllTables(self, _sheet, mincol=3):
        """
         Identifies and extracts table-like structures from a sheet.

         This method searches for contiguous blocks of rows that have at least `mincol`
         non-NaN values. These blocks are interpreted as tables. It iterates through
         the rows of the sheet.  When it encounters a row with at least `mincol`
         valid values, it marks the beginning of a potential table. When it finds a row
         with fewer than `mincol` valid values, it considers the preceding block of
         rows as a complete table and extracts it using the `_getTable` method.

         Args:
             _sheet (pandas.DataFrame): The sheet to search for tables within.
             mincol (int, optional): The minimum number of non-NaN values required
                 for a row to be considered part of a table. Defaults to 3.

         Returns:
             list: A list of pandas DataFrames, where each DataFrame represents
                 a table-like structure found in the sheet.  Returns an empty
                 list if no tables are found.
         """

        allTables = []
        init_table = -1

        for irow in range(_sheet.shape[0]):
            if len(_sheet.loc[irow].dropna()) >= mincol:
                if init_table == -1:
                    init_table = irow
            else:
                if init_table >= 0:
                    table = self._getTable(_sheet, init_table, irow)
                    allTables.append(table)
                    init_table = -1

        if init_table >= 0:
            table = self._getTable(_sheet, init_table, _sheet.shape[0])
            allTables.append(table)

        return allTables

    def getDataFrameBlocksByKeyName(self, df, key_name):
        """ used for sheet that contains different blocks of information,
        it splits the all data sheet into splited dataframes to be parser later on
        The Split is using the recognition of the first value on a cell that is repeated at the beginning
        of each block"""

        allDfBlocks = []

        identified_rows = self.getAllRowsFromKey(df, key_name)
        list_index_of_rows = list(map(lambda x: x[AdvancedExcel.ROW_INDEX], identified_rows))

        for index in range(len(list_index_of_rows)):
            ini = list_index_of_rows[index]
            if index + 1 > len(list_index_of_rows) - 1:
                end = df.shape[0]
            else:
                end = list_index_of_rows[index + 1]

            block = df.iloc[ini:end] \
                .dropna(axis=1, how='all') \
                .dropna(axis=0, how='all')

            block = block.reset_index(drop=True)
            allDfBlocks.append(block)

        return allDfBlocks

    def getDataFrameBlocksByKeyColumn(self, df, key_column):
        """
        Splits a DataFrame into blocks based on occurrences of a keyname.

        This method is designed for sheets containing multiple blocks of information,
        separated by a repeated `key_name` at the beginning of each block. It identifies
        the rows containing the `key_name` and uses their indices to split the DataFrame
        into individual blocks. Each block is then cleaned by removing rows and columns
        that are entirely NaN.

        Args:
            df (pandas.DataFrame): The DataFrame to split.
            key_name (str): The keyname that marks the beginning of each block.

        Returns:
            list: A list of pandas DataFrames, where each DataFrame represents a
                block of data. Returns an empty list if the `key_name` is not found.
        """

        allDfBlocks = []

        list_index_of_columns = self._getAllColumnsFromKey(df, key_column)

        for index in range(len(list_index_of_columns)):
            ini = list_index_of_columns[index]
            if index + 1 > len(list_index_of_columns) - 1:
                end = df.shape[1]
            else:
                end = list_index_of_columns[index + 1]

            block = df.iloc[:, ini:end] \
                .dropna(axis=0, how='all') \
                .dropna(axis=1, how='all')

            block = block.reset_index(drop=True)
            allDfBlocks.append(block)

        return allDfBlocks

    def replaceNaNInColumnByNaNStrings(self, df, columnName):
        """
        Replaces NaN (Not a Number) values in a specified column of a DataFrame with the string "NA".

        This method modifies the DataFrame in place by filling any NaN values in the column
        specified by `columnName` with the string "NA". This is often useful when you
        need to represent missing data with a specific string value, for example, when
        exporting the data to a format that doesn't handle NaN values well.

        Args:
            df (pandas.DataFrame): The DataFrame to modify.
            columnName (str): The name of the column in which to replace NaN values.

        Returns:
            pandas.DataFrame: The modified DataFrame (the modification is done in place).
        """
        df.loc[:, columnName] = df[columnName].fillna('NA')
        return df

    def replaceNaNStrings(self, df):
        """
        Replaces string representations of NaN values with actual NaN values in a DataFrame.

        This method searches for specific string values ('NA_NA', 'NA', 'NAN', 'na_na', 'na', 'nan')
        that are often used to represent missing data and replaces them with actual
        NaN (Not a Number) values. This ensures consistency in how missing data is
        handled within the DataFrame, allowing pandas functions to correctly interpret
        and process the missing values.

        Args:
            df (pandas.DataFrame): The DataFrame to modify.

        Returns:
            pandas.DataFrame: A new DataFrame with the string representations of NaN
                replaced by actual NaN values.  A copy of the DataFrame is created,
                so the original DataFrame is not modified in place.
        """
        df = pd.DataFrame(np.where(df.isin(['NA_NA', 'NA', 'NAN', "na_na", "na", "nan"]), np.nan, df), df.index,
                          df.columns)
        return df

    def fillForwardRowValues(self, df, columnNames=[]):
        """
       Fills forward missing values in specified columns of a DataFrame.

       In Excel files, merged cells can result in subsequent rows having empty
       values. This method fills these empty values with the value from the
       previous row within the specified columns.  This effectively propagates
       the value downwards until a new value is encountered.

       For example, if a column has the following values:
       `[val1, NaN, NaN, NaN, val2, NaN, NaN, ...]`
       this method will transform it to:
       `[val1, val1, val1, val1, val2, val2, val2, ...]`

       Args:
           df (pandas.DataFrame): The DataFrame to modify.
           columnNames (list, optional): A list of column names where forward fill
               should be applied. If None, defaults to an empty list, and no columns are filled.

       Returns:
           pandas.DataFrame: The modified DataFrame (the modification is done in place).
       """
        for cName in columnNames:
            if cName in df:
                df[cName] = df[cName].ffill()
        return df

    def removeEmptyRows(self, df):
        """
        Removes rows that contain only NaN (Not a Number) values from a DataFrame and resets the index.

        This method removes rows where *all* values are NaN. It then resets the DataFrame's
        index and drops the old index column (if it exists).  This ensures that the
        resulting DataFrame has a clean, sequential index.

        Args:
            df (pandas.DataFrame): The DataFrame to modify.

        Returns:
            pandas.DataFrame: A new DataFrame with the empty rows removed and the index reset.
            A copy of the DataFrame is created, so the original DataFrame is not modified in place.
        """
        df = df.dropna(axis=0, how='all')
        df = df.reset_index()
        df = df.drop(columns=['index'], errors='ignore')
        return df

    def removeEmptyColumns(self, df):
        """
       Removes empty or entirely whitespace-filled columns from a DataFrame.

       This method removes columns that contain either all NaN (Not a Number) values
       or all empty strings. It first removes columns with all NaN values using
       `dropna` and then removes columns where all values are empty strings ('').

       Args:
           df (pandas.DataFrame): The DataFrame to modify.

       Returns:
           pandas.DataFrame: A new DataFrame with the empty or whitespace-only columns removed.
           A copy of the DataFrame is created, so the original DataFrame is not modified in place.
       """
        df = df.dropna(axis=1, how='all')
        df = df.loc[:, (df != '').any(axis=0)]

        return df

def removeDuplicates(self, df):
    """
    Removes duplicate rows from a DataFrame and resets the index.

    This method removes any duplicate rows from the DataFrame.  It uses `drop_duplicates`
    with `ignore_index=True` to generate a new index without modifying the original.
    It then resets the index to ensure a clean, sequential index and drops the old index column.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with duplicate rows removed and the index reset.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.drop_duplicates(ignore_index=True)  # Remove duplicates and ignore original index
    df = df.reset_index()  # Reset the index
    df = df.drop(columns=['index'], errors='ignore')  # Drop the old index column (if it exists)
    return df

def removeRowsByRegEx(self, df, columnName, regex):
    """
    Removes rows from a DataFrame where a specified column matches a regular expression.

    This method uses `str.contains` to create a boolean mask indicating which rows
    match the given regular expression in the specified column. It then uses
    boolean indexing to keep only the rows that *do not* match the regex.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str): The name of the column to check against the regex.
        regex (str): The regular expression to match.

    Returns:
        pandas.DataFrame: A new DataFrame with rows matching the regex removed.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    filter = df[columnName].str.contains(regex, na=False)  # Create a boolean mask
    df = df[~filter]  # Keep rows that *do not* match the regex
    return df

def removeRowsNotInRegEx(self, df, columnName, regex):
    """
    Removes rows from a DataFrame where a specified column *does not* match a regular expression.

    This method is the inverse of `removeRowsByRegEx`. It uses `str.contains` to create a
    boolean mask indicating which rows match the given regular expression in the specified
    column. It then uses boolean indexing to keep *only* the rows that *do* match the regex.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str): The name of the column to check against the regex.
        regex (str): The regular expression to match.

    Returns:
        pandas.DataFrame: A new DataFrame with rows *not* matching the regex removed.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    filter = df[columnName].str.contains(regex, na=False)  # Create a boolean mask
    df = df[filter]  # Keep only rows that *do* match the regex
    return df

def removeRowsWhereValueInColumnIsEmpty(self, df, columnName):
    """
    Removes rows from a DataFrame where a specified column contains 'na' or null values.

    This method removes rows where the value in the specified `columnName` is either
    the string 'na' (case-insensitive) or a null (NaN) value.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str): The name of the column to check.

    Returns:
        pandas.DataFrame: A new DataFrame with rows containing 'na' or null in the
        specified column removed. A copy of the DataFrame is created, so the
        original DataFrame is not modified in place.
    """
    if columnName in df:  # Check if the column exists
        df = df[df[columnName].str.lower() != 'na']  # Remove rows with 'na' (case-insensitive)
        df = df[df[columnName].notnull()]  # Remove rows with null values
        df = df.reset_index()  # Reset the index
        df = df.drop(columns=['index'], errors='ignore')  # Drop the old index column
    return df

def removeRowsWhereValueIsEqualInColumn(self, df, columnName, rowValue):
    """
    Removes rows from a DataFrame where a specified column's value is equal to rowValue.

    This method removes rows where the value in the specified `columnName` is equal to the
    provided `rowValue`.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str): The name of the column to check.
        rowValue: The value to compare against.

    Returns:
        pandas.DataFrame: A new DataFrame with rows where the column value equals `rowValue`
        removed. A copy of the DataFrame is created, so the original DataFrame is not
        modified in place.
    """
    if columnName in df:  # Check if the column exists
        df = df[df[columnName] != rowValue]  # Remove rows where value equals rowValue
        df = df.reset_index()  # Reset the index
        df = df.drop(columns=['index'], errors='ignore')  # Drop the old index column
    return df
def transpose(self, df):
    """
    Transposes a DataFrame, sets the first row as the header, and performs formatting.

    This method transposes the input DataFrame, then uses the first row of the
    transposed DataFrame as the new header. It also converts the header to uppercase
    and replaces spaces in the header names with underscores.

    Args:
        df (pandas.DataFrame): The DataFrame to transpose and format.

    Returns:
        pandas.DataFrame: The transposed and formatted DataFrame. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.transpose().reset_index(drop=True)  # Transpose and reset index
    df = self.setRowAsHeader(df, index_row=0, uppercase=True, char_space_to="_")  # Set first row as header
    return df

def mergeRows(self, df, init_row, end_row, separator=''):
    """
    Merges the content of multiple rows into a single row.

    This method merges the content of rows from `init_row` (inclusive) to `end_row`
    (exclusive) into a single row.  It uses the specified `separator` to join
    the values within each column.  NaN values are replaced with the separator.
    The merged row is then inserted back into the DataFrame.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_row (int): The starting row index (inclusive).
        end_row (int): The ending row index (exclusive).
        separator (str, optional): The separator to use when joining values. Defaults to ''.

    Returns:
        pandas.DataFrame: A new DataFrame with the specified rows merged. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    dfHead = df[0:init_row]  # DataFrame before the rows to merge
    dfTail = df[end_row:]  # DataFrame after the rows to merge
    dfToMerge = df[init_row:end_row]  # DataFrame containing rows to be merged

    dfToMerge = pd.DataFrame(np.where(pd.isna(dfToMerge), separator, dfToMerge), dfToMerge.index, dfToMerge.columns) # Replace NaN with separator
    dfToMerge['TMP'] = 0  # Temporary column for grouping
    dfToMerge = dfToMerge.set_index('TMP')  # Set temporary column as index
    dfToMerge = dfToMerge.groupby(dfToMerge['TMP'], as_index=False).agg(' '.join)  # Merge rows using groupby and join
    dfToMerge = dfToMerge.drop('TMP', axis=1)  # Drop the temporary column

    df = pd.concat([dfHead, dfToMerge, dfTail], ignore_index=True)  # Concatenate the DataFrames
    df = df.reset_index(drop=True)  # Reset the index
    #df = df.drop(columns=['index'], errors='ignore')  # Drop the old index column. This is now redundant due to ignore_index=True in concat.

    return df

def stripAll(self, df):
    """
    Removes leading and trailing whitespace from all string values in a DataFrame.

    This method uses a regular expression to remove any whitespace characters
    at the beginning or end of each string value in the DataFrame.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with leading/trailing whitespace removed from all strings.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.replace(to_replace=r'^\s*(.*?)\s*$', value=r'\1', regex=True)
    return df

def removeAllCRLF(self, df, byValue=''):
    """
    Removes carriage return (CR) and line feed (LF) characters from string values in a DataFrame.

    This method replaces all occurrences of carriage return (`\r`) and line feed (`\n`)
    characters in the DataFrame with the value specified by `byValue`.  It then
    optionally replaces double occurrences of `byValue` with a single `byValue`.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        byValue (str, optional): The value to replace CR/LF characters with. Defaults to ''.

    Returns:
        pandas.DataFrame: A new DataFrame with CR/LF characters removed or replaced.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.replace({'\n': byValue, '\r': byValue}, regex=True)  # Replace CR/LF with byValue
    if byValue:  # Check if byValue is not empty
        df = df.replace({byValue + byValue: byValue}, regex=True)  # Replace double byValue
    return df

def replaceByDictionary(self, df, dictionary, column_to_apply):
    """
    Replaces values in a specified column of a DataFrame using a dictionary mapping.

    This method uses a dictionary to map existing values in the specified column
    to new values.  It also replaces the micro symbol (µ) with the letter 'u'
    before applying the dictionary mapping.  Any values that are not found
    as keys in the dictionary are left unchanged.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        dictionary (dict): A dictionary where keys are the values to replace and
            values are their replacements.
        column_to_apply (str): The name of the column to apply the replacement to.

    Returns:
        pandas.DataFrame: A new DataFrame with values replaced according to the dictionary.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    df.loc[:, column_to_apply] = (df[column_to_apply]
                                  .str.replace(r'µ', 'u')
                                  .map(dictionary)
                                  .fillna(df[column_to_apply]))
    return df

def replaceAll(self, df, to_replace, replace_by, columns_to_clean=None, columns_not_to_clean=None):
    """
    Replaces all occurrences of a specified value with another value in a DataFrame.

    This method replaces all occurrences of `to_replace` with `replace_by` within the
    DataFrame.  It offers flexibility in targeting specific columns:

    - If both `columns_to_clean` and `columns_not_to_clean` are None (or empty lists),
      the replacement is applied to *all* columns.
    - If `columns_to_clean` is provided, the replacement is applied *only* to the
      columns listed in this argument.
    - If `columns_not_to_clean` is provided, the replacement is applied to *all*
      columns *except* those listed in this argument.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        to_replace (str): The value to replace.
        replace_by (str): The value to replace with.
        columns_to_clean (list, optional): A list of column names to apply the
            replacement to. Defaults to None.
        columns_not_to_clean (list, optional): A list of column names to *exclude*
            from the replacement. Defaults to None.

    Returns:
        pandas.DataFrame: A new DataFrame with the replacements made. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """

    if columns_to_clean is None and columns_not_to_clean is None:  # Apply to all columns
        df = df.replace(to_replace=r'' + to_replace, value=r'' + replace_by, regex=True)
    elif columns_to_clean:  # Apply to specified columns
        for col in columns_to_clean:
            df[col] = df[col].astype(str)  # Ensure column is string type
            df[col] = df[col].replace(to_replace=r'' + to_replace, value=r'' + replace_by, regex=True)
    elif columns_not_to_clean:  # Apply to all but specified columns
        for col in df.columns:
            if col not in columns_not_to_clean:
                df[col] = df[col].astype(str)  # Ensure column is string type
                df[col] = df[col].replace(to_replace=r'' + to_replace, value=r'' + replace_by, regex=True)

    return df

def replaceAllByList(self, df, to_replace, columns_to_clean=None, columns_not_to_clean=None):
    """
    Replaces values in a DataFrame using a list of (pattern, replacement) tuples.

    This method iterates through a list of tuples, where each tuple defines a
    regular expression pattern and its corresponding replacement value.  For each
    tuple, it calls the `replaceAll` method to perform the replacement.  This allows
    for multiple replacements to be applied sequentially.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        to_replace (list of tuples): A list of tuples, where each tuple contains
            (pattern, replacement). The `pattern` is a regular expression.
        columns_to_clean (list, optional): A list of column names to apply the
            replacements to. Defaults to None.
        columns_not_to_clean (list, optional): A list of column names to *exclude*
            from the replacements. Defaults to None.

    Returns:
        pandas.DataFrame: A new DataFrame with the replacements made. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    for pattern, replacement in to_replace:  # Iterate through the list of (pattern, replacement) tuples
        df = self.replaceAll(df, pattern, replacement, columns_to_clean, columns_not_to_clean)

    return df

def replaceSpacesBySeparator(self, df, separator=';', columns_to_clean=None, columns_not_to_clean=None):
    """
    Replaces multiple spaces with a specified separator in a DataFrame.

    This method replaces multiple consecutive whitespace characters with the specified
    `separator`. It also replaces consecutive occurrences of the separator with a single
    separator.  This helps to normalize spacing within string values.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        separator (str, optional): The character to use as a separator. Defaults to ';'.
        columns_to_clean (list, optional): A list of column names to apply the
            replacement to. Defaults to None.
        columns_not_to_clean (list, optional): A list of column names to *exclude*
            from the replacement. Defaults to None.

    Returns:
        pandas.DataFrame: A new DataFrame with spaces replaced by the separator.
        A copy of the DataFrame is created, so the original DataFrame is not
        modified in place.
    """
    if columns_to_clean is None and columns_not_to_clean is None:  # Apply to all columns
        df = df.replace(to_replace=r'\s+', value=separator, regex=True) \
            .replace(to_replace=r'[' + separator + ']+', value=separator, regex=True)  # Replace multiple separators
    elif columns_to_clean:  # Apply to specified columns
        for col in columns_to_clean:
            df[col] = df[col].astype(str) # Ensure string type
            df[col] = df[col].replace(to_replace=r'\s+', value=separator, regex=True) \
                .replace(to_replace=r'[' + separator + ']+', value=separator, regex=True) # Replace multiple separators
    elif columns_not_to_clean:  # Apply to all but specified columns
        for col in df.columns:
            if col not in columns_not_to_clean:
                df[col] = df[col].astype(str)  # Ensure string type
                df[col] = df[col].replace(to_replace=r'\s+', value=separator, regex=True) \
                    .replace(to_replace=r'[' + separator + ']+', value=separator, regex=True)  # Replace multiple separators

    return df
def setColumn2Numeric(self, df, columnName):
    """
    Converts a specified column in a DataFrame to numeric type.

    This method uses `pd.to_numeric` to convert the specified column to a numeric
    data type.  This is useful when a column has been read as a string to preserve
    precision and needs to be converted for numerical operations.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str): The name of the column to convert.

    Returns:
        pandas.DataFrame: The modified DataFrame (the modification is done in place).
    """
    df[columnName] = pd.to_numeric(df[columnName])
    return df

def setListOfColumns2Numeric(self, df, columnNames=None):
    """
    Converts a list of columns in a DataFrame to numeric type.

    This method iterates through the provided list of column names and uses
    `pd.to_numeric` to convert each column to a numeric data type.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnNames (list, optional): A list of column names to convert. Defaults to None.

    Returns:
        pandas.DataFrame: The modified DataFrame (the modification is done in place).
    """
    if columnNames is None:
        columnNames = []

    for c_name in columnNames:
        df[c_name] = pd.to_numeric(df[c_name])
    return df

def setHeaderAsARow(self, df):
    """
    Moves the current header of a DataFrame to a new row at the top.

    This method takes the current header (column names) of the DataFrame and inserts
    it as a new row at the beginning of the DataFrame. The columns are then
    renamed with default integer indices.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with the original header as the first row.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    headers = df.columns.tolist()  # Get the current header
    df.loc[-1] = headers  # Insert the header as a new row at index -1
    df = df.sort_index().reset_index(drop=True)  # Sort by index and reset it
    return df

def setRowAsHeader(self, df, index_row, uppercase=False, char_space_to=" "):
    """
    Sets a specified row as the new header of a DataFrame.

    This method takes the values from the specified `index_row` and sets them as the
    new header (column names) of the DataFrame. The original row is then removed.
    It also offers options to convert the header to uppercase and replace spaces
    with a specified character.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        index_row (int): The index of the row to use as the header.
        uppercase (bool, optional): Whether to convert the header to uppercase. Defaults to False.
        char_space_to (str, optional): The character to replace spaces with in the header. Defaults to " ".

    Returns:
        pandas.DataFrame: A new DataFrame with the specified row as the header.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    new_header = pd.DataFrame(df.values[index_row]).T # Transpose to get a single row DataFrame for headers
    new_header = self.removeAllCRLF(new_header)  # Remove CR/LF characters
    new_header = self.stripAll(new_header)  # Remove leading/trailing whitespace
    if uppercase:
        new_header = new_header.apply(lambda x: x.astype(str).str.upper())  # Convert to uppercase
    if char_space_to != " ":
        new_header = new_header.replace(' ', char_space_to, regex=True)  # Replace spaces
    df.columns = new_header.iloc[0]  # Set the new header
    df = df.drop(index_row).reset_index(drop=True)  # Remove the old header row and reset index
    return df

def setRowAsHeaderByValue(self, df, column_to_search, value_to_match, uppercase=False, char_space_to=" "):
    """
    Sets a row as the header based on a value match in a specified column.

    This method searches for a specific value (`value_to_match`) in the specified
    `column_to_search`. If a match is found, the corresponding row is used as the
    new header using the same logic as the `setRowAsHeader` function.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_to_search (str): The name of the column to search in.
        value_to_match: The value to search for.
        uppercase (bool, optional): Whether to convert the header to uppercase. Defaults to False.
        char_space_to (str, optional): The character to replace spaces with in the header. Defaults to " ".

    Returns:
        pandas.DataFrame: A new DataFrame with the matching row as the header, or the
        original DataFrame if no match is found or the column does not exist. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    if column_to_search not in df.columns:
        return df  # Return original DataFrame if column doesn't exist

    index_row = df.index[df[column_to_search] == value_to_match].tolist() # Get the index as a list
    if index_row:  # Check if a match was found
        df = self.setRowAsHeader(df, index_row[0], uppercase, char_space_to)
    return df

def addHeaderAsARow(self, df):
    """
    Adds the current header of a DataFrame as a new row at the top.

    This method takes the current header (column names) of the DataFrame and inserts
    it as a new row at the beginning of the DataFrame. The columns are then
    renamed with default integer indices.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with the original header as the first row.
         A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = pd.concat([df.columns.to_frame().T, df], ignore_index=True)  # Concatenate header as a row
    df.columns = range(len(df.columns))  # Reset columns to default integer indices
    return df

def removeRowIfUniqueValue(self, df, init_row, end_row):
    """
    Removes rows containing only a single unique non-NaN value within a specified range.

    This method checks rows within the specified range (`init_row` to `end_row` inclusive)
    for rows where there is only one unique, non-NaN value. These rows are typically
    titles or other single-value rows in Excel files.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_row (int): The starting row index (inclusive).
        end_row (int): The ending row index (inclusive).

    Returns:
        pandas.DataFrame: A new DataFrame with the single-value rows removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    indexes_to_remove = []
    for index in range(init_row, end_row + 1):
        num_values = len(df.loc[index].dropna().unique()) # Count unique non-NaN values
        if num_values == 1:
            indexes_to_remove.append(index)

    df = df.drop(indexes_to_remove)
    df = df.reset_index(drop=True)
    return df
def removeRowsIfUniqueValueOnTop(self, df, offset):
    """
    Removes rows from the top of a DataFrame if they contain only one unique value.

    This method calls `removeRowIfUniqueValue` to remove rows from the beginning
    (top) of the DataFrame if they contain only one unique, non-NaN value.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        offset (int): The number of rows from the top to check.

    Returns:
        pandas.DataFrame: A new DataFrame with the single-value rows removed from the top.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    return self.removeRowIfUniqueValue(df, 0, offset)

def removeRowsIfUniqueValueOnBotton(self, df, offset):
    """
    Removes rows from the bottom of a DataFrame if they contain only one unique value.

    This method calls `removeRowIfUniqueValue` to remove rows from the end
    (bottom) of the DataFrame if they contain only one unique, non-NaN value.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        offset (int): The number of rows from the bottom to check.

    Returns:
        pandas.DataFrame: A new DataFrame with the single-value rows removed from the bottom.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    return self.removeRowIfUniqueValue(df, len(df.index) - offset, len(df.index))

def removeAllRowsFromAColumnNameWhenHasValue(self, df, init_row, columnName, rowValue, include_remove_row=True):
    """
    Removes rows from a DataFrame starting from a specified row when a value is found in a named column.

    This method finds the first occurrence of `rowValue` in the specified `columnName`
    starting from `init_row`.  It then removes all rows *from* that row (inclusive
    or exclusive, depending on `include_remove_row`) to the end of the DataFrame.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_row (int): The starting row index for the search.
        columnName (str): The name of the column to search in.
        rowValue: The value to search for.
        include_remove_row (bool, optional): If True (default), the row containing
            `rowValue` is *included* in the removed rows. If False, the row containing
            `rowValue` is *excluded* from the removed rows (rows *after* the matched row are removed).

    Returns:
        pandas.DataFrame: A new DataFrame with the rows removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    column_index = df.columns.get_loc(columnName)  # Get column index
    return self.removeAllRowsFromAColumnWhenHasValue(df, init_row, column_index, rowValue, include_remove_row)

def removeAllRowsFromAColumnWhenHasValue(self, df, init_row, index_column, value, include_remove_row=True):
    """
    Removes rows from a DataFrame starting from a specified row when a value is found in a column (by index).

    This method is similar to `removeAllRowsFromAColumnNameWhenHasValue`, but it uses the
    column index instead of the column name.  It finds the first occurrence of `value`
    in the specified column (by index) starting from `init_row`. It then removes all
    rows from that row (inclusive or exclusive, depending on `include_remove_row`) to
    the end of the DataFrame.  Handles `None` (NaN) values for `value` and the column.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_row (int): The starting row index for the search.
        index_column (int): The index of the column to search in.
        value: The value to search for.
        include_remove_row (bool, optional): If True (default), the row containing
            `value` is *included* in the removed rows. If False, the row containing
            `value` is *excluded* from the removed rows (rows *after* the matched row are removed).

    Returns:
        pandas.DataFrame: A new DataFrame with the rows removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    index_bottom = len(df.index)
    index_top = -1
    for index in range(init_row, index_bottom):
        cell_value = df.loc[index][index_column]
        if (cell_value == value) or \
                (pd.isnull(value) and pd.isnull(cell_value)): # Handle NaN comparison
            if include_remove_row:
                index_top = index
            else:
                index_top = index + 1
            break
    if index_top != -1:
        df = df.drop(range(index_top, index_bottom))
        df = df.reset_index(drop=True)
    return df

def removeAllRowsBeforeColumnValue(self, df, init_row, index_column, value):
    """
    Removes rows from a DataFrame *before* a specified value is found in a column.

    This method searches for the first occurrence of `value` in the specified `index_column`
    starting from `init_row`. It then removes all rows *before* that row (from `init_row` up to, but not including, the found row).
    Handles matching against a list of values, single values, and `None` (NaN) values.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_row (int): The starting row index for the search.
        index_column (int): The index of the column to search in.
        value: The value (or list of values) to search for.

    Returns:
        pandas.DataFrame: A new DataFrame with the rows removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    index_bottom = len(df.index)
    index_top = -1
    for index in range(init_row, index_bottom + 1):
        cell_value = df.loc[index][index_column]
        if (isinstance(value, list) and any(cell_value == v for v in value)) or \
                cell_value == value or \
                (isinstance(value, str) and pd.isnull(value) and pd.isnull(cell_value)): # Handle list, single, and NaN comparison
            index_top = index
            break
    if index_top != -1:
        df = df.drop(range(init_row, index_top))
        df = df.reset_index(drop=True)
    return df

def removeRowByIndex(self, df, index_row):
    """
    Removes a row from a DataFrame by its index.

    **Caution:** Using this method directly with integer row indices can be risky if
    the DataFrame's index has been manipulated (e.g., by sorting or filtering).  It's
    generally safer to use label-based indexing (`df.drop(label)`) if possible.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        index_row (int): The index of the row to remove.

    Returns:
        pandas.DataFrame: A new DataFrame with the row removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.drop(index_row)
    df = df.reset_index(drop=True)
    return df

    # -------------------------------------------------------------------------------------------------
    # COLUMNS
    # -------------------------------------------------------------------------------------------------
def removeAllColumnsWhereAnyRowHasAValue(self, df, value):
    """
    Removes columns from a DataFrame where *any* row contains a specified value.

    This method removes columns where at least one cell contains the specified `value`.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        value: The value to check for.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.loc[:, ~(df == value).any()]  # Remove columns where any row has the value
    return df

def removeAllColumnsWhereALLRowHasAValue(self, df, value):
    """
    Removes columns from a DataFrame where *all* rows contain a specified value.

    This method removes columns where *all* cells contain the specified `value`.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        value: The value to check for.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.loc[:, ~(df == value).all()]  # Remove columns where all rows have the value
    return df

def removeAllColumnsWhereALLRowAreEmpty(self, df):
    """
    Removes columns from a DataFrame where *all* rows are empty (NaN).

    This method removes columns that contain only NaN (Not a Number) values.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with the empty columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.dropna(how='all', axis=1)  # Remove columns where all values are NaN
    return df

def removeAllColumnsBeforeColumnValueHeader(self, df, header_value, skip_first=False):
    """
    Removes columns from a DataFrame *before* a column with a specified header value.

    This method searches for a column with the specified `header_value`. It then
    removes all columns *to the left* of that column (excluding the column with
    `header_value` itself). The `skip_first` parameter allows skipping the first
    occurrence of the header value in case it appears multiple times.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        header_value: The header value to search for.
        skip_first (bool, optional): Whether to skip the first occurrence of the
            header value. Defaults to False.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    header_was_found = False
    index_column = -1
    index = 0
    for cnames in df.columns:
        if header_was_found and index_column == -1:
            index_column = index
        elif cnames == header_value:
            if not skip_first:
                header_was_found = True
            else:
                skip_first = False
        index += 1  # More Pythonic increment
    if index_column != -1:
        df = df.iloc[:, range(index_column)]  # Keep columns up to (but not including) the target column
        df = df.reset_index(drop=True)
    return df

def removeAllColumnsAfterColumnValueHeader(self, df, header_value, skip_first=False):
    """
    Removes columns from a DataFrame *after* a column with a specified header value.

    This method searches for a column with the specified `header_value`.  It then
    removes all columns *to the right* of that column (excluding the column with
    `header_value` itself). The `skip_first` parameter allows skipping the first
    occurrence of the header value in case it appears multiple times.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        header_value: The header value to search for.
        skip_first (bool, optional): Whether to skip the first occurrence of the
            header value. Defaults to False.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    header_was_found = False
    index_column = -1
    index = 0
    for cnames in df.columns:
        if header_was_found and index_column == -1:
            index_column = index
        elif cnames == header_value:
            if not skip_first:
                header_was_found = True
            else:
                skip_first = False
        index += 1  # More Pythonic increment

    if index_column != -1:
        df = df.iloc[:, range(0, index_column + 1)] # Keep columns from the beginning up to and including the target column.
        df = df.reset_index(drop=True)
    return df
def moveToFirst(self, df, column_to_move_ahead):
    """
    Moves a specified column to the first position in a DataFrame.

    This method moves the column specified by `column_to_move_ahead` to the
    beginning (first position) of the DataFrame.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_to_move_ahead (str): The name of the column to move.

    Returns:
        pandas.DataFrame: A new DataFrame with the column moved to the first position.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
        Returns the original DataFrame if the specified column does not exist.
    """
    if column_to_move_ahead not in df.columns:
        return df # Return original if column doesn't exist

    col2move = df.pop(column_to_move_ahead)  # Remove the column and store it
    df.insert(0, column_to_move_ahead, col2move)  # Insert the column at the beginning
    return df
def melt(self, df, column_to_melt, column_name_id=None, column_name_values="VALUE"):
    """
    Transforms a DataFrame from a wide format to a long format (key-value pairs).

    This method performs a "melt" operation, converting a DataFrame with multiple
    columns (e.g., "value_x", "value_y", "value_z") into a key-value structure.
    The values from the specified columns are stacked into a single column
    (`column_name_values`), and a new column (`column_name_id` or the original
    `column_to_melt`) is created to store the corresponding column names (prefixed
    with the values from `column_to_melt`).

    For example, if the input DataFrame has columns ["parameter", "value_x", "value_y", "value_z"]
    and `column_to_melt` is "parameter" with values ["avg", "std", "%"], the output
    DataFrame will have columns (either `column_name_id` or `parameter`) and `column_name_values`.
    The values in the new column will be similar to: "avg_value_x", "avg_value_y", ..., "%_value_z".

    Args:
        df (pandas.DataFrame): The input DataFrame.
        column_to_melt (str): The name of the column containing the prefix for the new key column.
        column_name_id (str, optional): The name for the new ID column. If None, the
            original `column_to_melt` is used (with modified values). Defaults to None.
        column_name_values (str, optional): The name for the new values column. Defaults to "VALUE".

    Returns:
        pandas.DataFrame: A new DataFrame in long format. A copy of the DataFrame
        is created, so the original DataFrame is not modified in place.
    """
    df.loc[:, column_to_melt] = df.loc[:, column_to_melt].astype(str)  # Ensure string type for the column to melt
    df = df.melt(id_vars=[column_to_melt], var_name='column_name', value_name=column_name_values)  # Melt the DataFrame

    if column_name_id:  # Create the ID column if specified
        df[column_name_id] = df[column_to_melt] + '_' + df['column_name']  # Combine parameter and column names
        df.drop(columns=[column_to_melt], inplace=True)  # Drop the original parameter column
    else:  # Modify the original parameter column
        df[column_to_melt] = df[column_to_melt] + '_' + df['column_name']  # Combine parameter and column names
    df.drop(columns=['column_name'], inplace=True)  # Drop the temporary column name column

    return df
def melt_extended(self, df, column_to_generate_columns, column_name_id):
    """
    Performs a melt and pivot transformation to reshape a DataFrame.

    This method transforms a DataFrame from a wide format to a different wide format.
    One column (`column_to_generate_columns`) provides the values that will become
    the new columns. The other columns are "melted" into rows and then "pivoted"
    to create the new wide format.

    For example, if the input DataFrame has columns ["MTYPE", "P1", "P2"] and the values
    are:
    MTYPE = ["AVG", "STDEV", "N"]
    P1 = [0.2, 3.4, 5.5]
    P2 = [0.3, 3.3, 7.0]

    The output DataFrame will have columns ["PARAMETER", "AVG", "STDEV", "N"] with values:
    PARAMETER = ["P1", "P2"]
    AVG = [0.2, 0.3]
    STDEV = [3.4, 3.3]
    N = [5.5, 7.0]

    Args:
        df (pandas.DataFrame): The input DataFrame.  It should contain the column
            specified by `column_to_generate_columns` and other columns whose values
            are to be pivoted.
        column_to_generate_columns (str): The name of the column whose unique values
            will become the new column names.
        column_name_id (str): The name for the new ID column (which will contain the
            original column names that were melted).

    Returns:
        pandas.DataFrame: A new DataFrame in the reshaped wide format. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    columns_with_values = df.columns[df.columns != column_to_generate_columns].tolist()  # Columns to melt
    long_df = pd.melt(df, id_vars=[column_to_generate_columns], value_vars=columns_with_values,  # Melt
                      var_name="PARAMETERS", value_name="value")  # Rename value column to "value" for pivot_table compatibility
    wide_df = pd.pivot_table(long_df, index=["PARAMETERS"], columns=[column_to_generate_columns], values="value")  # Pivot
    wide_df[column_name_id] = wide_df.index  # Add ID column
    wide_df.reset_index(drop=True, inplace=True)  # Reset index
    wide_df = wide_df[[column_name_id] + [col for col in wide_df.columns if col not in [column_name_id]]]  # Reorder columns
    return wide_df
def transformAsNumeric(self, df, columnsToApply=None, columnsToAvoid=None, errors='coerce'):
    """
    Transforms specified columns in a DataFrame to numeric type, handling errors.

    This method attempts to convert the values in specified columns to numeric type.
    It offers flexibility in targeting specific columns:

    - If both `columnsToApply` and `columnsToAvoid` are None, the conversion is
      attempted on *all* columns.
    - If `columnsToApply` is provided, the conversion is applied *only* to the
      columns listed in this argument.
    - If `columnsToAvoid` is provided, the conversion is applied to *all*
      columns *except* those listed in this argument.

    The `errors` parameter controls how conversion errors are handled.  Using
    `errors='coerce'` (the default) will replace any non-numeric values with NaN.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnsToApply (list, optional): A list of column names to apply the
            conversion to. Defaults to None.
        columnsToAvoid (list, optional): A list of column names to *exclude*
            from the conversion. Defaults to None.
        errors (str, optional): How to handle conversion errors.  Defaults to 'coerce'.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns converted to numeric.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    if columnsToApply is None and columnsToAvoid is None:
        columns_to_convert = df.columns  # Convert all columns
    elif columnsToApply:
        columns_to_convert = columnsToApply  # Convert specified columns
    elif columnsToAvoid:
        columns_to_convert = df.columns.difference(columnsToAvoid)  # Convert all but avoided columns
    else:
        columns_to_convert = [] # Empty list if none of the above.

    df[columns_to_convert] = df[columns_to_convert].apply(pd.to_numeric, errors=errors, axis=1)
    return df

def fillColumnsWithLeftValue(self, df, index_row):
    """
    Fills NaN values in a row with the value from the nearest cell to the left.

    This method iterates through the columns of a specified row from left to right.
    If it encounters a NaN value, it fills it with the value from the nearest
    non-NaN cell to its left.  This is often useful for handling Excel files where
    cells have been merged or spanned.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        index_row (int): The index of the row to fill.

    Returns:
        pandas.DataFrame: A new DataFrame with the NaN values filled.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    left_value = None  # Store the value from the left
    df = df.reset_index(drop=True)  # Reset index for consistent access
    for index_column in range(len(df.columns)):
        current_value = df.loc[index_row, df.columns[index_column]] # More efficient way to access the cell value.
        if pd.isnull(current_value):
            df.loc[index_row, df.columns[index_column]] = left_value  # Fill with left value
        else:
            left_value = current_value  # Update left value

    return df

def removeColumnsByNameNaN(self, df):
    """
    Removes columns with NaN or "NAN" as their name.

    This method removes columns whose names are NaN (Not a Number) or the string "NAN".

    Args:
        df (pandas.DataFrame): The DataFrame to modify.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.loc[:, df.columns.notnull()]  # Remove columns with NaN names
    df = df.drop(columns=["NAN"], errors='ignore')  # Remove column named "NAN" (if it exists)
    return df
def removeColumnByName(self, df, key_name, if_column_exists=None):
    """
    Removes columns from a DataFrame based on a regular expression.

    This method removes columns whose names match the provided regular expression `key_name`.
    The `if_column_exists` parameter allows for conditional removal: if provided, the
    column is only dropped if the column with the `if_column_exists` name exists.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        key_name (str): The regular expression to match column names against.
        if_column_exists (str, optional): The name of a column that must exist for
            the removal to proceed. Defaults to None.

    Returns:
        pandas.DataFrame: A new DataFrame with the matching columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    if if_column_exists is None or if_column_exists in df.columns:
        columns_to_drop = list(df.filter(regex=key_name)) # Get the list of columns to drop
        df = df.drop(columns=columns_to_drop, errors='ignore') # Drop the columns, ignoring errors if they don't exist.
    return df

def removeColumnsByNames(self, df, column_names):
    """
    Removes multiple columns from a DataFrame based on a list of names.

    This method removes columns whose names are present in the `column_names` list.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_names (list): A list of column names to remove.

    Returns:
        pandas.DataFrame: A new DataFrame with the specified columns removed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    columns_to_drop = '|'.join(column_names)  # Create a regex pattern from the list of column names
    df = df.drop(columns=list(df.filter(regex=columns_to_drop)), errors='ignore') # Drop the columns.
    return df

def removeAllColumnsExceptListOfColumnNames(self, df, column_names):
    """
    Removes all columns from a DataFrame *except* those in a specified list.

    This method keeps only the columns listed in `column_names` and removes all other columns.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_names (list): A list of column names to keep.

    Returns:
        pandas.DataFrame: A new DataFrame with only the specified columns. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.drop(df.columns.difference(column_names), axis=1, errors='ignore') #Drop columns not in the list.
    return df

def mergeColumns(self, df, init_col, end_col, separator, name_of_column):
    """
    Merges the content of multiple columns into a single column.

    This method merges the values from columns `init_col` to `end_col` (inclusive)
    into a new column named `name_of_column`, using the specified `separator` to
    join the values.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        init_col (int): The starting column index (inclusive).
        end_col (int): The ending column index (inclusive).
        separator (str): The separator to use when joining values.
        name_of_column (str): The name of the new merged column.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns merged. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df[name_of_column] = df[df.columns[init_col:end_col + 1]] \
        .apply(lambda x: separator.join(x.dropna().astype(str)), axis=1)  # Merge columns
    columns_to_drop = df.columns[range(init_col, end_col + 1)].tolist() # Get the list of columns to drop.
    df = df.drop(columns=columns_to_drop, errors='ignore')  # Remove the original columns
    return df

def mergeColumnsByName(self, df, column_destine, column_left, column_right, remove_origin=True, separator='_'):
    """
    Merges content from two columns into a destination column.

    This method merges the values from `column_left` and `column_right` into
    `column_destine`, using the specified `separator`. If `remove_origin` is True
    (the default), the original `column_left` and `column_right` are removed.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_destine (str): The name of the destination column.
        column_left (str): The name of the left column to merge.
        column_right (str): The name of the right column to merge.
        remove_origin (bool, optional): Whether to remove the original columns. Defaults to True.
        separator (str, optional): The separator to use when joining values. Defaults to '_'.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns merged. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df[column_destine] = df[column_left].astype(str) + separator + df[column_right].astype(str)  # Merge columns, ensure strings
    if remove_origin:
        df = df.drop(columns=[column_right, column_left], errors='ignore')  # Remove original columns
    return df

def splitOnColumns(self, df, column_to_split, split_character="_", list_of_columns=None):
    """
    Splits a column into multiple columns based on a delimiter.

    This method splits the column specified by `column_to_split` into multiple
    columns using the `split_character` as a delimiter. The resulting columns are
    named according to the `list_of_columns` argument.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_to_split (str): The name of the column to split.
        split_character (str, optional): The character to split on. Defaults to "_".
        list_of_columns (list, optional): A list of names for the new columns. If None,
          new columns will be named 0, 1, 2, etc. Defaults to None.

    Returns:
        pandas.DataFrame: A new DataFrame with the column split. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """

    if list_of_columns is None:
        list_of_columns = range(len(df[column_to_split].iloc[0].split(split_character))) #Generate a list of numbers based on the split size.

    if df[column_to_split].str.contains(split_character).all(): #Only split if all rows contain the split character.
        df[list_of_columns] = df[column_to_split].str.split(split_character, expand=True)
    return df
def splitIn2ColumnsRight(self, df, column_to_split, columns_to_generate=None, split_character="_"):
    """
    Splits a column into two columns using the *rightmost* occurrence of a delimiter.

    This method splits the specified `column_to_split` into two new columns using
    the *rightmost* occurrence of the `split_character` as the delimiter.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_to_split (str): The name of the column to split.
        columns_to_generate (list, optional): A list of two names for the new columns.
            Defaults to ["LEFT_VALUE", "RIGHT_VALUE"].
        split_character (str, optional): The character to use as the delimiter. Defaults to "_".

    Returns:
        pandas.DataFrame: A new DataFrame with the column split. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    if columns_to_generate is None:
        columns_to_generate = ["LEFT_VALUE", "RIGHT_VALUE"]

    if df[column_to_split].str.contains(split_character).any(): #Only split if at least one row contains the split character
        df[columns_to_generate] = df[column_to_split].str.rsplit(split_character, n=1, expand=True)
    return df

def splitIn2ColumnsLeft(self, df, column_to_split, columns_to_generate=None, split_character="_"):
    """
    Splits a column into two columns using the *leftmost* occurrence of a delimiter.

    This method splits the specified `column_to_split` into two new columns using
    the *leftmost* occurrence of the `split_character` as the delimiter.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        column_to_split (str): The name of the column to split.
        columns_to_generate (list, optional): A list of two names for the new columns.
            Defaults to ["LEFT_VALUE", "RIGHT_VALUE"].
        split_character (str, optional): The character to use as the delimiter. Defaults to "_".

    Returns:
        pandas.DataFrame: A new DataFrame with the column split. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    if columns_to_generate is None:
        columns_to_generate = ["LEFT_VALUE", "RIGHT_VALUE"]

    if df[column_to_split].str.contains(split_character).any(): #Only split if at least one row contains the split character
        df[columns_to_generate] = df[column_to_split].str.split(split_character, n=1, expand=True)
    return df

def renameHeaders(self, df, dictionary):
    """
    Renames columns using a dictionary mapping.

    This method renames columns using the provided dictionary, where keys are the
    old column names and values are the new column names.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        dictionary (dict): A dictionary mapping old column names to new column names.

    Returns:
        pandas.DataFrame: The modified DataFrame (the modification is done in place).
    """
    df.rename(columns=dictionary, inplace=True)
    return df

def renameHeadersByRegExp(self, df, regexp, value=''):
    """
    Renames columns using a regular expression substitution.

    This method renames columns by applying a regular expression substitution.
    The `regexp` is used to search for parts of the column names, and `value` is
    used as the replacement.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        regexp (str): The regular expression to search for.
        value (str, optional): The replacement value. Defaults to ''.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns renamed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    df = df.rename(columns=lambda x: re.sub(regexp, value, x))
    return df

def renameHeadersByRegExpList(self, df, listOfRegExp, value=''):
    """
    Renames columns using a list of regular expression substitutions.

    This method applies a list of regular expression substitutions to the column
    names.  Each `regexp` in the `listOfRegExp` is used to search for parts of
    the column names, and `value` is used as the replacement.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        listOfRegExp (list): A list of regular expressions.
        value (str, optional): The replacement value. Defaults to ''.

    Returns:
        pandas.DataFrame: A new DataFrame with the columns renamed. A copy of the
        DataFrame is created, so the original DataFrame is not modified in place.
    """
    for regexp in listOfRegExp:
        df = self.renameHeadersByRegExp(df, regexp, value)
    return df

def caseHeaders(self, df, uppercase=True):
    """
    Converts column names to uppercase or lowercase.

    This method converts all column names to uppercase if `uppercase` is True,
    otherwise it converts them to lowercase.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        uppercase (bool, optional): Whether to convert to uppercase. Defaults to True.

    Returns:
        pandas.DataFrame: The modified DataFrame (the modification is done in place).
    """
    df.columns = df.columns.str.upper() if uppercase else df.columns.str.lower()
    return df

def caseColumnValues(self, df, columnName=None, title=False, uppercase=True):
    """
    Converts values in a specified column to title case, uppercase, or lowercase.

    This method converts the values in the specified `columnName` to title case
    (first letter uppercase, rest lowercase) if `title` is True. Otherwise, it
    converts them to uppercase if `uppercase` is True, or lowercase if `uppercase`
    is False.  It also strips leading/trailing whitespace from the column values.

    Args:
        df (pandas.DataFrame): The DataFrame to modify.
        columnName (str, optional): The name of the column to modify. Defaults to None.
        title (bool, optional): Whether to convert to title case. Defaults to False.
        uppercase (bool, optional): Whether to convert to uppercase (if title is False). Defaults to True.

    Returns:
        pandas.DataFrame: The modified DataFrame (the modification is done in place).
    """
    if columnName is not None:
        df = self.replaceNaNInColumnByNaNStrings(df, columnName)  # Handle potential NaN values
        df.loc[:, columnName] = df[columnName].str.strip()  # Remove leading/trailing whitespace

        if title:
            df.loc[:, columnName] = df[columnName].str.title()  # Convert to title case
        else:
            df.loc[:, columnName] = df[columnName].str.upper() if uppercase else df[columnName].str.lower()  # Convert to upper/lower case
    return df

    # -------------------------------------------------------------------------------------------------
    # LISTS
    # -------------------------------------------------------------------------------------------------

def mergeListDataFrames(self, df_list, column_id_name="GROUP_ID"):
    """
    Merges a list of DataFrames into a single DataFrame and assigns group IDs.

    This method concatenates a list of DataFrames vertically. It adds a "Group" column
    indicating the original DataFrame index for each row and then assigns a unique
    serial ID (`column_id_name`) to each row in the combined DataFrame.

    Args:
        df_list (list): A list of pandas DataFrames to merge.
        column_id_name (str, optional): The name for the new serial ID column. Defaults to "GROUP_ID".

    Returns:
        pandas.DataFrame: A new DataFrame containing all the merged data and the
        group ID column. A copy of the DataFrame is created, so the original
        DataFrames are not modified in place.
    """
    merged_df = pd.concat(df_list, keys=range(len(df_list)))  # Concatenate with group keys
    merged_df = merged_df.reset_index(level=0)  # Move group keys to a column
    merged_df = merged_df.rename(columns={'level_0': 'Group'})  # Rename the group key column
    merged_df = merged_df.reset_index(drop=True)  # Reset the index
    merged_df[column_id_name] = range(len(merged_df))  # Assign serial IDs
    return merged_df

    # -------------------------------------------------------------------------------------------------
    # "PRIVATE"
    # -------------------------------------------------------------------------------------------------
def _getAllColumnsFromKey(self, df, key_column):
    """
    Returns a list of indices for columns matching the specified key column name.

    This method searches for columns in the DataFrame whose names match `key_column`
    and returns a list of their integer indices.

    Args:
        df (pandas.DataFrame): The DataFrame to search.
        key_column (str): The name of the column to search for.

    Returns:
        list: A list of integer indices of the matching columns.  Returns an
        empty list if no matching columns are found.
    """
    key_indices = []
    for i, column in enumerate(df.columns):
        if column == key_column:
            key_indices.append(i)
    return key_indices

def _getNextValue(self, list_like, initial_position, limitOffset=0):
    """
    Searches for the next non-NaN value in a list-like object within a limited offset.

    This method searches for the first non-NaN value in the given `list_like` object
    (e.g., a pandas Series) starting from `initial_position`.  The search is limited
    by `limitOffset` if it's greater than 0.

    Args:
        list_like (pandas.Series or list-like): The list-like object to search.
        initial_position (int): The starting index for the search.
        limitOffset (int, optional): The maximum offset from `initial_position` to
            search. If 0, searches to the end of the list. Defaults to 0.

    Returns:
        The next non-NaN value found, or None if no such value is found within the limit.
    """
    max_offset = len(list_like) if limitOffset <= 0 else initial_position + limitOffset
    for index in range(initial_position, min(max_offset, len(list_like))): # Ensure index does not go beyond list bounds.
        value = list_like.iloc[index] if hasattr(list_like, 'iloc') else list_like[index] # Handles both Series and list-like.
        if not pd.isna(value):
            return value
    return None # Return None if no value is found.

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
        return [sheet.name for sheet in wb.sheets] # Return a list of sheet names.

def _getTable(self, _sheet, init_table, end_table):
    """
    Extracts a table (DataFrame) from a sheet within specified row boundaries.

    This method extracts a portion of the input sheet (`_sheet`) between `init_table`
    and `end_table` (exclusive) as a new DataFrame. It then cleans the table by
    removing rows and columns that are entirely NaN, sets the first row as the header
    (after cleaning and formatting it), and resets the index.

    Args:
        _sheet (pandas.DataFrame): The sheet (DataFrame) to extract the table from.
        init_table (int): The starting row index (inclusive).
        end_table (int): The ending row index (exclusive).

    Returns:
        pandas.DataFrame: A new DataFrame representing the extracted table.
        A copy of the DataFrame is created, so the original DataFrame is not modified in place.
    """
    table = _sheet.iloc[init_table:end_table].dropna(axis=1, how='all').dropna(axis=0, how='all') # Extract and clean

    # Set the header from the first row after cleaning and formatting:
    table.columns = [str(s).strip().upper().replace(' ', '_') for s in table.iloc[0]]
    table = table.drop(0).reset_index(drop=True)  # Remove the first row (old header) and reset index
    return table

def _headerColumnsAreEmpty(self, columns):
    """
    Checks if all column names in a Series start with "Unnamed:".

    Args:
        columns (pandas.Series): The Series containing the column names.

    Returns:
        bool: True if all column names start with "Unnamed:", False otherwise.
    """
    return columns.str.contains('^Unnamed:').all()

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

    if hasattr(cell, 'hyperlink') and hasattr(cell.hyperlink, 'target'):
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

OpenpyxlReader._convert_cell = _convert_cell  # Apply the monkey patch
