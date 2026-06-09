import pandas as pd


class LookupMixin:
    """
    Mixin for key/value lookup and navigation inside sheets
    (get next value to the right of a key, find all rows containing a key, etc.).
    """

    # Note: the three special keys are exposed on the final AdvancedExcel class
    # as class attributes (ROW_INDEX, COL_INDEX, DATA). Internally we use the
    # literal strings for the dict keys so the mixins stay self-contained.

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
            allResults.append(self._getNextValue(row["data"], row["col_index"] + 1, limitoffset))

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
                if none of the keynames are found, or if no value is found within the
                offset limit for the first found key.
        """
        allResults = []
        for keyname in keynames:
            data_rows = self.getAllRowsFromKey(_sheet, keyname)
            if len(data_rows) > 0:
                row = data_rows[0]
                allResults.append(
                    self._getNextValue(row["data"], row["col_index"] + 1, limitoffset)
                )
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
                - "data" (pandas.Series): The entire row containing the key.
                - "row_index" (int): The row index of the key.
                - "col_index" (int): The column index of the key.
                Returns an empty list if the `keyname` is not found.
        """

        dataRows = []
        key_lower = keyname.lower().strip()

        # Improved version: still correct first-match-per-row semantics,
        # but avoids unnecessary NaN checks and uses direct string compare.
        # For repeated calls on the same large sheet, a cached lowercased view
        # would yield further gains.
        for row_idx in range(_sheet.shape[0]):
            row = _sheet.iloc[row_idx]
            for col_idx, val in enumerate(row):
                if isinstance(val, str) and val.lower().strip() == key_lower:
                    dataRows.append({
                        "data": _sheet.loc[row_idx],
                        "row_index": row_idx,
                        "col_index": col_idx,
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
        for index in range(
            initial_position, min(max_offset, len(list_like))
        ):  # Ensure index does not go beyond list bounds.
            value = (
                list_like.iloc[index] if hasattr(list_like, "iloc") else list_like[index]
            )  # Handles both Series and list-like.
            if not pd.isna(value):
                return value
        return None  # Return None if no value is found.
