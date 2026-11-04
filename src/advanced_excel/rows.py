import numpy as np
import pandas as pd


class RowMixin:
    """
    Mixin containing all row-oriented operations:
    removing/filtering rows by various criteria, merging rows,
    setting rows as headers, forward-fill on rows, etc.
    """

    def remove_duplicates(self, df):
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
        df = df.drop(columns=["index"], errors="ignore")  # Drop the old index column (if it exists)
        return df

    def remove_rows_by_regex(self, df, columnName, regex):
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

    def remove_rows_not_in_regex(self, df, columnName, regex):
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

    def remove_rows_where_value_in_column_is_empty(self, df, columnName):
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
            df = df[df[columnName].str.lower() != "na"]  # Remove rows with 'na' (case-insensitive)
            df = df[df[columnName].notnull()]  # Remove rows with null values
            df = df.reset_index()  # Reset the index
            df = df.drop(columns=["index"], errors="ignore")  # Drop the old index column
        return df

    def remove_rows_where_value_is_equal_in_column(self, df, columnName, rowValue):
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
            df = df.drop(columns=["index"], errors="ignore")  # Drop the old index column
        return df

    def merge_rows(self, df, init_row, end_row, separator=""):
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

        dfToMerge = pd.DataFrame(
            np.where(pd.isna(dfToMerge), separator, dfToMerge), dfToMerge.index, dfToMerge.columns
        )  # Replace NaN with separator
        dfToMerge["TMP"] = 0  # Temporary column for grouping
        grouped = dfToMerge.groupby("TMP", as_index=False).agg(lambda x: separator.join(x))
        grouped = grouped.drop(columns=["TMP"], errors="ignore")

        df = pd.concat([dfHead, grouped, dfTail], ignore_index=True)  # Concatenate the DataFrames
        df = df.reset_index(drop=True)  # Reset the index
        # df = df.drop(columns=['index'], errors='ignore')  # Drop the old index column. This is now redundant due to ignore_index=True in concat.

        return df

    def remove_row_if_unique_value(self, df, init_row, end_row):
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
        n = len(df)
        for index in range(init_row, end_row + 1):
            if index >= n:
                break
            num_values = len(df.loc[index].dropna().unique())  # Count unique non-NaN values
            if num_values == 1:
                indexes_to_remove.append(index)

        if indexes_to_remove:
            df = df.drop(indexes_to_remove)
        df = df.reset_index(drop=True)
        return df

    def remove_rows_if_unique_value_on_top(self, df, offset):
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
        return self.remove_row_if_unique_value(df, 0, offset)

    def remove_rows_if_unique_value_on_bottom(self, df, offset):
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
        return self.remove_row_if_unique_value(df, len(df.index) - offset, len(df.index))

    def remove_all_rows_from_a_column_name_when_has_value(
        self, df, init_row, columnName, rowValue, include_remove_row=True
    ):
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
        return self.remove_all_rows_from_a_column_when_has_value(
            df, init_row, column_index, rowValue, include_remove_row
        )

    def remove_all_rows_from_a_column_when_has_value(
        self, df, init_row, index_column, value, include_remove_row=True
    ):
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
            cell_value = df.loc[index].iloc[index_column]
            if (cell_value == value) or (
                pd.isnull(value) and pd.isnull(cell_value)
            ):  # Handle NaN comparison
                index_top = index if include_remove_row else index + 1
                break
        if index_top != -1:
            df = df.drop(range(index_top, index_bottom))
            df = df.reset_index(drop=True)
        return df

    def remove_all_rows_before_column_value(self, df, init_row, index_column, value):
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
        for index in range(init_row, index_bottom):
            cell_value = df.loc[index].iloc[index_column]
            if (
                (isinstance(value, list) and any(cell_value == v for v in value))
                or cell_value == value
                or (isinstance(value, str) and pd.isnull(value) and pd.isnull(cell_value))
            ):  # Handle list, single, and NaN comparison
                index_top = index
                break
        if index_top != -1:
            df = df.drop(range(init_row, index_top))
            df = df.reset_index(drop=True)
        return df

    def remove_row_by_index(self, df, index_row):
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

    # Also include the row-related methods that lived in the original class body
    def remove_empty_rows(self, df):
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
        df = df.dropna(axis=0, how="all")
        df = df.reset_index()
        df = df.drop(columns=["index"], errors="ignore")
        return df

    def fill_forward_row_values(self, df, columnNames=None):
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
        for cName in columnNames or []:
            if cName in df:
                df[cName] = df[cName].ffill()
        return df

    def set_header_as_a_row(self, df):
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
        header_row = pd.DataFrame([headers], columns=df.columns, dtype=object)
        df = pd.concat([header_row, df], ignore_index=True)
        return df

    def set_row_as_header(self, df, index_row, uppercase=False, char_space_to=" "):
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
        new_header = pd.DataFrame(
            df.values[index_row]
        ).T  # Transpose to get a single row DataFrame for headers
        new_header = self.remove_all_crlf(new_header)  # Remove CR/LF characters
        new_header = self.strip_all(new_header)  # Remove leading/trailing whitespace
        if uppercase:
            new_header = new_header.apply(
                lambda x: x.astype(str).str.upper()
            )  # Convert to uppercase
        if char_space_to != " ":
            new_header = new_header.replace(" ", char_space_to, regex=True)  # Replace spaces
        df.columns = new_header.iloc[0]  # Set the new header
        df = df.drop(index_row).reset_index(drop=True)  # Remove the old header row and reset index
        return df

    def set_row_as_header_by_value(
        self, df, column_to_search, value_to_match, uppercase=False, char_space_to=" "
    ):
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

        index_row = df.index[
            df[column_to_search] == value_to_match
        ].tolist()  # Get the index as a list
        if index_row:  # Check if a match was found
            df = self.set_row_as_header(df, index_row[0], uppercase, char_space_to)
        return df

    def add_header_as_a_row(self, df):
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
        df = pd.concat(
            [df.columns.to_frame().T, df], ignore_index=True
        )  # Concatenate header as a row
        df.columns = range(len(df.columns))  # Reset columns to default integer indices
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
        df = self.set_row_as_header(
            df, index_row=0, uppercase=True, char_space_to="_"
        )  # Set first row as header
        return df
