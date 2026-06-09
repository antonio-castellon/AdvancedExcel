class BlockTableMixin:
    """
    Mixin for detecting and extracting multiple tables / blocks / entities
    inside a single sheet (repeated key sections, "Batch Number" style blocks, etc.).
    This is the core of the "advanced" multi-table Excel handling.
    """

    def get_all_tables(self, _sheet, mincol=3):
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

    def get_dataframe_blocks_by_key_name(self, df, key_name):
        """used for sheet that contains different blocks of information,
        it splits the all data sheet into splited dataframes to be parser later on
        The Split is using the recognition of the first value on a cell that is repeated at the beginning
        of each block"""

        allDfBlocks = []

        identified_rows = self.get_all_rows_from_key(df, key_name)
        list_index_of_rows = list(map(lambda x: x["row_index"], identified_rows))

        for index in range(len(list_index_of_rows)):
            ini = list_index_of_rows[index]
            if index + 1 > len(list_index_of_rows) - 1:
                end = df.shape[0]
            else:
                end = list_index_of_rows[index + 1]

            block = df.iloc[ini:end].dropna(axis=1, how="all").dropna(axis=0, how="all")

            block = block.reset_index(drop=True)
            allDfBlocks.append(block)

        return allDfBlocks

    def get_dataframe_blocks_by_key_column(self, df, key_column):
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

            block = df.iloc[:, ini:end].dropna(axis=0, how="all").dropna(axis=1, how="all")

            block = block.reset_index(drop=True)
            allDfBlocks.append(block)

        return allDfBlocks

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
        table = (
            _sheet.iloc[init_table:end_table].dropna(axis=1, how="all").dropna(axis=0, how="all")
        )  # Extract and clean

        # Set the header from the first row after cleaning and formatting:
        table.columns = [str(s).strip().upper().replace(" ", "_") for s in table.iloc[0]]
        table = table.drop(0).reset_index(
            drop=True
        )  # Remove the first row (old header) and reset index
        return table

    def _headerColumnsAreEmpty(self, columns):
        """
        Checks if all column names in a Series start with "Unnamed:".

        Args:
            columns (pandas.Series): The Series containing the column names.

        Returns:
            bool: True if all column names start with "Unnamed:", False otherwise.
        """
        return columns.str.contains("^Unnamed:").all()
