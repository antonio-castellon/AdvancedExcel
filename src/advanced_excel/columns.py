import pandas as pd


class ColumnMixin:
    """
    Mixin containing all column-oriented operations:
    removing/filtering columns by various criteria, merging/splitting columns,
    moving columns, setting columns to numeric, etc.
    """

    def remove_empty_columns(self, df):
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
        df = df.dropna(axis=1, how="all")
        df = df.loc[:, (df != "").any(axis=0)]

        return df

    def remove_all_columns_where_any_row_has_a_value(self, df, value):
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

    def remove_all_columns_where_all_row_has_a_value(self, df, value):
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

    def remove_all_columns_where_all_row_are_empty(self, df):
        """
        Removes columns from a DataFrame where *all* rows are empty (NaN).

        This method removes columns that contain only NaN (Not a Number) values.

        Args:
            df (pandas.DataFrame): The DataFrame to modify.

        Returns:
            pandas.DataFrame: A new DataFrame with the empty columns removed. A copy of the
            DataFrame is created, so the original DataFrame is not modified in place.
        """
        df = df.dropna(how="all", axis=1)  # Remove columns where all values are NaN
        return df

    def remove_all_columns_before_column_value_header(self, df, header_value, skip_first=False):
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
        index_column = -1
        for idx, cname in enumerate(df.columns):
            if cname == header_value:
                if not skip_first:
                    index_column = idx  # position of the target column
                    break
                else:
                    skip_first = False
        if index_column != -1:
            df = df.iloc[:, :index_column]  # strictly before the target (excluding it)
            df = df.reset_index(drop=True)
        return df

    def remove_all_columns_after_column_value_header(self, df, header_value, skip_first=False):
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
        index_column = -1
        for idx, cname in enumerate(df.columns):
            if cname == header_value:
                if not skip_first:
                    index_column = idx  # position of the target column
                    break
                else:
                    skip_first = False
        if index_column != -1:
            df = df.iloc[:, : index_column + 1]  # up to and including the target
            df = df.reset_index(drop=True)
        return df

    def move_to_first(self, df, column_to_move_ahead):
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
            return df  # Return original if column doesn't exist

        col2move = df.pop(column_to_move_ahead)  # Remove the column and store it
        df.insert(0, column_to_move_ahead, col2move)  # Insert the column at the beginning
        return df

    def set_column_to_numeric(self, df, columnName):
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

    def set_list_of_columns_to_numeric(self, df, columnNames=None):
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

    def remove_columns_by_name_nan(self, df):
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
        df = df.drop(columns=["NAN"], errors="ignore")  # Remove column named "NAN" (if it exists)
        return df

    def remove_column_by_name(self, df, key_name, if_column_exists=None):
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
            columns_to_drop = list(df.filter(regex=key_name))  # Get the list of columns to drop
            df = df.drop(
                columns=columns_to_drop, errors="ignore"
            )  # Drop the columns, ignoring errors if they don't exist.
        return df

    def remove_columns_by_names(self, df, column_names):
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
        columns_to_drop = "|".join(
            column_names
        )  # Create a regex pattern from the list of column names
        df = df.drop(
            columns=list(df.filter(regex=columns_to_drop)), errors="ignore"
        )  # Drop the columns.
        return df

    def remove_all_columns_except_list_of_column_names(self, df, column_names):
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
        df = df.drop(
            df.columns.difference(column_names), axis=1, errors="ignore"
        )  # Drop columns not in the list.
        return df

    def merge_columns(self, df, init_col, end_col, separator, name_of_column):
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
        df[name_of_column] = df[df.columns[init_col : end_col + 1]].apply(
            lambda x: separator.join(x.dropna().astype(str)), axis=1
        )  # Merge columns
        columns_to_drop = df.columns[
            range(init_col, end_col + 1)
        ].tolist()  # Get the list of columns to drop.
        df = df.drop(columns=columns_to_drop, errors="ignore")  # Remove the original columns
        return df

    def merge_columns_by_name(
        self, df, column_destine, column_left, column_right, remove_origin=True, separator="_"
    ):
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
        df[column_destine] = (
            df[column_left].astype(str) + separator + df[column_right].astype(str)
        )  # Merge columns, ensure strings
        if remove_origin:
            df = df.drop(
                columns=[column_right, column_left], errors="ignore"
            )  # Remove original columns
        return df

    def split_on_columns(self, df, column_to_split, split_character="_", list_of_columns=None):
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
            list_of_columns = range(
                len(df[column_to_split].iloc[0].split(split_character))
            )  # Generate a list of numbers based on the split size.

        if (
            df[column_to_split].str.contains(split_character).all()
        ):  # Only split if all rows contain the split character.
            df[list_of_columns] = df[column_to_split].str.split(split_character, expand=True)
        return df

    def split_in_2_columns_right(
        self, df, column_to_split, columns_to_generate=None, split_character="_"
    ):
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

        if (
            df[column_to_split].str.contains(split_character).any()
        ):  # Only split if at least one row contains the split character
            df[columns_to_generate] = df[column_to_split].str.rsplit(
                split_character, n=1, expand=True
            )
        return df

    def split_in_2_columns_left(
        self, df, column_to_split, columns_to_generate=None, split_character="_"
    ):
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

        if (
            df[column_to_split].str.contains(split_character).any()
        ):  # Only split if at least one row contains the split character
            df[columns_to_generate] = df[column_to_split].str.split(
                split_character, n=1, expand=True
            )
        return df
