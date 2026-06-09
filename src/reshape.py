import pandas as pd
import numpy as np
import re


class ReshapeMixin:
    """
    Mixin for reshape / wide-long / melt / pivot style operations and
    merging lists of DataFrames.
    """

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
