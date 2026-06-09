import pandas as pd
import numpy as np
import re


class CleaningMixin:
    """
    Mixin for value/string/NaN/header/dtype cleaning and transformation operations.
    Includes replace*, strip, case, rename, numeric conversion helpers, etc.
    """

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
                df[col] = df[col].astype(str)  # Ensure string type
                df[col] = df[col].replace(to_replace=r'\s+', value=separator, regex=True) \
                    .replace(to_replace=r'[' + separator + ']+', value=separator, regex=True)  # Replace multiple separators
        elif columns_not_to_clean:  # Apply to all but specified columns
            for col in df.columns:
                if col not in columns_not_to_clean:
                    df[col] = df[col].astype(str)  # Ensure string type
                    df[col] = df[col].replace(to_replace=r'\s+', value=separator, regex=True) \
                        .replace(to_replace=r'[' + separator + ']+', value=separator, regex=True)  # Replace multiple separators

        return df

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
            columns_to_convert = []  # Empty list if none of the above.

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
            current_value = df.loc[index_row, df.columns[index_column]]  # More efficient way to access the cell value.
            if pd.isnull(current_value):
                df.loc[index_row, df.columns[index_column]] = left_value  # Fill with left value
            else:
                left_value = current_value  # Update left value

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
