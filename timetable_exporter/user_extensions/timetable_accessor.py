import importlib
import pandas as pd
from pandas.api.extensions import register_dataframe_accessor


def _resolve_column_name(df: pd.DataFrame, name: str) -> str:
    if name in df.columns:
        return name

    if name is None:
        raise KeyError("Column name is None")

    target = str(name)
    target_stripped = target.strip()

    # Exact match after stripping
    stripped_map = {}
    for col in df.columns:
        if isinstance(col, str):
            stripped_map.setdefault(col.strip(), col)
    if target_stripped in stripped_map:
        return stripped_map[target_stripped]

    # Case-insensitive match after stripping
    folded_map = {}
    for col in df.columns:
        if isinstance(col, str):
            folded_map.setdefault(col.strip().casefold(), col)
    folded_key = target_stripped.casefold()
    if folded_key in folded_map:
        return folded_map[folded_key]

    raise KeyError(f"Column not found: {name}. Available columns: {list(df.columns)}")

@register_dataframe_accessor("timetable")
class TimetableAccessor:
    def __init__(self, pandas_obj):
        self._obj = pandas_obj

    def filter(self, filters, exact_match=False):
        """
        Filters the DataFrame based on the provided filters.
        :param filters: A dictionary where keys are column names and values are the filter values.
                        If the value is a list, it will filter for any of the values in the list.
        :param exact_match: If True, filters for exact matches. If False, uses string contains.
        :return: A filtered DataFrame.
        """
        df = self._obj.copy()
        # Apply custom filters to the DataFrame if any
        if filters is None:
            return df
        
        try:
            for column, values in filters.items():
                resolved_column = _resolve_column_name(df, column)
                if isinstance(values, list):
                    if exact_match:
                        df = df[df[resolved_column].isin(values)]
                    else:
                        # Ensure the column is of string type before using .str.contains
                        df = df[df[resolved_column].astype(str).str.contains('|'.join(values), case=False, na=False)]
                else:
                    if exact_match:
                        df = df[df[resolved_column] == values]
                    else:
                        # Ensure the column is of string type before using .str.contains
                        df = df[df[resolved_column].astype(str).str.contains(values, case=False, na=False)]

        except KeyError as e:
            raise KeyError(f"Column not found in DataFrame: {e}")
        except Exception as e:
            raise Exception(f"Error applying filters: {e}")
        return df
    
    # function for calling an objects internal method on each row of a column
    def call_internal_method(self, method_name: str, column: str, *args, **kwargs):
        """
        Calls an internal method of the DataFrame on each row.
        :param method_name: The name of the method to call.
        :param args: Positional arguments to pass to the method.
        :param kwargs: Keyword arguments to pass to the method.

        :return: A DataFrame with the results of the method call.
        """ 
        df = self._obj.copy()
        # Check if the column exists in the DataFrame
        resolved_column = _resolve_column_name(df, column)
        # Check if the method is callable

        # Check if the method exists for at least one element in the column
        if not all(hasattr(x, method_name) for x in df[resolved_column].dropna()):
            raise AttributeError(f"Method '{method_name}' not found for elements in column '{resolved_column}'.")

        df[resolved_column] = df[resolved_column].apply(lambda x: getattr(x, method_name)(*args, **kwargs))
        return df
    
    def rename_columns(self, mapper, **kwargs):
        """
        Renames the columns of the DataFrame.
        :param mapper: A dictionary mapping old column names to new column names.
        :param kwargs: Additional arguments to pass to the rename method.
        :return: A DataFrame with renamed columns.
        """
        df = self._obj.copy()
        return df.rename(columns=mapper, **kwargs)