import pandas as pd
from datetime import datetime


def _resolve_column_name(df: pd.DataFrame, name: str) -> str:
    if name in df.columns:
        return name
    if name is None:
        raise ValueError("Column name is None")
    target = str(name).strip()
    for col in df.columns:
        if isinstance(col, str) and col.strip() == target:
            return col
    folded = target.casefold()
    for col in df.columns:
        if isinstance(col, str) and col.strip().casefold() == folded:
            return col
    raise ValueError(f"Column '{name}' does not exist in the DataFrame. Available columns: {list(df.columns)}")

def combine_date_time(self, date_col: str, time_col: str, datetime_col=None, tz=None, drop_invalid: bool = False, keep_source: bool = False) -> pd.DataFrame:
    """
    Combine date and time columns into a single datetime column.
    """
    if datetime_col is None:
        datetime_col = f"{date_col}_{time_col}"

    df = self._obj.copy()

    date_col_resolved = _resolve_column_name(df, date_col)
    time_col_resolved = _resolve_column_name(df, time_col)

    # Combine the date and time columns into a single datetime column
    combined = pd.to_datetime(
        df[date_col_resolved].astype(str) + ' ' + df[time_col_resolved].astype(str),
        errors='coerce'
    )
    # If a timezone is provided, localize the datetime column to that timezone
    if tz:
        combined = combined.dt.tz_localize(tz, ambiguous='NaT', nonexistent='shift_forward')

    # Check for any NaT values that may have resulted from invalid date/time combinations
    if combined.isnull().any():
        if drop_invalid:
            df = df.loc[~combined.isnull()].copy()
            combined = combined.loc[~combined.isnull()]
        else:
            raise ValueError("Invalid date/time combination found in the DataFrame.")
    
    df[datetime_col] = combined
    # Optionally, drop the original date and time columns
    if not keep_source:
        df.drop(columns=[date_col_resolved, time_col_resolved], inplace=True)
    # Update the original DataFrame with the combined datetime column
    
    return df

# this function extrapolates date ranges such as [6/3-17/4, 1/5-29/5]
# to a list of dates [6/3,13/3,20/3,27/3,3/4,10/4,17/4,1/5,8/5,15/5,22/5,29/5]
def extrapolate_date_range(date_range: str, year: int=None, format=r"%d/%m/%Y", frequency="7D") -> list[datetime]:
    """
    Extrapolate date ranges to a list of dates.
    """
    # Split the range (be robust to stray whitespace)
    start_str, end_str = map(str.strip, date_range.split('-'))
    start = append_year(start_str, year=year, format=format)
    end = append_year(end_str, year=year, format=format)
    
    # Generate and return the list of dates
    return pd.date_range(start=start, end=end, freq=frequency).tolist()

def extrapolate_date_ranges(date_ranges: str, delim=',', year: int=None, format=r"%d/%m/%Y", frequency="7D") -> list[datetime]:
    if date_ranges is None or (isinstance(date_ranges, float) and pd.isna(date_ranges)):
        return []

    all_dates = []
    for date_range in str(date_ranges).split(delim):
        date_range = date_range.strip()

        # Check if the date range is not empty
        if not date_range:
            continue
        # Check if single date not range
        if '-' not in date_range:
            date = append_year(date_range.strip(), year=year, format=format)
            all_dates.append(date)
        else:
            # Extrapolate the date range
            dates = extrapolate_date_range(date_range.strip(), year=year, format=format, frequency=frequency)
            all_dates.extend(dates)
    
    return all_dates


def append_year(date_str: str, year: int=None, format=r"%d/%m/%y") -> datetime:
    """
    Append a year to a date string.
    """
    if date_str is None or (isinstance(date_str, float) and pd.isna(date_str)):
        raise ValueError("Missing date string")

    date_str = str(date_str).strip()
    if not date_str:
        raise ValueError("Empty date string")
    try :
        date_obj = datetime.strptime(date_str, format)
    except ValueError:
        raise ValueError(f"Date string '{date_str}' does not match format '{format}'.")
    if year is not None:
        date_obj = date_obj.replace(year=year)
    return date_obj
    
    

# this function turns each row of a Dataframe with a list of dates into a row for each date
def expand_dates(self, dates_col: str , date_col: str=None, year=None, format=r"%d/%m/%Y", frequency="7D") -> pd.DataFrame:
    """
    Expand a DataFrame with a list of dates into separate rows for each date.
    """
    if date_col is None:
        date_col = dates_col
    df = self._obj.copy()

    dates_col_resolved = _resolve_column_name(df, dates_col)
    
    # extrapolate the date ranges
    df[date_col] = df[dates_col_resolved].apply(lambda x: extrapolate_date_ranges(x, year=year, format=format, frequency=frequency))
    # Explode the date column into separate rows
    # df[date_col] = df[date_col].apply(lambda x: pd.to_datetime(x).strftime(r"%d/%m/%Y"))
    exploded_df = df.explode(date_col)
    exploded_df.reset_index(drop=True, inplace=True)
    return exploded_df