# -*- coding: utf-8 -*-
"""
main.py - Excel to DataFrame
Created by NCagle
2024-07-27
      _
   __(.)<
~~~⋱___)~~~

<Description>
"""

# ~~~ Standard library imports ~~~
import sys
from pathlib import Path
import string
import random

# ~~~ Third-party library imports ~~~
from pprint import pprint
import pandas as pd
import geopandas as gpd


# Convert a column letter to an index
def col_to_idx(col: str, zero_idx: bool=True) -> int:
    """
    Convert an Excel-style column letter to an integer index. The columns can
    either be indexed at zero or one.
    Index at zero: 'A' -> 0
    Index at one: 'A' -> 1

    Args:
        col (str): The Excel-style column letter (e.g., "A", "Z", "AA").
        zero_idx (bool): If True, return a zero-based index; if False, return a one-based index.

    Returns:
        idx (int): The zero-based or one-based column index as an integer.
    """
    idx = 0
    for char in col:
        if char in string.ascii_letters:
            # Convert the column letter to a one-based index
            # 'A' -> 1, 'B' -> 2, ..., 'Z' -> 26, 'AA' -> 27, etc.
            # The formula (ord(char.upper()) - ord("A")) + 1 converts the character to its equivalent integer value
            # Multiplying current index by 26 accumulates the index for multi-letter columns
            idx = idx * 26 + (ord(char.upper()) - ord("A")) + 1

    if zero_idx:
        # Subtract 1 to convert to zero-based indexing for returned value
        idx -= 1

    return idx


# Convert an index to a column letter
def idx_to_col(idx: int, zero_idx: bool=True) -> str:
    """
    Convert an integer index to an Excel-style column letter. The columns can
    either be indexed at zero or one.
    Index at zero: 'A' -> 0
    Index at one: 'A' -> 1

    Args:
        idx (int): The zero-based or one-based column index.
        zero_idx (bool): If True, treat the index value as zero-based; if False, treat it as one-based.

    Returns:
        col (str): The Excel-style column letter.

    Equivalent logic without recursion:
        ```
        while idx >= 0:
            quo, rem = divmod(idx, 26)
            col = chr(rem + 65) + col
            idx = quo - 1
        ```

    Citation:
        Refactored by Nat Cagle on 2024-07-27
        Original code by Giancarlo Sportelli on 2016-06-03
        https://stackoverflow.com/a/37604105
        CC BY-SA 4.0
    """
    col = ""

    if not zero_idx:
        # Subtract 1 to convert to one-based indexing for calculation
        idx -= 1

    if idx < 0:
        # Return an empty string for negative indices (invalid input)
        return col

    # Calculate the quotient and remainder for the current character
    # Example: 27 -> 'AA': (27 / 26 = 1, 27 % 26 = 1) -> 'A' + 'A'
    quo, rem = divmod(idx, 26)  # 26 english ASCII letters

    # Recursive call to handle the conversion until quotient is 0
    # chr(rem + 65) converts 0 -> 'A', 1 -> 'B', etc.
    # Concatenating the unicode character accumulates multi-letter columns
    col = idx_to_col(quo - 1) + chr(rem + 65)  # chr(65) == "A"

    return col


# Load the selected data from the excel sheet and serialize it as a pickle
def load_and_pickle(range_address="$C$1:$D$5", active_sheet="Sheet1"):
    """
    _summary_

    Default parameter values are examples for testing and referenced by comments.
    """
    # Read the bounds of the sheet
    _, max_col = pd.read_excel("workbook.xlsm", sheet_name=active_sheet).shape

    # Parse the range_address into column and row start and end
    range_start, range_end = range_address.split(":")  # "$C$1", "$D$5"
    _, col_start, row_start = range_start.split("$")  # "", "C", "1"
    _, col_end, row_end = range_end.split("$")  # "", "D", "5"

    # Calculate the column indices
    # Pandas dataframes use zero-based indexing
    col_start_index = col_to_idx(col_start)  # 2
    col_end_index = col_to_idx(col_end)  # 3

    # Coerce the row numbers to integers and subtract 1 for zero-based indexing
    row_start = int(row_start) - 1  # 0
    row_end = int(row_end) - 1  # 4

    # Adjust the column indices if they are out of bounds
    # Row indices don't have a maximum bound for loading into a dataframe
    if col_start_index >= max_col:
        print("Selected starting column is out of bounds.")
        return False
    if col_end_index >= max_col:
        col_end_index = max_col - 1

    if col_start_index > col_end_index or row_start > row_end:
        print("Selected range is out of bounds.")
        return False

    # Construct the pandas `read_excel()` parameters to only load the selected range of data
    usecols = range(col_start_index, col_end_index + 1)  # range(2, 4)
    # skiprows = range(row_start - 1)  # range(0, 0)
    skiprows = range(row_start)  # range(0, 0)
    # nrows = row_end - row_start + 1  # 5
    nrows = row_end - row_start  # 4

    # Read the specified range from the sheet
    data_selection = pd.read_excel(
        "workbook.xlsm",
        sheet_name=active_sheet,
        usecols=usecols,
        skiprows=skiprows,
        nrows=nrows
    )

    print(data_selection)

    # Slice the DataFrame to get the selected range
    # data_selection = spreadsheet.iloc[row_start:row_end+1, col_start_index:col_end_index+1]

    data_selection.to_pickle("workbook_range_df.pkl")
    # data_selection.to_csv("workbook_range_df.txt")  # For quickly checking output

    return True


def main():
    if len(sys.argv) >= 3:
        range_address = sys.argv[1]
        print(f"Range Address: {range_address}")
        active_sheet = sys.argv[2]
        print(f"Active Sheet: {active_sheet}")

        with open("args.txt", "w", encoding="utf-8") as file:
            _ = file.write("__Last Script Call Arguments__\n\n")
            _ = file.write(f"Script Path ({type(sys.argv[0]).__name__}): {sys.argv[0]}\n")
            _ = file.write(f"Range Address ({type(range_address).__name__}): {range_address}\n")
            _ = file.write(f"Active Sheet ({type(active_sheet).__name__}): {active_sheet}\n")

            if len(sys.argv) > 3:
                _ = file.write("\n__Unexpected Arguments__\n\n")
                for idx, arg in enumerate(sys.argv):
                    if idx not in [0, 1, 2]:
                        _ = file.write(f"Argument {idx}: {arg}\n")

        try:
            # Load the selected data from the excel sheet and serialize it as a pickle
            export_success = load_and_pickle(range_address, active_sheet)
        except Exception as e:
            with open("ERROR EXPORTING DATAFRAME.txt", "w", encoding="utf-8") as file:
                _ = file.write("__ERROR EXPORTING DATAFRAME__\n\n")
                _ = file.write("An error occurred while trying to export the selected range "
                    + f"'{range_address}' on sheet '{active_sheet}' to a dataframe.\n\n")
                _ = file.write(f"Error:\n{e}\n")

        if not export_success:
            with open("ERROR EXPORTING DATAFRAME.txt", "w", encoding="utf-8") as file:
                _ = file.write("__ERROR EXPORTING DATAFRAME__\n\n")
                _ = file.write("An error occurred while trying to export the selected range "
                    + f"'{range_address}' on sheet '{active_sheet}' to a dataframe.\n\n")
                _ = file.write("This file can be deleted.\n")
                _ = file.write("Please check your selection and try again.\n")
    else:
        raise ValueError(f"Missing argument in execution command.\nArguments:\n{sys.argv}")


if __name__ == "__main__":
    main()

# py main.py "$C$1:$D$5" "Sheet1"
