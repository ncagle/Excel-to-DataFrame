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
import random

# ~~~ Third-party library imports ~~~
import pandas as pd
import geopandas as gpd

# Setting module options
# warnings.filterwarnings("once")  # "ignore"
# pd.set_option("display.max_columns", None)
# pd.set_option("display.float_format", lambda x: f"{x:,.3f}")

# <Description>
def func():
    pass


def main():
    spreadsheet = pd.read_excel("workbook.xlsm")
    print(spreadsheet)
    pd.to_pickle(spreadsheet, "workbook_range_df.pkl")


if __name__ == "__main__":
    main()
