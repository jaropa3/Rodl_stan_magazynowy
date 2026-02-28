import pandas as pd
import sys
from transform import time_complexity_to_df
from extract import read_input
from config import FILE_PATH
import cProfile
import pstats

def main():
    
    df_dane = read_input(FILE_PATH)
    df_dane = df_dane.replace(r"^\s*$", pd.NA, regex=True)
    df_dane = df_dane.dropna(how="all").reset_index(drop=True)

    col = df_dane.iloc[:, 0]

    filtered = col.str.strip()
    filtered = filtered[filtered.str.match(r"^[A-Z0-9]{1,3}\d{5,}")]

    print(filtered)
if __name__ == "__main__":

    main()