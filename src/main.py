import pandas as pd
import sys
from transform import time_complexity_to_df
from extract import read_input, filter_data, parse_row
from config import FILE_PATH
from openpyxl import load_workbook
import cProfile
import pstats
import re

# wb = load_workbook(FILE_PATH)
# ws = wb.active

# for merged in list(ws.merged_cells.ranges):
#     min_col, min_row, max_col, max_row = merged.bounds
#     value = ws.cell(min_row, min_col).value
#     ws.unmerge_cells(str(merged))

#     # wypełnij wszystkie komórki
#     for r in range(min_row, max_row + 1):
#         for c in range(min_col, max_col + 1):
#             ws.cell(r, c).value = value

# wb.save("normalized.xlsx")

# df = pd.read_excel("normalized.xlsx")

def main():
    
    df_dane = read_input(FILE_PATH)
    df_dane = filter_data(df_dane)
    #pattern = re.compile(r'^\d+[A-Z]+\s+\d+\s+.+\s+\d+\s+\d+\.\d+\s+\d+\.\d+')
    df_dane = pd.DataFrame([parse_row(r) for r in df_dane])
    #rows = [l for l in lines if pattern.match(l)]
    print(df_dane)   
if __name__ == "__main__":

    main()