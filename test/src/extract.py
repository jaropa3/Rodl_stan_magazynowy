import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog, messagebox



def read_input(file_path):
    all_sheets = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)
    
    combined_df = pd.concat(all_sheets, ignore_index=True)
    return combined_df

