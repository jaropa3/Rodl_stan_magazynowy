import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook
import re

def read_input(file_path):
    df = pd.read_excel(file_path, header=None, dtype=str)
    return (
        df.fillna("")
        .agg(" ".join, axis=1)
        .str.strip()
        .tolist()
    )

def filter_data(lines):
    pat = re.compile(r'^\d+[A-Z]+\s+\d+')
    return [l for l in lines if pat.match(l)]
    #df = df.ffill()
    
def parse_report(path):
    lines = read_input(path)
    rows = filter_data(lines)
    return pd.DataFrame([read_input(r) for r in rows])

def parse_row(row: str):
    parts = re.split(r'\s+', row)

    item = parts[0]
    vendor = parts[1]

    # opis = wszystko do momentu liczby count
    i = 2
    desc_parts = []
    while not parts[i].isdigit():
        desc_parts.append(parts[i])
        i += 1

    description = " ".join(desc_parts)

    count = int(parts[i]); i += 1
    cost = float(parts[i]); i += 1
    ext_cost = float(parts[i]); i += 1
    retail = float(parts[i]); i += 1
    ext_retail = float(parts[i].replace(".", "0."))  # fix ".00"

    return [item, vendor, description, count, cost, ext_cost, retail, ext_retail]