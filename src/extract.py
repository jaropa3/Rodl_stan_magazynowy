import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog, messagebox
import re

custom_sheet_names = {
    "642": "Ursus",
    "642 Dec2": "Ursus - dodane",
    "739": "Piaseczno",
    "739 Dec3": "Piaseczno - dodane",
    "1052": "Kraków",
    "1052 Dec4": "Kraków - dodane",
    "1053": "Wrocław",
    "1053 Dec4": "Wrocław - dodane",
    "1626": "Annopol",
    "1626 Dec3": "Annopol - dodane"
}

def read_input(file_path):
    all_sheets = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str)
    df_list = []

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        
        # jeśli mamy własną nazwę, użyj jej, inaczej weź nazwę z Excela
        custom_name = custom_sheet_names.get(sheet_name, sheet_name)
        df["Nazwa sklepu"] = custom_name
        
        df_list.append(df)

    combined_df = pd.concat(df_list, ignore_index=True)
    return combined_df
    

ROW_PATTERN = re.compile(
    r"""
    ^\s*
    (?P<sku>[A-Z0-9]+)
    \s+
    (?P<vendor>\d+)
    \s+
    (?P<desc>.*?)
    \s+
    (?P<count>\d+)
    \s+
    (?P<cost>[-\d.,]+)
    \s+
    (?P<ext_cost>[-\d.,]+)
    \s+
    (?P<retail>[-\d.,]+)
    \s+
    (?P<ext_retail>[-.\d,]+)
    \s*$
    """,
    re.VERBOSE,
)

ROW_PATTERN_Transactions = re.compile(r"""
    ^\s*
    (?P<Date>\d{1,2}/\d{1,2}/\d{2})      # Data mm/dd/yy
    \s+
    (?P<Item>\d+)                         # Item (numer)
    \s+
    (?P<Description>.+?)                  # Description - wszystko do CO
    \s+
    (?P<CO>[A-Z0-9]+(?:\s*-\s*)?)        # CO - jedno słowo lub znak
    \s+
    (?P<SIZ>\S+)                          # SIZ - jedno słowo
    \s+
    (?P<Qty>-?\d+)                          # Ilość
    \s+
    (?P<Cost>[-\d.,]+)                    # Cost
    \s+
    (?P<Price>[-\d.,]+)                   # Price
    \s+
    (?P<Batch>\d+)                        # Batch
    \s+
    (?P<Typ>\S+)                           # Typ - jedno słowo/znak
    \s+
    (?P<Ref>\d+)                           # Ref
    \s+
    (?P<Per_Year>\d+)                      # Per_Year
    \s+
    (?P<Year>\d{4})                        # Rok
    \s*$
""", re.VERBOSE)

def parse_money(x: str) -> float:
    if x is None:
        return 0.0

    x = str(x).strip()

    # brak wartości
    if x in {"", "-", "--", "NA", "N/A"}:
        return 0.0

    # księgowe minusy (123.45)
    neg = x.startswith("(") and x.endswith(")")
    x = x.strip("()")

    # ".99" -> "0.99"
    if x.startswith("."):
        x = "0" + x

    # usuń spacje
    x = x.replace(" ", "")

    # normalizacja separatorów
    if "," in x and "." in x and x.rfind(",") > x.rfind("."):
        # EU: 1.234,56
        x = x.replace(".", "").replace(",", ".")
    else:
        x = x.replace(",", "")

    try:
        val = float(x)
    except ValueError:
        raise ValueError(f"Bad numeric value: {x}")

    return -val if neg else val

def parse_report(path):
    df = read_input(path)  # teraz df to DataFrame z kolumną sheet_name
    records = []
    excel_cols = df.columns[:-1]  

    for _, row in df.iterrows():
        # 1. Połącz wszystkie kolumny Excela w jeden string do regexu
        #    pomijamy kolumnę 'sheet_name', bo ją zachowamy osobno
        line_str = " ".join(row[excel_cols].astype(str))

        # 2. Dopasuj regex
        m = ROW_PATTERN.match(line_str)
        if not m:
            continue

        g = m.groupdict()

        # 3. Dodaj rekord do listy
        records.append({
            "item": g["sku"],
            "vendor": int(g["vendor"]),
            "description": g["desc"],
            "count": int(g["count"]),
            "cost": parse_money(g["cost"]),
            "ext_cost": parse_money(g["ext_cost"]),
            "retail": parse_money(g["retail"]),
            "ext_retail": parse_money(g["ext_retail"]),
            "Nazwa sklepu": row[df.columns[-1]],
        })

    return pd.DataFrame(records)

def parse_report_transactions(path):
    df = read_input(path)  # teraz df to DataFrame z kolumną sheet_name
    records = []
    excel_cols = df.columns[:-1]  

    for _, row in df.iterrows():
        # 1. Połącz wszystkie kolumny Excela w jeden string do regexu
        #    pomijamy kolumnę 'sheet_name', bo ją zachowamy osobno
        line_str = " ".join(row[excel_cols].astype(str))

        # 2. Dopasuj regex
        m = ROW_PATTERN_Transactions.match(line_str)
        if not m:
            continue

        g = m.groupdict()

        records.append({
    "Date": g["Date"],
    "Item": int(g["Item"]),
    "Description": g["Description"],
    "CO": g["CO"],
    "SIZ": g["SIZ"],
    "Qty": int(g["Qty"]),
    "Cost": parse_money(g["Cost"]),
    "Price": parse_money(g["Price"]),
    "Batch": int(g["Batch"]),
    "Typ": g["Typ"],
    "Ref": int(g["Ref"]),
    "Per_Year": int(g["Per_Year"]),
    "Year": int(g["Year"]),
    "Nazwa sklepu": row["Nazwa sklepu"]
    })

    return pd.DataFrame(records)

def is_number(x: str) -> bool:
    x = x.replace(",", "").replace(".", "")
    return x.isdigit()

def parse_row(row: str):
    parts = re.split(r'\s+', row.strip())

    item = parts[0]
    vendor = parts[1]

    # opis = wszystko do pierwszej liczby (count)
    desc_parts = []
    i = 2
    while i < len(parts) and not is_number(parts[i]):
        desc_parts.append(parts[i])
        i += 1

    description = " ".join(desc_parts)

    # teraz liczby
    count = int(parts[i]); i += 1
    cost = parse_money(parts[i]); i += 1
    ext_cost = parse_money(parts[i]); i += 1
    retail = parse_money(parts[i]); i += 1
    ext_retail = parse_money(parts[i]); i += 1

    return [
        item,
        vendor,
        description,
        count,
        cost,
        ext_cost,
        retail,
        ext_retail
    ]