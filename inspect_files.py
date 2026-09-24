import pandas as pd
import sys
import os

files = [
    "GRADE DE JULHO - 14ª VERSÃO.xlsm",
    "GRADE DE EVENTOS COMBATE 2026 - JULHO  (V4).xlsx",
    "PPV 2026 (Jul 10ª versão).xlsx"
]

for file in files:
    filepath = os.path.join("c:\\Users\\ligomes\\Downloads\\atualizacao_solution", file)
    print(f"--- Analyzing: {file} ---")
    try:
        xl = pd.ExcelFile(filepath)
        print(f"Sheet names: {xl.sheet_names}")
        
        for sheet in xl.sheet_names[:1]: # Check first sheet or relevant sheets
            df = xl.parse(sheet, nrows=10)
            print(f"\nSheet: {sheet}")
            print(f"Columns: {df.columns.tolist()}")
            print(df.head(5).to_string())
    except Exception as e:
        print(f"Error reading {file}: {e}")
    print("\n" + "="*50 + "\n")
