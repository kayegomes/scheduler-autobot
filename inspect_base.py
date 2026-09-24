import pandas as pd
import sys

filepath = r"C:\Users\ligomes\Downloads\atualizacao_solution\Check_Pre_Envio_Gerado.xlsx"

try:
    df = pd.read_excel(filepath, nrows=5)
    print("Columns in Check_Pre_Envio_Gerado.xlsx:")
    print(df.columns.tolist())
    print("\nFirst row:")
    print(df.head(1).to_dict('records'))
except Exception as e:
    print(f"Error reading file: {e}")
