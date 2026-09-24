import pandas as pd

df = pd.read_excel('GRADE DE MAIO - 9ª VERSÃO.xlsm', nrows=5)
print("Columns:")
print(df.columns.tolist())
print("\nFirst 2 rows:")
print(df.head(2).to_dict('records'))
