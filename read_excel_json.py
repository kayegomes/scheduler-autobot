import pandas as pd
import json

df = pd.read_excel('GRADE DE MAIO - 9ª VERSÃO.xlsm', nrows=5)
data = {
    'columns': df.columns.tolist(),
    'first_rows': df.head(2).to_dict('records')
}
with open('output.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
