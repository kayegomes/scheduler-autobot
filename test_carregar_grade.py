import os
import sys

sys.path.append(os.path.join(os.getcwd(), 'core'))

from core.match_eventos import carregar_grade

files = [
    "GRADE DE JULHO - 14ª VERSÃO.xlsm",
    "GRADE DE EVENTOS COMBATE 2026 - JULHO  (V4).xlsx",
    "PPV 2026 (Jul 10ª versão).xlsx"
]

for f in files:
    filepath = os.path.join(os.getcwd(), f)
    print(f"\n--- Testing carregar_grade on {f} ---")
    df = carregar_grade(filepath)
    if df is not None:
        print("Columns:", df.columns.tolist())
        print("First row:", df.head(1).to_dict('records'))
        print("Shapes:", df.shape)
    else:
        print("carregar_grade returned None")
