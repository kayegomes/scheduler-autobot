import logging
from typing import Any, Dict, List

import pandas as pd

from core.columns import normalizar_colunas, normalizar_valores

logger = logging.getLogger(__name__)

# Colunas de controle que nunca devem virar "alteração" num e-mail.
COLUNAS_IGNORADAS = {
    'temp_id', 'occurrence', 'merge_key', 'funcionario', 'data',
    'data_raw', 'id', 'grade_id', 'row display', 'parent #',
}


def standardize_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza colunas/valores e descarta linhas sem funcionário."""
    if df is None or df.empty:
        return pd.DataFrame()

    df = normalizar_valores(normalizar_colunas(df))

    for col in ('funcionario', 'data'):
        if col not in df.columns:
            df[col] = ''

    df['funcionario'] = df['funcionario'].astype(str).str.strip()
    df = df[~df['funcionario'].str.lower().isin(['', 'nan', 'none'])]
    df['data'] = df['data'].astype(str).str.strip()

    return df.fillna('')


def _chave(df: pd.DataFrame) -> pd.Series:
    """Identidade de uma linha: funcionário + data + evento.

    Incluir o evento evita que reordenar a planilha (ou inserir uma linha no
    meio) desalinhe as ocorrências e gere alterações que não existiram.
    """
    evento = df['evento'].astype(str) if 'evento' in df.columns else pd.Series([''] * len(df), index=df.index)
    return df['funcionario'].astype(str) + '|' + df['data'].astype(str) + '|' + evento


def compare_schedules(df_old: pd.DataFrame, df_new: pd.DataFrame, grade_id: int) -> List[Dict[str, Any]]:
    df_old = standardize_df(df_old)
    df_new = standardize_df(df_new)

    changes: List[Dict[str, Any]] = []

    if df_new.empty:
        logger.warning("Escala nova vazia após normalização. Nenhuma comparação feita.")
        return changes

    if df_old.empty:
        for _, row in df_new.iterrows():
            changes.append({
                'grade_id': grade_id,
                'funcionario': row['funcionario'],
                'data': row.get('data', ''),
                'tipo': 'ENTRADA',
                'campo': '*',
                'valor_antigo': '',
                'valor_novo': 'Adicionado à escala',
            })
        return changes

    df_old = df_old.copy()
    df_new = df_new.copy()

    df_old['temp_id'] = _chave(df_old)
    df_new['temp_id'] = _chave(df_new)

    df_old['occurrence'] = df_old.groupby('temp_id').cumcount()
    df_new['occurrence'] = df_new.groupby('temp_id').cumcount()

    df_old['merge_key'] = df_old['temp_id'] + '_' + df_old['occurrence'].astype(str)
    df_new['merge_key'] = df_new['temp_id'] + '_' + df_new['occurrence'].astype(str)

    # Compara as colunas presentes nos DOIS lados. Antes só as de df_new eram
    # consideradas, então uma coluna removida da planilha nunca era reportada.
    cols_to_compare = sorted(
        (set(df_new.columns) & set(df_old.columns)) - COLUNAS_IGNORADAS
    )

    merged = pd.merge(df_old, df_new, on='merge_key', how='outer',
                      suffixes=('_old', '_new'), indicator=True)

    for _, row in merged.iterrows():
        func = row.get('funcionario_new') if pd.notna(row.get('funcionario_new')) else row.get('funcionario_old')
        data = row.get('data_new') if pd.notna(row.get('data_new')) else row.get('data_old')

        if pd.isna(func) or str(func).strip() == '':
            continue

        if row['_merge'] == 'left_only':
            changes.append({
                'grade_id': grade_id,
                'funcionario': func,
                'data': data,
                'tipo': 'SAIDA',
                'campo': _evento_da_linha(row, '_old'),
                'valor_antigo': 'Na escala',
                'valor_novo': 'Removido da escala',
            })
        elif row['_merge'] == 'right_only':
            changes.append({
                'grade_id': grade_id,
                'funcionario': func,
                'data': data,
                'tipo': 'ENTRADA',
                'campo': _evento_da_linha(row, '_new'),
                'valor_antigo': '',
                'valor_novo': 'Adicionado à escala',
            })
        elif row['_merge'] == 'both':
            for col in cols_to_compare:
                col_old, col_new = f"{col}_old", f"{col}_new"
                if col_old not in row or col_new not in row:
                    continue

                val_old = '' if pd.isna(row[col_old]) else str(row[col_old]).strip()
                val_new = '' if pd.isna(row[col_new]) else str(row[col_new]).strip()

                if val_old == val_new:
                    continue

                if val_new.upper() == 'CANCELADO':
                    changes.append({
                        'grade_id': grade_id,
                        'funcionario': func,
                        'data': data,
                        'tipo': 'SAIDA',
                        'campo': _evento_da_linha(row, '_old') or col.title(),
                        'valor_antigo': val_old,
                        'valor_novo': val_new,
                    })
                else:
                    changes.append({
                        'grade_id': grade_id,
                        'funcionario': func,
                        'data': data,
                        'tipo': 'ALTERACAO',
                        'campo': col.title(),
                        'valor_antigo': val_old,
                        'valor_novo': val_new,
                    })

    return changes


def _evento_da_linha(row: pd.Series, sufixo: str) -> str:
    """Nome do evento da linha, usado como rótulo humano da alteração."""
    for chave in (f'evento{sufixo}', 'evento'):
        if chave in row and pd.notna(row[chave]):
            valor = str(row[chave]).strip()
            if valor and valor.lower() not in ('nan', 'none'):
                return valor
    return ''
