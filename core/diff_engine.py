import pandas as pd
import numpy as np
from typing import List, Dict, Any

def standardize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    
    df = df.copy()
    # Normaliza colunas
    df.columns = df.columns.str.strip().str.lower()
    
    required_cols = ['funcionario', 'data']
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''
            
    # Remove linhas onde funcionário é vazio
    df['funcionario'] = df['funcionario'].astype(str).str.strip()
    df = df[df['funcionario'] != '']
    df['funcionario'] = df['funcionario'].replace('nan', '')
    df = df[df['funcionario'] != '']
    
    df['data'] = df['data'].astype(str).str.strip()
    df = df.fillna('')
    
    # Remove duplicatas se houver exatamente a mesma pessoa e data (mantém ultima)
    # df = df.drop_duplicates(subset=['funcionario', 'data'], keep='last')
    
    return df

def compare_schedules(df_old: pd.DataFrame, df_new: pd.DataFrame, grade_id: int) -> List[Dict[str, Any]]:
    df_old = standardize_df(df_old)
    df_new = standardize_df(df_new)
    
    changes = []
    
    if df_old.empty:
        # Se não há escala anterior, vamos considerar entrada na escala
        for _, row in df_new.iterrows():
            changes.append({
                'grade_id': grade_id,
                'funcionario': row['funcionario'],
                'data': row.get('data', ''),
                'tipo': 'ENTRADA',
                'campo': '*',
                'valor_antigo': '',
                'valor_novo': 'Adicionado à escala'
            })
        return changes

    # Adiciona um ID unico temporário para faciliar merge caso haja multiplas entradas para func+data
    df_old['temp_id'] = df_old['funcionario'] + '_' + df_old['data']
    df_new['temp_id'] = df_new['funcionario'] + '_' + df_new['data']
    
    # Merge com outer join para achar exclusividades e comunalidades
    # Para lidar com duplicações, deixamos de indexar mas avisamos. Ideal é deduplicar ou numerar itens.
    df_old['occurrence'] = df_old.groupby('temp_id').cumcount()
    df_new['occurrence'] = df_new.groupby('temp_id').cumcount()
    
    df_old['merge_key'] = df_old['temp_id'] + '_' + df_old['occurrence'].astype(str)
    df_new['merge_key'] = df_new['temp_id'] + '_' + df_new['occurrence'].astype(str)
    
    # Identificar colunas comparáveis
    cols_to_compare = [c for c in df_new.columns if c not in ['temp_id', 'occurrence', 'merge_key', 'funcionario', 'data']]
    
    merged = pd.merge(df_old, df_new, on='merge_key', how='outer', suffixes=('_old', '_new'), indicator=True)
    
    for _, row in merged.iterrows():
        func = row.get('funcionario_new') if pd.notna(row.get('funcionario_new')) else row.get('funcionario_old')
        data = row.get('data_new') if pd.notna(row.get('data_new')) else row.get('data_old')
        
        # Filtra nans
        if pd.isna(func) or func == '':
            continue
            
        if row['_merge'] == 'left_only':
            changes.append({
                'grade_id': grade_id,
                'funcionario': func,
                'data': data,
                'tipo': 'SAIDA',
                'campo': '*',
                'valor_antigo': 'Na escala',
                'valor_novo': 'Removido da escala'
            })
        elif row['_merge'] == 'right_only':
            changes.append({
                'grade_id': grade_id,
                'funcionario': func,
                'data': data,
                'tipo': 'ENTRADA',
                'campo': '*',
                'valor_antigo': '',
                'valor_novo': 'Adicionado à escala'
            })
        elif row['_merge'] == 'both':
            # Verifica alteração em cada coluna
            for col in cols_to_compare:
                col_old = f"{col}_old"
                col_new = f"{col}_new"
                
                # if exists in both
                if col_old in row and col_new in row:
                    val_old = str(row[col_old]) if pd.notna(row[col_old]) else ''
                    val_new = str(row[col_new]) if pd.notna(row[col_new]) else ''
                    
                    if val_old != val_new:
                        if val_new == 'CANCELADO':
                            # Encontra a coluna de evento/programa real
                            ev_col = next((c for c in row.index if c.endswith('_old') and ('evento' in c.lower() or 'programa' in c.lower() or 'descri' in c.lower())), None)
                            ev_name = str(row.get(ev_col, 'Evento')) if ev_col else 'Evento'
                            
                            changes.append({
                                'grade_id': grade_id,
                                'funcionario': func,
                                'data': data,
                                'tipo': 'SAIDA',
                                'campo': ev_name,
                                'valor_antigo': val_old,
                                'valor_novo': val_new
                            })
                        else:
                            changes.append({
                                'grade_id': grade_id,
                                'funcionario': func,
                                'data': data,
                                'tipo': 'ALTERACAO',
                                'campo': col.title(),
                                'valor_antigo': val_old,
                                'valor_novo': val_new
                            })
                        
    return changes
