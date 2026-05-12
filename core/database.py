import sqlite3
import pandas as pd
from typing import List, Dict, Any, Tuple
import datetime
from config import DB_PATH, DEFAULT_CONFIG

class DatabaseManager:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        self._create_tables()
        self._init_config()

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def _create_tables(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # Tabela de Configurações
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS config (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            ''')
            
            # Tabela de Grades (Cabeçalho/Importação)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS grades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    import_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    filename TEXT,
                    source_email_id TEXT
                )
            ''')
            
            # Tabela de Itens da Grade
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS grade_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    grade_id INTEGER,
                    funcionario TEXT,
                    data TEXT,
                    horario TEXT,
                    evento TEXT,
                    sonora TEXT,
                    raw_data TEXT, -- JSON com todas as colunas para flexibilidade
                    FOREIGN KEY (grade_id) REFERENCES grades (id)
                )
            ''')
            
            # Tabela de Alterações (Histórico)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS changes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    grade_id INTEGER,
                    funcionario TEXT,
                    data_escala TEXT,
                    tipo TEXT, -- 'ENTRADA', 'SAIDA', 'ALTERACAO'
                    campo TEXT,
                    valor_antigo TEXT,
                    valor_novo TEXT,
                    data_deteccao DATETIME DEFAULT CURRENT_TIMESTAMP,
                    email_enviado INTEGER DEFAULT 0,
                    FOREIGN KEY (grade_id) REFERENCES grades (id)
                )
            ''')
            
            # Tabela de Mapeamento Funcionário -> Email
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS funcionarios (
                    nome TEXT PRIMARY KEY,
                    email TEXT
                )
            ''')
            conn.commit()

    def _init_config(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            for key, value in DEFAULT_CONFIG.items():
                cursor.execute('INSERT OR IGNORE INTO config (key, value) VALUES (?, ?)', (key, str(value)))
            conn.commit()

    def get_config(self) -> Dict[str, str]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT key, value FROM config')
            return {row[0]: row[1] for row in cursor.fetchall()}

    def update_config(self, key: str, value: Any):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE config SET value = ? WHERE key = ?', (str(value), key))
            conn.commit()

    def is_email_processed(self, email_id: str) -> bool:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT 1 FROM grades WHERE source_email_id = ?', (email_id,))
            return cursor.fetchone() is not None

    def save_new_grade(self, df: pd.DataFrame, filename: str, email_id: str) -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT INTO grades (filename, source_email_id) VALUES (?, ?)', (filename, email_id))
            grade_id = cursor.lastrowid
            
            # Prepare data for insertion
            items = []
            for _, row in df.iterrows():
                funcionario = row.get('funcionario', '')
                data = str(row.get('data', ''))
                horario = str(row.get('horario', ''))
                evento = str(row.get('evento', ''))
                sonora = str(row.get('sonora', ''))
                raw_data = row.to_json()
                
                items.append((grade_id, funcionario, data, horario, evento, sonora, raw_data))
                
            cursor.executemany('''
                INSERT INTO grade_items (grade_id, funcionario, data, horario, evento, sonora, raw_data)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', items)
            conn.commit()
            return grade_id

    def save_changes(self, changes: List[Dict[str, Any]]):
        if not changes:
            return
            
        with self.get_connection() as conn:
            cursor = conn.cursor()
            records = [
                (
                    c['grade_id'], c['funcionario'], c.get('data', ''), c['tipo'], 
                    c.get('campo', ''), str(c.get('valor_antigo', '')), str(c.get('valor_novo', ''))
                ) for c in changes
            ]
            cursor.executemany('''
                INSERT INTO changes (grade_id, funcionario, data_escala, tipo, campo, valor_antigo, valor_novo)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', records)
            conn.commit()

    def get_last_grade_df(self) -> pd.DataFrame:
        with self.get_connection() as conn:
            query = "SELECT id FROM grades ORDER BY id DESC LIMIT 1"
            last_grade = pd.read_sql_query(query, conn)
            
            if last_grade.empty:
                return pd.DataFrame()
                
            grade_id = last_grade.iloc[0]['id']
            query_items = f"SELECT * FROM grade_items WHERE grade_id = {grade_id}"
            df = pd.read_sql_query(query_items, conn)
            
            # Reconstruct original DataFrame from raw_data if needed, but for diff we need it
            if not df.empty:
                df_reconstructed = pd.read_json(df['raw_data'].to_json(orient='records'))
                # Actually raw_data is a string of JSON in each row
                df_expanded = df['raw_data'].apply(pd.read_json, typ='series')
                return df_expanded
            return pd.DataFrame()

    def get_employee_email(self, name: str) -> str:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT email FROM funcionarios WHERE nome = ?', (name,))
            result = cursor.fetchone()
            return result[0] if result else None
            
    def update_employee_emails(self, mapping_dict: Dict[str, str]):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            records = [(name, email) for name, email in mapping_dict.items()]
            cursor.executemany('''
                INSERT INTO funcionarios (nome, email) 
                VALUES (?, ?)
                ON CONFLICT(nome) DO UPDATE SET email=excluded.email
            ''', records)
            conn.commit()

    def get_recent_changes(self, limit: int = 100) -> pd.DataFrame:
            with self.get_connection() as conn:
                query = "SELECT * FROM changes ORDER BY data_deteccao DESC LIMIT ?"
                return pd.read_sql_query(query, conn, params=(limit,))
                
    def get_dashboard_stats(self) -> Dict[str, Any]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            stats = {}
            # Alterações na semana (considerando últimos 7 dias para simplificar SQLite)
            cursor.execute("SELECT COUNT(*) FROM changes WHERE data_deteccao >= date('now', '-7 days')")
            stats['alteracoes_semana'] = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(DISTINCT funcionario) FROM changes WHERE data_deteccao >= date('now', '-7 days')")
            stats['funcionarios_afetados'] = cursor.fetchone()[0]
            
            cursor.execute("SELECT MAX(import_date) FROM grades")
            stats['ultima_execucao'] = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM changes WHERE email_enviado = 1")
            stats['emails_enviados'] = cursor.fetchone()[0]
            
            return stats
            
    def mark_email_sent(self, change_ids: List[int]):
         with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.executemany("UPDATE changes SET email_enviado = 1 WHERE id = ?", [(c_id,) for c_id in change_ids])
            conn.commit()
