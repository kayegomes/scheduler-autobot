import datetime
import json
import logging
import os
import sqlite3
import unicodedata
from contextlib import contextmanager
from typing import Any, Dict, List, Optional

import pandas as pd

from config import DB_PATH, DEFAULT_CONFIG
from core.columns import preparar_escala

logger = logging.getLogger(__name__)

# Versão do schema. Incrementar sempre que uma migração for adicionada.
SCHEMA_VERSION = 3

# Corte do match aproximado de nomes. Acima de 0.85 (valor anterior) pessoas
# diferentes colidiam: "marcos antonio sa" x "marcio antonio sa" dá 0.94.
FUZZY_CUTOFF = 0.92


def _chave_nome(nome: Any) -> str:
    """Nome sem acento, minúsculo e com espaços colapsados.

    É o que resolve de verdade o caso do README (nome escrito sem acento):
    vira match EXATO, sem depender de heurística.
    """
    s = "".join(
        ch for ch in unicodedata.normalize("NFKD", str(nome)) if not unicodedata.combining(ch)
    )
    return " ".join(s.lower().split())


class DatabaseManager:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        self._create_tables()
        self._init_config()
        self._migrar()

    @contextmanager
    def get_connection(self):
        """Conexão com commit/rollback automáticos e — o que faltava — close().

        O `with sqlite3.connect(...)` original só controla a transação; a
        conexão continuava aberta até o GC recolher.
        """
        conn = sqlite3.connect(self.db_path, timeout=30)
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _create_tables(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()

            cursor.execute("PRAGMA journal_mode=WAL")

            # Tabela de Configurações
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS config (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            ''')

            # Tabela de Grades (Cabeçalho/Importação)
            # status: 'pending' enquanto o ciclo não terminou, 'done' depois de
            # notificar. Só grades 'done' servem de base de comparação.
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS grades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    import_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                    filename TEXT,
                    source_email_id TEXT,
                    status TEXT DEFAULT 'done'
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
                    data_deteccao DATETIME,
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

            cursor.execute('CREATE INDEX IF NOT EXISTS idx_items_grade ON grade_items (grade_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_changes_grade ON changes (grade_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_grades_email ON grades (source_email_id)')

            # Colunas acrescentadas depois da v1
            colunas = {r[1] for r in cursor.execute("PRAGMA table_info(grades)").fetchall()}
            if 'status' not in colunas:
                # Registros que já existiam foram processados por completo.
                cursor.execute("ALTER TABLE grades ADD COLUMN status TEXT DEFAULT 'done'")
                logger.info("Migração: coluna 'status' adicionada em grades.")
            if 'baseline_id' not in colunas:
                # Qual escala serviu de base naquele processamento. Permite
                # reavaliar um arquivo recusado quando a base muda.
                cursor.execute("ALTER TABLE grades ADD COLUMN baseline_id INTEGER")
                logger.info("Migração: coluna 'baseline_id' adicionada em grades.")

            conn.commit()

    def _init_config(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            for key, value in DEFAULT_CONFIG.items():
                cursor.execute('INSERT OR IGNORE INTO config (key, value) VALUES (?, ?)', (key, str(value)))
            conn.commit()

    # ------------------------------------------------------------------
    # Migrações
    # ------------------------------------------------------------------

    def _migrar(self):
        versao = int(self.get_config().get('schema_version', 1) or 1)
        if versao >= SCHEMA_VERSION:
            return

        logger.info(f"Migrando banco da versão {versao} para {SCHEMA_VERSION}...")
        self._backup(versao)
        if versao < 2:
            self._migrar_raw_data()
        self.update_config('schema_version', SCHEMA_VERSION)
        logger.info("Migração concluída.")

    def _backup(self, versao_origem: int):
        """Cópia do banco antes de qualquer migração que reescreva dados."""
        import shutil

        if not os.path.exists(self.db_path):
            return
        destino = (
            f"{self.db_path}.v{versao_origem}."
            f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.bak"
        )
        try:
            shutil.copy2(self.db_path, destino)
            logger.info(f"Backup do banco criado em: {destino}")
        except OSError as e:
            logger.warning(f"Não foi possível criar backup antes da migração: {e}")

    def _migrar_raw_data(self):
        """Renormaliza o `raw_data` das grades já existentes.

        As grades antigas foram gravadas com `Timestamp.to_json()`, que grava
        data como epoch em ms. Sem esta migração, a primeira comparação depois
        da correção veria "1782864000000" de um lado e "01/07/2026" do outro e
        dispararia SAIDA+ENTRADA para a escala inteira.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            grade_ids = [r[0] for r in cursor.execute("SELECT id FROM grades").fetchall()]

            total = 0
            for grade_id in grade_ids:
                linhas = cursor.execute(
                    "SELECT id, raw_data FROM grade_items WHERE grade_id = ?", (grade_id,)
                ).fetchall()
                if not linhas:
                    continue

                registros = []
                ids = []
                for item_id, raw in linhas:
                    try:
                        registros.append(json.loads(raw) if raw else {})
                        ids.append(item_id)
                    except (ValueError, TypeError):
                        continue

                if not registros:
                    continue

                df = preparar_escala(pd.DataFrame(registros))
                atualizacoes = [
                    (df.iloc[i].to_json(force_ascii=False), ids[i]) for i in range(len(ids))
                ]
                cursor.executemany("UPDATE grade_items SET raw_data = ? WHERE id = ?", atualizacoes)
                total += len(atualizacoes)

            conn.commit()
            if total:
                logger.info(f"Migração: {total} itens de grade renormalizados.")

    # ------------------------------------------------------------------
    # Configuração
    # ------------------------------------------------------------------

    def get_config(self) -> Dict[str, str]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT key, value FROM config')
            return {row[0]: row[1] for row in cursor.fetchall()}

    def update_config(self, key: str, value: Any):
        """Grava a configuração criando a chave se ela ainda não existir.

        O UPDATE puro anterior descartava em silêncio qualquer chave fora do
        DEFAULT_CONFIG — era por isso que "Termos Imunes" e "Última
        Verificação" nunca salvavam.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO config (key, value) VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
            ''', (key, str(value)))
            conn.commit()

    # ------------------------------------------------------------------
    # Grades
    # ------------------------------------------------------------------

    def get_baseline_id(self) -> Optional[int]:
        """Id da última grade concluída — a que serve de base de comparação."""
        with self.get_connection() as conn:
            row = conn.execute(
                "SELECT id FROM grades WHERE status = 'done' ORDER BY id DESC LIMIT 1"
            ).fetchone()
            return row[0] if row else None

    def is_email_processed(self, email_id: str, modo_teste: bool = False) -> bool:
        """True para e-mails já concluídos ou recusados contra a base atual.

        Um arquivo recusado (grade de outro período, por exemplo) não deve ser
        rebaixado e reavaliado a cada ciclo. Mas a recusa vale só enquanto a
        base for a mesma: se a escala vigente mudar, o arquivo volta à fila,
        porque aí ele pode passar a fazer sentido.

        Execuções de homologação ('test') contam apenas enquanto o Modo Teste
        estiver ligado — evitam gerar um relatório igual a cada ciclo, mas não
        podem fazer o e-mail ser pulado quando o envio real for ativado.
        """
        baseline = self.get_baseline_id()
        baseline = baseline if baseline is not None else -1
        estados = ["'ignored'"] + (["'test'"] if modo_teste else [])

        with self.get_connection() as conn:
            row = conn.execute(
                f"""SELECT 1 FROM grades
                    WHERE source_email_id = ?
                      AND (status = 'done'
                           OR (status IN ({','.join(estados)})
                               AND IFNULL(baseline_id, -1) = ?))
                    LIMIT 1""",
                (email_id, baseline)
            ).fetchone()
            return row is not None

    def registrar_execucao_teste(self, email_id: str, filename: str,
                                 baseline_id: Optional[int], pasta: str = '') -> int:
        """Anota um ciclo de homologação sem gravar itens nem alterações.

        A linha fica com status 'test': `get_last_grade_df()` só enxerga
        'done', então uma homologação nunca vira base de comparação.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """INSERT INTO grades (import_date, filename, source_email_id, status, baseline_id)
                   VALUES (?, ?, ?, 'test', ?)""",
                (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                 f"{filename} [TESTE: {pasta}]" if pasta else f"{filename} [TESTE]",
                 email_id, baseline_id)
            )
            conn.commit()
            return cursor.lastrowid

    def registrar_email_ignorado(self, email_id: str, filename: str, motivo: str,
                                 baseline_id: Optional[int]):
        """Marca um anexo como recusado para a base atual, sem criar itens."""
        with self.get_connection() as conn:
            conn.execute(
                """INSERT INTO grades (import_date, filename, source_email_id, status, baseline_id)
                   VALUES (?, ?, ?, 'ignored', ?)""",
                (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                 f"{filename} [IGNORADO: {motivo}]", email_id, baseline_id)
            )
            conn.commit()

    def descartar_grades_pendentes(self, email_id: str):
        """Remove grades incompletas de uma tentativa anterior do mesmo e-mail."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            ids = [r[0] for r in cursor.execute(
                "SELECT id FROM grades WHERE source_email_id = ? AND status = 'pending'",
                (email_id,)
            ).fetchall()]
            for gid in ids:
                cursor.execute("DELETE FROM grade_items WHERE grade_id = ?", (gid,))
                cursor.execute("DELETE FROM changes WHERE grade_id = ?", (gid,))
                cursor.execute("DELETE FROM grades WHERE id = ?", (gid,))
            conn.commit()
            if ids:
                logger.info(f"Descartadas {len(ids)} grade(s) pendente(s) de tentativa anterior.")

    def save_new_grade(self, df: pd.DataFrame, filename: str, email_id: str,
                       status: str = 'done', baseline_id: Optional[int] = None) -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            local_now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute(
                'INSERT INTO grades (import_date, filename, source_email_id, status, baseline_id) '
                'VALUES (?, ?, ?, ?, ?)',
                (local_now, filename, email_id, status, baseline_id)
            )
            grade_id = cursor.lastrowid

            # Normaliza nomes E valores: é o que garante que reler do banco
            # devolva exatamente o que foi gravado.
            df_norm = preparar_escala(df)

            items = []
            for _, row in df_norm.iterrows():
                items.append((
                    grade_id,
                    str(row.get('funcionario', '')),
                    str(row.get('data', '')),
                    str(row.get('inicio', row.get('horario', ''))),
                    str(row.get('evento', '')),
                    str(row.get('sonora', '')),
                    row.to_json(force_ascii=False),
                ))

            cursor.executemany('''
                INSERT INTO grade_items (grade_id, funcionario, data, horario, evento, sonora, raw_data)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', items)
            conn.commit()
            return grade_id

    def marcar_grade_concluida(self, grade_id: int):
        with self.get_connection() as conn:
            conn.execute("UPDATE grades SET status = 'done' WHERE id = ?", (grade_id,))
            conn.commit()

    def get_last_grade_df(self) -> pd.DataFrame:
        """Última escala COMPLETAMENTE processada — a base de comparação."""
        with self.get_connection() as conn:
            last_grade = pd.read_sql_query(
                "SELECT id FROM grades WHERE status = 'done' ORDER BY id DESC LIMIT 1", conn
            )
            if last_grade.empty:
                return pd.DataFrame()

            grade_id = int(last_grade.iloc[0]['id'])
            df = pd.read_sql_query(
                "SELECT raw_data FROM grade_items WHERE grade_id = ? ORDER BY id", conn,
                params=(grade_id,)
            )
            if df.empty:
                return pd.DataFrame()

            registros = []
            for raw in df['raw_data']:
                try:
                    registros.append(json.loads(raw) if raw else {})
                except (ValueError, TypeError):
                    registros.append({})

            return preparar_escala(pd.DataFrame(registros))

    # ------------------------------------------------------------------
    # Alterações
    # ------------------------------------------------------------------

    def save_changes(self, changes: List[Dict[str, Any]]) -> List[int]:
        """Grava as alterações e devolve os IDs criados, na mesma ordem."""
        if not changes:
            return []

        with self.get_connection() as conn:
            cursor = conn.cursor()
            # Hora local, para bater com import_date e last_check_time. O
            # CURRENT_TIMESTAMP do SQLite é UTC e deixava o histórico 3h adiantado.
            agora = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ids = []
            for c in changes:
                cursor.execute('''
                    INSERT INTO changes (grade_id, funcionario, data_escala, tipo, campo,
                                         valor_antigo, valor_novo, data_deteccao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    c['grade_id'], c['funcionario'], c.get('data', ''), c['tipo'],
                    c.get('campo', ''), str(c.get('valor_antigo', '')),
                    str(c.get('valor_novo', '')), agora,
                ))
                ids.append(cursor.lastrowid)
            conn.commit()
            return ids

    def mark_email_sent(self, change_ids: List[int]):
        if not change_ids:
            return
        with self.get_connection() as conn:
            conn.executemany(
                "UPDATE changes SET email_enviado = 1 WHERE id = ?",
                [(int(c_id),) for c_id in change_ids]
            )
            conn.commit()

    def get_recent_changes(self, limit: int = 100) -> pd.DataFrame:
        with self.get_connection() as conn:
            return pd.read_sql_query(
                "SELECT * FROM changes ORDER BY data_deteccao DESC, id DESC LIMIT ?",
                conn, params=(limit,)
            )

    # ------------------------------------------------------------------
    # Funcionários
    # ------------------------------------------------------------------

    def get_employee_email(self, name: str) -> Optional[str]:
        """E-mail do funcionário. Match exato sem acento; fuzzy só com trava.

        O match aproximado antigo (cutoff 0.85, sem validação) podia mandar a
        escala de uma pessoa para outra: "Marcos Antonio Sá" casava com
        "Marcio Antonio Sá" a 0.94. Agora exige primeiro E último nome iguais,
        e sempre registra em log quando não foi exato.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT nome, email FROM funcionarios')
            rows = cursor.fetchall()

        if not rows:
            return None

        mapa = {_chave_nome(r[0]): r[1] for r in rows if r[0]}
        alvo = _chave_nome(name)

        if alvo in mapa:
            return mapa[alvo]

        from difflib import get_close_matches
        candidatos = get_close_matches(alvo, list(mapa.keys()), n=1, cutoff=FUZZY_CUTOFF)
        if not candidatos:
            return None

        candidato = candidatos[0]
        tokens_alvo, tokens_cand = alvo.split(), candidato.split()
        if (
            tokens_alvo and tokens_cand
            and tokens_alvo[0] == tokens_cand[0]
            and tokens_alvo[-1] == tokens_cand[-1]
        ):
            logger.warning(
                f"Match APROXIMADO de funcionário: '{name}' -> '{candidato}' "
                f"({mapa[candidato]}). Confira o mapeamento de e-mails."
            )
            return mapa[candidato]

        logger.warning(
            f"Match aproximado DESCARTADO por segurança: '{name}' parecia '{candidato}', "
            f"mas primeiro/último nome divergem. Nenhum e-mail enviado."
        )
        return None

    def get_employee_mapping(self) -> Dict[str, str]:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nome, email FROM funcionarios ORDER BY nome")
            return {row[0]: row[1] for row in cursor.fetchall()}

    def update_employee_emails(self, mapping_dict: Dict[str, str], substituir: bool = True):
        """Grava o mapeamento. Com `substituir`, o que sumiu da tela é removido.

        Antes só havia INSERT/UPDATE: apagar uma linha na interface não
        desfazia o vínculo, e a pessoa continuava recebendo e-mail.
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()
            if substituir:
                manter = list(mapping_dict.keys())
                if manter:
                    marcadores = ",".join("?" * len(manter))
                    cursor.execute(f"DELETE FROM funcionarios WHERE nome NOT IN ({marcadores})", manter)
                else:
                    cursor.execute("DELETE FROM funcionarios")

            cursor.executemany('''
                INSERT INTO funcionarios (nome, email)
                VALUES (?, ?)
                ON CONFLICT(nome) DO UPDATE SET email=excluded.email
            ''', list(mapping_dict.items()))
            conn.commit()

    # ------------------------------------------------------------------
    # Dashboard
    # ------------------------------------------------------------------

    def update_last_check_time(self):
        """Atualiza o timestamp da última verificação com horário local."""
        self.update_config('last_check_time', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

    def get_dashboard_stats(self) -> Dict[str, Any]:
        with self.get_connection() as conn:
            cursor = conn.cursor()

            stats = {}
            seven_days_ago = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
            cursor.execute("SELECT COUNT(*) FROM changes WHERE data_deteccao >= ?", (seven_days_ago,))
            stats['alteracoes_semana'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(DISTINCT funcionario) FROM changes WHERE data_deteccao >= ?", (seven_days_ago,))
            stats['funcionarios_afetados'] = cursor.fetchone()[0]

            cursor.execute("SELECT value FROM config WHERE key = 'last_check_time'")
            row = cursor.fetchone()
            if row and row[0]:
                stats['ultima_execucao'] = row[0]
            else:
                cursor.execute("SELECT MAX(import_date) FROM grades")
                stats['ultima_execucao'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM changes WHERE email_enviado = 1")
            stats['emails_enviados'] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM changes WHERE email_enviado = 0")
            stats['emails_pendentes'] = cursor.fetchone()[0]

            return stats
