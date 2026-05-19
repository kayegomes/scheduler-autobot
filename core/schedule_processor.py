import os
import logging
import pandas as pd
from typing import List, Dict, Any
from traceback import format_exc
from .database import DatabaseManager
from .outlook_monitor import OutlookMonitor
from .diff_engine import compare_schedules
from .email_sender import EmailSender
from config import ATTACHMENTS_DIR

logger = logging.getLogger(__name__)

class ScheduleProcessor:
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager
        
    def process_cycle(self):
        # 1. Recupera Configurações
        config = self.db.get_config()
        keyword = config.get('mail_subject_keyword', 'Nova Escala')
        folder = config.get('outlook_folder', 'Caixa de Entrada')
        test_mode = str(config.get('test_mode', '1')) == '1'
        subj = config.get('notification_subject', 'Atualização na sua Escala')
        immune_keys_str = config.get('immune_keywords', 'VIAGEM, FOLGA, OFF, REUNIAO, GRAVACAO, MEDICO, FERIAS')
        immune_keys = [k.strip().upper() for k in immune_keys_str.split(',') if k.strip()]
        
        logger.info(f"Iniciando ciclo de processamento. Keyword: {keyword}")
        
        # 2. Busca e-mails
        monitor = OutlookMonitor(folder_name=folder, keyword=keyword, save_dir=ATTACHMENTS_DIR)
        emails_encontrados = monitor.search_new_emails()
        
        if not emails_encontrados:
            logger.info("Nenhum novo e-mail encontrado.")
            return

        for email_id, att_path in emails_encontrados:
            if self.db.is_email_processed(email_id):
                logger.debug(f"E-mail já processado anteriormente. Ignorando...")
                continue
                
            try:
                # 3. Lê planilha nova
                logger.info(f"Processando arquivo recebido: {att_path}")
                from core.match_eventos import carregar_grade, buscar_evento_na_grade
                
                # 4. Lê antiga do BD primeiro para usar como base se for cruzamento
                df_old = self.db.get_last_grade_df()
                
                grade_tv = carregar_grade(att_path)
                if grade_tv is not None and not df_old.empty:
                    logger.info("Arquivo identificado como Grade de TV. Iniciando cruzamento (Match)...")
                    df_new = df_old.copy()
                    
                    # Padroniza colunas do df_old para o match se necessário
                    # O buscar_evento_na_grade busca por 'Data' e 'Evento/Programa'
                    col_map = {c.lower(): c for c in df_new.columns}
                    data_col = col_map.get('data', 'Data')
                    evento_col = col_map.get('evento', col_map.get('evento/programa', col_map.get('evento/descricao', 'Evento/Programa')))
                    
                    for idx, row in df_new.iterrows():
                        # Cria um dict fake row pro match
                        match_row = {
                            'Data': row.get(data_col),
                            'Evento/Programa': row.get(evento_col),
                            'Início': row.get(col_map.get('inicio', 'inicio')),
                            'Fim': row.get(col_map.get('fim', 'fim'))
                        }
                        
                        match_info = buscar_evento_na_grade(match_row, grade_tv)
                        if match_info:
                            # Encontrou correspondência na nova grade, atualiza os horários
                            df_new.at[idx, col_map.get('inicio', 'inicio')] = match_info.get('horario_inicio', '')
                            df_new.at[idx, col_map.get('fim', 'fim')] = match_info.get('horario_fim', '')
                            # Opcional: pre e pos
                        else:
                            # Se não encontrou, e não for folga/viagem, pode ter caido
                            ev_upper = str(row.get(evento_col, '')).upper()
                            if not any(k in ev_upper for k in immune_keys):
                                # Evento caiu da grade
                                df_new.at[idx, col_map.get('inicio', 'inicio')] = "CANCELADO"
                else:
                    logger.info("Arquivo lido como Escala Comum.")
                    if att_path.endswith('.csv'):
                        df_new = pd.read_csv(att_path)
                    else:
                        df_new = pd.read_excel(att_path)
                    
                # O df_old já foi lido acima
                
                # 5. Salva nova escala no banco
                filename = os.path.basename(att_path)
                grade_id = self.db.save_new_grade(df_new, filename, email_id)
                logger.info(f"Grade {grade_id} salva com sucesso.")
                
                # 6. Compara
                logger.info("Executando comparação (Diff) de escalas...")
                changes = compare_schedules(df_old, df_new, grade_id)
                
                if not changes:
                    logger.info("Nenhuma alteração detectada nesta versão.")
                    continue
                    
                # Salva alterações
                self.db.save_changes(changes)
                logger.info(f"{len(changes)} alterações detectadas e salvas.")
                
                # 7. Dispara e-mails
                self._send_notifications(changes, df_new, test_mode, subj)
                
            except Exception as e:
                logger.error(f"Erro processando arquivo {att_path}:\n{format_exc()}")

    def _send_notifications(self, changes: List[Dict[str, Any]], df_new: pd.DataFrame, test_mode: bool, base_subject: str):
        # Agrupa mudanças por funcionário
        grouped = {}
        for c in changes:
            func = c['funcionario']
            if func not in grouped:
                grouped[func] = []
            grouped[func].append(c)
            
        sender = EmailSender(test_mode=test_mode)
        
        for func, func_changes in grouped.items():
            email_address = self.db.get_employee_email(func)
            if not email_address:
                logger.warning(f"E-mail não encontrado no banco para: {func}. Notificação não enviada.")
                continue
                
            # Filtra a escala completa apenas deste funcionário para enviar no email
            df_func = df_new[df_new['funcionario'].str.lower() == func.lower()] if df_new is not None else None
            success = sender.send_notification(email_address, func, func_changes, df_func, subject=base_subject)
            # Para fins de agilidade, marcaremos o lote todo como enviado no nível da grade,
            # mas o ideal seria capturar os IDs individuais das changes.
                
        # Atualiza a flag de enviado (simplificado: marcamos todos associados a esta grade)
        if changes:
            grade_id = changes[0]['grade_id']
            with self.db.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE changes SET email_enviado=1 WHERE grade_id=?", (grade_id,))
                conn.commit()
