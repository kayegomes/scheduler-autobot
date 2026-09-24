import os
import logging
from typing import Callable, Optional, Tuple
import datetime

logger = logging.getLogger(__name__)

EXTENSOES_PLANILHA = ('.xlsx', '.xlsm', '.csv', '.xls')

class OutlookMonitor:
    def __init__(self, folder_name: str, keyword: str, save_dir: str):
        self.folder_name = folder_name
        self.keyword_raw = keyword
        # Suporta múltiplas palavras-chave separadas por vírgula
        if isinstance(keyword, str):
            self.keywords = [k.strip().lower() for k in keyword.split(',') if k.strip()]
        elif isinstance(keyword, list):
            self.keywords = [k.strip().lower() for k in keyword if k.strip()]
        else:
            self.keywords = ["grade"]
        self.save_dir = os.path.abspath(save_dir)
        os.makedirs(self.save_dir, exist_ok=True)
        
    def get_dispatch(self):
        try:
            import win32com.client
            return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        except ImportError:
            logger.error("pywin32 não instalado ou não suportado.")
            return None
        except Exception as e:
            logger.error(f"Erro ao conectar ao Outlook: {str(e)}")
            return None

    def search_new_emails(self, ja_processado: Optional[Callable[[str], bool]] = None) -> list:
        """Retorna lista com tuplas (email_id, attachment_path).

        Busca emails na pasta configurada que contenham qualquer uma das palavras-chave no assunto
        e possuam anexos de planilha (.xlsx, .xlsm, .csv, .xls).

        `ja_processado` é consultado ANTES de gravar o anexo em disco. Sem esse
        filtro, todo ciclo rebaixava os anexos dos 50 últimos e-mails e a pasta
        `data/attachments` crescia indefinidamente.
        """
        outlook = self.get_dispatch()
        if not outlook:
            return []

        try:
            # 6 = Inbox
            folder = outlook.GetDefaultFolder(6)
            
            # Se a pasta alvo não for a Inbox padrão, busca subpastas ou pelo nome
            if self.folder_name.lower() not in ["inbox", "caixa de entrada"]:
                try:
                    folder = folder.Folders(self.folder_name)
                    logger.info(f"Usando subpasta do Outlook: '{self.folder_name}'")
                except Exception:
                    logger.warning(f"Pasta '{self.folder_name}' não encontrada, usando Caixa de Entrada.")

            messages = folder.Items
            messages.Sort("[ReceivedTime]", True) # Mais recentes primeiro
            
            found_emails = []
            keyword_match_count = 0
            no_attachment_count = 0
            no_spreadsheet_count = 0
            already_done_count = 0

            for msg in list(messages)[:50]: # Limita ultimos 50 emails
                try:
                    subject = msg.Subject or ""
                    subject_lower = subject.lower()
                    if any(kw in subject_lower for kw in self.keywords):
                        keyword_match_count += 1
                        entry_id = msg.EntryID

                        if ja_processado is not None and ja_processado(entry_id):
                            already_done_count += 1
                            continue

                        if msg.Attachments.Count > 0:
                            spreadsheet_found = False
                            for att in msg.Attachments:
                                if not str(att.FileName).lower().endswith(EXTENSOES_PLANILHA):
                                    continue

                                filename = f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{att.FileName}"
                                att_path = os.path.join(self.save_dir, filename)
                                att.SaveAsFile(att_path)
                                found_emails.append((entry_id, att_path))
                                spreadsheet_found = True
                                # Considera apenas o 1o anexo util deste email
                                break

                            if not spreadsheet_found:
                                no_spreadsheet_count += 1
                                att_names = [att.FileName for att in msg.Attachments]
                                logger.info(f"Email com keyword encontrado mas sem planilha. Assunto: '{subject}'. Anexos: {att_names}")
                        else:
                            no_attachment_count += 1
                            logger.debug(f"Email com keyword encontrado mas sem anexos. Assunto: '{subject}'")
                except Exception as e:
                    logger.error(f"Erro lendo mensagem: {str(e)}")
                    continue

            logger.info(
                f"Varredura concluída: {keyword_match_count} email(s) com palavras-chave {self.keywords} encontrado(s). "
                f"{len(found_emails)} com planilha nova, {already_done_count} já processado(s), "
                f"{no_attachment_count} sem anexos, {no_spreadsheet_count} com anexos não-planilha."
            )

            return found_emails

        except Exception as e:
            logger.error(f"Erro na varredura: {str(e)}")
            return []

    def limpar_anexos_antigos(self, dias: int = 30):
        """Remove anexos baixados há mais de `dias` dias."""
        if not dias or dias <= 0:
            return
        limite = datetime.datetime.now() - datetime.timedelta(days=dias)
        removidos = 0
        try:
            for nome in os.listdir(self.save_dir):
                caminho = os.path.join(self.save_dir, nome)
                try:
                    if not os.path.isfile(caminho):
                        continue
                    if datetime.datetime.fromtimestamp(os.path.getmtime(caminho)) < limite:
                        os.remove(caminho)
                        removidos += 1
                except OSError:
                    continue
        except OSError as e:
            logger.warning(f"Não foi possível limpar anexos antigos: {e}")
            return

        if removidos:
            logger.info(f"Limpeza: {removidos} anexo(s) com mais de {dias} dias removido(s).")
