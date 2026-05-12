import os
import logging
from typing import Optional, Tuple
import datetime

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class OutlookMonitor:
    def __init__(self, folder_name: str, keyword: str, save_dir: str):
        self.folder_name = folder_name
        self.keyword = keyword.lower()
        self.save_dir = save_dir
        
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

    def search_new_emails(self) -> list:
        # Retorna lista com tuplas (email_id, attachment_path)
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
                except Exception:
                    logger.warning(f"Pasta {self.folder_name} não encontrada, usando Caixa de Entrada.")

            messages = folder.Items
            messages.Sort("[ReceivedTime]", True) # Mais recentes primeiro
            
            found_emails = []
            
            for msg in list(messages)[:50]: # Limita ultimos 50 emails
                try:
                    subject = msg.Subject or ""
                    if self.keyword in subject.lower():
                        entry_id = msg.EntryID
                        
                        if msg.Attachments.Count > 0:
                            for att in msg.Attachments:
                                filename = f"{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}_{att.FileName}"
                                att_path = os.path.join(self.save_dir, filename)
                                
                                # Verifica extensão
                                if filename.lower().endswith(('.xlsx', '.csv', '.xls')):
                                    att.SaveAsFile(att_path)
                                    found_emails.append((entry_id, att_path))
                                    # Considera apenas o 1o anexo util deste email
                                    break
                except Exception as e:
                    logger.error(f"Erro lendo mensagem: {str(e)}")
                    continue
                    
            return found_emails

        except Exception as e:
            logger.error(f"Erro na varredura: {str(e)}")
            return []
