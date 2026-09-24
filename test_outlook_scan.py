import logging
import sys
from core.outlook_monitor import OutlookMonitor
from core.database import DatabaseManager

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

db = DatabaseManager()
config = db.get_config()

keyword = config.get('mail_subject_keyword', 'GRADE, PPV, ESCALA, NOVA ESCALA')
folder = config.get('outlook_folder', 'Caixa de Entrada')

print(f"=== TESTANDO OUTLOOK MONITOR ===")
print(f"Buscando e-mails com keyword: '{keyword}' na pasta: '{folder}'")

monitor = OutlookMonitor(folder_name=folder, keyword=keyword, save_dir='data/attachments')
emails = monitor.search_new_emails()

print(f"\nResultado da busca:")
print(f"E-mails encontrados com anexos válidos: {len(emails)}")
for email_id, path in emails:
    print(f" - Email ID: {email_id[:25]}... -> Anexo: {path}")
