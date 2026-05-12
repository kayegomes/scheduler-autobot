import os

# Caminho absoluto para o banco de dados e arquivos locais
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data", "escala.db")
ATTACHMENTS_DIR = os.path.join(BASE_DIR, "data", "attachments")

# Cria pastas se não existirem
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# Configurações padrão sugeridas (serão sobrepostas pelo que estiver no DB)
DEFAULT_CONFIG = {
    "mail_subject_keyword": "Nova Escala",
    "outlook_folder": "Caixa de Entrada",
    "check_interval_min": 15,
    "test_mode": 1, # 1 para Sim (testes), 0 para Não (produção)
    "employee_mapping_path": "",
    "notification_subject": "Atualização na sua Escala"
}

# Colunas esperadas na planilha (minímas)
EXPECTED_COLUMNS = ['funcionario', 'data', 'horario']
