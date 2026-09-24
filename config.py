import os

# Caminho absoluto para o banco de dados e arquivos locais
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "data", "escala.db")
ATTACHMENTS_DIR = os.path.join(BASE_DIR, "data", "attachments")

# Cria pastas se não existirem
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# Pasta onde os e-mails de homologação (Modo Teste) são gravados.
# Fica ao lado do projeto, e não no diretório de onde o app foi iniciado.
TEST_EMAILS_DIR = os.path.join(BASE_DIR, "emails_teste")

# Dias de retenção dos anexos baixados do Outlook.
ATTACHMENT_RETENTION_DAYS = 30

# Configurações padrão sugeridas (serão sobrepostas pelo que estiver no DB)
# IMPORTANTE: toda chave usada pela interface precisa estar aqui OU ser gravada
# via update_config(), que agora faz upsert.
DEFAULT_CONFIG = {
    "mail_subject_keyword": "GRADE, PPV, ESCALA, NOVA ESCALA",
    "outlook_folder": "Caixa de Entrada",
    "check_interval_min": 15,
    "test_mode": 1, # 1 para Sim (testes), 0 para Não (produção)
    "employee_mapping_path": "",
    "notification_subject": "Atualização na sua Escala",
    "immune_keywords": "VIAGEM, FOLGA, OFF, REUNIAO, GRAVACAO, MEDICO, FERIAS",
    "last_check_time": "",
    # Onde o Modo Teste grava os relatórios de homologação. Vazio = usa
    # TEST_EMAILS_DIR (emails_teste/ ao lado do projeto).
    "test_output_dir": "",
    # Trava de segurança: acima deste número de funcionários afetados num único
    # ciclo, o envio é bloqueado e fica pendente de revisão manual. Protege
    # contra disparo em massa quando a base de comparação está defasada
    # (ex.: grade de outro mês). 0 desativa a trava.
    "max_funcionarios_por_ciclo": 25,
}

# Colunas esperadas na planilha (minímas)
EXPECTED_COLUMNS = ['funcionario', 'data', 'horario']
