"""Homologação manual: simula um ciclo completo sem tocar no Outlook.

Uso:
    python run_test.py [arquivo_recebido.xlsx]

Diferente da versão anterior, este script NÃO reimplementa o fluxo: ele chama
o mesmo `ScheduleProcessor` que roda em produção. Reescrever a lógica aqui foi
o que permitiu que o cruzamento quebrado passasse despercebido — o teste
"passava" porque testava uma cópia do código, não o código.
"""

import logging
import os
import sys

from config import TEST_EMAILS_DIR
from core.columns import carregar_escala, detectar_tipo_planilha
from core.database import DatabaseManager
from core.schedule_processor import ScheduleProcessor

logging.basicConfig(level=logging.INFO, format='%(levelname)s | %(message)s')

DB_TESTE = os.path.join('data', 'test_scheduler.db')
ESCALA_BASE = 'Check_Pre_Envio_Gerado.xlsx'


def main():
    arquivo = sys.argv[1] if len(sys.argv) > 1 else 'GRADE DE MAIO - 9ª VERSÃO.xlsm'

    if not os.path.exists(arquivo):
        print(f"[ERRO] Arquivo não encontrado: {arquivo}")
        return 1
    if not os.path.exists(ESCALA_BASE):
        print(f"[ERRO] Escala base não encontrada: {ESCALA_BASE}")
        return 1

    if os.path.exists(DB_TESTE):
        os.remove(DB_TESTE)

    db = DatabaseManager(DB_TESTE)
    db.update_config('test_mode', '1')
    # Sem limite: em homologação queremos ver todos os e-mails que sairiam.
    db.update_config('max_funcionarios_por_ciclo', '0')

    print(f"\n{'=' * 70}\n1. Carregando escala base: {ESCALA_BASE}\n{'=' * 70}")
    df_base = carregar_escala(ESCALA_BASE)
    db.save_new_grade(df_base, os.path.basename(ESCALA_BASE), 'carga_manual')
    pessoas = sorted(df_base['funcionario'].unique())
    print(f"   {len(df_base)} linhas | {len(pessoas)} funcionários")

    # Mapeia todo mundo para um e-mail de teste, senão nada é notificado.
    db.update_employee_emails({p: 'homologacao@teste.local' for p in pessoas})

    print(f"\n{'=' * 70}\n2. Simulando recebimento: {arquivo}\n{'=' * 70}")
    print(f"   Tipo detectado: {detectar_tipo_planilha(arquivo)}")

    processor = ScheduleProcessor(db)
    config = db.get_config()
    immune = [k.strip().upper() for k in config['immune_keywords'].split(',') if k.strip()]

    processor._processar_arquivo(
        email_id='homologacao',
        att_path=arquivo,
        immune_keys=immune,
        test_mode=True,
        subj=config['notification_subject'],
        max_funcs=0,
    )

    print(f"\n{'=' * 70}\n3. Resultado\n{'=' * 70}")
    # Em Modo Teste nada vai para o banco: o resultado está na pasta do ciclo.
    pastas = sorted(
        (os.path.join(TEST_EMAILS_DIR, d) for d in os.listdir(TEST_EMAILS_DIR)),
        key=os.path.getmtime
    ) if os.path.isdir(TEST_EMAILS_DIR) else []

    if not pastas:
        print("   Nenhum relatório gerado (nenhuma alteração detectada).")
        return 0

    ultima = pastas[-1]
    resumo = os.path.join(ultima, "RESUMO.txt")
    if os.path.exists(resumo):
        with open(resumo, encoding="utf-8") as f:
            print("\n".join("   " + l for l in f.read().splitlines()))

    print(f"\n   Pasta do ciclo: {ultima}")
    emails = os.path.join(ultima, "emails")
    if os.path.isdir(emails):
        print(f"   {len(os.listdir(emails))} preview(s) HTML em {emails}")

    return 0


if __name__ == '__main__':
    sys.exit(main())
