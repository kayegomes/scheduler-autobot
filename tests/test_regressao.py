"""Regressão das falhas encontradas na análise do fluxo.

Cada teste aqui corresponde a um defeito real que estava em produção. Rodar:

    python -m unittest discover -s tests -v
"""

import os
import re
import sys
import tempfile
import unittest

import pandas as pd

_AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(_AQUI))  # raiz do projeto
sys.path.insert(0, _AQUI)                   # permite 'import fixtures' em
                                            # 'unittest tests.test_regressao'

from core.columns import (  # noqa: E402
    carregar_escala, detectar_tipo_planilha, normalizar_colunas, preparar_escala,
)
from core.database import DatabaseManager  # noqa: E402
from core.diff_engine import compare_schedules  # noqa: E402
from core.email_sender import EmailSender  # noqa: E402

import fixtures  # noqa: E402

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Planilhas sintéticas, criadas uma vez por execução. Reproduzem as manhas do
# formato real (ver tests/fixtures.py), então a suíte roda inteira num clone
# limpo — sem depender de arquivos com dados de funcionários.
ESCALA = ""
GRADE = ""
GRADE_ESPECIALIZADA = ""
GRADE_OUTRO_PERIODO = ""
_TMP_FIXTURES = ""

# As planilhas de produção continuam servindo de validação extra quando estão
# presentes na raiz do projeto, mas nenhum teste depende delas.
ESCALA_REAL = os.path.join(RAIZ, "Check_Pre_Envio_Gerado.xlsx")
GRADE_REAL = os.path.join(RAIZ, "GRADE DE EVENTOS COMBATE 2026 - JULHO  (V4).xlsx")
precisa_planilhas_reais = unittest.skipUnless(
    os.path.exists(ESCALA_REAL) and os.path.exists(GRADE_REAL),
    "planilhas de produção ausentes (dados internos, não versionados)"
)


def setUpModule():
    global ESCALA, GRADE, GRADE_ESPECIALIZADA, GRADE_OUTRO_PERIODO, _TMP_FIXTURES
    _TMP_FIXTURES = tempfile.mkdtemp(prefix="fixtures_escala_")
    ESCALA = fixtures.criar_escala(os.path.join(_TMP_FIXTURES, "escala.xlsx"))
    GRADE = fixtures.criar_grade_tv(os.path.join(_TMP_FIXTURES, "grade_tv.xlsx"))
    GRADE_ESPECIALIZADA = fixtures.criar_grade_especializada(
        os.path.join(_TMP_FIXTURES, "grade_combate.xlsx"))
    GRADE_OUTRO_PERIODO = fixtures.criar_grade_de_outro_periodo(
        os.path.join(_TMP_FIXTURES, "grade_outubro.xlsx"))


def tearDownModule():
    import shutil
    shutil.rmtree(_TMP_FIXTURES, ignore_errors=True)


def _db_temporario() -> DatabaseManager:
    caminho = os.path.join(tempfile.mkdtemp(prefix="escala_test_"), "teste.db")
    return DatabaseManager(db_path=caminho)


class TestNormalizacaoDeColunas(unittest.TestCase):
    """Crítico #3 e #4: colunas acentuadas e apelidos não resolvidos."""

    def test_acento_vira_coluna_canonica(self):
        df = normalizar_colunas(pd.DataFrame({"Início": ["10:00"], "Nome ": ["Ana"]}))
        self.assertIn("inicio", df.columns)
        self.assertIn("funcionario", df.columns)

    def test_atribuicao_nao_cria_coluna_fantasma(self):
        df = preparar_escala(pd.DataFrame({"Início": ["10:00", "11:00"], "Nome": ["Ana", "Bia"]}))
        df.at[0, "inicio"] = "14:00"
        self.assertEqual(list(df.columns).count("inicio"), 1)
        self.assertEqual(df.at[0, "inicio"], "14:00")
        self.assertEqual(df.at[1, "inicio"], "11:00")

    def test_data_raw_nao_vira_data(self):
        df = normalizar_colunas(pd.DataFrame({"Data": ["01/07/2026"], "Data_raw": ["x"]}))
        self.assertIn("data", df.columns)
        self.assertIn("data_raw", df.columns)

    def test_escala_real_resolve_todas_as_colunas_chave(self):
        df = carregar_escala(ESCALA)
        for col in ("funcionario", "data", "inicio", "fim", "evento", "pre"):
            self.assertIn(col, df.columns, f"coluna canônica ausente: {col}")


class TestSerializacao(unittest.TestCase):
    """Crítico #1: datas viravam epoch e quebravam a chave de comparação."""

    def test_ida_e_volta_preserva_data(self):
        db = _db_temporario()
        df = pd.DataFrame({
            "Nome": ["Ana"],
            "Data": [pd.Timestamp("2026-07-01")],
            "Início": [pd.Timestamp("2026-07-01 14:30")],
        })
        db.save_new_grade(df, "base.xlsx", "id1")
        recarregado = db.get_last_grade_df()

        self.assertEqual(recarregado.at[0, "data"], "01/07/2026")
        self.assertEqual(recarregado.at[0, "inicio"], "14:30")

    def test_planilha_identica_nao_gera_alteracao(self):
        db = _db_temporario()
        df = carregar_escala(ESCALA)
        db.save_new_grade(df, "base.xlsx", "id1")

        changes = compare_schedules(db.get_last_grade_df(), carregar_escala(ESCALA), 2)
        self.assertEqual(changes, [], f"{len(changes)} alterações falsas com arquivo idêntico")

    def test_alteracao_real_e_detectada(self):
        db = _db_temporario()
        df = carregar_escala(ESCALA)
        db.save_new_grade(df, "base.xlsx", "id1")

        df_novo = carregar_escala(ESCALA)
        df_novo.at[0, "inicio"] = "23:45"

        changes = compare_schedules(db.get_last_grade_df(), df_novo, 2)
        self.assertEqual(len(changes), 1)
        self.assertEqual(changes[0]["tipo"], "ALTERACAO")
        self.assertEqual(changes[0]["valor_novo"], "23:45")

    def test_migracao_converte_epoch_legado(self):
        """Grades gravadas antes da correção precisam ser renormalizadas."""
        db = _db_temporario()
        with db.get_connection() as conn:
            conn.execute(
                "INSERT INTO grades (import_date, filename, source_email_id, status) "
                "VALUES ('2026-01-01 00:00:00', 'legado.xlsx', 'antigo', 'done')"
            )
            conn.execute(
                "INSERT INTO grade_items (grade_id, funcionario, data, raw_data) VALUES (1, 'Ana', '', ?)",
                ['{"funcionario":"Ana","data":1782864000000,"inicio":"14:30:00"}'],
            )
            conn.execute("UPDATE config SET value='1' WHERE key='schema_version'")

        DatabaseManager(db_path=db.db_path)  # dispara a migração
        recarregado = DatabaseManager(db_path=db.db_path).get_last_grade_df()
        self.assertEqual(recarregado.at[0, "data"], "01/07/2026")
        self.assertEqual(recarregado.at[0, "inicio"], "14:30")


class TestDeteccaoDeTipo(unittest.TestCase):
    """Crítico #2: a escala real era confundida com grade de TV."""

    def test_escala_real_e_escala(self):
        self.assertEqual(detectar_tipo_planilha(ESCALA), "escala")

    def test_grade_real_e_grade(self):
        self.assertEqual(detectar_tipo_planilha(GRADE), "grade_tv")

    def test_carregar_grade_nao_fatia_escala(self):
        from core.match_eventos import carregar_grade
        grade = carregar_grade(ESCALA)
        # Com o bug, o achatamento descartava a coluna Nome.
        self.assertIsNotNone(grade)
        self.assertIn("NOME", [str(c).upper().strip() for c in grade.columns])


class TestConfiguracao(unittest.TestCase):
    """Alto #6: update_config descartava chaves inexistentes."""

    def test_cria_chave_nova(self):
        db = _db_temporario()
        db.update_config("immune_keywords", "VIAGEM, FOLGA")
        self.assertEqual(db.get_config()["immune_keywords"], "VIAGEM, FOLGA")

    def test_last_check_time_persiste(self):
        db = _db_temporario()
        db.update_last_check_time()
        self.assertTrue(db.get_config().get("last_check_time"))
        self.assertTrue(db.get_dashboard_stats()["ultima_execucao"])


class TestMapeamentoDeFuncionarios(unittest.TestCase):
    """Alto #9 e Médio #16."""

    def test_nome_sem_acento_e_match_exato(self):
        db = _db_temporario()
        db.update_employee_emails({"Marina Tavares": "marina@exemplo.com"})
        self.assertEqual(db.get_employee_email("Marina Tavares"), "marina@exemplo.com")

    def test_pessoa_diferente_nao_recebe_email_alheio(self):
        db = _db_temporario()
        db.update_employee_emails({"Marcio Antonio Sa": "marcio@exemplo.com"})
        self.assertIsNone(db.get_employee_email("Marcos Antonio Sa"))

        db.update_employee_emails({"Joana Pedro Alves": "joana@exemplo.com"})
        self.assertIsNone(db.get_employee_email("Joao Pedro Alves"))

    def test_remover_linha_apaga_vinculo(self):
        db = _db_temporario()
        db.update_employee_emails({"Ana": "ana@exemplo.com", "Bia": "bia@exemplo.com"})
        db.update_employee_emails({"Ana": "ana@exemplo.com"})
        self.assertIsNone(db.get_employee_email("Bia"))


class TestEmailSender(unittest.TestCase):
    """Médio #11: prefixo [:5] fazia datas diferentes casarem."""

    def test_datas_diferentes_nao_casam(self):
        self.assertFalse(EmailSender._datas_iguais("2026-07-01", "2026-12-25"))
        self.assertFalse(EmailSender._datas_iguais("01/07/2026", "01/07/2025"))
        self.assertFalse(EmailSender._datas_iguais("01/07/2026", "02/07/2026"))

    def test_datas_iguais_casam(self):
        self.assertTrue(EmailSender._datas_iguais("01/07/2026", "01/07/2026"))
        self.assertTrue(EmailSender._datas_iguais("2026-07-01", "01/07/2026"))

    @staticmethod
    def _tabela(html_body):
        """Extrai as linhas da tabela como listas de células."""
        linhas = []
        for tr in re.findall(r"<tr>(.*?)</tr>", html_body, re.S)[1:]:
            celulas = re.findall(r"<td class='\w*'>(.*?)</td>", tr, re.S)
            if celulas:
                linhas.append([c.strip() for c in celulas])
        return linhas

    def _escala_da_semana(self):
        return preparar_escala(pd.DataFrame([
            {"Nome": "Rui Campos", "Data": "21/09/2026", "Início": "23:30", "Fim": "01:00",
             "Evento": "TROCA DE PASSES"},
            {"Nome": "Rui Campos", "Data": "23/09/2026", "Início": "11:30", "Fim": "12:00",
             "Pré": "09:45", "Evento": "COPA DO MUNDO FEMININA SUB-20"},
            {"Nome": "Rui Campos", "Data": "25/09/2026", "Início": "15:45", "Fim": "17:45",
             "Evento": "HUNGRIA X UCRÂNIA"},
            {"Nome": "Rui Campos", "Data": "22/09/2026", "Início": "00:00", "Fim": "00:00",
             "Evento": "", "Descrição": "Day Off / Folga"},
            {"Nome": "Rui Campos", "Data": "24/09/2026", "Início": "00:00", "Fim": "00:00",
             "Evento": "", "Descrição": "Comp Day / Folga Compensatória"},
        ]))

    def test_tabela_sai_em_ordem_de_data(self):
        """A planilha agrupa por tipo de atividade: saía 21, 23, 25, 22, 24."""
        html_body = EmailSender(test_mode=True)._generate_html_body(
            "Rui Campos", [], self._escala_da_semana()
        )
        datas = [l[2] for l in self._tabela(html_body)]
        self.assertEqual(datas, ["21/09/2026", "22/09/2026", "23/09/2026",
                                 "24/09/2026", "25/09/2026"])

    def test_ordena_por_hora_dentro_do_dia(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "21/09/2026", "Início": "20:00", "Evento": "B"},
            {"Nome": "Ana", "Data": "21/09/2026", "Início": "08:00", "Evento": "A"},
        ]))
        html_body = EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
        self.assertEqual([l[5] for l in self._tabela(html_body)], ["08:00", "20:00"])

    def test_folga_mostra_rotulo_e_traco_nos_horarios(self):
        html_body = EmailSender(test_mode=True)._generate_html_body(
            "Rui Campos", [], self._escala_da_semana()
        )
        folgas = [l for l in self._tabela(html_body) if l[2] in ("22/09/2026", "24/09/2026")]
        self.assertEqual(len(folgas), 2)
        for linha in folgas:
            self.assertEqual(linha[7], "FOLGA")
            self.assertEqual(linha[4], "-", "Pré")
            self.assertEqual(linha[5], "-", "Início")
            self.assertEqual(linha[6], "-", "Fim")

    def test_evento_real_a_meia_noite_mantem_horario(self):
        """Só folga vira '-': um evento que começa 00:00 continua 00:00."""
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "21/09/2026", "Início": "00:00", "Fim": "02:00",
             "Evento": "NFL - JOGO DA MADRUGADA"},
        ]))
        linha = self._tabela(
            EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
        )[0]
        self.assertEqual(linha[5], "00:00")
        self.assertEqual(linha[7], "NFL - JOGO DA MADRUGADA")

    def test_plataforma_sportv_e_normalizada(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "21/09/2026", "Início": "10:00", "Evento": "A",
             "Plataforma": "Sportv 2"},
            {"Nome": "Ana", "Data": "22/09/2026", "Início": "10:00", "Evento": "B",
             "Plataforma": "Sportv 3"},
            {"Nome": "Ana", "Data": "23/09/2026", "Início": "10:00", "Evento": "C",
             "Plataforma": "Sportv"},
            {"Nome": "Ana", "Data": "24/09/2026", "Início": "10:00", "Evento": "D",
             "Plataforma": "TV Globo - REDE"},
            {"Nome": "Ana", "Data": "25/09/2026", "Início": "10:00", "Evento": "E",
             "Plataforma": "Combate"},
        ]))
        plataformas = [
            l[1] for l in self._tabela(
                EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
            )
        ]
        self.assertEqual(
            plataformas, ["Sportv", "Sportv", "Sportv", "TV Globo - REDE", "Combate"],
            "canais Sportv viram 'Sportv'; outras famílias passam intactas"
        )

    def test_elenco_nao_repete_o_destinatario(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Carlos Vieira", "Data": "23/09/2026", "Início": "11:30",
             "Evento": "COPA DO MUNDO",
             "Elenco": "Carlos Vieira ; Rui Campos ; Sofia Dantas"},
        ]))
        linha = self._tabela(
            EmailSender(test_mode=True)._generate_html_body("Carlos Vieira", [], df)
        )[0]
        self.assertEqual(linha[10], "Rui Campos ; Sofia Dantas")

    def test_elenco_ignora_acento_ao_remover_o_destinatario(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Sofia Dantas", "Data": "23/09/2026", "Início": "11:30",
             "Evento": "X", "Elenco": "Sofia Dantas ; Rui Campos"},
        ]))
        linha = self._tabela(
            EmailSender(test_mode=True)._generate_html_body("Sofia Dantas", [], df)
        )[0]
        self.assertEqual(linha[10], "Rui Campos")

    def test_local_vem_de_local_de_locucao(self):
        """A base tem 'Local de locução' (Offtube) e 'Local' (o estádio)."""
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "23/09/2026", "Início": "11:30", "Evento": "X",
             "Local de locução": "Offtube ION", "Local": "Internacional - Polônia"},
        ]))
        linha = self._tabela(
            EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
        )[0]
        self.assertEqual(linha[9], "Offtube ION")

    def test_local_cai_para_local_quando_nao_ha_locucao(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "23/09/2026", "Início": "11:30", "Evento": "X",
             "Local": "ION"},
        ]))
        linha = self._tabela(
            EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
        )[0]
        self.assertEqual(linha[9], "ION")

    def test_linha_de_folga_nao_sai_em_branco(self):
        """Folga tem 'evento' vazio: sem fallback a linha ia em branco, 00:00 a 00:00."""
        changes = [{"data": "22/09/2026", "tipo": "ALTERACAO", "campo": "Inicio",
                    "valor_antigo": "19:00", "valor_novo": "18:30"}]
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Diego", "Data": "22/09/2026", "Início": "18:30", "Fim": "19:00",
             "Evento": "PILOTO - CIRCUITO MUNDIAL"},
            {"Nome": "Diego", "Data": "21/09/2026", "Início": "00:00", "Fim": "00:00",
             "Evento": "", "Descrição": "Day Off / Folga"},
        ]))
        html = EmailSender(test_mode=True)._generate_html_body("Diego", changes, df)
        linha_folga = [l for l in self._tabela(html) if l[2] == "21/09/2026"][0]
        self.assertEqual(linha_folga[7], "FOLGA", "a linha de folga precisa ser identificada")

    def test_destaque_so_na_linha_alterada(self):
        changes = [{
            "data": "17/07/2026", "tipo": "ALTERACAO", "campo": "Inicio",
            "valor_antigo": "19:30", "valor_novo": "23:45",
        }]
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "17/07/2026", "Início": "23:45", "Fim": "22:10"},
            {"Nome": "Ana", "Data": "25/12/2026", "Início": "10:00", "Fim": "12:00"},
        ]))
        html = EmailSender(test_mode=True)._generate_html_body("Ana", changes, df)
        linhas = [l for l in html.split("<tr>") if "23:45" in l or "10:00" in l]
        self.assertTrue(any("highlight" in l for l in linhas if "23:45" in l))
        self.assertFalse(any("highlight" in l for l in linhas if "10:00" in l))


class _SenderFalso:
    """Registra os envios em vez de falar com o Outlook."""

    def __init__(self, test_mode=True, falhar_para=()):
        self.enviados = []
        self.falhar_para = set(falhar_para)

    def send_notification(self, to_address, employee_name, changes, df_schedule=None,
                          subject="", sem_destinatario=False):
        if employee_name in self.falhar_para:
            return False
        self.enviados.append(employee_name)
        return True
class TestCicloDeProcessamento(unittest.TestCase):
    """Críticos #2/#5 e Alto #7/#8, no caminho real do ScheduleProcessor."""

    def setUp(self):
        import core.schedule_processor as sp
        self.sp = sp
        self.db = _db_temporario()
        self.processor = sp.ScheduleProcessor(self.db)
        self._sender_original = sp.EmailSender

    def tearDown(self):
        self.sp.EmailSender = self._sender_original

    def _instalar_sender(self, **kwargs):
        sender = _SenderFalso(**kwargs)
        self.sp.EmailSender = lambda test_mode=True, preview_dir=None: sender
        return sender

    def _base_e_arquivo_novo(self):
        """Carrega a escala real como base e devolve um arquivo com 1 mudança."""
        df = carregar_escala(ESCALA)
        self.db.save_new_grade(df, "base.xlsx", "carga_manual")

        alvo = df.iloc[0]["funcionario"]
        self.db.update_employee_emails({alvo: "alvo@exemplo.com"})

        df_novo = carregar_escala(ESCALA)
        df_novo.at[0, "inicio"] = "23:45"
        destino = os.path.join(tempfile.mkdtemp(), "nova_escala.xlsx")
        df_novo.to_excel(destino, index=False)
        return alvo, destino

    def test_escala_recebida_e_realmente_processada(self):
        alvo, arquivo = self._base_e_arquivo_novo()
        sender = self._instalar_sender()

        self.processor._processar_arquivo("email-1", arquivo, ["FOLGA"], False, "Assunto", 25)

        changes = self.db.get_recent_changes()
        self.assertEqual(len(changes), 1, "a alteração real precisa ser detectada")
        self.assertEqual(changes.iloc[0]["valor_novo"], "23:45")
        self.assertEqual(sender.enviados, [alvo])
        self.assertEqual(int(changes.iloc[0]["email_enviado"]), 1)

    def test_falha_no_envio_deixa_email_reprocessavel(self):
        alvo, arquivo = self._base_e_arquivo_novo()
        self._instalar_sender(falhar_para=[alvo])

        self.processor._processar_arquivo("email-2", arquivo, ["FOLGA"], False, "Assunto", 25)

        # A alteração fica registrada como NÃO enviada...
        changes = self.db.get_recent_changes()
        self.assertEqual(int(changes.iloc[0]["email_enviado"]), 0)

    def test_erro_no_meio_do_ciclo_nao_queima_o_email(self):
        _, arquivo = self._base_e_arquivo_novo()

        def explode(*a, **k):
            raise RuntimeError("falha simulada no envio")

        self.sp.EmailSender = explode
        with self.assertRaises(RuntimeError):
            self.processor._processar_arquivo("email-3", arquivo, ["FOLGA"], False, "Assunto", 25)

        # O e-mail continua elegível e a grade pendente não virou base.
        self.assertFalse(self.db.is_email_processed("email-3"))
        self.db.descartar_grades_pendentes("email-3")
        self.assertEqual(len(self.db.get_last_grade_df()), len(carregar_escala(ESCALA)))

    def test_trava_de_envio_em_massa(self):
        self.db.save_new_grade(carregar_escala(ESCALA), "base.xlsx", "carga_manual")
        sender = self._instalar_sender()

        # Escala completamente diferente -> todo mundo entra e sai.
        df_novo = pd.DataFrame([
            {"Nome": f"Pessoa {i}", "Data": "01/08/2026", "Início": "10:00", "Fim": "12:00"}
            for i in range(40)
        ])
        destino = os.path.join(tempfile.mkdtemp(), "outro_mes.xlsx")
        df_novo.to_excel(destino, index=False)

        self.processor._processar_arquivo("email-4", destino, ["FOLGA"], False, "Assunto", 25)

        self.assertEqual(sender.enviados, [], "envio em massa deveria ter sido bloqueado")
        self.assertGreater(len(self.db.get_recent_changes()), 0, "alterações ficam no histórico")

    def test_grade_tv_sem_base_nao_vira_escala(self):
        sender = self._instalar_sender()
        self.processor._processar_arquivo("email-5", GRADE, ["FOLGA"], False, "Assunto", 25)

        self.assertTrue(self.db.get_last_grade_df().empty)
        self.assertEqual(sender.enviados, [])

    def test_grade_tv_de_outro_periodo_nao_cancela_escala(self):
        """Nenhum match + cancelamento em massa = grade errada, aborta."""
        self.db.save_new_grade(carregar_escala(ESCALA), "base.xlsx", "carga_manual")
        sender = self._instalar_sender()

        self.processor._processar_arquivo("email-6", GRADE_ESPECIALIZADA, ["FOLGA"], False, "Assunto", 25)

        changes = self.db.get_recent_changes()
        cancelados = [c for _, c in changes.iterrows() if str(c["valor_novo"]).upper() == "CANCELADO"]
        self.assertEqual(cancelados, [], "a escala inteira não pode ser cancelada por engano")
        self.assertEqual(sender.enviados, [])


class TestGradeSemEfeito(unittest.TestCase):
    """Defeitos observados no log de 24/09: grades no-op soterrando a base."""

    def setUp(self):
        import core.schedule_processor as sp
        self.sp = sp
        self.db = _db_temporario()
        self.processor = sp.ScheduleProcessor(self.db)
        self._sender_original = sp.EmailSender
        sp.EmailSender = lambda test_mode=True, preview_dir=None: _SenderFalso()

    def tearDown(self):
        self.sp.EmailSender = self._sender_original

    def test_grade_de_outro_periodo_nao_cria_grade_copia(self):
        """A grade de outubro não toca a escala de julho: nada deve ser salvo.

        No log real, seis grades idênticas à base foram criadas assim — e a
        última soterrou a escala que o usuário tinha acabado de carregar.
        """
        base_id = self.db.save_new_grade(carregar_escala(ESCALA), "base.xlsx", "carga_manual")

        self.processor._processar_arquivo("email-x", GRADE_OUTRO_PERIODO, ["FOLGA"], False, "Assunto", 25)

        self.assertEqual(self.db.get_baseline_id(), base_id,
                         "a base de comparação não pode mudar por uma grade sem efeito")

    def test_carga_manual_sobrevive_ao_ciclo(self):
        """Sequência exata do log: base antiga -> carga manual -> grade no-op."""
        self.db.save_new_grade(carregar_escala(ESCALA), "base_antiga.xlsx", "carga_manual")

        nova = pd.DataFrame([
            {"Nome": "Paulo Reis", "Data": "21/09/2026", "Início": "10:00",
             "Fim": "12:00", "Evento/Descrição": "JOGO A"},
        ])
        id_manual = self.db.save_new_grade(nova, "escala_setembro.xlsx", "carga_manual")

        self.processor._processar_arquivo("email-y", GRADE_OUTRO_PERIODO, ["FOLGA"], False, "Assunto", 25)

        self.assertEqual(self.db.get_baseline_id(), id_manual,
                         "a escala carregada manualmente foi soterrada")
        self.assertEqual(len(self.db.get_last_grade_df()), 1)

    def test_arquivo_recusado_nao_reprocessa(self):
        """Os 4 arquivos Combate eram rebaixados e recusados a cada 2 minutos."""
        self.db.save_new_grade(carregar_escala(ESCALA), "base.xlsx", "carga_manual")

        self.assertFalse(self.db.is_email_processed("email-z"))
        self.processor._processar_arquivo("email-z", GRADE_OUTRO_PERIODO, ["FOLGA"], False, "Assunto", 25)
        self.assertTrue(self.db.is_email_processed("email-z"),
                        "arquivo recusado deve parar de ser rebaixado")

    def test_recusa_e_reavaliada_quando_a_base_muda(self):
        """A recusa vale para aquela base; com escala nova o arquivo volta à fila."""
        self.db.save_new_grade(carregar_escala(ESCALA), "base.xlsx", "carga_manual")
        self.processor._processar_arquivo("email-w", GRADE_OUTRO_PERIODO, ["FOLGA"], False, "Assunto", 25)
        self.assertTrue(self.db.is_email_processed("email-w"))

        self.db.save_new_grade(
            pd.DataFrame([{"Nome": "Ana", "Data": "01/10/2026", "Início": "10:00"}]),
            "escala_outubro.xlsx", "carga_manual"
        )
        self.assertFalse(self.db.is_email_processed("email-w"),
                         "com outra base, o arquivo precisa ser reavaliado")
class TestCruzamentoComGradeReal(unittest.TestCase):
    """Defeitos revelados pelo teste com a escala de 21/09 a 25/09."""

    def setUp(self):
        import core.schedule_processor as sp
        self.sp = sp
        self.db = _db_temporario()
        self.processor = sp.ScheduleProcessor(self.db)

    def _cruzar(self, df_escala, grade_path=None, immune=("VIAGEM", "FOLGA")):
        # Resolvido em tempo de chamada: as fixtures só existem depois do
        # setUpModule, e um argumento default é avaliado no import.
        grade_path = grade_path or GRADE
        return self.processor._cruzar_com_grade_tv(grade_path, df_escala, list(immune))

    def test_grade_nao_apaga_horario_que_nao_informa(self):
        """A grade do Sportv não traz FIM; sobrescrever com vazio apagava 146 horários."""
        import core.schedule_processor as sp

        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "01/07/2026", "Início": "10:00",
             "Fim": "12:15", "Evento/Descrição": "JOGO A"},
        ]))
        import core.match_eventos as me
        original = me.buscar_evento_na_grade
        me.buscar_evento_na_grade = lambda row, grade: {
            'horario_inicio': '11:00', 'horario_fim': ''  # grade sem FIM
        }
        try:
            me_datas = sp._datas_da_grade
            sp._datas_da_grade = lambda g: {"01/07/2026"}
            try:
                resultado = self._cruzar(df)
            finally:
                sp._datas_da_grade = me_datas
        finally:
            me.buscar_evento_na_grade = original

        self.assertIsNotNone(resultado)
        self.assertEqual(resultado.at[0, "inicio"], "11:00", "o início informado deve ser aplicado")
        self.assertEqual(resultado.at[0, "fim"], "12:15", "o fim NÃO pode ser apagado pela omissão")

    def test_pre_acompanha_o_inicio_quando_a_grade_muda_o_horario(self):
        """Caso real: Paulo Mancha, Pré 13:30 / Início 14:00 (consistente) na
        base. A grade de TV antecipou o jogo para Início 13:00, mas o código
        só atualizava Início/Fim — o Pré ficava congelado em 13:30, que virou
        POSTERIOR ao novo início. A convocação aparecia depois do jogo começar.
        """
        import core.schedule_processor as sp
        import core.match_eventos as me

        df = preparar_escala(pd.DataFrame([
            {"Nome": "Paulo Mancha", "Data": "11/10/2026", "Pré": "13:30",
             "Início": "14:00", "Fim": "17:00",
             "Evento/Descrição": "NFL - LAS VEGAS RAIDERS X NEW ENGLAND PATRIOTS"},
        ]))

        original = me.buscar_evento_na_grade
        me.buscar_evento_na_grade = lambda row, grade: {
            'pre_jogo': '12:30', 'pos_jogo': '',
            'horario_inicio': '13:00', 'horario_fim': '17:00',
        }
        datas_orig = sp._datas_da_grade
        sp._datas_da_grade = lambda g: {"11/10/2026"}
        try:
            resultado = self._cruzar(df)
        finally:
            me.buscar_evento_na_grade = original
            sp._datas_da_grade = datas_orig

        self.assertIsNotNone(resultado)
        self.assertEqual(resultado.at[0, "inicio"], "13:00")
        self.assertEqual(resultado.at[0, "pre"], "12:30",
                         "o Pré precisa acompanhar o novo Início informado pela grade")
        # Consistência: pré sempre antes (ou igual) do início, nunca depois.
        self.assertLessEqual(resultado.at[0, "pre"], resultado.at[0, "inicio"])

    def test_pre_antigo_e_preservado_quando_a_grade_nao_informa(self):
        """Mesma regra de 'a grade manda no que informa, não no que omite',
        agora para PRE: se a grade não trouxer convocação para o evento, o
        valor da escala anterior é mantido (não é apagado nem zerado)."""
        import core.schedule_processor as sp
        import core.match_eventos as me

        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "01/07/2026", "Pré": "09:30",
             "Início": "10:00", "Fim": "12:00", "Evento/Descrição": "JOGO A"},
        ]))

        original = me.buscar_evento_na_grade
        me.buscar_evento_na_grade = lambda row, grade: {
            'pre_jogo': '', 'pos_jogo': '',
            'horario_inicio': '10:00', 'horario_fim': '12:00',
        }
        datas_orig = sp._datas_da_grade
        sp._datas_da_grade = lambda g: {"01/07/2026"}
        try:
            resultado = self._cruzar(df)
        finally:
            me.buscar_evento_na_grade = original
            sp._datas_da_grade = datas_orig

        self.assertEqual(resultado.at[0, "pre"], "09:30")

    def test_termo_imune_em_outra_coluna_protege_a_linha(self):
        """'Day Off / Folga' vem em 'descricao', com 'evento' vazio."""
        from core.schedule_processor import _contexto_da_linha

        linha = pd.Series({
            "funcionario": "Ana", "data": "22/09/2026", "evento": "",
            "descricao": "Day Off / Folga", "tipo de atividade": "Other Time Off",
        })
        contexto = _contexto_da_linha(linha)
        self.assertIn("FOLGA", contexto)
        self.assertTrue(any(k in contexto for k in ["VIAGEM", "FOLGA"]))

    def test_linha_sem_evento_nunca_e_cancelada(self):
        df = preparar_escala(pd.DataFrame([
            {"Nome": "Ana", "Data": "01/07/2026", "Início": "00:00", "Fim": "00:00",
             "Evento/Descrição": "", "Descrição": "Bloqueio de agenda"},
        ]))
        import core.schedule_processor as sp
        import core.match_eventos as me

        original = me.buscar_evento_na_grade
        me.buscar_evento_na_grade = lambda row, grade: None
        datas_orig = sp._datas_da_grade
        sp._datas_da_grade = lambda g: {"01/07/2026"}
        try:
            resultado = self._cruzar(df, immune=("VIAGEM",))
        finally:
            me.buscar_evento_na_grade = original
            sp._datas_da_grade = datas_orig

        # Sem candidatos reais, o cruzamento não tem efeito e é ignorado —
        # o que importa é que nada virou CANCELADO.
        if resultado is not None:
            self.assertNotIn("CANCELADO", resultado["inicio"].astype(str).str.upper().tolist())


class TestLeituraDeGradeDensa(unittest.TestCase):
    """A grade mensal do Sportv perdia a data em 98% das linhas."""

    def test_datas_sobrevivem_ao_achatamento_de_blocos(self):
        """A DATA de cada bloco é vazia: a data da linha mora só no prefixo."""
        from core.match_eventos import carregar_grade

        grade = carregar_grade(GRADE)
        self.assertIsNotNone(grade)

        datas = pd.to_datetime(grade["DATA"], errors="coerce").dropna()
        proporcao = len(datas) / len(grade)
        self.assertGreater(proporcao, 0.90,
                           f"só {proporcao:.0%} das linhas têm data; o prefixo do bloco se perdeu")
        self.assertEqual(datas.dt.month.unique().tolist(), [fixtures.MES])

    def test_cabecalho_fora_da_primeira_linha_e_encontrado(self):
        from core.match_eventos import carregar_grade

        grade = carregar_grade(GRADE)
        self.assertIn("EVENTO/CAMPEONATO", grade.columns)
        self.assertTrue((grade["INICIO"].astype(str).str.strip() != "").any(),
                        "os horários do bloco precisam ser lidos")

    def test_blocos_lado_a_lado_sao_empilhados(self):
        """4 blocos de canal na mesma linha viram 4 linhas na grade."""
        from core.match_eventos import carregar_grade

        grade = carregar_grade(GRADE)
        eventos = set(grade["EVENTO/CAMPEONATO"].astype(str))
        for _, _, esperado in fixtures.GRADE_PADRAO:
            self.assertIn(esperado, eventos, f"evento perdido no achatamento: {esperado}")

    def test_desempate_por_proximidade_de_horario(self):
        """Mesmo programa exibido 2x no dia: vence a exibição mais próxima."""
        from core.match_eventos import buscar_na_grade

        grade = pd.DataFrame([
            {"DATA": pd.Timestamp("2026-09-21"), "EVENTO/CAMPEONATO": "SPORTV NEWS",
             "JOGO": "", "INICIO": "07:00", "FIM": "08:00", "PRE": "", "POS": ""},
            {"DATA": pd.Timestamp("2026-09-21"), "EVENTO/CAMPEONATO": "SPORTV NEWS",
             "JOGO": "", "INICIO": "21:00", "FIM": "22:00", "PRE": "", "POS": ""},
        ])
        _, _, inicio, _, _ = buscar_na_grade(
            grade, "21/09/2026", "SPORTV NEWS", inicio_rel="20:30"
        )
        self.assertEqual(inicio, "21:00", "deveria casar com a exibição da noite")

        _, _, inicio, _, _ = buscar_na_grade(
            grade, "21/09/2026", "SPORTV NEWS", inicio_rel="06:40"
        )
        self.assertEqual(inicio, "07:00", "deveria casar com a exibição da manhã")


class TestModoTeste(unittest.TestCase):
    """Homologação em pasta própria, sem tocar no banco de produção."""

    def setUp(self):
        import core.schedule_processor as sp
        self.sp = sp
        self.db = _db_temporario()
        self.processor = sp.ScheduleProcessor(self.db)
        self.saida = tempfile.mkdtemp(prefix="relatorios_teste_")
        self.processor.pasta_saida_teste = self.saida

        self.df_base = carregar_escala(ESCALA)
        self.base_id = self.db.save_new_grade(self.df_base, "base.xlsx", "carga_manual")
        self.db.update_employee_emails(
            {n: f"{n.split()[0].lower()}@teste.local" for n in self.df_base['funcionario'].unique()}
        )

        # Altera o horário de DUAS pessoas diferentes, para a homologação
        # gerar dois previews distintos.
        df_novo = carregar_escala(ESCALA)
        self.afetados = []
        for idx in df_novo.index:
            func = df_novo.at[idx, "funcionario"]
            if func in self.afetados or not str(df_novo.at[idx, "evento"]).strip():
                continue
            df_novo.at[idx, "inicio"] = "23:45"
            self.afetados.append(func)
            if len(self.afetados) == 2:
                break
        self.assertEqual(len(self.afetados), 2, "a fixture precisa ter 2+ funcionários")

        self.arquivo = os.path.join(tempfile.mkdtemp(), "nova.xlsx")
        df_novo.to_excel(self.arquivo, index=False)

    def _rodar(self, max_funcs=25):
        self.processor._processar_arquivo(
            "email-teste", self.arquivo, ["FOLGA"], True, "Assunto", max_funcs
        )
        pastas = [os.path.join(self.saida, d) for d in os.listdir(self.saida)]
        return pastas[0] if pastas else None

    def test_gera_pasta_com_relatorio_e_previews(self):
        pasta = self._rodar()
        self.assertIsNotNone(pasta, "o ciclo de teste deveria criar uma pasta")

        self.assertTrue(os.path.exists(os.path.join(pasta, "RESUMO.txt")))
        self.assertTrue(os.path.exists(os.path.join(pasta, "alteracoes.xlsx")))

        emails = os.path.join(pasta, "emails")
        self.assertTrue(os.path.isdir(emails))
        self.assertEqual(len(os.listdir(emails)), 2, "um preview por funcionário afetado")

        planilha = pd.read_excel(os.path.join(pasta, "alteracoes.xlsx"))
        self.assertEqual(len(planilha), 2)
        self.assertIn("23:45", planilha["Valor novo"].astype(str).tolist())

    def test_nome_da_pasta_identifica_o_ciclo(self):
        pasta = os.path.basename(self._rodar())
        self.assertRegex(pasta, r"^\d{8}_\d{6}_nova$")

    def test_nao_grava_alteracoes_no_banco(self):
        self._rodar()
        self.assertTrue(self.db.get_recent_changes().empty,
                        "homologação não pode poluir o histórico de produção")

    def test_nao_mexe_na_base_de_comparacao(self):
        self._rodar()
        self.assertEqual(self.db.get_baseline_id(), self.base_id)
        self.assertEqual(len(self.db.get_last_grade_df()), len(self.df_base))

    def test_nao_conta_como_email_enviado(self):
        self._rodar()
        self.assertEqual(self.db.get_dashboard_stats()["emails_enviados"], 0)

    def test_limite_nao_bloqueia_em_homologacao_mas_avisa(self):
        pasta = self._rodar(max_funcs=1)
        emails = os.path.join(pasta, "emails")
        self.assertEqual(len(os.listdir(emails)), 2, "em teste, todos os previews saem")

        with open(os.path.join(pasta, "RESUMO.txt"), encoding="utf-8") as f:
            resumo = f.read()
        self.assertIn("BLOQUEADO", resumo, "o relatório precisa avisar sobre o limite")

    def test_teste_nao_impede_processamento_real_depois(self):
        """Homologar e depois desligar o Modo Teste não pode pular o e-mail."""
        self._rodar()
        self.assertTrue(self.db.is_email_processed("email-teste", modo_teste=True))
        self.assertFalse(self.db.is_email_processed("email-teste", modo_teste=False),
                         "com envio real ligado, o e-mail precisa voltar à fila")

    def test_preview_sai_mesmo_sem_mapeamento_de_email(self):
        """A pasta emails/ vinha vazia quando ninguém estava mapeado."""
        self.db.update_employee_emails({}, substituir=True)

        pasta = self._rodar()
        arquivos = os.listdir(os.path.join(pasta, "emails"))
        self.assertEqual(len(arquivos), 2, "em homologação o preview sai sem destinatário")
        self.assertTrue(all(a.startswith("SEM_EMAIL_") for a in arquivos), arquivos)

        conteudo = open(os.path.join(pasta, "emails", arquivos[0]), encoding="utf-8").read()
        self.assertIn("SEM DESTINATÁRIO", conteudo)

        with open(os.path.join(pasta, "RESUMO.txt"), encoding="utf-8") as f:
            self.assertIn("Sem e-mail     : 2", f.read())

    def test_producao_nao_gera_nada_sem_email(self):
        """Fora da homologação, quem não tem e-mail apenas não é notificado."""
        self.db.update_employee_emails({}, substituir=True)
        self.processor._processar_arquivo(
            "email-prod", self.arquivo, ["FOLGA"], False, "Assunto", 25
        )
        changes = self.db.get_recent_changes()
        self.assertEqual(int(changes["email_enviado"].sum()), 0,
                         "sem destinatário não pode contar como enviado")

    def test_pasta_configuravel(self):
        from core.test_report import resolver_pasta_saida
        self.assertEqual(resolver_pasta_saida({"test_output_dir": "D:/x"}, "padrao"), "D:/x")
        self.assertEqual(resolver_pasta_saida({"test_output_dir": ""}, "padrao"), "padrao")
        self.assertEqual(resolver_pasta_saida({}, "padrao"), "padrao")


class TestCargaManualExclusiva(unittest.TestCase):
    """A carga manual não pode rodar junto com o ciclo."""

    def test_carga_durante_ciclo_e_recusada(self):
        import threading
        import core.schedule_processor as sp

        processor = sp.ScheduleProcessor(_db_temporario())
        liberar = threading.Event()
        entrou = threading.Event()

        processor._process_cycle = lambda: (entrou.set(), liberar.wait(timeout=5))
        t = threading.Thread(target=processor.process_cycle, daemon=True)
        t.start()
        entrou.wait(timeout=5)

        with self.assertRaises(sp.CicloEmAndamento):
            with processor.exclusivo(timeout=0.1):
                pass

        liberar.set()
        t.join(timeout=5)

        # Depois do ciclo, a carga passa normalmente.
        with processor.exclusivo(timeout=1):
            pass


class TestConcorrencia(unittest.TestCase):
    """Alto #10: agendador + botão manual disparando juntos."""

    def test_ciclo_concorrente_e_ignorado(self):
        import threading
        import core.schedule_processor as sp

        db = _db_temporario()
        processor = sp.ScheduleProcessor(db)

        entrou = threading.Event()
        liberar = threading.Event()
        execucoes = []

        def ciclo_lento():
            execucoes.append(1)
            entrou.set()
            liberar.wait(timeout=5)

        processor._process_cycle = ciclo_lento

        t = threading.Thread(target=processor.process_cycle, daemon=True)
        t.start()
        entrou.wait(timeout=5)

        processor.process_cycle()  # disparo concorrente
        liberar.set()
        t.join(timeout=5)

        self.assertEqual(len(execucoes), 1)


@precisa_planilhas_reais
class TestPlanilhasDeProducao(unittest.TestCase):
    """Validação extra contra os arquivos reais, quando estão na raiz.

    As fixtures sintéticas cobrem o comportamento; esta classe confirma que
    elas continuam fiéis ao formato que a operação realmente envia. Se uma
    planilha nova mudar de layout, é aqui que aparece primeiro.
    """

    def test_tipos_sao_classificados_corretamente(self):
        self.assertEqual(detectar_tipo_planilha(ESCALA_REAL), "escala")
        self.assertEqual(detectar_tipo_planilha(GRADE_REAL), "grade_tv")

    def test_colunas_canonicas_da_escala_real(self):
        df = carregar_escala(ESCALA_REAL)
        for col in ("funcionario", "data", "inicio", "fim", "evento"):
            self.assertIn(col, df.columns, f"coluna canônica ausente: {col}")

    def test_escala_real_identica_nao_gera_alteracao(self):
        db = _db_temporario()
        db.save_new_grade(carregar_escala(ESCALA_REAL), "base.xlsx", "id1")
        changes = compare_schedules(
            db.get_last_grade_df(), carregar_escala(ESCALA_REAL), 2
        )
        self.assertEqual(changes, [], f"{len(changes)} alterações falsas")

    def test_escala_real_nao_e_fatiada_como_grade(self):
        from core.match_eventos import carregar_grade
        grade = carregar_grade(ESCALA_REAL)
        self.assertIsNotNone(grade)
        self.assertIn("NOME", [str(c).upper().strip() for c in grade.columns])


if __name__ == "__main__":
    unittest.main(verbosity=2)
