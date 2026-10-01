"""Testes de unidade das partes que a suíte de regressão não exercitava.

A regressão cobre o fluxo ponta a ponta; aqui ficam os conversores, utilitários
e scripts isolados — onde um caso de borda passa despercebido com facilidade.
"""

import datetime
import os
import sys
import tempfile
import unittest

import pandas as pd

_AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(_AQUI))
sys.path.insert(0, _AQUI)

from core import columns  # noqa: E402
from core.database import DatabaseManager  # noqa: E402
from core.diff_engine import compare_schedules  # noqa: E402
from core.email_sender import EmailSender, _minutos, _normalizar_plataforma  # noqa: E402
from core.outlook_monitor import OutlookMonitor  # noqa: E402
from core.test_report import RelatorioTeste, _slug, resolver_pasta_saida  # noqa: E402


def _db_temporario() -> DatabaseManager:
    return DatabaseManager(db_path=os.path.join(tempfile.mkdtemp(), "t.db"))


class TestConversaoDeValores(unittest.TestCase):
    """`core/columns.py`: é daqui que vem a estabilidade do diff."""

    def test_hora_de_varios_formatos(self):
        casos = [
            ("14:30", "14:30"), ("14:30:00", "14:30"), ("9:05", "09:05"),
            (datetime.time(7, 0), "07:00"),
            (pd.Timestamp("2026-07-17 19:30"), "19:30"),
            (datetime.timedelta(hours=22, minutes=15), "22:15"),
            (0.8125, "19:30"),  # Excel guarda hora como fração do dia
            ("CANCELADO", "CANCELADO"),
            ("", ""), ("-", ""), ("nan", ""), (None, ""), (float("nan"), ""),
        ]
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(columns._fmt_hora(entrada), esperado)

    def test_data_de_varios_formatos(self):
        casos = [
            ("01/07/2026", "01/07/2026"), ("1/7/2026", "01/07/2026"),
            ("2026-07-01", "01/07/2026"),
            (pd.Timestamp("2026-07-01 19:30"), "01/07/2026"),
            (datetime.date(2026, 7, 1), "01/07/2026"),
            (1782864000000, "01/07/2026"),    # epoch ms — grades antigas
            ("1782864000000", "01/07/2026"),
            ("", ""), ("nan", ""), (None, ""),
        ]
        for entrada, esperado in casos:
            with self.subTest(entrada=entrada):
                self.assertEqual(columns._fmt_data(entrada), esperado)

    def test_normalizacao_e_idempotente(self):
        """Normalizar duas vezes não pode mudar o valor, senão o diff acusa."""
        df = pd.DataFrame({
            "Data": [pd.Timestamp("2026-07-01")],
            "Início": [pd.Timestamp("2026-07-01 14:30")],
            "WO#": [2490841.0],
            "Evento": ["JOGO"],
        })
        uma = columns.preparar_escala(df)
        duas = columns.preparar_escala(uma)
        pd.testing.assert_frame_equal(uma, duas)

    def test_float_inteiro_nao_ganha_casa_decimal(self):
        df = columns.preparar_escala(pd.DataFrame({"Nome": ["Ana"], "WO#": [2490841.0]}))
        self.assertEqual(df.at[0, "wo#"], "2490841")

    def test_prioridade_de_apelidos(self):
        """Com dois candidatos, vence o de maior prioridade."""
        df = columns.normalizar_colunas(pd.DataFrame(
            columns=["Atividade/Descrição", "Descrição", "Nome", "Elenco"]
        ))
        self.assertIn("evento", df.columns)
        self.assertIn("descricao", df.columns)   # o perdedor é preservado
        self.assertIn("funcionario", df.columns)
        self.assertIn("elenco", df.columns)      # não virou 'funcionario'

    def test_coluna_canonica_existente_nao_e_sobrescrita(self):
        df = columns.normalizar_colunas(pd.DataFrame(columns=["Evento", "Descrição"]))
        self.assertIn("evento", df.columns)
        self.assertIn("descricao", df.columns)


class TestDiffEngine(unittest.TestCase):

    def _par(self, antigo, novo):
        return compare_schedules(pd.DataFrame(antigo), pd.DataFrame(novo), 1)

    def test_coluna_removida_e_reportada(self):
        """Antes, cols_to_compare saía só de df_new e perdia o que sumiu."""
        changes = self._par(
            [{"Nome": "Ana", "Data": "01/07/2026", "Evento": "X", "Local": "ION"}],
            [{"Nome": "Ana", "Data": "01/07/2026", "Evento": "X"}],
        )
        self.assertEqual(changes, [], "coluna ausente não é alteração de valor")

    def test_reordenar_linhas_nao_gera_alteracao(self):
        """A chave inclui o evento: inserir linha no meio não desalinha."""
        base = [
            {"Nome": "Ana", "Data": "01/07/2026", "Evento": "A", "Início": "10:00"},
            {"Nome": "Ana", "Data": "01/07/2026", "Evento": "B", "Início": "20:00"},
        ]
        self.assertEqual(self._par(base, list(reversed(base))), [])

    def test_linha_inserida_no_meio_so_gera_entrada(self):
        antigo = [
            {"Nome": "Ana", "Data": "01/07/2026", "Evento": "A", "Início": "10:00"},
            {"Nome": "Ana", "Data": "01/07/2026", "Evento": "C", "Início": "20:00"},
        ]
        novo = antigo[:1] + [
            {"Nome": "Ana", "Data": "01/07/2026", "Evento": "B", "Início": "15:00"}
        ] + antigo[1:]
        changes = self._par(antigo, novo)
        self.assertEqual([c["tipo"] for c in changes], ["ENTRADA"])

    def test_funcionario_vazio_e_descartado(self):
        changes = self._par(
            [{"Nome": "Ana", "Data": "01/07/2026", "Evento": "A"}],
            [{"Nome": "Ana", "Data": "01/07/2026", "Evento": "A"},
             {"Nome": "", "Data": "01/07/2026", "Evento": "B"},
             {"Nome": "nan", "Data": "01/07/2026", "Evento": "C"}],
        )
        self.assertEqual(changes, [])

    def test_escala_nova_vazia_nao_apaga_ninguem(self):
        changes = self._par([{"Nome": "Ana", "Data": "01/07/2026", "Evento": "A"}], [])
        self.assertEqual(changes, [], "planilha vazia não pode gerar SAIDA em massa")


class TestFormatacaoDoEmail(unittest.TestCase):

    def test_minutos(self):
        self.assertEqual(_minutos("08:30"), 510)
        self.assertEqual(_minutos("00:00"), 0)
        self.assertEqual(_minutos("-"), 0)
        self.assertEqual(_minutos(None), 0)

    def test_plataforma(self):
        for entrada in ("Sportv", "Sportv 2", "Sportv 3", "SPORTV 4", "sportv5"):
            self.assertEqual(_normalizar_plataforma(entrada), "Sportv", entrada)
        for entrada in ("TV Globo - REDE", "Combate", "Premiere 3", "GE.com 01"):
            self.assertEqual(_normalizar_plataforma(entrada), entrada)
        self.assertEqual(_normalizar_plataforma(""), "")

    def test_dia_da_semana_derivado_da_data(self):
        self.assertEqual(EmailSender._format_dia_semana("21/09/2026"), "segunda-feira")
        self.assertEqual(EmailSender._format_dia_semana("25/09/2026"), "sexta-feira")

    def test_dia_textual_da_planilha_tem_prioridade(self):
        self.assertEqual(
            EmailSender._format_dia_semana("21/09/2026", "quarta-feira"), "quarta-feira"
        )

    def test_dia_invalido_nao_derruba(self):
        self.assertEqual(EmailSender._format_dia_semana("lixo", ""), "-")

    def test_produto_descarta_sufixos_inuteis(self):
        """'COPA.../2026/NA' deve exibir só a primeira parte útil."""
        df = columns.preparar_escala(pd.DataFrame([{
            "Nome": "Ana", "Data": "01/07/2026", "Início": "10:00", "Evento": "X",
            "Produto": "COPA DO MUNDO DE TENIS DE MESA/2026/NA",
        }]))
        html = EmailSender(test_mode=True)._generate_html_body("Ana", [], df)
        self.assertIn("COPA DO MUNDO DE TENIS DE MESA", html)
        self.assertNotIn("/2026/NA", html)

    def test_saudacao_por_horario(self):
        from core.email_sender import _saudacao
        self.assertEqual(_saudacao(datetime.datetime(2026, 9, 21, 9)), "bom dia")
        self.assertEqual(_saudacao(datetime.datetime(2026, 9, 21, 15)), "boa tarde")
        self.assertEqual(_saudacao(datetime.datetime(2026, 9, 21, 21)), "boa noite")

    def test_escala_sem_registro_nao_quebra(self):
        html = EmailSender(test_mode=True)._generate_html_body("Ana", [], None)
        self.assertIn("Nenhum registro", html)


class TestRelatorioTeste(unittest.TestCase):

    def test_slug_remove_prefixo_do_anexo(self):
        self.assertEqual(_slug("20260924161319_GRADE DE SETEMBRO.xlsm"), "GRADE_DE_SETEMBRO")
        self.assertEqual(_slug("escala.xlsx"), "escala")
        self.assertEqual(_slug(""), "ciclo")

    def test_pasta_de_saida_configuravel(self):
        self.assertEqual(resolver_pasta_saida({"test_output_dir": "D:/x"}, "p"), "D:/x")
        self.assertEqual(resolver_pasta_saida({"test_output_dir": "  "}, "p"), "p")
        self.assertEqual(resolver_pasta_saida({}, "p"), "p")

    def test_resumo_sem_alteracoes_nao_gera_planilha(self):
        raiz = tempfile.mkdtemp()
        rel = RelatorioTeste(raiz, "vazio.xlsx")
        rel.salvar([], pd.DataFrame(), {"descricao": "grade 1"},
                   {"enviados": [], "sem_email": [], "falhas": []})
        self.assertTrue(os.path.exists(os.path.join(rel.pasta, "RESUMO.txt")))
        self.assertFalse(os.path.exists(os.path.join(rel.pasta, "alteracoes.xlsx")))

    def test_resumo_lista_quem_ficou_sem_email(self):
        raiz = tempfile.mkdtemp()
        rel = RelatorioTeste(raiz, "x.xlsx")
        rel.salvar(
            [{"funcionario": "Ana", "data": "01/07/2026", "tipo": "ALTERACAO",
              "campo": "Inicio", "valor_antigo": "10:00", "valor_novo": "11:00"}],
            pd.DataFrame([{"a": 1}]), {"descricao": "grade 1"},
            {"enviados": [], "sem_email": ["Ana"], "falhas": []},
            avisos=["limite estourado"],
        )
        with open(os.path.join(rel.pasta, "RESUMO.txt"), encoding="utf-8") as f:
            texto = f.read()
        self.assertIn("Sem e-mail     : 1", texto)
        self.assertIn("- Ana", texto)
        self.assertIn("limite estourado", texto)
        planilha = pd.read_excel(os.path.join(rel.pasta, "alteracoes.xlsx"))
        self.assertEqual(planilha.at[0, "Valor novo"], "11:00")


class TestLimpezaDeAnexos(unittest.TestCase):

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.monitor = OutlookMonitor("Caixa de Entrada", "grade", self.dir)

    def _arquivo(self, nome, dias_atras):
        caminho = os.path.join(self.dir, nome)
        with open(caminho, "w") as f:
            f.write("x")
        quando = (datetime.datetime.now() - datetime.timedelta(days=dias_atras)).timestamp()
        os.utime(caminho, (quando, quando))
        return caminho

    def test_remove_apenas_os_antigos(self):
        velho = self._arquivo("velho.xlsx", 45)
        novo = self._arquivo("novo.xlsx", 2)

        self.monitor.limpar_anexos_antigos(dias=30)

        self.assertFalse(os.path.exists(velho))
        self.assertTrue(os.path.exists(novo))

    def test_zero_dias_desativa_a_limpeza(self):
        velho = self._arquivo("velho.xlsx", 999)
        self.monitor.limpar_anexos_antigos(dias=0)
        self.assertTrue(os.path.exists(velho))

    def test_pasta_inexistente_nao_derruba(self):
        monitor = OutlookMonitor("x", "y", self.dir)
        monitor.save_dir = os.path.join(self.dir, "nao_existe")
        monitor.limpar_anexos_antigos(dias=1)  # não deve levantar

    def test_keywords_aceita_lista_e_string(self):
        self.assertEqual(
            OutlookMonitor("x", "GRADE, PPV", self.dir).keywords, ["grade", "ppv"]
        )
        self.assertEqual(
            OutlookMonitor("x", ["Grade", "PPV"], self.dir).keywords, ["grade", "ppv"]
        )
        self.assertEqual(OutlookMonitor("x", None, self.dir).keywords, ["grade"])


class TestIntervaloDoAgendador(unittest.TestCase):
    """Um valor inválido na tela matava a thread do agendador em silêncio."""

    def setUp(self):
        import main
        self.fn = main._intervalo_valido

    def test_valores_validos(self):
        self.assertEqual(self.fn("15"), 15)
        self.assertEqual(self.fn(5), 5)
        self.assertEqual(self.fn("  7 "), 7)

    def test_valores_invalidos_caem_no_padrao(self):
        for ruim in ("", "abc", None, "2,5", []):
            self.assertEqual(self.fn(ruim, padrao=15), 15, repr(ruim))

    def test_intervalo_minimo(self):
        self.assertEqual(self.fn("0"), 1)
        self.assertEqual(self.fn("-5"), 1)


class TestScriptCorrigirBaseline(unittest.TestCase):
    """Marca grades duplicadas como 'ignored' e devolve a base correta."""

    def test_detecta_copia_e_devolve_a_base(self):
        import importlib.util
        db = _db_temporario()

        escala = pd.DataFrame([
            {"Nome": "Ana", "Data": "21/09/2026", "Início": "10:00", "Evento": "A"}
        ])
        id_boa = db.save_new_grade(escala, "escala_setembro.xlsx", "manual")
        # Duas cópias idênticas, como as grades no-op do ciclo automático
        db.save_new_grade(db.get_last_grade_df(), "grade_a.xlsx", "e1")
        db.save_new_grade(db.get_last_grade_df(), "grade_b.xlsx", "e2")

        self.assertNotEqual(db.get_baseline_id(), id_boa, "pré-condição: base soterrada")

        spec = importlib.util.spec_from_file_location(
            "corrigir_baseline",
            os.path.join(os.path.dirname(_AQUI), "scripts", "corrigir_baseline.py")
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)

        import sqlite3
        con = sqlite3.connect(db.db_path)
        assinaturas, duplicadas, vistas = {}, [], {}
        for (gid,) in con.execute("SELECT id FROM grades ORDER BY id"):
            sig = mod.assinatura(con, gid)
            if sig in vistas:
                duplicadas.append(gid)
            else:
                vistas[sig] = gid
        con.close()

        self.assertEqual(len(duplicadas), 2, "as duas cópias devem ser detectadas")

        with db.get_connection() as conn:
            conn.executemany("UPDATE grades SET status='ignored' WHERE id=?",
                             [(g,) for g in duplicadas])
        self.assertEqual(db.get_baseline_id(), id_boa)


class TestScriptGerarMapeamento(unittest.TestCase):

    def test_sugestao_de_endereco(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "gerar_mapeamento",
            os.path.join(os.path.dirname(_AQUI), "scripts", "gerar_mapeamento.py")
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)

        self.assertEqual(mod.sugerir("Marina Tavares", "exemplo.com"),
                         "marina.tavares@exemplo.com")
        self.assertEqual(mod.sugerir("José da Silva Neto", "exemplo.com"),
                         "jose.neto@exemplo.com")
        self.assertEqual(mod.sugerir("Pelezinho", "exemplo.com"), "pelezinho@exemplo.com")
        self.assertEqual(mod.sugerir("   ", "exemplo.com"), "")


if __name__ == "__main__":
    unittest.main(verbosity=2)
