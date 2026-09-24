"""Relatório de homologação — saída do Modo Teste em pasta própria.

Em Modo Teste nada é enviado e nada é gravado no banco de produção. Tudo que o
ciclo produziria vai para uma pasta datada, com os e-mails que sairiam e a
planilha das alterações detectadas, para conferência antes de ligar o envio.
"""

from __future__ import annotations

import datetime
import logging
import os
import re
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)

# Prefixo de timestamp que o OutlookMonitor acrescenta ao salvar o anexo.
_PREFIXO_ANEXO = re.compile(r"^\d{14}_")


def _slug(nome: str, limite: int = 48) -> str:
    base = _PREFIXO_ANEXO.sub("", os.path.splitext(str(nome))[0])
    base = re.sub(r"[^\w\s-]", "", base, flags=re.UNICODE).strip()
    base = re.sub(r"[\s_]+", "_", base)
    return base[:limite].strip("_") or "ciclo"


class RelatorioTeste:
    """Uma pasta por ciclo de homologação."""

    def __init__(self, raiz: str, arquivo: str, quando: Optional[datetime.datetime] = None):
        quando = quando or datetime.datetime.now()
        self.arquivo = os.path.basename(arquivo)
        self.pasta = os.path.join(
            raiz, f"{quando:%Y%m%d_%H%M%S}_{_slug(self.arquivo)}"
        )
        self.emails_dir = os.path.join(self.pasta, "emails")
        os.makedirs(self.emails_dir, exist_ok=True)
        self.inicio = quando

    # ------------------------------------------------------------------

    def salvar(
        self,
        changes: List[Dict[str, Any]],
        df_new: Optional[pd.DataFrame],
        baseline: Dict[str, Any],
        envio: Dict[str, List[str]],
        avisos: Optional[List[str]] = None,
    ) -> str:
        """Grava RESUMO.txt e alteracoes.xlsx. Devolve o caminho da pasta."""
        self._salvar_planilha(changes)
        self._salvar_resumo(changes, df_new, baseline, envio, avisos or [])
        logger.info(f"[MODO TESTE] Relatório do ciclo em: {self.pasta}")
        return self.pasta

    # ------------------------------------------------------------------

    def _salvar_planilha(self, changes: List[Dict[str, Any]]):
        if not changes:
            return
        destino = os.path.join(self.pasta, "alteracoes.xlsx")
        df = pd.DataFrame([
            {
                "Funcionário": c.get("funcionario", ""),
                "Data": c.get("data", ""),
                "Tipo": c.get("tipo", ""),
                "Campo": c.get("campo", ""),
                "Valor antigo": c.get("valor_antigo", ""),
                "Valor novo": c.get("valor_novo", ""),
            }
            for c in changes
        ])
        try:
            with pd.ExcelWriter(destino, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Alterações")
                (
                    df.groupby(["Funcionário", "Tipo"]).size()
                    .rename("Qtde").reset_index()
                    .to_excel(writer, index=False, sheet_name="Por funcionário")
                )
        except Exception as e:
            logger.error(f"[MODO TESTE] Falha ao gravar {destino}: {e}")

    def _salvar_resumo(self, changes, df_new, baseline, envio, avisos):
        afetados = sorted({c.get("funcionario", "") for c in changes})
        tipos: Dict[str, int] = {}
        for c in changes:
            tipos[c.get("tipo", "?")] = tipos.get(c.get("tipo", "?"), 0) + 1

        linhas = [
            "RELATÓRIO DE HOMOLOGAÇÃO — MODO TESTE",
            "=" * 62,
            "Nenhum e-mail foi enviado e nada foi gravado no banco de produção.",
            "",
            f"Gerado em      : {self.inicio:%d/%m/%Y %H:%M:%S}",
            f"Arquivo        : {self.arquivo}",
            f"Escala base    : {baseline.get('descricao', '-')}",
            f"Linhas na nova : {len(df_new) if df_new is not None else 0}",
            "",
            "ALTERAÇÕES",
            "-" * 62,
            f"Total          : {len(changes)}",
        ]
        for tipo, n in sorted(tipos.items()):
            linhas.append(f"  {tipo:<12} : {n}")
        linhas += [
            f"Funcionários   : {len(afetados)}",
            "",
            "NOTIFICAÇÕES QUE SERIAM ENVIADAS",
            "-" * 62,
            f"Com e-mail     : {len(envio.get('enviados', []))}",
            f"Sem e-mail     : {len(envio.get('sem_email', []))}",
            f"Com falha      : {len(envio.get('falhas', []))}",
        ]

        if envio.get("sem_email"):
            linhas += ["", "Sem e-mail no mapeamento (não seriam notificados):"]
            linhas += [f"  - {n}" for n in sorted(envio["sem_email"])]

        if avisos:
            linhas += ["", "AVISOS", "-" * 62] + [f"  ! {a}" for a in avisos]

        if afetados:
            linhas += ["", "FUNCIONÁRIOS AFETADOS", "-" * 62]
            linhas += [f"  - {n}" for n in afetados]

        linhas += [
            "",
            "ARQUIVOS",
            "-" * 62,
            "  emails/          previews HTML de cada notificação",
            "  alteracoes.xlsx  todas as alterações detectadas",
        ]

        destino = os.path.join(self.pasta, "RESUMO.txt")
        try:
            with open(destino, "w", encoding="utf-8") as f:
                f.write("\n".join(linhas) + "\n")
        except OSError as e:
            logger.error(f"[MODO TESTE] Falha ao gravar {destino}: {e}")


def resolver_pasta_saida(config: Dict[str, str], padrao: str) -> str:
    """Pasta configurada em 'test_output_dir', ou o padrão do projeto."""
    escolhida = str(config.get("test_output_dir", "") or "").strip()
    return escolhida or padrao
