"""Normalização canônica de colunas e valores das planilhas de escala.

Este módulo é a ÚNICA fonte de verdade sobre nomes de coluna no projeto.
Antes dele cada módulo (`schedule_processor`, `database`, `diff_engine`,
`email_sender`) mantinha a sua própria lista de apelidos, todas diferentes —
nenhuma cobria ao mesmo tempo `Atividade/Descrição`, `Início` e `Equipe`, que
são exatamente os nomes usados nas planilhas reais.

Dois problemas são resolvidos aqui:

1. **Nomes**: `Início` e `inicio` passam a ser a mesma coluna (`inicio`), então
   `df.at[idx, 'inicio']` atualiza a coluna real em vez de criar uma fantasma.
2. **Valores**: datas e horários viram string canônica (`dd/mm/aaaa`, `HH:MM`)
   *antes* de serem gravados. Sem isso o `to_json()` do banco transformava
   `Timestamp` em epoch (`1782864000000`) e a chave de comparação do diff nunca
   mais casava, fazendo uma planilha idêntica gerar SAIDA+ENTRADA para todo
   mundo.
"""

from __future__ import annotations

import datetime
import logging
import re
import unicodedata
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)

__all__ = [
    "sem_acento",
    "normalizar_nome_coluna",
    "normalizar_colunas",
    "coluna",
    "normalizar_valores",
    "preparar_escala",
    "carregar_escala",
    "detectar_tipo_planilha",
    "ALIASES",
    "CAMPOS_HORA",
    "CAMPOS_DATA",
]


# ---------------------------------------------------------------------------
# Apelidos conhecidos -> nome canônico
# ---------------------------------------------------------------------------
# A ORDEM DENTRO DE CADA LISTA É PRIORIDADE. Quando a planilha traz mais de um
# candidato (ex.: `Atividade/Descrição` e `Descrição`), vence o primeiro da
# lista; os demais permanecem com o nome normalizado original.

ALIASES: Dict[str, List[str]] = {
    "funcionario": [
        "funcionario", "nome", "escalado", "talento", "profissional",
        "colaborador", "resource", "equipe", "elenco",
    ],
    "data": ["data", "data escala", "data_escala", "data da escala", "date"],
    "dia": ["dia", "dia da semana", "dia semana"],
    "inicio": ["inicio", "hora inicio", "horario inicio", "horario", "start", "hora"],
    "fim": ["fim", "termino", "hora fim", "end"],
    "pre": ["pre", "convocacao", "pre jogo"],
    "pos": ["pos", "pos jogo"],
    "evento": [
        "evento/descricao", "evento/programa", "evento", "atividade/descricao",
        "atividade", "programa", "descricao", "tipo atividade",
    ],
    "event_group": ["event group", "event_group", "grupo de evento"],
    "produto": [
        "produto", "produto (wo/shift)", "produto (wo/quick hold)", "produto (wo)",
    ],
    "local": [
        "local de locucao", "local locucao", "local narracao", "local de narracao",
        "local de gravacao", "local gravacao", "local",
    ],
    "plataforma": ["plataforma", "canal", "canal (master room)"],
    "coordenador": ["coordenador"],
    "produtor": ["produtor"],
    "narrador": ["narrador"],
    "comentarista": ["comentarista"],
    "reporter": ["reporter"],
    "elenco": ["elenco"],
    "sonora": ["sonora"],
    "funcao": ["funcao"],
    "status": ["status"],
}

# Campos cujo valor deve ser normalizado como horário (HH:MM).
CAMPOS_HORA = {"inicio", "fim", "pre", "pos"}

# Campos cujo valor deve ser normalizado como data (dd/mm/aaaa).
CAMPOS_DATA = {"data"}

# Colunas que jamais devem ser promovidas a canônicas mesmo contendo o termo.
# `Data_raw` existe na planilha real e é o que fazia `carregar_grade()` achar
# que havia dois blocos de grade lado a lado.
BLOQUEADAS = {"data_raw", "data raw"}

_VAZIOS = {"", "-", "nan", "none", "nat", "null", "<na>"}


# ---------------------------------------------------------------------------
# Nomes de coluna
# ---------------------------------------------------------------------------

def sem_acento(texto: Any) -> str:
    """Remove acentos preservando o restante do texto."""
    s = str(texto)
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch)
    )


def normalizar_nome_coluna(nome: Any) -> str:
    """`'  Início '` -> `'inicio'`, `'Atividade/Descrição'` -> `'atividade/descricao'`."""
    s = sem_acento(nome).strip().lower()
    return re.sub(r"\s+", " ", s)


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza os nomes das colunas e promove apelidos conhecidos a canônicos.

    Colunas desconhecidas são mantidas (apenas normalizadas), para que o diff
    continue comparando tudo que a planilha traz.
    """
    if df is None or df.empty and len(df.columns) == 0:
        return df

    df = df.copy()
    df.columns = [normalizar_nome_coluna(c) for c in df.columns]

    # Colunas duplicadas após a normalização: mantém a primeira.
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    atuais = set(df.columns)
    renomear: Dict[str, str] = {}
    ja_usados: set = set()

    for canonico, apelidos in ALIASES.items():
        if canonico in atuais:
            # Já veio com o nome certo; nada a promover.
            ja_usados.add(canonico)
            continue
        for apelido in apelidos:
            if apelido in BLOQUEADAS:
                continue
            if apelido in atuais and apelido not in renomear and apelido not in ja_usados:
                renomear[apelido] = canonico
                ja_usados.add(apelido)
                break

    if renomear:
        df = df.rename(columns=renomear)
        df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    return df


def coluna(df: pd.DataFrame, canonico: str) -> Optional[str]:
    """Devolve o nome real da coluna canônica no DataFrame, ou None."""
    if df is None or canonico in BLOQUEADAS:
        return None
    if canonico in df.columns:
        return canonico
    for apelido in ALIASES.get(canonico, []):
        if apelido in df.columns:
            return apelido
    return None


# ---------------------------------------------------------------------------
# Valores
# ---------------------------------------------------------------------------

def _eh_vazio(v: Any) -> bool:
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    return str(v).strip().lower() in _VAZIOS


def _fmt_hora(v: Any) -> str:
    """Qualquer representação de horário -> `'HH:MM'` (ou string original)."""
    if _eh_vazio(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        return v.strftime("%H:%M")
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M")
    if isinstance(v, datetime.timedelta):
        total = int(v.total_seconds()) // 60
        return f"{(total // 60) % 24:02d}:{total % 60:02d}"
    if isinstance(v, float) and 0 < v < 1:
        # Excel guarda hora como fração do dia quando a célula não é formatada.
        minutos = round(v * 24 * 60)
        return f"{(minutos // 60) % 24:02d}:{minutos % 60:02d}"

    s = str(v).strip()
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return s


def _fmt_data(v: Any) -> str:
    """Qualquer representação de data -> `'dd/mm/aaaa'`.

    Aceita epoch em milissegundos, que é como as grades antigas ficaram
    gravadas no banco por causa do `to_json()`. Isso faz a migração do histórico
    acontecer sozinha.
    """
    if _eh_vazio(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, datetime.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if abs(v) > 1e11:  # epoch em ms (legado)
            try:
                return pd.Timestamp(int(v), unit="ms").strftime("%d/%m/%Y")
            except (ValueError, OverflowError, OSError):
                return str(v)
        return str(v)

    s = str(v).strip()
    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", s):
        d, m, a = s.split("/")
        return f"{int(d):02d}/{int(m):02d}/{a}"
    if re.fullmatch(r"-?\d{12,}", s):  # epoch em ms que chegou como string
        try:
            return pd.Timestamp(int(s), unit="ms").strftime("%d/%m/%Y")
        except (ValueError, OverflowError, OSError):
            return s
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%d/%m/%Y")
    except (ValueError, TypeError, pd.errors.ParserError):
        return s


def _fmt_generico(v: Any) -> str:
    if _eh_vazio(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%d/%m/%Y")
        return v.strftime("%d/%m/%Y %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M")
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def normalizar_valores(df: pd.DataFrame) -> pd.DataFrame:
    """Converte todo o DataFrame para strings canônicas e estáveis.

    Depois desta função, salvar no banco e reler devolve exatamente os mesmos
    valores — que é o que torna a comparação do diff confiável.
    """
    if df is None or df.empty:
        return df

    df = df.copy()
    for col in df.columns:
        if col in CAMPOS_HORA:
            df[col] = df[col].map(_fmt_hora)
        elif col in CAMPOS_DATA:
            df[col] = df[col].map(_fmt_data)
        else:
            df[col] = df[col].map(_fmt_generico)
    return df


def preparar_escala(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza nomes e valores — o ponto de entrada único para qualquer escala."""
    if df is None:
        return pd.DataFrame()
    return normalizar_valores(normalizar_colunas(df))


# ---------------------------------------------------------------------------
# Leitura de arquivos
# ---------------------------------------------------------------------------

def carregar_escala(caminho: str) -> pd.DataFrame:
    """Lê uma escala individual (.xlsx/.xlsm/.xls/.csv) já normalizada."""
    if str(caminho).lower().endswith(".csv"):
        df = pd.read_csv(caminho)
    else:
        df = pd.read_excel(caminho)
    return preparar_escala(df)


# Termos que identificam a coluna de pessoa numa escala individual. Uma grade
# de TV não tem nenhuma delas — é essa a diferença que distingue os dois tipos.
_TERMOS_PESSOA = {
    "nome", "funcionario", "escalado", "talento", "profissional",
    "colaborador", "resource", "equipe", "elenco",
}
_TERMOS_GRADE = {
    "jogo", "campeonato", "evento/campeonato", "mandante", "visitante",
    "evento", "evento/programa", "programa", "transmissao",
}


def detectar_tipo_planilha(caminho: str) -> str:
    """Classifica o anexo em `'escala'`, `'grade_tv'` ou `'desconhecido'`.

    A decisão é tomada sobre o cabeçalho BRUTO do arquivo. Antes, o sistema
    dependia de `carregar_grade()` devolver None para reconhecer uma escala —
    mas ela nunca devolvia, porque o achatamento de "blocos lado a lado"
    descartava justamente a coluna `Nome`. O resultado era a escala recebida ser
    tratada como grade de TV e descartada inteira.
    """
    caminho_l = str(caminho).lower()

    if caminho_l.endswith(".csv"):
        try:
            cabecalho = [normalizar_nome_coluna(c) for c in pd.read_csv(caminho, nrows=0).columns]
            return "escala" if _tem_pessoa(cabecalho) else "desconhecido"
        except Exception as e:
            logger.warning(f"Não foi possível ler o cabeçalho de {caminho}: {e}")
            return "desconhecido"

    try:
        import openpyxl

        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
        try:
            achou_grade = False
            for aba in wb.worksheets:
                for i, linha in enumerate(aba.iter_rows(values_only=True), start=1):
                    if i > 25:
                        break
                    celulas = [normalizar_nome_coluna(c) for c in linha if c is not None]
                    if not celulas:
                        continue
                    if _tem_pessoa(celulas):
                        return "escala"
                    # Grade de TV: tem data + descrição de evento, mas nenhuma
                    # coluna de pessoa (é a programação do canal, não a escala).
                    if "data" in celulas and any(t in celulas for t in _TERMOS_GRADE):
                        achou_grade = True
            return "grade_tv" if achou_grade else "desconhecido"
        finally:
            wb.close()
    except Exception as e:
        logger.warning(f"Não foi possível classificar {caminho}: {e}")
        return "desconhecido"


def _tem_pessoa(celulas: List[str]) -> bool:
    return any(c in _TERMOS_PESSOA for c in celulas)
