from __future__ import annotations

# =============================================================================
# match_eventos.py  —  Matching inteligente Relatório ↔ Grade GE TV
# =============================================================================
#
# COMO INTEGRAR AO PROJETO:
# --------------------------
# 1. Copie este arquivo para a pasta do projeto.
#
# 2. No gerar_html_escala.py, troque a função de carregamento e busca:
#
#       from match_eventos import carregar_grade, buscar_na_grade
#       grade = carregar_grade("GRADE_EVENTOS_...xlsx")
#
# 3. Assinatura ESTENDIDA de buscar_na_grade (backward-compatible):
#
#       pre, pos, inicio, fim = buscar_na_grade(
#           grade,
#           data_str   = "23/02/2026",    # igual à original
#           evento_str = "FLAMENGO X VASCO",
#           # Novos parâmetros opcionais — forneça quando disponíveis:
#           event_group = row['Event Group'],  # coluna do relatório
#           inicio_rel  = row['Início'],        # horário do profissional no rel
#           fim_rel     = row['Fim'],
#       )
#
# 4. diagnosticar_match() aceita os mesmos parâmetros extras.
#
# =============================================================================
# ESTRATÉGIA DE MATCHING (3 sinais combinados)
# =============================================================================
# Os 3 sinais são calculados por linha da grade filtrada pelo mesmo dia:
#
#   SINAL 1 — TEXTO (peso base 50%)
#     Compara "Evento/Programa" do REL com "JOGO + EVENTO/CAMPEONATO" da grade.
#     Dois modos adaptativos:
#       • Confronto (tem " X "): estratégia de times domina.
#       • Programa (sem " X "):  token_sort + LCS dominam; 1º token igual
#                                garante score mínimo 0.50.
#
#   SINAL 2 — EVENT GROUP (peso bônus até +0.30)
#     Compara "Event Group" do relatório com "EVENTO/CAMPEONATO" da grade.
#     Jaccard normalizado — captura "COPA DO BRASIL DE FUTEBOL" vs "Copa do Brasil".
#     Quando score de texto já é alto (≥0.60) o bônus é reduzido (evita dupla
#     contagem). Quando texto é fraco mas Event Group bate bem, o bônus é maior.
#
#   SINAL 3 — SOBREPOSIÇÃO DE HORÁRIO (bônus binário +0.20)
#     Verifica se o horário do profissional no relatório (Início..Fim) cai dentro
#     da janela PRÉ..FIM do evento na grade.
#     Resolve casos de adversário indefinido onde o texto falha mas o horário
#     confirma o match (ex: "FORTALEZA X MAGUARY" vs "FORTALEZA X ADV. INDEFINIDO").
#
# Score final = min(s_texto + bônus_eg + bônus_hora, 1.0)
# Threshold padrão = 0.42
#
# =============================================================================
# DESCOBERTAS DOS ARQUIVOS REAIS
# =============================================================================
# • Grade lida via openpyxl (aba "GRADE GE TV 2026") — pandas perde os horários.
# • Grade tem sufixos de estado: "FLAMENGO - RJ" → normalizar remove " - RJ".
# • "Event Group" do relatório = campeonato completo (ex: "COPA DO BRASIL DE FUTEBOL")
#   → comparar com EVENTO/CAMPEONATO da grade é muito eficaz.
# • Eventos da Saudi League, F1, etc. não existem na grade → retornar vazio.
# =============================================================================

import re
import unicodedata
from difflib import SequenceMatcher
from typing import Optional, Tuple

import pandas as pd
from datetime import datetime as _dt, timedelta

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

THRESHOLD_PADRAO = 0.42

def safe_print(msg: str):
    """Imprime mensagens tratando erros de encoding no console do Windows."""
    try:
        print(msg)
    except Exception:
        try:
            # Fallback removendo caracteres incompatíveis (cp1252)
            clean_msg = str(msg).encode('ascii', 'replace').decode('ascii')
            print(clean_msg)
        except Exception:
            pass



# ---------------------------------------------------------------------------
# Normalização
# ---------------------------------------------------------------------------

_SUFIXOS = re.compile(
    r"\s*[-–]\s*(?:AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|"
    r"RJ|RN|RS|RO|RR|SC|SP|SE|TO|BOL|ARG|VEN|VEM|COL|URU|PAR|ECU|CHI|PER)\b",
    re.IGNORECASE,
)
_SEP    = re.compile(r"\s+(?:VS\.?|×|x)\s+", re.IGNORECASE)
_TRACO  = re.compile(r"\s*[-–—]\s*")
_DESCAR = re.compile(
    r"\b(?:CS|4K|UHD|AO\s+VIVO|HD|EXIBICAO|EXIBIÇÃO|TRANSMISSAO|TRANSMISSÃO|"
    r"JOGO\s+\d+|FASE\s+UNICA|FASE\s+ÚNICA|\d+[ªº]\s+FASE|RODADA\s+\d+|"
    r"2A\s+FASE|3A\s+FASE)\b",
    re.IGNORECASE,
)


def _ascii(t: str) -> str:
    return unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()


def normalizar(texto: Optional[str]) -> str:
    """Normalização completa: ASCII, remoção de sufixos de estado, padronização de separadores."""
    if not texto or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    t = _ascii(str(texto).upper().strip())
    t = _SUFIXOS.sub("", t)
    t = _SEP.sub(" X ", t)
    t = _TRACO.sub(" ", t)
    t = _DESCAR.sub(" ", t)
    t = re.sub(r"[^\w\s]", " ", t)
    # Especial: normalizar GETV para GE TV
    t = re.sub(r"\bGETV\b", "GE TV", t)
    return re.sub(r"\s+", " ", t).strip()


def _texto_grade(campeonato: str, jogo: str) -> str:
    """Combina JOGO + EVENTO/CAMPEONATO em texto normalizado único."""
    j = normalizar(jogo)
    e = normalizar(campeonato)
    return f"{j} {e}".strip() if j and j not in ("NAN", "NONE", "") else e


# ---------------------------------------------------------------------------
# Sinal 1 — Similaridade de texto
# ---------------------------------------------------------------------------

def _jaccard(a: str, b: str) -> float:
    wa, wb = set(a.split()), set(b.split())
    return len(wa & wb) / len(wa | wb) if (wa and wb) else 0.0


def _token_sort(a: str, b: str) -> float:
    return _jaccard(" ".join(sorted(a.split())), " ".join(sorted(b.split())))


def _lcs(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def _confronto(a: str, b: str) -> float:
    """Compara os times num padrão 'TIME1 X TIME2'."""
    pat = re.compile(r"(.+?)\s+X\s+(.+)")
    ma, mb = pat.search(a), pat.search(b)
    if not ma or not mb:
        return 0.0
    ta = {ma.group(1).strip(), ma.group(2).strip()}
    tb = {mb.group(1).strip(), mb.group(2).strip()}
    adv_a = any("ADV" in t or "INDEFINIDO" in t for t in ta)
    adv_b = any("ADV" in t or "INDEFINIDO" in t for t in tb)
    comuns = ta & tb
    if len(comuns) == 2:
        return 1.0
    if len(comuns) == 1:
        return 0.7 if (adv_a or adv_b) else 0.5
        
    # Camada extra: verificar se pelo menos um time bate parcialmente
    # (ex: INTERNACIONAL vs INTER)
    for t1 in ta:
        for t2 in tb:
            if t1 and t2 and (t1 in t2 or t2 in t1 or _jaccard(t1, t2) > 0.6):
                # Se um time bate, e o outro é comum ou similar
                outros_a = ta - {t1}
                outros_b = tb - {t2}
                if outros_a and outros_b:
                    oa3, ob3 = list(outros_a)[0], list(outros_b)[0]
                    if oa3 in ob3 or ob3 in oa3 or _jaccard(oa3, ob3) > 0.6:
                        return 0.85
                return 0.45
    return 0.0


def _bonus_comp(a: str, b: str) -> float:
    COMP = {
        "LIBERTADORES", "SULAMERICANA", "COPA", "BRASIL", "BRASILEIRAO",
        "CARIOCA", "MINEIRO", "GAUCHO", "PAULISTA", "PERNAMBUCANO",
        "FINAL", "SEMIFINAL", "QUARTAS", "OITAVAS", "CHAMPIONS",
        "MUNDIAL", "SUPERCOPA", "MGM", "SLAM", "SURF", "JUDO", "JUDÔ",
        "VOLEY", "VOLEI", "TENIS", "MESA", "SKATE", "FORMULA", "NFL",
    }
    wa, wb = set(a.split()), set(b.split())
    return min(len((wa & wb) & COMP) * 0.12, 0.36)


def _score_texto(a: str, b: str) -> float:
    """Score de texto puro entre dois textos normalizados."""
    if not a or not b: return 0.0
    if a == b: return 1.0
    # Substring check
    if len(a) > 6 and (a in b or b in a): return 0.9
    if a in b or b in a:
        return 0.92

    tem_conf = " X " in a or " X " in b

    if tem_conf:
        score = (
            _confronto(a, b) * 0.45
            + _token_sort(a, b) * 0.25
            + _lcs(a, b) * 0.15
            + _jaccard(a, b) * 0.10
            + _bonus_comp(a, b) * 0.05
        )
    else:
        lista_a, lista_b = a.split(), b.split()
        primeiro_token_igual = (
            bool(lista_a and lista_b)
            and lista_a[0] == lista_b[0]
            and len(lista_a[0]) > 3
        )
        score_base = (
            _token_sort(a, b) * 0.50
            + _lcs(a, b) * 0.30
            + _jaccard(a, b) * 0.20
        )
        score = (
            max(score_base, 0.50) if primeiro_token_igual
            else _token_sort(a, b) * 0.40 + _lcs(a, b) * 0.25 + _jaccard(a, b) * 0.15 + _bonus_comp(a, b) * 0.20
        )

    return round(min(score, 1.0), 4)


# ---------------------------------------------------------------------------
# Sinal 2 — Event Group vs EVENTO/CAMPEONATO
# ---------------------------------------------------------------------------

def _score_event_group(event_group: Optional[str], campeonato_grade: Optional[str]) -> float:
    """
    Compara o 'Event Group' do relatório com 'EVENTO/CAMPEONATO' da grade.
    Retorna Jaccard [0,1]. Ex: "COPA DO BRASIL DE FUTEBOL" vs "Copa do Brasil" → 0.60.
    """
    if not event_group or not campeonato_grade:
        return 0.0
    a = normalizar(event_group)
    b = normalizar(campeonato_grade)
    
    # Check substring exata (se um está contido no outro)
    if a and b and (a in b or b in a) and len(a) > 4:
        return 0.86
        
    if not a or not b:
        return 0.0
    if a == b or a in b or b in a:
        return 1.0
    return round(_jaccard(a, b), 4)


# ---------------------------------------------------------------------------
# Sinal 3 — Sobreposição de horário
# ---------------------------------------------------------------------------

def _hora_min(v) -> Optional[int]:
    """Converte horário (string HH:MM ou datetime.time) para minutos desde meia-noite."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if hasattr(v, "hour"):       # datetime.time
        return v.hour * 60 + v.minute
    s = str(v).strip()
    if s in {"", "-", "nan", "NaN", "None", "NAT", "NaT", "nat", "<NA>"}:
        return None
    try:
        parts = s.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    except (ValueError, IndexError):
        return None


def _sobreposicao(
    inicio_rel: Optional[str],
    fim_rel: Optional[str],
    pre_grade,
    fim_grade,
) -> bool:
    """
    Retorna True se o horário do profissional (inicio_rel..fim_rel) sobrepõe
    a janela PRÉ..FIM do evento na grade.
    Trata cruzamento de meia-noite em ambos os lados.
    """
    ir = _hora_min(inicio_rel)
    fr = _hora_min(fim_rel)
    pg = _hora_min(pre_grade)
    fg = _hora_min(fim_grade)

    if ir is None:
        return False
    if pg is None:
        pg = _hora_min(fim_grade)
        if pg is None:
            return False
    if fg is None:
        return False

    if fr is None:
        fr = ir + 120

    if fg < pg:
        fg += 1440
    if fr < ir:
        fr += 1440
    for offset in (0, 1440, -1440):
        ir_ = ir + offset
        fr_ = fr + offset
        if ir_ < fg and fr_ > pg:
            return True
    return False


# ---------------------------------------------------------------------------
# Função principal de score combinado
# ---------------------------------------------------------------------------

def calcular_score_combinado(
    evento_str: Optional[str],
    campeonato_grade: str,
    jogo_grade: str,
    event_group: Optional[str] = None,
    inicio_rel: Optional[str] = None,
    fim_rel: Optional[str] = None,
    pre_grade = None,
    fim_grade_val = None,
) -> float:
    """
    Calcula score combinado (3 sinais) entre um evento do REL e uma linha da grade.
    Retorna float [0, 1].
    """
    ev_n  = normalizar(evento_str)
    tg_n  = _texto_grade(campeonato_grade, jogo_grade)
    s_txt = _score_texto(ev_n, tg_n)

    s_eg = _score_event_group(event_group, campeonato_grade)
    peso_eg = 0.30 if s_txt < 0.40 else (0.20 if s_txt < 0.60 else 0.10)
    bonus_eg = s_eg * peso_eg

    # --- NOVO: Penalidade ABSOLUTA por conflito de times (Evita misturar Flamengo e Palmeiras) ---
    def _conflito_de_times(a, b):
        # Lista expandida de times recorrentes na Libertadores e Sul-Americana
        times = [
            "FLAMENGO", "PALMEIRAS", "SANTOS", "SAO PAULO", "CORINTHIANS", "FLU", 
            "VASCO", "BOTAFOGO", "GREMIO", "INTER", "ATLETICO", "CRUZEIRO",
            "JUNIOR", "BARRANQUILLA", "CUSCO", "CHILE", "COLO", "CATHOLICA", "URUGUAI", 
            "PENAROL", "NACIONAL", "LIBERTAD", "CERRO", "BOLIVAR", "STRONGEST", 
            "BARCELONA", "LIGA", "QUITO", "MILLONARIOS", "CARACAS", "TOLIMA"
        ]
        ev_tokens = set(a.split())
        tg_tokens = set(b.split())
        
        # Se um cita um time e o outro cita OUTRO time da lista, é conflito total
        times_no_rel = [t for t in times if t in ev_tokens]
        times_na_grade = [t for t in times if t in tg_tokens]
        
        if times_no_rel and times_na_grade:
            if not any(t in times_na_grade for t in times_no_rel):
                return True
        return False

    if _conflito_de_times(ev_n, tg_n):
        s_txt = 0.05 # Força o descarte quase total do match por inconsistência de times

    tem_hora = inicio_rel is not None
    bonus_hora = (
        0.20 if (tem_hora and _sobreposicao(inicio_rel, fim_rel, pre_grade, fim_grade_val))
        else 0.0
    )

    return round(min(s_txt + bonus_eg + bonus_hora, 1.0), 4)


# ---------------------------------------------------------------------------
# API pública — manter compatibilidade com código anterior
# ---------------------------------------------------------------------------

def calcular_similaridade(
    texto_rel: Optional[str],
    texto_grade: Optional[str],
    threshold: float = 0.0,
) -> float:
    """
    Compatibilidade com a assinatura original.
    Usa apenas sinal de texto (sem Event Group / horário).
    Para melhor precisão use calcular_score_combinado() diretamente.
    """
    if not texto_rel or not texto_grade:
        return 0.0
    a = normalizar(texto_rel)
    b = normalizar(texto_grade)
    if not a or not b:
        return 0.0
    score = _score_texto(a, b)
    return 0.0 if (threshold and score < threshold) else score


# ---------------------------------------------------------------------------
# Carregamento correto da grade
# ---------------------------------------------------------------------------

def carregar_grade(caminho: str) -> Optional[pd.DataFrame]:
    """
    Lê a grade de eventos usando openpyxl.
    Procura o cabeçalho dinamicamente nas primeiras 20 linhas.
    Preserva os tipos datetime.time dos horários que pandas.read_excel() perde.
    Retorna DataFrame com colunas normalizadas ou None em caso de erro.
    """
    try:
        import openpyxl

        wb = openpyxl.load_workbook(caminho, data_only=True)
        aba = None
        ano_str = str(_dt.now().year)

        # 1. Tentar os nomes exatos preferidos (Evita 'v. Rasc')
        for nome_prioridade in [f"GRADE GE TV {ano_str}", f"GRADE GE TV {ano_str} 1 POR DIA"]:
            if nome_prioridade in wb.sheetnames:
                aba = wb[nome_prioridade]
                break
        
        # 2. Se não achar, procura qualquer aba que tenha GRADE e GE TV e não seja rascunho
        if aba is None:
            for nome in wb.sheetnames:
                nome_upper = nome.upper()
                if "GRADE" in nome_upper and "GE TV" in nome_upper and "RASC" not in nome_upper:
                    aba = wb[nome]
                    break
                    
        # 3. Fallback
        if aba is None:
            aba = wb.active

        # Procurar cabeçalho dinamicamente nas primeiras 20 linhas
        header_row_idx = None
        headers = []
        rows = []
        for i, row in enumerate(aba.iter_rows(values_only=True), start=1):
            rows.append(row)
            if i > 20: break
            row_vals = [str(c).upper().strip() if c else "" for c in row]
            if any("DATA" in v for v in row_vals) and (
                any("IN" in v for v in row_vals) or 
                any("JOGO" in v for v in row_vals) or 
                any("START" in v for v in row_vals) or 
                any("HORA" in v for v in row_vals) or
                any("EVENTO" in v for v in row_vals)
            ):
                header_row_idx = i
                
                # Normalizar cabeçalhos: maiúsculas, sem acentos, sem espaços extras
                for idx, c in enumerate(row):
                    if c:
                        # Remover acentos básicos dos cabeçalhos pra padronizar "INÍCIO" -> "INICIO"
                        h = str(c).upper().strip()
                        import unicodedata
                        h = "".join(ch for ch in unicodedata.normalize('NFKD', h) if not unicodedata.combining(ch))
                        headers.append(h)
                    else:
                        headers.append(f"COL_{idx}")
                break

        if header_row_idx is None:
            safe_print("[ERRO] Nenhuma linha de cabeçalho contendo 'DATA' encontrada nas primeiras 20 linhas.")
            return None

        # Ler dados 
        rows = []
        for row in aba.iter_rows(min_row=header_row_idx + 1, values_only=True):
            rows.append(row)
        
        # Deduplicar e empilhar (flatten) blocos lado a lado se houver múltiplos "DATA"
        # Isso acontece em grades da TV fechada (ex: Sportv 1, 2, 3 lado a lado)
        #
        # A comparação precisa ser EXATA. Com `"DATA" in h`, a coluna `Data_raw`
        # da escala individual contava como um segundo bloco: a planilha era
        # fatiada a partir da 1ª coluna DATA e perdia `Nome`, `Tipo Atividade` e
        # `Descrição` — o que fazia a escala ser confundida com grade de TV e
        # descartada inteira.
        data_indices = [idx for idx, h in enumerate(headers) if h.strip() == "DATA"]
        
        if len(data_indices) > 1:
            safe_print(f"Detectados {len(data_indices)} blocos lado a lado. Planificando...")
            flattened_rows = []

            # Colunas à ESQUERDA do primeiro bloco são contexto comum da linha
            # (ex.: 'DATA GRADE', 'DIA DA SEMANA'). Descartá-las fazia a grade
            # de setembro perder a data em 98% das linhas: as colunas DATA de
            # cada bloco só são preenchidas esporadicamente, enquanto a data
            # real da linha mora justamente nesse prefixo.
            prefix_headers = headers[:data_indices[0]]

            for row in rows:
                prefixo = dict(zip(prefix_headers, row[:data_indices[0]]))

                for i in range(len(data_indices)):
                    start_idx = data_indices[i]
                    end_idx = data_indices[i+1] if i+1 < len(data_indices) else len(headers)

                    sub_row = row[start_idx:end_idx]
                    sub_headers = headers[start_idx:end_idx]

                    # Só adiciona se houver algum valor útil (excluir blocos inteiramente vazios)
                    # Só adiciona se a "DATA" ou "INÍCIO" deste bloco não estiver vazia
                    if len(sub_row) > 0 and (sub_row[0] or (len(sub_row) > 1 and sub_row[1])):
                        # O bloco vence o prefixo quando os dois trazem a mesma coluna.
                        registro = dict(prefixo)
                        registro.update(dict(zip(sub_headers, sub_row)))
                        flattened_rows.append(registro)

            df = pd.DataFrame(flattened_rows)
        else:
            df = pd.DataFrame(rows, columns=headers)
        
        # Como achatamos (flatten), agora garantimos que não há colunas duplicadas
        df = df.loc[:, ~df.columns.duplicated()]
        
        # O cabeçalho no df agora não tem acentos. (ex: 'DATA', 'PRE', 'POS', 'INICIO', 'FIM')
        if "DATA" in df.columns:
            # Completa a data do bloco com a data comum da linha quando o bloco
            # não a repete ('DATA GRADE', 'DATA REAL'). Sem isso a grade tem
            # evento, mas o match não acha porque a linha ficou sem data.
            alternativas = [
                c for c in df.columns
                if c != "DATA" and "DATA" in c and "SEMANA" not in c
            ]
            for alt in alternativas:
                faltando = df["DATA"].isna() | (df["DATA"].astype(str).str.strip() == "")
                if not faltando.any():
                    break
                df.loc[faltando, "DATA"] = df.loc[faltando, alt]

            df["DATA"] = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
            
            # --- FIX: Tratar anos 1900 (Excel bug when year is omitted) ---
            # Se encontrar ano 1900, assume-se o ano da execução ou o ano mais frequente na grade
            # Para este projeto, vamos usar o ano atual do sistema se for 1900.
            current_year = _dt.now().year
            # Se estamos em Janeiro/Fevereiro e a grade tem datas de Dezembro, pode ser o ano anterior, 
            # mas simplificando para o ano atual resolve 99% dos casos de rascunhos.
            # Verificamos se há algum ano 2026, 2027 etc na grade para usar como referência
            referencia = df["DATA"].dropna()
            anos_validos = referencia[referencia.dt.year > 1901].dt.year
            ano_base = anos_validos.max() if not anos_validos.empty else current_year
            
            mask_1900 = (df["DATA"].dt.year == 1900)
            if mask_1900.any():
                df.loc[mask_1900, "DATA"] = df.loc[mask_1900, "DATA"].apply(
                    lambda x: x.replace(year=int(ano_base)) if pd.notnull(x) else x
                )
        else:
            # Fallback caso a coluna seja chamada de 'DATA DO EVENTO'
            data_col = next((c for c in df.columns if "DATA" in c), None)
            if data_col:
                df["DATA"] = pd.to_datetime(df[data_col], errors="coerce", dayfirst=True)

        def _h(v) -> str:
            if pd.isna(v):  # Isso checa nativamente strings vazias do pandas (NaN, NaT, None)
                return ""
            if hasattr(v, "strftime"):
                try:
                    return v.strftime("%H:%M")
                except ValueError:
                    return ""
            s = str(v).strip()
            return "" if s in {"-", "", "nan", "NaN", "None"} else s

        # Procurar as colunas de horário        # Mapeamento inteligente de colunas
        cols_final = {}
        for target, aliases in {
            "PRE": ["PRE", "CONVOCACAO"],
            "POS": ["POS"],
            "INICIO": ["INICIO", "START", "HORA"],
            "FIM": ["FIM", "END", "TERMINO"]
        }.items():
            found = None
            # Tentar cada alias na lista (ordem de prioridade)
            for alias in aliases:
                # 1. Match exato
                found = next((c for c in df.columns if c == alias), None)
                if found: break
                # 2. Match parcial (não-duracao)
                found = next((c for c in df.columns if alias in c and "DURACAO" not in c), None)
                if found: break
                # 3. Match parcial (com duracao)
                found = next((c for c in df.columns if alias in c), None)
                if found: break
            
            if found:
                cols_final[target] = found

        def safe_time_delta(t_str, base_time, is_subtract=True):
            try:
                # Se não tem base, não tem o que calcular
                if not base_time or base_time == "": return ""
                
                # Parse base time (HH:MM)
                bh, bm = map(int, base_time.split(':'))
                base_dt = _dt.combine(_dt.today(), _dt.min.time().replace(hour=bh, minute=bm))
                
                # Parse delta (pode ser HH:MM ou HH:MM:SS)
                dh, dm = 0, 0
                parts = t_str.split(':')
                dh = int(parts[0])
                dm = int(parts[1]) if len(parts) > 1 else 0
                
                delta = timedelta(hours=dh, minutes=dm)
                
                res_dt = base_dt - delta if is_subtract else base_dt + delta
                return res_dt.strftime("%H:%M")
            except:
                return ""

        # Processar colunas principais primeiro
        for k in ["INICIO", "FIM"]:
            if k in cols_final:
                df[k] = df[cols_final[k]].apply(_h)
            else:
                df[k] = ""

        # Processar PRE e POS com detecção de duração
        for k in ["PRE", "POS"]:
            if k in cols_final:
                col_name = cols_final[k]
                is_duracao = "DURACAO" in col_name or "DURA" in col_name
                
                if is_duracao:
                    # Se for duração, calcular baseado em INICIO ou FIM
                    base_key = "INICIO" if k == "PRE" else "FIM"
                    df[k] = df.apply(lambda row: safe_time_delta(_h(row[col_name]), row[base_key], is_subtract=(k == "PRE")), axis=1)
                else:
                    df[k] = df[col_name].apply(_h)
            else:
                df[k] = ""

        # Garantir EVENTO/CAMPEONATO e JOGO (mesma lógica flexível)
        if "EVENTO/CAMPEONATO" not in df.columns:
            campeonato_col = next((c for c in df.columns if "CAMPEONATO" in c or "EVENTO" in c), None)
            df["EVENTO/CAMPEONATO"] = df[campeonato_col].fillna("").astype(str) if campeonato_col else ""
            
        if "JOGO" not in df.columns:
            jogo_col = next((c for c in df.columns if "JOGO" in c or "PARTIDA" in c), None)
            
            if "MANDANTE" in df.columns and "VISITANTE" in df.columns:
                j_str = df["MANDANTE"].fillna("").astype(str) + " X " + df["VISITANTE"].fillna("").astype(str)
                df["JOGO"] = j_str.replace("^ X $", "", regex=True)
            elif "OBSERVACAO" in df.columns:
                df["JOGO"] = df["OBSERVACAO"].fillna("").astype(str)
            elif "COMENTARIOS" in df.columns:
                df["JOGO"] = df["COMENTARIOS"].fillna("").astype(str)
            elif jogo_col:
                df["JOGO"] = df[jogo_col].fillna("").astype(str)
            else:
                df["JOGO"] = ""

        safe_print(
            f"[OK] Grade carregada via scan dinâmico: {len(df)} linhas lidas."
        )
        wb.close()
        return df

    except Exception as e:
        import traceback
        safe_print(f"[ERRO] Erro ao carregar grade: {e}")
        safe_print(traceback.format_exc())
        return None


# ---------------------------------------------------------------------------
# buscar_na_grade — versão estendida (backward-compatible)
# ---------------------------------------------------------------------------

def buscar_na_grade(
    grade_eventos: Optional[pd.DataFrame],
    data_str: str,
    evento_escala: Optional[str],
    threshold: float = THRESHOLD_PADRAO,
    event_group: Optional[str] = None,
    inicio_rel: Optional[str] = None,
    fim_rel: Optional[str] = None,
) -> Tuple[str, str, str, str, str]:
    """
    Busca horários na grade para um evento do relatório usando 3 sinais:
      1. Similaridade de texto (Evento/Programa vs JOGO+EVENTO/CAMPEONATO)
      2. Event Group vs EVENTO/CAMPEONATO (se fornecido)
      3. Sobreposição de horário (se inicio_rel fornecido)

    Retorna: (pre_jogo, pos_jogo, inicio, fim, grade_id)
    """
    vazio: Tuple[str, str, str, str, str] = ("", "", "", "", "")

    if grade_eventos is None or grade_eventos.empty:
        return vazio
    if not evento_escala or (isinstance(evento_escala, float) and pd.isna(evento_escala)):
        return vazio

    # --- OTIMIZAÇÃO: Termos que NÃO constam na grade de eventos da TV --- 
    ev_upper = str(evento_escala).upper()
    termos_pular = [
        "VIAGEM", "FOLGA", "GRAV", "REUNIAO", "REUNIÃO", "TREINO", "OFF",
        "PODCAST", "PROMODAY", "MKT", "INSTITUCIONAL", "COMERCIAL", "ENTREVISTA"
    ]
    if any(k in ev_upper for k in termos_pular):
        return vazio

    try:
        # Se data_str já é um Timestamp ou datetime (não string), converter diretamente
        if hasattr(data_str, 'date'):
            data_alvo = data_str.date() if callable(data_str.date) else data_str.date
        elif hasattr(data_str, 'to_pydatetime'):
            data_alvo = data_str.to_pydatetime().date()
        else:
            data_alvo = _dt.strptime(str(data_str).strip(), "%d/%m/%Y").date()
    except (ValueError, TypeError):
        safe_print(f"⚠️  Data inválida: '{data_str}'")
        return vazio

    if "DATA" not in grade_eventos.columns:
        safe_print("❌ Coluna DATA ausente na grade.")
        return vazio

    grade = grade_eventos.copy()
    if not hasattr(grade["DATA"].iloc[0], "date"):
        grade["DATA"] = pd.to_datetime(grade["DATA"], errors="coerce")
    grade["_dt"] = grade["DATA"].dt.date
    dia = grade[grade["_dt"] == data_alvo]

    if dia.empty:
        safe_print(f"ℹ️  Grade sem eventos para {data_str}.")
        return vazio

    melhor_score = 0.0
    melhor_linha = None
    melhor_dist = None

    # Distância até o horário que o profissional já tem na escala. Serve de
    # critério de desempate: um mesmo programa vai ao ar várias vezes no dia
    # (SporTV News às 07:00 e às 21:00) e todas as exibições têm score de texto
    # idêntico. Sem desempate vencia a primeira da planilha, o que produzia
    # "mudança de horário" de 14 horas.
    ref_min = _hora_min(inicio_rel)

    for _, linha in dia.iterrows():
        campeonato = str(linha.get("EVENTO/CAMPEONATO", "") or "")
        jogo       = str(linha.get("JOGO", "") or "")
        pre_g      = linha.get("PRE", "")
        fim_g      = linha.get("FIM", "")

        score = calcular_score_combinado(
            evento_str      = evento_escala,
            campeonato_grade= campeonato,
            jogo_grade      = jogo,
            event_group     = event_group,
            inicio_rel      = inicio_rel,
            fim_rel         = fim_rel,
            pre_grade       = pre_g,
            fim_grade_val   = fim_g,
        )
        # Distância (em minutos) entre o horário desta linha da grade e o que
        # o profissional já tem na escala.
        dist = None
        if ref_min is not None:
            ini_g = _hora_min(linha.get("INICIO", ""))
            if ini_g is not None:
                bruta = abs(ini_g - ref_min)
                dist = min(bruta, 1440 - bruta)

        if score > melhor_score:
            melhor_score, melhor_linha, melhor_dist = score, linha, dist
        elif (
            score == melhor_score
            and melhor_linha is not None
            and dist is not None
            and (melhor_dist is None or dist < melhor_dist)
        ):
            # Mesmo score de texto: fica a exibição mais próxima do horário
            # que já estava na escala.
            melhor_linha, melhor_dist = linha, dist

    if melhor_linha is None or melhor_score < threshold:
        if melhor_score >= 0.2:
            safe_print(f"⚠️  Match muito fraco para a grade: '{evento_escala}' | {data_str} (heurística={melhor_score:.3f})")
        else:
            safe_print(f"⚠️  Sem match na grade: '{evento_escala}' | {data_str} (melhor={melhor_score:.3f})")
        return vazio
    else:
        safe_print(
            f"✅ Match Rápido: '{melhor_linha.get('EVENTO/CAMPEONATO', '')} | "
            f"{melhor_linha.get('JOGO', '')}' (score={melhor_score:.3f})"
        )
    
    # Gerar ID único combinando campeonato, jogo e horário de início da grade
    # SÓ GERA ID SE O SCORE FOR ALTO (Confiança de que é o mesmo slot da grade)
    # Match GEMINI (0.99) ou Match Rápido bom (>= 0.8)
    if melhor_score >= 0.8:
        g_campeonato = str(melhor_linha.get('EVENTO/CAMPEONATO', '')).strip()
        g_jogo = str(melhor_linha.get('JOGO', '')).strip()
        g_inicio = str(melhor_linha.get('INICIO', '')).strip()
        grade_id = f"{g_campeonato} | {g_jogo} | {g_inicio}".strip()
    else:
        grade_id = "" # Match fraco não gera ID de agrupamento soberano

    def _c(col: str) -> str:
        v = melhor_linha.get(col, "")
        if not v:
            return ""
        s = str(v).strip()
        return "" if s in {"-", "", "nan", "NaN", "None"} else s

    return _c("PRE"), _c("POS"), _c("INICIO"), _c("FIM"), grade_id

def buscar_evento_na_grade(row, grade_eventos):
    """Alias para buscar_na_grade usado pelo gerador HTML."""
    if grade_eventos is None: return None
    data = row.get('Data', '')
    evento = row.get('Evento/Programa', '')
    event_group = row.get('Event Group', None)
    
    # Prioriza horarios do registro se existirem
    inicio_rel = str(row.get('Início', row.get('Inicio', ''))).strip()
    fim_rel = str(row.get('Fim', row.get('FIM', ''))).strip()
    
    p, po, i, f, gid = buscar_na_grade(grade_eventos, data, evento, event_group=event_group, inicio_rel=inicio_rel, fim_rel=fim_rel)
    if not i: return None
    return {'pre_jogo': p, 'pos_jogo': po, 'horario_inicio': i, 'horario_fim': f, 'id_exclusivo': gid}

def combinar_horarios_grade_2018(pre_g, pos_g, ini_g, fim_g, ini_18, fim_18, ev, data):
    """Combina horários da grade com horários do registro 2018 (escala individual)."""
    # A grade é soberana para tudo: INICIO, FIM, PRE, POS
    # Se não temos grade, os do 2018 ganham.
    if not ini_g and not fim_g:
        return "", ini_18, fim_18, ""
    return pre_g, ini_g, fim_g, pos_g

# ---------------------------------------------------------------------------
# Diagnóstico interativo
# ---------------------------------------------------------------------------

def diagnosticar_match(
    evento_a: str,
    evento_b_campeonato: str,
    evento_b_jogo: str = "",
    event_group: Optional[str] = None,
    inicio_rel: Optional[str] = None,
    fim_rel: Optional[str] = None,
    pre_grade = None,
    fim_grade_val = None,
) -> None:
    """Imprime breakdown completo de todos os sinais."""
    a    = normalizar(evento_a)
    tg   = _texto_grade(evento_b_campeonato, evento_b_jogo)
    s_txt = _score_texto(a, tg)
    s_eg  = _score_event_group(event_group, evento_b_campeonato)
    sob   = _sobreposicao(inicio_rel, fim_rel, pre_grade, fim_grade_val)
    peso_eg   = 0.30 if s_txt < 0.40 else (0.20 if s_txt < 0.60 else 0.10)
    bonus_eg  = s_eg * peso_eg
    bonus_hora= 0.20 if (inicio_rel and sob) else 0.0
    score_final = round(min(s_txt + bonus_eg + bonus_hora, 1.0), 4)

    safe_print("=" * 65)
    safe_print(f"REL evento    : {evento_a}  →  norm: {a}")
    safe_print(f"GRADE campeon.: {evento_b_campeonato}")
    safe_print(f"GRADE jogo    : {evento_b_jogo}  →  combinado: {tg}")
    safe_print(f"REL EventGroup: {event_group}")
    safe_print(f"Horário REL   : {inicio_rel} → {fim_rel}")
    safe_print(f"Janela grade  : PRÉ={pre_grade} FIM={fim_grade_val}")
    safe_print("-" * 65)
    safe_print(f"  Sinal 1 TEXTO      : {s_txt:.4f}")
    safe_print(f"  Sinal 2 EVENT GROUP: {s_eg:.4f}  (peso={peso_eg:.2f} → bônus={bonus_eg:.4f})")
    safe_print(f"  Sinal 3 HORÁRIO    : sobreposição={sob}  (bônus={bonus_hora:.2f})")
    safe_print(f"  ► SCORE FINAL      : {score_final:.4f}")
    safe_print(f"  ► {'✅ BATE' if score_final >= THRESHOLD_PADRAO else '❌ NÃO BATE'}  (threshold={THRESHOLD_PADRAO})")
    print("=" * 65)