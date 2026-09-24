import datetime
import logging
import os
import re
from typing import Any, Dict, List

import pandas as pd

from config import TEST_EMAILS_DIR
from core.columns import normalizar_colunas, sem_acento

logger = logging.getLogger(__name__)

DIAS_PT = [
    'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira',
    'sexta-feira', 'sábado', 'domingo',
]

_VAZIOS = {'', 'nan', 'none', 'nat', '-'}


def _saudacao(agora: datetime.datetime = None) -> str:
    hora = (agora or datetime.datetime.now()).hour
    if hora < 12:
        return "bom dia"
    if hora < 18:
        return "boa tarde"
    return "boa noite"


def _chave_pessoa(nome: Any) -> str:
    """Nome comparável: sem acento, minúsculo, espaços colapsados."""
    return " ".join(sem_acento(nome).lower().split())


def _normalizar_plataforma(valor: Any) -> str:
    """'Sportv 2', 'Sportv 3' -> 'Sportv'.

    O canal específico interessa à operação, não a quem recebe a escala.
    Plataformas de outras famílias ('TV Globo - REDE', 'Combate') passam
    intactas.
    """
    v = _texto(valor)
    if not v:
        return ''
    if re.fullmatch(r"sportv\s*\d*", sem_acento(v).strip(), re.IGNORECASE):
        return "Sportv"
    return v


def _minutos(valor: Any) -> int:
    """'HH:MM' -> minutos desde a meia-noite. Serve para ordenar a tabela."""
    m = re.search(r"(\d{1,2}):(\d{2})", str(valor or ""))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else 0


def _texto(valor: Any) -> str:
    """Valor legível ou string vazia."""
    if valor is None:
        return ''
    try:
        if pd.isna(valor):
            return ''
    except (TypeError, ValueError):
        pass
    s = str(valor).strip()
    return '' if s.lower() in _VAZIOS else s


class EmailSender:
    def __init__(self, test_mode: bool = True, preview_dir: str = None):
        self.test_mode = test_mode
        # Pasta do ciclo de homologação. Sem ela, cai na pasta geral.
        self.preview_dir = preview_dir or TEST_EMAILS_DIR

    def get_outlook_app(self):
        try:
            import win32com.client
            return win32com.client.Dispatch('outlook.application')
        except Exception as e:
            logger.error(f"Erro ao iniciar app do Outlook: {str(e)}")
            return None

    def send_notification(self,
                          to_address: str,
                          employee_name: str,
                          changes: List[Dict[str, Any]],
                          df_schedule: Any = None,
                          subject: str = "Aviso: Atualização na sua Escala de Trabalho",
                          sem_destinatario: bool = False) -> bool:

        # Em homologação o preview é gerado mesmo sem e-mail cadastrado: o
        # objetivo é conferir o conteúdo, e a pasta vinha vazia justamente
        # quando o mapeamento ainda não estava preenchido.
        if not to_address and not (self.test_mode and sem_destinatario):
            logger.warning(f"Não há e-mail para {employee_name}. Pulando envio.")
            return False

        html_body = self._generate_html_body(employee_name, changes, df_schedule)
        if sem_destinatario:
            html_body = self._aviso_sem_destinatario(employee_name) + html_body
        text_body = self._generate_text_body(employee_name, changes)

        if self.test_mode:
            destino = to_address or "(sem e-mail cadastrado)"
            logger.info(f"[TEST MODO] E-mail simulado para: {destino} | Assunto: {subject}")

            os.makedirs(self.preview_dir, exist_ok=True)
            safe_name = "".join(
                c for c in employee_name if c.isalnum() or c in (' ', '_')
            ).rstrip().replace(' ', '_') or "sem_nome"
            if sem_destinatario:
                safe_name = f"SEM_EMAIL_{safe_name}"
            filename = os.path.join(self.preview_dir, f"{safe_name}.html")
            if os.path.exists(filename):
                # Mesmo nome no mesmo ciclo (homônimos): não sobrescreve.
                filename = os.path.join(
                    self.preview_dir,
                    f"{safe_name}_{datetime.datetime.now():%H%M%S%f}.html"
                )

            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(html_body)
                logger.info(f"[TEST MODO] Preview HTML salvo em: {filename}")
                return True
            except OSError as e:
                logger.error(f"[TEST MODO] Erro ao salvar arquivo HTML: {e}")
                return False

        outlook_app = self.get_outlook_app()
        if not outlook_app:
            return False

        try:
            mail = outlook_app.CreateItem(0)
            mail.To = to_address
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Body = text_body
            mail.Send()
            logger.info(f"E-mail enviado via Outlook para: {to_address} ({employee_name})")
            return True
        except Exception as e:
            logger.error(f"Erro enviando email para {to_address}: {str(e)}")
            return False

    # ------------------------------------------------------------------
    # Formatação
    # ------------------------------------------------------------------

    @staticmethod
    def _format_data(val_str: Any) -> str:
        v = _texto(val_str)
        if not v:
            return '-'
        if ' ' in v and ('-' in v or '/' in v):
            v = v.split(' ')[0]
        if '-' in v and len(v) == 10:
            partes = v.split('-')
            if len(partes) == 3:
                return f"{partes[2]}/{partes[1]}/{partes[0]}"
        return v

    @classmethod
    def _format_dia_semana(cls, data_str: Any, dia_raw: Any = "") -> str:
        bruto = _texto(dia_raw)
        if bruto and not bruto.isdigit() and len(bruto) < 20:
            if 'feira' in bruto.lower() or bruto.lower() in ('sábado', 'domingo', 'sabado'):
                return bruto.lower()
        try:
            dt = datetime.datetime.strptime(cls._format_data(data_str), "%d/%m/%Y")
            return DIAS_PT[dt.weekday()]
        except (ValueError, TypeError):
            return bruto if bruto else "-"

    @classmethod
    def _ordenar_por_data(cls, df: pd.DataFrame) -> pd.DataFrame:
        """Ordena a escala por data e hora de início.

        A tabela saía na ordem da planilha de origem, que agrupa por tipo de
        atividade — as folgas iam para o fim e a escala aparecia como
        21, 23, 25, 22, 24.
        """
        if df is None or df.empty or 'data' not in df.columns:
            return df

        chaves = pd.DataFrame(index=df.index)
        chaves['dia'] = pd.to_datetime(
            df['data'].map(cls._format_data), format='%d/%m/%Y', errors='coerce'
        )
        chaves['hora'] = (
            df['inicio'].map(_minutos) if 'inicio' in df.columns else 0
        )
        # Datas ilegíveis vão para o fim, sem derrubar o restante da tabela.
        ordenadas = chaves.sort_values(['dia', 'hora'], na_position='last')
        return df.loc[ordenadas.index]

    @classmethod
    def _datas_iguais(cls, d1: Any, d2: Any) -> bool:
        """Compara duas datas já normalizadas.

        A versão anterior aceitava `d1[:5] == d2[:5]` como prova de igualdade.
        Com datas ISO isso fazia '2026-07-01' casar com '2026-12-25' (prefixo
        '2026-'), e a escala inteira saía destacada de amarelo.
        """
        f1, f2 = cls._format_data(d1), cls._format_data(d2)
        return f1 == f2 and f1 != '-'

    @staticmethod
    def _format_hora(val_str: Any) -> str:
        v = _texto(val_str)
        if not v:
            return '-'
        if v.upper() == 'CANCELADO':
            return 'CANCELADO'
        if ' ' in v and ':' in v:
            v = v.split(' ')[-1]
        if len(v) >= 4 and ':' in v:
            partes = v.split(':')
            return f"{partes[0].zfill(2)}:{partes[1]}"
        return v

    # ------------------------------------------------------------------

    def _generate_html_body(self, name: str, changes: List[Dict[str, Any]], df_schedule: Any) -> str:
        first_name = name.split()[0] if name else name

        intro_texts = []
        for c in changes:
            data_str = str(c.get('data', ''))
            data_short = data_str[:5] if len(data_str) >= 5 else data_str
            dia_sem = self._format_dia_semana(data_str)

            if c['tipo'] == 'SAIDA':
                ev_name = _texto(c.get('campo'))
                if not ev_name or ev_name.lower() in ('inicio', 'início', 'fim', '*'):
                    ev_name = _texto(c.get('valor_antigo')) or 'EVENTO'
                intro_texts.append(
                    f"<span style='background-color: yellow; font-weight: bold;'>"
                    f"CAIU {ev_name}, na sua escala da {dia_sem}, {data_short}.</span> "
                    f"Segue escala atualizada."
                )
            elif c['tipo'] == 'ENTRADA':
                intro_texts.append(
                    f"<span style='background-color: yellow; font-weight: bold;'>"
                    f"NOVA ESCALA adicionada para a {dia_sem}, {data_short}.</span>"
                )
            else:
                campo_lbl = _texto(c.get('campo')) or 'horário'
                if campo_lbl.lower() in ('inicio', 'início', 'fim', 'horario', 'horário', 'pre', 'pré'):
                    intro_texts.append(
                        f"<span style='background-color: yellow; font-weight: bold;'>"
                        f"Houve alteração no horário e a definição do confronto na sua escala "
                        f"da {dia_sem}, {data_short}.</span>"
                    )
                else:
                    intro_texts.append(
                        f"<span style='background-color: yellow; font-weight: bold;'>"
                        f"Houve alteração de {campo_lbl} na sua escala da {dia_sem}, {data_short}.</span>"
                    )

        intro_html = "<br>".join(intro_texts) if intro_texts else "Sua escala de trabalho foi atualizada."

        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Calibri, Arial, sans-serif; color: #000; font-size: 14px; margin: 15px; }}
                table {{ border-collapse: collapse; width: 100%; max-width: 1200px; font-size: 12px; margin-top: 15px; }}
                th, td {{ border: 1px solid #a0a0a0; padding: 4px 6px; text-align: left; vertical-align: middle; }}
                th {{ background-color: #f2f2f2; font-weight: normal; }}
                .highlight {{ background-color: #ffff00; font-weight: bold; }}
            </style>
        </head>
        <body>
            <p>Oi {first_name}, {_saudacao()}! Tudo bem?</p>
            <p>{intro_html}</p>
            <p>Segue a escala atualizada:</p>
            <table>
                <tr>
                    <th>Nome</th>
                    <th>Plataforma</th>
                    <th>Data</th>
                    <th>Dia</th>
                    <th>Pré</th>
                    <th>Início</th>
                    <th>Fim</th>
                    <th>Evento/Descrição</th>
                    <th>Produto</th>
                    <th>Local</th>
                    <th>Elenco</th>
                    <th>Coordenador</th>
                    <th>Produtor</th>
                </tr>
        """

        def is_cell_changed(data_val, col_name):
            col_c = col_name.lower()
            for c in changes:
                if not self._datas_iguais(c.get('data', ''), data_val):
                    continue
                c_tipo = c.get('tipo', '')
                c_campo = str(c.get('campo', '')).strip().lower()

                if c_tipo == 'ENTRADA':
                    if col_c in ('inicio', 'fim', 'pre', 'evento', 'data', 'dia'):
                        return True
                elif c_tipo == 'ALTERACAO':
                    if any(k in c_campo for k in ('horario', 'horário', 'inicio', 'início', 'fim', 'confronto')):
                        if col_c in ('inicio', 'fim', 'pre', 'evento'):
                            return True
                    # Só compara nomes de campo quando há nome — `'' in col_c`
                    # é sempre verdadeiro e destacava a linha inteira.
                    if c_campo and (col_c in c_campo or c_campo in col_c):
                        return True
            return False

        def get_produto(row_dict):
            for key in ('produto', 'produto (wo/quick hold)', 'produto (wo/shift)', 'event_group'):
                val = _texto(row_dict.get(key))
                if val:
                    partes = [
                        p.strip() for p in val.split('/')
                        if p.strip() and p.strip().upper() != 'NA' and not p.strip().isdigit()
                    ]
                    return partes[0] if partes else val
            return ''

        def get_evento(row_dict):
            """Descrição da atividade, olhando além da coluna de evento.

            Folgas e bloqueios de agenda não têm 'evento' preenchido — o texto
            mora em 'descricao'. Sem este fallback a pessoa recebia a escala
            com linhas em branco marcando 00:00 às 00:00.
            """
            for chave in ('evento', 'descricao', 'row display',
                          'atividade/descricao', 'tipo de atividade'):
                valor = _texto(row_dict.get(chave))
                if valor:
                    return valor
            return ''

        def eh_folga(row_dict) -> bool:
            """Linha de folga: 'Day Off / Folga', 'Comp Day / Folga Compensatória'...

            Nessas linhas o horário 00:00–00:00 é só preenchimento, não um
            compromisso à meia-noite.
            """
            texto = sem_acento(get_evento(row_dict)).upper()
            return any(t in texto for t in ('FOLGA', 'DAY OFF', 'TIME OFF'))

        chave_destinatario = _chave_pessoa(name)

        def get_local(row_dict):
            """Local de locução — onde a pessoa trabalha.

            A base tem duas colunas: 'Local de locução' ('Offtube SP') e
            'Local' (o estádio do evento). Quem recebe a escala precisa da
            primeira.
            """
            for chave in ('local de locucao', 'local locucao', 'local de narracao',
                          'local narracao', 'local de gravacao', 'local gravacao',
                          'local'):
                valor = _texto(row_dict.get(chave))
                if valor:
                    return valor
            return ''

        def get_elenco(row_dict):
            """Os OUTROS integrantes da equipe.

            O próprio destinatário vinha repetido na lista — ele já sabe que
            está escalado; o que a coluna comunica é com quem ele trabalha.
            """
            bruto = _texto(row_dict.get('elenco'))
            if not bruto:
                bruto = ' ; '.join(
                    v for v in (
                        _texto(row_dict.get(k))
                        for k in ('narrador', 'comentarista', 'reporter')
                    ) if v
                )

            pessoas, vistos = [], set()
            for parte in bruto.split(';'):
                pessoa = parte.strip()
                if not pessoa:
                    continue
                chave = _chave_pessoa(pessoa)
                if chave == chave_destinatario or chave in vistos:
                    continue
                vistos.add(chave)
                pessoas.append(pessoa)
            return ' ; '.join(pessoas)

        if df_schedule is not None and not df_schedule.empty:
            df_active = normalizar_colunas(df_schedule)

            # Linhas canceladas saem da tabela: o aviso "CAIU ..." já está no topo.
            if 'inicio' in df_active.columns:
                df_active = df_active[
                    df_active['inicio'].astype(str).str.strip().str.upper() != 'CANCELADO'
                ]

            df_active = self._ordenar_por_data(df_active)

            if df_active.empty:
                html += "<tr><td colspan='13'>Nenhum registro ativo na escala.</td></tr>"
            else:
                for _, row in df_active.iterrows():
                    row_dict = {str(k).strip().lower(): v for k, v in row.items()}

                    data_val = self._format_data(row_dict.get('data', ''))
                    dia_val = self._format_dia_semana(data_val, row_dict.get('dia', ''))
                    pre_val = _texto(row_dict.get('pre')) or '-'
                    inicio_val = self._format_hora(row_dict.get('inicio', ''))
                    fim_val = self._format_hora(row_dict.get('fim', ''))
                    evento_val = get_evento(row_dict)

                    if eh_folga(row_dict):
                        # Folga não tem horário: 00:00 às 00:00 é preenchimento.
                        evento_val = 'FOLGA'
                        pre_val = inicio_val = fim_val = '-'
                    produto_val = get_produto(row_dict)
                    local_val = get_local(row_dict)
                    elenco_val = get_elenco(row_dict)
                    plataforma_val = _normalizar_plataforma(
                        row_dict.get('plataforma') or row_dict.get('canal')
                    )
                    coord_val = _texto(row_dict.get('coordenador'))
                    prod_val = _texto(row_dict.get('produtor'))

                    def td(val, col_name=""):
                        # Os valores já chegam formatados; aqui só tratamos
                        # nulos, preservando o '-' usado como placeholder.
                        texto = '' if val is None else str(val)
                        if texto.strip().lower() in ('nan', 'none', 'nat'):
                            texto = ''
                        cls = "highlight" if is_cell_changed(data_val, col_name) else ""
                        return f"<td class='{cls}'>{texto}</td>"

                    html += f"""
                    <tr>
                        {td(name, 'nome')}
                        {td(plataforma_val, 'plataforma')}
                        {td(data_val, 'data')}
                        {td(dia_val, 'dia')}
                        {td(pre_val, 'pre')}
                        {td(inicio_val, 'inicio')}
                        {td(fim_val, 'fim')}
                        {td(evento_val, 'evento')}
                        {td(produto_val, 'produto')}
                        {td(local_val, 'local')}
                        {td(elenco_val, 'elenco')}
                        {td(coord_val, 'coordenador')}
                        {td(prod_val, 'produtor')}
                    </tr>
                    """
        else:
            html += "<tr><td colspan='13'>Nenhum registro de escala encontrado.</td></tr>"

        html += """
            </table>
            <br>
            <p>Qualquer dúvida, estamos à disposição.</p>
        </body>
        </html>
        """
        return html

    @staticmethod
    def _aviso_sem_destinatario(nome: str) -> str:
        return (
            "<div style=\"font-family: Calibri, Arial, sans-serif; background:#fff3cd; "
            "border:2px solid #e0a800; padding:10px 14px; margin:12px 0;\">"
            "<b>PRÉVIA — SEM DESTINATÁRIO</b><br>"
            f"Não há e-mail cadastrado para <b>{nome}</b>. Com o envio real ligado, "
            "esta notificação <b>não sairia</b>.<br>"
            "Preencha em Configurações → Mapeamento de E-mails."
            "</div>"
        )

    def _generate_text_body(self, name: str, changes: List[Dict[str, Any]]) -> str:
        text = f"Olá, {name}!\n\nSua escala de trabalho foi atualizada. Detalhes:\n"
        for c in changes:
            text += (
                f"- Data: {c.get('data', '')} | {c.get('tipo', '')} | "
                f"Campo: {c.get('campo', '')} | Antigo: {c.get('valor_antigo', '')} "
                f"-> Novo: {c.get('valor_novo', '')}\n"
            )
        text += "\nQualquer dúvida, estamos à disposição."
        return text
