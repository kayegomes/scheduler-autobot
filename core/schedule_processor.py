import logging
import os
import threading
from contextlib import contextmanager
from traceback import format_exc
from typing import Any, Dict, List, Optional

import pandas as pd

from config import ATTACHMENTS_DIR, ATTACHMENT_RETENTION_DAYS, TEST_EMAILS_DIR
from core.columns import carregar_escala, detectar_tipo_planilha
from core.database import DatabaseManager
from core.diff_engine import compare_schedules
from core.email_sender import EmailSender
from core.outlook_monitor import OutlookMonitor
from core.test_report import RelatorioTeste, resolver_pasta_saida

logger = logging.getLogger(__name__)

# Proporção mínima de eventos da escala que a grade precisa confirmar para que
# os não encontrados possam ser tratados como cancelados.
TAXA_MINIMA_MATCH_GRADE = 0.30


class CicloEmAndamento(RuntimeError):
    """Uma operação manual foi tentada enquanto o ciclo automático rodava."""


def _datas_da_grade(grade_tv: pd.DataFrame) -> set:
    """Datas (dd/mm/aaaa) efetivamente presentes na grade de TV."""
    if 'DATA' not in grade_tv.columns:
        return set()
    datas = pd.to_datetime(grade_tv['DATA'], errors='coerce').dropna()
    return set(datas.dt.strftime('%d/%m/%Y'))


# Colunas que descrevem o que a pessoa está fazendo. Um termo imune pode
# aparecer em qualquer uma delas, dependendo do layout da planilha.
CAMPOS_CONTEXTO = (
    'evento', 'descricao', 'row display', 'tipo de atividade',
    'sub-atividade', 'sub-atividade (shift)', 'atividade', 'funcao', 'notas',
)


def _contexto_da_linha(row) -> str:
    """Texto em maiúsculas com tudo que descreve a atividade da linha."""
    partes = []
    for campo in CAMPOS_CONTEXTO:
        valor = row.get(campo)
        if valor is None:
            continue
        texto = str(valor).strip()
        if texto and texto.lower() not in ('nan', 'none', 'nat'):
            partes.append(texto)
    return ' | '.join(partes).upper()


class ScheduleProcessor:
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager
        # Impede que o agendador automático e o botão "Sincronizar Agora"
        # rodem ao mesmo tempo: os dois passariam juntos pela checagem de
        # e-mail já processado e mandariam a notificação duas vezes.
        self._lock = threading.Lock()
        # Releída da configuração a cada ciclo; o padrão vale para chamadas
        # diretas a _processar_arquivo (homologação via run_test.py).
        self.pasta_saida_teste = TEST_EMAILS_DIR

    @property
    def em_execucao(self) -> bool:
        return self._lock.locked()

    @contextmanager
    def exclusivo(self, timeout: float = 0):
        """Bloqueia o ciclo durante uma operação manual (carga de escala base).

        Sem isso, clicar em "Carregar Escala Base Inicial" no meio de um ciclo
        fazia a escala recém-carregada ser soterrada pela grade que o ciclo
        salvava logo depois, usando o `df_old` que ele tinha lido antes.
        """
        if not self._lock.acquire(timeout=timeout):
            raise CicloEmAndamento("Há um ciclo de sincronização em andamento.")
        try:
            yield
        finally:
            self._lock.release()

    def process_cycle(self):
        if not self._lock.acquire(blocking=False):
            logger.info("Ciclo já em andamento. Ignorando disparo concorrente.")
            return
        try:
            self._process_cycle()
        finally:
            self._lock.release()

    # ------------------------------------------------------------------

    def _process_cycle(self):
        config = self.db.get_config()
        keyword = config.get('mail_subject_keyword', 'Nova Escala')
        folder = config.get('outlook_folder', 'Caixa de Entrada')
        test_mode = str(config.get('test_mode', '1')) == '1'
        subj = config.get('notification_subject', 'Atualização na sua Escala')
        immune_keys_str = config.get('immune_keywords') or 'VIAGEM, FOLGA, OFF, REUNIAO, GRAVACAO, MEDICO, FERIAS'
        immune_keys = [k.strip().upper() for k in immune_keys_str.split(',') if k.strip()]
        max_funcs = _inteiro(config.get('max_funcionarios_por_ciclo'), 25)

        self.pasta_saida_teste = resolver_pasta_saida(config, TEST_EMAILS_DIR)

        logger.info(
            f"Iniciando ciclo de processamento. Keyword: '{keyword}', Pasta: '{folder}'"
            + (f" | MODO TESTE -> {self.pasta_saida_teste}" if test_mode else "")
        )
        self.db.update_last_check_time()

        monitor = OutlookMonitor(folder_name=folder, keyword=keyword, save_dir=ATTACHMENTS_DIR)
        # O filtro vai junto: antes o anexo era baixado toda vez, para os 50
        # últimos e-mails, e só depois o sistema via que já tinha processado.
        emails_encontrados = monitor.search_new_emails(
            ja_processado=lambda eid: self.db.is_email_processed(eid, modo_teste=test_mode)
        )

        if not emails_encontrados:
            logger.info("Nenhum novo e-mail com anexo de planilha encontrado.")
            monitor.limpar_anexos_antigos(ATTACHMENT_RETENTION_DAYS)
            return

        for email_id, att_path in emails_encontrados:
            try:
                self._processar_arquivo(
                    email_id, att_path, immune_keys, test_mode, subj, max_funcs
                )
            except Exception:
                logger.error(f"Erro processando arquivo {att_path}:\n{format_exc()}")
                # A grade pendente é descartada para que a próxima tentativa
                # comece limpa e o e-mail volte a ser elegível.
                self.db.descartar_grades_pendentes(email_id)

        monitor.limpar_anexos_antigos(ATTACHMENT_RETENTION_DAYS)

    # ------------------------------------------------------------------

    def _processar_arquivo(self, email_id: str, att_path: str, immune_keys: List[str],
                           test_mode: bool, subj: str, max_funcs: int):
        logger.info(f"Processando arquivo recebido: {att_path}")
        filename = os.path.basename(att_path)

        # Restos de uma tentativa anterior que falhou no meio do caminho.
        self.db.descartar_grades_pendentes(email_id)

        baseline_id = self.db.get_baseline_id()
        df_old = self.db.get_last_grade_df()
        tipo = detectar_tipo_planilha(att_path)
        logger.info(f"Planilha classificada como: {tipo}")

        def recusar(motivo: str):
            """Recusa definitiva para esta base: não reprocessa a cada ciclo."""
            self.db.registrar_email_ignorado(email_id, filename, motivo, baseline_id)

        if tipo == 'grade_tv':
            if df_old.empty:
                logger.warning(
                    "Grade de TV recebida sem escala base cadastrada. Uma grade de TV "
                    "não contém funcionários e não pode virar a escala base — carregue "
                    "a escala pelo botão 'Carregar Escala Base Inicial'. Arquivo ignorado."
                )
                recusar("sem escala base")
                return
            df_new = self._cruzar_com_grade_tv(att_path, df_old, immune_keys)
            if df_new is None:
                recusar("grade não aplicável à escala vigente")
                return
        else:
            df_new = carregar_escala(att_path)

        if 'funcionario' not in df_new.columns:
            logger.error(
                f"A planilha '{filename}' não tem coluna de funcionário "
                f"reconhecível (colunas: {list(df_new.columns)[:12]}). Arquivo ignorado."
            )
            recusar("sem coluna de funcionário")
            return

        df_new = df_new[df_new['funcionario'].astype(str).str.strip() != '']
        if df_new.empty:
            logger.error("Nenhuma linha com funcionário preenchido. Arquivo ignorado.")
            recusar("nenhuma linha com funcionário")
            return

        # --- Modo Teste: nada entra no histórico de produção -------------
        if test_mode:
            self._homologar(email_id, filename, df_old, df_new, baseline_id, subj, max_funcs)
            return

        # A grade nasce 'pending': só vira base de comparação depois que as
        # notificações saírem. Antes, o e-mail era marcado como processado
        # aqui, e qualquer falha adiante perdia a notificação para sempre.
        grade_id = self.db.save_new_grade(
            df_new, filename, email_id, status='pending', baseline_id=baseline_id
        )
        logger.info(f"Grade {grade_id} salva ({len(df_new)} linhas) aguardando notificação.")

        logger.info("Executando comparação (Diff) de escalas...")
        changes = compare_schedules(df_old, df_new, grade_id)

        if not changes:
            logger.info("Nenhuma alteração detectada nesta versão.")
            self.db.marcar_grade_concluida(grade_id)
            return

        change_ids = self.db.save_changes(changes)
        for c, cid in zip(changes, change_ids):
            c['id'] = cid
        logger.info(f"{len(changes)} alterações detectadas e salvas.")

        afetados = {c['funcionario'] for c in changes}
        if max_funcs and len(afetados) > max_funcs:
            logger.error(
                f"ENVIO BLOQUEADO: {len(afetados)} funcionários seriam notificados de uma vez "
                f"(limite configurado: {max_funcs}). Isso costuma indicar base de comparação "
                f"defasada — por exemplo uma grade de outro mês comparada com a anterior. "
                f"As {len(changes)} alterações ficaram registradas no histórico como NÃO "
                f"enviadas; revise e ajuste 'max_funcionarios_por_ciclo' se for legítimo."
            )
            self.db.marcar_grade_concluida(grade_id)
            return

        self._send_notifications(changes, df_new, test_mode, subj)
        self.db.marcar_grade_concluida(grade_id)

    # ------------------------------------------------------------------

    def _homologar(self, email_id: str, filename: str, df_old: pd.DataFrame,
                   df_new: pd.DataFrame, baseline_id: Optional[int],
                   subj: str, max_funcs: int):
        """Ciclo de Modo Teste: tudo vai para uma pasta, nada para o banco."""
        relatorio = RelatorioTeste(self.pasta_saida_teste, filename)

        logger.info("Executando comparação (Diff) de escalas...")
        changes = compare_schedules(df_old, df_new, grade_id=0)
        logger.info(f"[MODO TESTE] {len(changes)} alteração(ões) detectada(s).")

        avisos = []
        afetados = {c['funcionario'] for c in changes}
        if max_funcs and len(afetados) > max_funcs:
            # Em homologação o limite não bloqueia: o objetivo é justamente
            # ver tudo que sairia. Mas o aviso vai para o relatório.
            avisos.append(
                f"{len(afetados)} funcionários seriam notificados de uma vez, acima do limite "
                f"de {max_funcs}. Com o Modo Teste desligado, este ciclo seria BLOQUEADO."
            )
            logger.warning(f"[MODO TESTE] {avisos[-1]}")

        envio = {'enviados': [], 'sem_email': [], 'falhas': []}
        if changes:
            envio = self._send_notifications(
                changes, df_new, True, subj, preview_dir=relatorio.emails_dir
            )

        descricao_base = (
            f"grade {baseline_id} ({len(df_old)} linhas)" if baseline_id else "nenhuma"
        )
        pasta = relatorio.salvar(
            changes, df_new, {'descricao': descricao_base}, envio, avisos
        )

        # Registra só para não refazer o mesmo relatório a cada ciclo. Vale
        # enquanto o Modo Teste estiver ligado e a base for a mesma.
        self.db.registrar_execucao_teste(email_id, filename, baseline_id, os.path.basename(pasta))

    # ------------------------------------------------------------------

    def _cruzar_com_grade_tv(self, att_path: str, df_old: pd.DataFrame,
                             immune_keys: List[str]) -> Optional[pd.DataFrame]:
        """Aplica os horários de uma grade de TV sobre a escala vigente.

        Concluir que um evento "caiu" só é legítimo quando a grade recebida de
        fato cobre aquele evento. Uma grade especializada (Combate, PPV) traz
        um recorte da programação: sem essa distinção, os 85 eventos de futebol
        e vôlei da escala eram cancelados só por não estarem numa grade de MMA.
        """
        from core.match_eventos import buscar_evento_na_grade, carregar_grade

        grade_tv = carregar_grade(att_path)
        if grade_tv is None or grade_tv.empty:
            logger.warning("Não foi possível ler a grade de TV. Arquivo ignorado.")
            return None

        # Sem coluna de evento não há como saber o que casou e o que caiu.
        # Marcar tudo como CANCELADO nessa situação era o pior desfecho possível.
        if 'evento' not in df_old.columns:
            logger.error(
                "A escala base não tem coluna de evento reconhecível; o cruzamento com a "
                "grade de TV cancelaria a escala inteira por engano. Arquivo ignorado."
            )
            return None

        datas_cobertas = _datas_da_grade(grade_tv)
        if not datas_cobertas:
            logger.warning("A grade de TV não tem datas legíveis. Arquivo ignorado.")
            return None

        df_new = df_old.copy()
        for col in ('inicio', 'fim', 'pre', 'pos'):
            if col not in df_new.columns:
                df_new[col] = ''

        n_match = 0
        candidatos = 0          # linhas que a grade poderia confirmar ou derrubar
        fora_de_cobertura = 0
        sem_evento = 0
        pendentes_cancelamento = []

        for idx, row in df_new.iterrows():
            data_linha = str(row.get('data', '')).strip()
            evento = str(row.get('evento', '')).strip()

            # Folga, viagem e afins podem estar descritas em qualquer coluna de
            # contexto — na escala real "Day Off / Folga" vem em 'descricao',
            # com 'evento' vazio. Procurar o termo só em 'evento' deixava 197
            # folgas passarem direto para a fila de cancelamento.
            imune = any(k in _contexto_da_linha(row) for k in immune_keys)

            match_info = buscar_evento_na_grade({
                'Data': row.get('data'),
                'Evento/Programa': row.get('evento'),
                'Event Group': row.get('event_group'),
                'Início': row.get('inicio'),
                'Fim': row.get('fim'),
            }, grade_tv)

            if match_info:
                n_match += 1
                candidatos += 1
                # A grade manda no que ela informa, não no que ela omite.
                # A grade do Sportv não preenche FIM: sobrescrever com vazio
                # apagava o horário real e gerava 146 falsos "mudou o horário".
                #
                # PRE e POS entram na mesma regra. Antes só INICIO/FIM eram
                # aplicados: quando a grade trazia um novo horário de início,
                # o PRÉ antigo (calculado para o horário anterior) ficava para
                # trás — "Pré 13:30" com "Início 13:00" é a convocação chegando
                # DEPOIS do início do evento. Com PRE/POS também vindos do
                # match, os dois horários mudam juntos.
                for campo, chave in (
                    ('inicio', 'horario_inicio'), ('fim', 'horario_fim'),
                    ('pre', 'pre_jogo'), ('pos', 'pos_jogo'),
                ):
                    valor = str(match_info.get(chave, '') or '').strip()
                    if valor:
                        df_new.at[idx, campo] = valor
                continue

            if imune:
                continue

            # Sem nome de evento não dá para afirmar que "caiu da grade" — não
            # há o que procurar. Linhas assim são folgas, reuniões e bloqueios
            # de agenda, e ficam de fora inclusive da taxa de acerto.
            if not evento:
                sem_evento += 1
                continue

            # A grade não tem NADA nesse dia: ela não cobre esse período, então
            # o silêncio dela não significa que o evento caiu.
            if data_linha not in datas_cobertas:
                fora_de_cobertura += 1
                continue

            candidatos += 1
            pendentes_cancelamento.append(idx)

        taxa = (n_match / candidatos) if candidatos else 0.0
        logger.info(
            f"Cruzamento: {n_match}/{candidatos} evento(s) casado(s) (taxa {taxa:.0%}), "
            f"{len(pendentes_cancelamento)} candidato(s) a CANCELADO, "
            f"{fora_de_cobertura} linha(s) em datas fora da grade (preservadas), "
            f"{sem_evento} linha(s) sem evento/imunes (preservadas)."
        )

        # A grade não toca nenhuma data da escala vigente: não há o que aplicar.
        # Salvar o resultado aqui criaria uma cópia idêntica da base e, pior,
        # ela passaria a ser a nova base — soterrando uma escala carregada
        # manualmente enquanto o ciclo rodava.
        if candidatos == 0:
            logger.warning(
                f"Grade de TV ignorada: nenhuma das {len(df_new)} linhas da escala vigente "
                f"cai em data coberta por esta grade. A escala atual cobre outro período."
            )
            return None

        # Taxa de acerto baixa demais indica grade de outro período ou de um
        # recorte diferente da programação — cancelar nesse cenário é ruído.
        if pendentes_cancelamento and taxa < TAXA_MINIMA_MATCH_GRADE:
            logger.error(
                f"Cruzamento abortado: apenas {taxa:.0%} dos {candidatos} eventos da escala "
                f"foram encontrados na grade (mínimo {TAXA_MINIMA_MATCH_GRADE:.0%}). A grade "
                f"provavelmente não corresponde a este período ou cobre outro recorte da "
                f"programação. Arquivo ignorado para não cancelar a escala inteira."
            )
            return None

        if n_match == 0 and not pendentes_cancelamento:
            logger.info("Grade de TV não alterou nada na escala vigente. Nada a salvar.")
            return None

        for idx in pendentes_cancelamento:
            df_new.at[idx, 'inicio'] = "CANCELADO"

        return df_new

    # ------------------------------------------------------------------

    def _send_notifications(self, changes: List[Dict[str, Any]], df_new: pd.DataFrame,
                            test_mode: bool, base_subject: str,
                            preview_dir: Optional[str] = None) -> Dict[str, List[str]]:
        grouped: Dict[str, List[Dict[str, Any]]] = {}
        for c in changes:
            grouped.setdefault(c['funcionario'], []).append(c)

        sender = EmailSender(test_mode=test_mode, preview_dir=preview_dir)
        enviados: List[int] = []
        notificados: List[str] = []
        sem_email: List[str] = []
        falhas: List[str] = []

        tem_funcionario = df_new is not None and 'funcionario' in df_new.columns

        for func, func_changes in grouped.items():
            email_address = self.db.get_employee_email(func)
            if not email_address:
                sem_email.append(func)
                # Em produção não há o que enviar. Em homologação o preview
                # sai mesmo assim, marcado como SEM_EMAIL_, para revisão.
                if not test_mode:
                    continue

            df_func = None
            if tem_funcionario:
                df_func = df_new[
                    df_new['funcionario'].astype(str).str.strip().str.lower() == str(func).strip().lower()
                ]

            try:
                ok = sender.send_notification(
                    email_address, func, func_changes, df_func, subject=base_subject,
                    sem_destinatario=not email_address
                )
            except Exception as e:
                logger.error(f"Erro inesperado notificando {func}: {e}")
                ok = False

            if not email_address:
                continue  # preview gerado, mas ninguém foi notificado

            if ok:
                # Só as alterações efetivamente notificadas são marcadas. Antes,
                # um UPDATE em bloco marcava a grade inteira como enviada —
                # inclusive quem nem tinha e-mail cadastrado.
                notificados.append(func)
                enviados.extend(c['id'] for c in func_changes if c.get('id'))
            else:
                falhas.append(func)

        # Em homologação nada foi enviado de verdade: marcar como enviado
        # inflaria o contador do dashboard com e-mails que nunca saíram.
        if not test_mode:
            self.db.mark_email_sent(enviados)

        prefixo = "[MODO TESTE] " if test_mode else ""
        verbo = "simulada(s)" if test_mode else "enviada(s)"
        logger.info(
            f"{prefixo}Notificações: {len(notificados)} {verbo}, "
            f"{len(falhas)} com falha, {len(sem_email)} sem e-mail cadastrado."
        )
        if sem_email:
            logger.warning(f"{prefixo}Sem e-mail no mapeamento: {', '.join(sorted(sem_email))}")
        if falhas:
            logger.warning(f"{prefixo}Falha no envio: {', '.join(sorted(falhas))}")

        return {'enviados': notificados, 'sem_email': sem_email, 'falhas': falhas}


def _inteiro(valor: Any, padrao: int) -> int:
    try:
        return int(str(valor).strip())
    except (TypeError, ValueError):
        return padrao
