import threading
import time
import schedule
import logging
from traceback import format_exc
from core.database import DatabaseManager
from core.schedule_processor import ScheduleProcessor
from gui.main_window import run_gui

# Setup main logging
logging.basicConfig(
    filename='automacao_escala.log',
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
# Also output to the console
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger("").addHandler(console)

logger = logging.getLogger(__name__)

INTERVALO_PADRAO_MIN = 15


def _intervalo_valido(valor, padrao=INTERVALO_PADRAO_MIN) -> int:
    """Intervalo em minutos, tolerante a lixo digitado na tela de configuração."""
    try:
        minutos = int(str(valor).strip())
    except (TypeError, ValueError):
        logger.warning(f"Intervalo inválido em configuração: {valor!r}. Usando {padrao} min.")
        return padrao
    if minutos < 1:
        logger.warning(f"Intervalo {minutos} min é baixo demais. Usando 1 min.")
        return 1
    return minutos


class AppOrchestrator:
    def __init__(self):
        self.db = DatabaseManager()
        self.processor = ScheduleProcessor(self.db)
        self.is_running = True
        self._intervalo_atual = None

    def run_scheduler(self):
        """Loop do agendador, relendo o intervalo a cada volta.

        Antes o intervalo era lido uma única vez no boot: mudar o valor na tela
        de Configurações não tinha efeito até reiniciar o app. E um valor
        não-numérico derrubava esta thread em silêncio — a automação parava
        sem nenhum sinal na interface.
        """
        while self.is_running:
            try:
                intervalo = _intervalo_valido(self.db.get_config().get('check_interval_min'))

                if intervalo != self._intervalo_atual:
                    schedule.clear()
                    schedule.every(intervalo).minutes.do(self._executar_ciclo)
                    self._intervalo_atual = intervalo
                    logger.info(f"Agendador ativo com intervalo de {intervalo} minutos.")

                schedule.run_pending()
            except Exception:
                logger.error(f"Erro no laço do agendador (o agendador continua):\n{format_exc()}")

            time.sleep(1)

    def _executar_ciclo(self):
        try:
            self.processor.process_cycle()
        except Exception:
            logger.error(f"Erro no ciclo automático:\n{format_exc()}")

    def trigger_manual(self):
        if self.processor.em_execucao:
            logger.info("Já existe um ciclo em andamento. Disparo manual ignorado.")
            return False
        logger.info("Usuário disparou processamento manual.")
        t = threading.Thread(target=self._executar_ciclo, daemon=True)
        t.start()
        return True
        
    def start(self):
        # Start scheduler loop in a background thread
        t = threading.Thread(target=self.run_scheduler)
        t.daemon = True
        t.start()
        
        # Start GUI in the main thread
        run_gui(self.db, self.trigger_manual, self.processor)
            
if __name__ == '__main__':
    app = AppOrchestrator()
    app.start()
