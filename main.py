import threading
import time
import schedule
import logging
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

class AppOrchestrator:
    def __init__(self):
        self.db = DatabaseManager()
        self.processor = ScheduleProcessor(self.db)
        self.is_running = True

    def run_scheduler(self):
        # Check initial config interval
        config = self.db.get_config()
        interval = int(config.get('check_interval_min', 15))
        
        schedule.every(interval).minutes.do(self.processor.process_cycle)
        
        logger.info(f"Agendador iniciado com intervalo de {interval} minutos.")
        
        while self.is_running:
            schedule.run_pending()
            time.sleep(1)

    def trigger_manual(self):
        logger.info("Usuário disparou processamento manual.")
        t = threading.Thread(target=self.processor.process_cycle)
        t.daemon = True
        t.start()
        
    def start(self):
        # Start scheduler loop in a background thread
        t = threading.Thread(target=self.run_scheduler)
        t.daemon = True
        t.start()
        
        # Start GUI in the main thread
        run_gui(self.db, self.trigger_manual)
            
if __name__ == '__main__':
    app = AppOrchestrator()
    app.start()
