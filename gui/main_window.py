import logging
from contextlib import contextmanager

import customtkinter as ctk

from core.database import DatabaseManager
from core.schedule_processor import CicloEmAndamento
from .dashboard_tab import DashboardTab
from .changes_tab import ChangesTab
from .settings_tab import SettingsTab

logger = logging.getLogger(__name__)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

INTERVALO_REFRESH_MS = 10000


class MainWindow(ctk.CTk):
    def __init__(self, db: DatabaseManager, trigger_manual_callback, processor=None):
        super().__init__()
        self.title("Scheduler AutoBot - Gestão de Escalas")
        self.geometry("1100x750")

        self.db = db
        self.trigger_manual = trigger_manual_callback
        self.processor = processor
        self._refresh_job = None

        # Main Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Top bar
        self.top_frame = ctk.CTkFrame(self, height=50, corner_radius=0)
        self.top_frame.grid(row=0, column=0, sticky="ew")

        self.title_label = ctk.CTkLabel(
            self.top_frame, text="Monitoramento de Escalas de Trabalho",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.title_label.pack(side="left", padx=20, pady=10)

        self.btn_sync = ctk.CTkButton(self.top_frame, text="Sincronizar Agora", command=self.on_sync)
        self.btn_sync.pack(side="right", padx=20, pady=10)

        self.btn_load_base = ctk.CTkButton(
            self.top_frame, text="Carregar Escala Base Inicial", fg_color="#2ecc71",
            hover_color="#27ae60", command=self.on_load_base
        )
        self.btn_load_base.pack(side="right", padx=10, pady=10)

        # Tabview
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 20))

        self.tab_dash = self.tabview.add("Dashboard")
        self.tab_hist = self.tabview.add("Histórico de Alterações")
        self.tab_cfg = self.tabview.add("Configurações")

        # Init views
        self.dashboard_view = DashboardTab(self.tab_dash, self.db)
        self.changes_view = ChangesTab(self.tab_hist, self.db)
        self.settings_view = SettingsTab(self.tab_cfg, self.db)

        self._agendar_refresh()

    # ------------------------------------------------------------------

    def _agendar_refresh(self, delay_ms: int = INTERVALO_REFRESH_MS):
        """Mantém UMA única cadeia de atualização.

        Antes, `refresh_data` sempre se reagendava e `on_sync` agendava mais
        uma: cada clique em Sincronizar deixava para trás um laço de 5s extra,
        redesenhando os gráficos do matplotlib para sempre.
        """
        if self._refresh_job is not None:
            try:
                self.after_cancel(self._refresh_job)
            except ValueError:
                pass
        self._refresh_job = self.after(delay_ms, self.refresh_data)

    def refresh_data(self):
        try:
            self.dashboard_view.load_data()
            self.changes_view.load_data()
        except Exception:
            logger.exception("Falha ao atualizar a interface.")
        finally:
            self._refresh_job = None
            self._agendar_refresh()

    # ------------------------------------------------------------------

    def on_sync(self):
        iniciou = self.trigger_manual()
        if iniciou is False:
            self.btn_sync.configure(text="Já sincronizando...")
            self.after(3000, lambda: self.btn_sync.configure(text="Sincronizar Agora"))
            return

        self.btn_sync.configure(state="disabled", text="Sincronizando...")
        self.after(5000, lambda: self.btn_sync.configure(state="normal", text="Sincronizar Agora"))
        self._agendar_refresh(2000)

    def on_load_base(self):
        from tkinter import filedialog, messagebox
        import os
        from core.columns import carregar_escala

        filepath = filedialog.askopenfilename(
            title="Selecione a Escala Base Inicial",
            filetypes=[("Planilhas e CSV", "*.xlsx *.xlsm *.xls *.csv")]
        )
        if not filepath:
            return

        try:
            df = carregar_escala(filepath)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo:\n{e}")
            return

        # Sem coluna de funcionário a escala não serve de base: o diff
        # descartaria todas as linhas e o próximo ciclo trataria a escala real
        # como se todo mundo tivesse entrado agora.
        if 'funcionario' not in df.columns:
            messagebox.showerror(
                "Planilha inválida",
                "Não encontrei uma coluna de funcionário nesta planilha.\n\n"
                "Esperado algo como: Nome, Funcionário, Escalado ou Talento.\n"
                f"Colunas lidas: {', '.join(list(df.columns)[:10])}"
            )
            return

        df = df[df['funcionario'].astype(str).str.strip() != '']
        if df.empty:
            messagebox.showerror("Planilha vazia", "Nenhuma linha com funcionário preenchido.")
            return

        pessoas = df['funcionario'].nunique()
        if not messagebox.askyesno(
            "Confirmar carga da escala base",
            f"Arquivo: {os.path.basename(filepath)}\n"
            f"Linhas: {len(df)}   |   Funcionários: {pessoas}\n\n"
            "Esta planilha passa a ser a BASE DE COMPARAÇÃO. A próxima grade recebida "
            "será comparada contra ela.\n\nConfirmar?"
        ):
            return

        try:
            # A carga precisa ser exclusiva: se o ciclo automático estiver
            # rodando, ele já leu a base antiga e vai gravar por cima desta.
            with self._carga_exclusiva():
                grade_id = self.db.save_new_grade(df, os.path.basename(filepath), "carga_manual")
            messagebox.showinfo(
                "Sucesso",
                f"Escala Base carregada! (ID: {grade_id})\n"
                f"Linhas: {len(df)} | Funcionários: {pessoas}"
            )
            self.refresh_data()
        except CicloEmAndamento:
            messagebox.showwarning(
                "Sincronização em andamento",
                "Há um ciclo de sincronização rodando agora.\n\n"
                "Aguarde ele terminar (veja o log) e carregue a escala de novo — "
                "senão a escala carregada seria sobrescrita pelo ciclo em curso."
            )
        except Exception as e:
            logger.exception("Falha ao salvar escala base.")
            messagebox.showerror("Erro", f"Falha ao salvar a escala base:\n{e}")

    @contextmanager
    def _carga_exclusiva(self):
        if self.processor is None:
            yield
        else:
            with self.processor.exclusivo(timeout=2):
                yield


def run_gui(db_manager: DatabaseManager, trigger_manual_callback, processor=None):
    app = MainWindow(db_manager, trigger_manual_callback, processor)
    app.mainloop()
