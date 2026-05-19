import customtkinter as ctk
from core.database import DatabaseManager
from .dashboard_tab import DashboardTab
from .changes_tab import ChangesTab
from .settings_tab import SettingsTab

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class MainWindow(ctk.CTk):
    def __init__(self, db: DatabaseManager, trigger_manual_callback):
        super().__init__()
        self.title("Scheduler AutoBot - Gestão de Escalas")
        self.geometry("1100x750")
        
        self.db = db
        self.trigger_manual = trigger_manual_callback
        
        # Main Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Top bar
        self.top_frame = ctk.CTkFrame(self, height=50, corner_radius=0)
        self.top_frame.grid(row=0, column=0, sticky="ew")
        
        self.title_label = ctk.CTkLabel(self.top_frame, text="Monitoramento de Escalas de Trabalho", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(side="left", padx=20, pady=10)
        
        self.btn_sync = ctk.CTkButton(self.top_frame, text="Sincronizar Agora", command=self.on_sync)
        self.btn_sync.pack(side="right", padx=20, pady=10)
        
        self.btn_load_base = ctk.CTkButton(self.top_frame, text="Carregar Escala Base Inicial", fg_color="#2ecc71", hover_color="#27ae60", command=self.on_load_base)
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
        
        # Start auto refresh
        self.after(5000, self.refresh_data)
        
    def on_sync(self):
        self.btn_sync.configure(state="disabled", text="Sincronizando...")
        self.trigger_manual()
        
        # Re-enables button after 5 seconds to prevent spam
        self.after(5000, lambda: self.btn_sync.configure(state="normal", text="Sincronizar Agora"))
        # Force refresh immediately
        self.after(2000, self.refresh_data)

    def on_load_base(self):
        from tkinter import filedialog, messagebox
        import pandas as pd
        import os
        
        filepath = filedialog.askopenfilename(
            title="Selecione a Escala Base Inicial",
            filetypes=[("Excel/CSV files", "*.xlsx *.csv")]
        )
        if filepath:
            try:
                if filepath.endswith('.csv'):
                    df = pd.read_csv(filepath)
                else:
                    df = pd.read_excel(filepath)
                    
                filename = os.path.basename(filepath)
                grade_id = self.db.save_new_grade(df, filename, "carga_manual")
                messagebox.showinfo("Sucesso", f"Escala Base carregada com sucesso! (ID: {grade_id})\nLinhas importadas: {len(df)}")
                self.refresh_data()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao carregar o arquivo:\n{e}")

    def refresh_data(self):
        # Update UI components
        try:
            self.dashboard_view.load_data()
            self.changes_view.load_data()
        except:
            pass
        self.after(5000, self.refresh_data)

def run_gui(db_manager: DatabaseManager, trigger_manual_callback):
    app = MainWindow(db_manager, trigger_manual_callback)
    app.mainloop()
