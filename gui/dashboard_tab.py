import customtkinter as ctk
import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from core.database import DatabaseManager
import matplotlib

# Set matplotlib to use a non-interactive backend by default to avoid issues
matplotlib.use("TkAgg")

class DashboardTab:
    def __init__(self, parent: ctk.CTkFrame, db: DatabaseManager):
        self.parent = parent
        self.db = db
        
        self.parent.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.parent.grid_rowconfigure(1, weight=1)
        
        # Cards
        self.cards = {}
        self.build_card("Alterações na Semana", "alteracoes_semana", 0, 0)
        self.build_card("Funcionários Afetados", "funcionarios_afetados", 0, 1)
        self.build_card("Última Verificação", "ultima_execucao", 0, 2)
        self.build_card("Total E-mails Enviados", "emails_enviados", 0, 3)
        
        # Charts Area (Left)
        self.frame_charts = ctk.CTkFrame(self.parent)
        self.frame_charts.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
        
        # Top 5 Area (Right)
        self.frame_top5 = ctk.CTkFrame(self.parent)
        self.frame_top5.grid(row=1, column=3, sticky="nsew", padx=10, pady=10)
        
        lbl_top5 = ctk.CTkLabel(self.frame_top5, text="Top 5 Mais Afetados", font=ctk.CTkFont(weight="bold", size=16))
        lbl_top5.pack(pady=20)
        
        self.top5_labels = []
        for _ in range(5):
            lbl = ctk.CTkLabel(self.frame_top5, text="-", font=ctk.CTkFont(size=14))
            lbl.pack(pady=5, anchor="w", padx=20)
            self.top5_labels.append(lbl)
            
        self.canvas_widget = None
        self._assinatura_grafico = None

        # Initial load
        self.load_data()

    def build_card(self, title, key, col, row=0):
        frame = ctk.CTkFrame(self.parent, corner_radius=10)
        frame.grid(row=row, column=col, sticky="nsew", padx=10, pady=(0, 10))
        
        lbl_title = ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=14))
        lbl_title.pack(pady=(15, 0))
        
        lbl_val = ctk.CTkLabel(frame, text="--", font=ctk.CTkFont(size=32, weight="bold"))
        lbl_val.pack(pady=(5, 15))
        
        self.cards[key] = lbl_val

    def load_data(self):
        # Update Cards Data
        stats = self.db.get_dashboard_stats()
        self.cards['alteracoes_semana'].configure(text=str(stats.get('alteracoes_semana', 0)))
        self.cards['funcionarios_afetados'].configure(text=str(stats.get('funcionarios_afetados', 0)))
        
        last = stats.get('ultima_execucao')
        self.cards['ultima_execucao'].configure(text=last[:16] if last else "Nunca")
        self.cards['emails_enviados'].configure(text=str(stats.get('emails_enviados', 0)))
        
        # Pull data for charts
        df = self.db.get_recent_changes(limit=1000)
        if df is None or df.empty:
            return
            
        # Top 5 list
        top5 = df['funcionario'].value_counts().head(5)
        for lbl in self.top5_labels:
            lbl.configure(text="-") # reset
            
        for i, (name, count) in enumerate(top5.items()):
            if i < len(self.top5_labels):
                self.top5_labels[i].configure(text=f"{i+1}. {name} ({count} alterações)")
            
        # Redesenha os gráficos só quando os dados mudaram. Antes, a
        # atualização periódica reconstruía a figura do matplotlib do zero a
        # cada ciclo, mesmo sem nenhuma alteração nova.
        assinatura = (len(df), str(df['data_deteccao'].iloc[0]) if len(df) else '')
        if assinatura != self._assinatura_grafico:
            self._assinatura_grafico = assinatura
            self.draw_charts(df)

    def draw_charts(self, df):
        if self.canvas_widget:
            self.canvas_widget.destroy()

        fig = Figure(figsize=(8, 4), dpi=100)
        fig.patch.set_facecolor('#2b2b2b') # CTk dark mode generic bg
        
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        
        for ax in [ax1, ax2]:
            ax.set_facecolor('#2b2b2b')
            ax.tick_params(colors='white')
            for spine in ax.spines.values(): 
                spine.set_color('white')
        
        # Gráfico de Pizza de Tipos
        types = df['tipo'].value_counts()
        colors = ['#f39c12', '#2ecc71', '#e74c3c'] # Laranja, Verde, Vermelho
        ax1.pie(types.values, labels=types.index, autopct='%1.1f%%', textprops={'color': 'white'}, colors=colors[:len(types)])
        ax1.set_title('Distribuição por Tipo', color='white')
        
        # Gráfico de Barras dos últimos 7 dias com alterações
        try:
            df['dia'] = pd.to_datetime(df['data_deteccao']).dt.date
            days = df['dia'].value_counts().sort_index()[-7:]
            
            ax2.bar([d.strftime('%d/%m') for d in days.index], days.values, color='#3B8ED0')
            ax2.set_title('Alterações Detectadas (Últimos Dias)', color='white')
            # Rotaciona labels para caber
            ax2.tick_params(axis='x', rotation=45)
        except Exception:
            pass # Ignora erros de data parsing no gráfico
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.frame_charts)
        self.canvas_widget = canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        canvas.draw()
