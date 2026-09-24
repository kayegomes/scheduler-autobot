import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from typing import Optional
from core.database import DatabaseManager

class ChangesTab:
    def __init__(self, parent: ctk.CTkFrame, db: DatabaseManager):
        self.parent = parent
        self.db = db
        
        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_rowconfigure(1, weight=1)
        
        # Filtros
        self.filter_frame = ctk.CTkFrame(self.parent)
        self.filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.lbl_func = ctk.CTkLabel(self.filter_frame, text="Funcionário:")
        self.lbl_func.pack(side="left", padx=(10, 5), pady=10)
        self.entry_func = ctk.CTkEntry(self.filter_frame, width=150)
        self.entry_func.pack(side="left", padx=5)
        self.entry_func.bind("<KeyRelease>", lambda e: self.apply_filter())
        
        self.btn_export = ctk.CTkButton(self.filter_frame, text="Exportar p/ Excel", command=self.export_excel)
        self.btn_export.pack(side="right", padx=10)
        
        # Tabela (Treeview)
        # Usamos ttk.Treeview com um estilo adaptado
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                        background="#2b2b2b",
                        foreground="white",
                        rowheight=25,
                        fieldbackground="#2b2b2b",
                        bordercolor="#343638",
                        borderwidth=0)
        style.map('Treeview', background=[('selected', '#1f538d')])
        style.configure("Treeview.Heading",
                        background="#565b5e",
                        foreground="white",
                        relief="flat")
        style.map("Treeview.Heading", background=[('active', '#343638')])
        
        columns = ("Data Detecção", "Funcionário", "Data Escala", "Tipo", "Campo", "Antigo", "Novo", "Enviado")
        self.tree = ttk.Treeview(self.parent, columns=columns, show="headings")
        
        for col in columns:
            self.tree.heading(col, text=col)
            # Definir larguras
            width = 150 if col in ("Data Detecção", "Funcionário") else 100
            if col in ("Antigo", "Novo"): width = 200
            self.tree.column(col, width=width, anchor="w")
            
        self.tree.grid(row=1, column=0, sticky="nsew")
        
        # Scrollbar
        self.scrollbar = ctk.CTkScrollbar(self.parent, command=self.tree.yview)
        self.scrollbar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=self.scrollbar.set)
        
        # Dados originais para filtrar
        self.all_data = []
        self._assinatura = None
        self.load_data()

    def load_data(self):
        df = self.db.get_recent_changes(limit=200)
        registros = [] if df is None or df.empty else df.fillna('').to_dict('records')

        # A atualização periódica reconstruía a tabela inteira a cada ciclo,
        # jogando o scroll de volta para o topo enquanto o usuário navegava.
        # Só redesenha quando os dados realmente mudaram.
        assinatura = (len(registros), tuple(r.get('id') for r in registros[:50]),
                      tuple(r.get('email_enviado') for r in registros[:50]))
        if assinatura == self._assinatura:
            return

        self._assinatura = assinatura
        self.all_data = registros
        self.apply_filter()

    def apply_filter(self):
        term = self.entry_func.get().lower()

        # Limpa tabela
        for item in self.tree.get_children():
            self.tree.delete(item)

        filtered = [r for r in self.all_data if term in str(r.get('funcionario', '')).lower()]

        for r in filtered:
            data_detec = str(r.get('data_deteccao', ''))[:16] # tira segundos e resto
            enviado = "Sim" if r.get('email_enviado') else "Não"
            vals = (data_detec, r.get('funcionario', ''), r.get('data_escala', ''),
                    r.get('tipo', ''), r.get('campo', ''), r.get('valor_antigo', ''),
                    r.get('valor_novo', ''), enviado)
            self.tree.insert("", "end", values=vals)

    def export_excel(self):
        import pandas as pd
        from tkinter import filedialog
        
        if not self.all_data:
            return
            
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Exportar Histórico"
        )
        
        if filepath:
            df = pd.DataFrame(self.all_data)
            df.to_excel(filepath, index=False)
