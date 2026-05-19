import customtkinter as ctk
from core.database import DatabaseManager

class SettingsTab:
    def __init__(self, parent: ctk.CTkFrame, db: DatabaseManager):
        self.parent = parent
        self.db = db
        
        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_columnconfigure(1, weight=1)
        
        # Load configs
        self.configs = self.db.get_config()
        
        # Config UI Elements
        row = 0
        self.create_label_entry("Palavra-chave do Assunto:", "mail_subject_keyword", row)
        row += 1
        
        self.create_label_entry("Pasta do Outlook:", "outlook_folder", row)
        row += 1
        
        self.create_label_entry("Assunto do E-mail Enviado:", "notification_subject", row)
        row += 1
        
        self.create_label_entry("Intervalo Verificação (min):", "check_interval_min", row)
        row += 1
        
        self.create_label_entry("Termos Imunes (separados por vírgula):", "immune_keywords", row)
        row += 1
        
        # Modo Teste Checkbox
        self.test_mode_var = ctk.StringVar(value=str(self.configs.get("test_mode", "1")))
        self.check_test = ctk.CTkCheckBox(self.parent, text="Modo Teste (Simula envio no console)", variable=self.test_mode_var, onvalue="1", offvalue="0")
        self.check_test.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=10)
        row += 1
        
        # Email Mapping Text Area (Simple way to assign emails in MVP)
        lbl_mapping = ctk.CTkLabel(self.parent, text="Mapeamento de E-mails (Nome=email@dominio):")
        lbl_mapping.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=(10, 0))
        row += 1
        
        self.txt_mapping = ctk.CTkTextbox(self.parent, height=150)
        self.txt_mapping.grid(row=row, column=0, columnspan=2, sticky="ew", padx=20, pady=5)
        self.load_emails()
        row += 1
        
        # Save Button
        self.btn_save = ctk.CTkButton(self.parent, text="Salvar Configurações", command=self.save_configs)
        self.btn_save.grid(row=row, column=1, sticky="e", padx=20, pady=20)
        
        self.entries = {}
        for key, entry in self._entry_refs:
            self.entries[key] = entry
            
    _entry_refs = []

    def create_label_entry(self, label_text, config_key, row):
        lbl = ctk.CTkLabel(self.parent, text=label_text)
        lbl.grid(row=row, column=0, sticky="w", padx=20, pady=10)
        
        val = str(self.configs.get(config_key, ""))
        entry = ctk.CTkEntry(self.parent, width=300)
        entry.insert(0, val)
        entry.grid(row=row, column=1, sticky="ew", padx=20, pady=10)
        
        self._entry_refs.append((config_key, entry))

    def load_emails(self):
        with self.db.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nome, email FROM funcionarios")
            text = "\n".join([f"{nome}={email}" for nome, email in cursor.fetchall()])
            self.txt_mapping.insert("0.0", text)

    def save_configs(self):
        # Save basic configs
        for key, entry in self.entries.items():
            self.db.update_config(key, entry.get())
            
        self.db.update_config("test_mode", self.test_mode_var.get())
        
        # Save emails
        mapping_text = self.txt_mapping.get("0.0", "end").strip()
        lines = mapping_text.split('\n')
        email_dict = {}
        for line in lines:
            if '=' in line:
                nome, email = line.split("=", 1)
                email_dict[nome.strip()] = email.strip()
                
        if email_dict:
            self.db.update_employee_emails(email_dict)
            
        # UI Feedback
        self.btn_save.configure(text="Salvo com Sucesso!", fg_color="green")
        self.parent.after(2000, lambda: self.btn_save.configure(text="Salvar Configurações", fg_color=["#3B8ED0", "#1F6AA5"]))
