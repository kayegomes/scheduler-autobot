import logging

import customtkinter as ctk

from core.database import DatabaseManager

logger = logging.getLogger(__name__)


class SettingsTab:
    def __init__(self, parent: ctk.CTkFrame, db: DatabaseManager):
        self.parent = parent
        self.db = db

        # Instância, não classe. Como atributo de classe, a lista era
        # compartilhada entre todas as SettingsTab criadas no processo: uma
        # segunda aba guardava widgets já destruídos e quebrava ao salvar.
        self.entries = {}

        self.parent.grid_columnconfigure(0, weight=1)
        self.parent.grid_columnconfigure(1, weight=1)

        self.configs = self.db.get_config()

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

        self.create_label_entry("Máx. funcionários por ciclo (0 = sem limite):",
                                "max_funcionarios_por_ciclo", row)
        row += 1

        self.create_label_entry("Pasta dos relatórios de teste (vazio = padrão):",
                                "test_output_dir", row)
        row += 1

        # Modo Teste Checkbox
        self.test_mode_var = ctk.StringVar(value=str(self.configs.get("test_mode", "1")))
        self.check_test = ctk.CTkCheckBox(
            self.parent,
            text="Modo Teste (não envia e-mail e não grava no histórico; gera relatório em pasta)",
            variable=self.test_mode_var, onvalue="1", offvalue="0"
        )
        self.check_test.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=10)
        row += 1

        self.lbl_pasta = ctk.CTkLabel(
            self.parent, text=self._texto_pasta(), text_color="gray", anchor="w"
        )
        self.lbl_pasta.grid(row=row, column=0, columnspan=2, sticky="w", padx=20)
        row += 1

        self.btn_abrir = ctk.CTkButton(
            self.parent, text="Abrir pasta de relatórios", width=200,
            fg_color="transparent", border_width=1, command=self.abrir_pasta_relatorios
        )
        self.btn_abrir.grid(row=row, column=0, sticky="w", padx=20, pady=(5, 10))
        row += 1

        lbl_mapping = ctk.CTkLabel(
            self.parent,
            text="Mapeamento de E-mails (Nome=email@dominio) — linhas removidas aqui são apagadas:"
        )
        lbl_mapping.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=(10, 0))
        row += 1

        self.txt_mapping = ctk.CTkTextbox(self.parent, height=150)
        self.txt_mapping.grid(row=row, column=0, columnspan=2, sticky="ew", padx=20, pady=5)
        self.load_emails()
        row += 1

        self.btn_save = ctk.CTkButton(self.parent, text="Salvar Configurações", command=self.save_configs)
        self.btn_save.grid(row=row, column=1, sticky="e", padx=20, pady=20)

    def _pasta_relatorios(self) -> str:
        from config import TEST_EMAILS_DIR
        from core.test_report import resolver_pasta_saida
        return resolver_pasta_saida(self.db.get_config(), TEST_EMAILS_DIR)

    def _texto_pasta(self) -> str:
        return f"Relatórios de homologação em: {self._pasta_relatorios()}"

    def abrir_pasta_relatorios(self):
        import os
        import subprocess
        from tkinter import messagebox

        pasta = self._pasta_relatorios()
        if not os.path.isdir(pasta):
            messagebox.showinfo(
                "Pasta ainda não existe",
                f"Nenhum ciclo em Modo Teste rodou ainda.\n\nEla será criada em:\n{pasta}"
            )
            return
        try:
            os.startfile(pasta)  # noqa: S606 - Windows
        except AttributeError:
            subprocess.Popen(["xdg-open", pasta])
        except OSError as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")

    def create_label_entry(self, label_text, config_key, row):
        lbl = ctk.CTkLabel(self.parent, text=label_text)
        lbl.grid(row=row, column=0, sticky="w", padx=20, pady=10)

        entry = ctk.CTkEntry(self.parent, width=300)
        entry.insert(0, str(self.configs.get(config_key, "")))
        entry.grid(row=row, column=1, sticky="ew", padx=20, pady=10)

        self.entries[config_key] = entry

    def load_emails(self):
        mapeamento = self.db.get_employee_mapping()
        texto = "\n".join(f"{nome}={email}" for nome, email in mapeamento.items())
        self.txt_mapping.delete("0.0", "end")
        self.txt_mapping.insert("0.0", texto)

    def save_configs(self):
        from tkinter import messagebox

        for key, entry in self.entries.items():
            self.db.update_config(key, entry.get())

        self.db.update_config("test_mode", self.test_mode_var.get())

        mapping_text = self.txt_mapping.get("0.0", "end").strip()
        email_dict = {}
        invalidas = []
        for numero, linha in enumerate(mapping_text.split('\n'), start=1):
            linha = linha.strip()
            if not linha:
                continue
            if '=' not in linha:
                invalidas.append(f"linha {numero}: {linha[:40]}")
                continue
            nome, email = linha.split("=", 1)
            nome, email = nome.strip(), email.strip()
            if not nome or '@' not in email:
                invalidas.append(f"linha {numero}: {linha[:40]}")
                continue
            email_dict[nome] = email

        if invalidas:
            messagebox.showwarning(
                "Linhas ignoradas no mapeamento",
                "Formato esperado: Nome=email@dominio\n\n" + "\n".join(invalidas[:10])
            )

        self.db.update_employee_emails(email_dict, substituir=True)
        self.load_emails()
        self.lbl_pasta.configure(text=self._texto_pasta())

        self.btn_save.configure(text="Salvo com Sucesso!", fg_color="green")
        self.parent.after(
            2000,
            lambda: self.btn_save.configure(text="Salvar Configurações", fg_color=["#3B8ED0", "#1F6AA5"])
        )
