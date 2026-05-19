import logging
from typing import Dict, List, Any

logger = logging.getLogger(__name__)

class EmailSender:
    def __init__(self, test_mode: bool = True):
        self.test_mode = test_mode
    
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
                          subject: str = "Aviso: Atualização na sua Escala de Trabalho"):
        
        if not to_address:
            logger.warning(f"Não há e-mail para {employee_name}. Pulando envio.")
            return False
            
        html_body = self._generate_html_body(employee_name, changes, df_schedule)
        text_body = self._generate_text_body(employee_name, changes)
        
        if self.test_mode:
            import os
            from datetime import datetime
            
            logger.info(f"[TEST MODO] E-mail simulado para: {to_address}")
            logger.info(f"Assunto: {subject}")
            
            # Opção 1: Salva o e-mail em uma pasta local para preview
            preview_dir = "emails_teste"
            os.makedirs(preview_dir, exist_ok=True)
            
            safe_name = "".join(c for c in employee_name if c.isalnum() or c in (' ', '_')).rstrip().replace(' ', '_')
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(preview_dir, f"escala_{safe_name}_{timestamp}.html")
            
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write(html_body)
                logger.info(f"[TEST MODO] Preview HTML salvo em: {filename}")
            except Exception as e:
                logger.error(f"[TEST MODO] Erro ao salvar arquivo HTML: {e}")
                
            return True

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

    def _generate_html_body(self, name: str, changes: List[Dict[str, Any]], df_schedule: Any) -> str:
        # Build intro text based on changes
        intro_texts = []
        for c in changes:
            if c['tipo'] == 'SAIDA':
                intro_texts.append(f"<span style='background-color: yellow; font-weight: bold;'>CAIU {c.get('campo', 'Evento')}</span> na sua escala do dia {c['data']}.")
            elif c['tipo'] == 'ENTRADA':
                intro_texts.append(f"<span style='background-color: yellow; font-weight: bold;'>NOVA ESCALA</span> adicionada para o dia {c['data']}.")
            else:
                intro_texts.append(f"<span style='background-color: yellow; font-weight: bold;'>Houve alteração de {c['campo']}</span> na sua escala do dia {c['data']}.")

        intro_html = "<br>".join(intro_texts) if intro_texts else "Sua escala de trabalho foi atualizada."

        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Calibri, Arial, sans-serif; color: #000; font-size: 14px; }}
                table {{ border-collapse: collapse; width: 100%; max-width: 1200px; font-size: 12px; margin-top: 15px; }}
                th, td {{ border: 1px solid #a0a0a0; padding: 4px; text-align: left; vertical-align: middle; }}
                th {{ background-color: #f2f2f2; font-weight: normal; }}
                .highlight {{ background-color: #ffff00; font-weight: bold; }}
            </style>
        </head>
        <body>
            <p>Oi {name}, boa tarde! Tudo bem?</p>
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
        
        # Helper to check if a specific cell changed
        def get_cell_class(data, col_name):
            for c in changes:
                if c['data'] == data:
                    if c['tipo'] == 'ENTRADA':
                        return 'highlight'
                    if c['tipo'] == 'ALTERACAO':
                        # Simplification: match column names loosely
                        if col_name.lower() in c['campo'].lower() or c['campo'].lower() in col_name.lower():
                            return 'highlight'
            return ''

        # Render schedule table
        import pandas as pd
        if df_schedule is not None and not df_schedule.empty:
            # Pega as colunas disponíveis, tentando mapear para as desejadas
            cols = [c for c in df_schedule.columns if c.lower() != 'funcionario']
            
            for _, row in df_schedule.iterrows():
                data_val = str(row.get('data', ''))
                dia_val = str(row.get('dia', row.get('dia da semana', '')))
                
                # Check for row-level highlighting if it's an 'ENTRADA'
                is_new_row = any(c['tipo'] == 'ENTRADA' and c['data'] == data_val for c in changes)
                row_class = "highlight" if is_new_row else ""
                
                def td(val, col_name=""):
                    cls = "highlight" if get_cell_class(data_val, col_name) and not is_new_row else ""
                    return f"<td class='{cls}'>{val if pd.notna(val) else ''}</td>"

                html += f"""
                <tr class="{row_class}">
                    {td(name, 'nome')}
                    {td(row.get('plataforma', row.get('canal', '')), 'plataforma')}
                    {td(data_val, 'data')}
                    {td(dia_val, 'dia')}
                    {td(row.get('pre', row.get('pré', '')), 'pre')}
                    {td(row.get('inicio', row.get('início', '')), 'inicio')}
                    {td(row.get('fim', ''), 'fim')}
                    {td(row.get('evento/descricao', row.get('evento/programa', row.get('evento', ''))), 'evento')}
                    {td(row.get('produto', ''), 'produto')}
                    {td(row.get('local', ''), 'local')}
                    {td(row.get('elenco', ''), 'elenco')}
                    {td(row.get('coordenador', ''), 'coordenador')}
                    {td(row.get('produtor', ''), 'produtor')}
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
        
    def _generate_text_body(self, name: str, changes: List[Dict[str, Any]]) -> str:
        text = f"Olá, {name}!\n\nSua escala de trabalho foi atualizada. Detalhes:\n"
        for c in changes:
            text += f"- Data: {c['data']} | {c['tipo']} | Campo: {c['campo']} | Antigo: {c['valor_antigo']} -> Novo: {c['valor_novo']}\n"
        text += "\nQualquer dúvida, estamos à disposição."
        return text
