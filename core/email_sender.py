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
                          subject: str = "Aviso: Atualização na sua Escala de Trabalho"):
        
        if not to_address:
            logger.warning(f"Não há e-mail para {employee_name}. Pulando envio.")
            return False
            
        html_body = self._generate_html_body(employee_name, changes)
        text_body = self._generate_text_body(employee_name, changes)
        
        if self.test_mode:
            logger.info(f"[TEST MODO] E-mail simulado para: {to_address}")
            logger.info(f"Assunto: {subject}")
            logger.debug(f"Corpo: \n{text_body}")
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

    def _generate_html_body(self, name: str, changes: List[Dict[str, Any]]) -> str:
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; color: #333; }}
                table {{ border-collapse: collapse; width: 100%; max-width: 800px; }}
                th, td {{ padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }}
                th {{ background-color: #f2f2f2; }}
                .tag-entrada {{ color: #2ecc71; font-weight: bold; }}
                .tag-saida {{ color: #e74c3c; font-weight: bold; }}
                .tag-alt {{ color: #f39c12; font-weight: bold; }}
            </style>
        </head>
        <body>
            <h3>Olá, {name}!</h3>
            <p>Sua escala de trabalho possui novas atualizações. Seguem os detalhes abaixo:</p>
            <table>
                <tr>
                    <th>Data</th>
                    <th>Tipo</th>
                    <th>Campo Alterado</th>
                    <th>Valor Anterior</th>
                    <th>Novo Valor</th>
                </tr>
        """
        
        for c in changes:
            tipo = c['tipo']
            c_class = "tag-alt"
            if tipo == "ENTRADA": c_class = "tag-entrada"
            elif tipo == "SAIDA": c_class = "tag-saida"
            
            html += f"""
                <tr>
                    <td>{c['data']}</td>
                    <td class="{c_class}">{tipo}</td>
                    <td>{c['campo']}</td>
                    <td>{c['valor_antigo']}</td>
                    <td>{c['valor_novo']}</td>
                </tr>
            """
            
        html += """
            </table>
            <br>
            <p>Atenciosamente,<br>Equipe de Planejamento/Escala</p>
        </body>
        </html>
        """
        return html
        
    def _generate_text_body(self, name: str, changes: List[Dict[str, Any]]) -> str:
        text = f"Olá, {name}!\n\nSua escala de trabalho foi atualizada. Detalhes:\n"
        for c in changes:
            text += f"- Data: {c['data']} | {c['tipo']} | Campo: {c['campo']} | Antigo: {c['valor_antigo']} -> Novo: {c['valor_novo']}\n"
        text += "\nAtenciosamente,\nEquipe de Escala"
        return text
