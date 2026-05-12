# Scheduler AutoBot - Gestão e Automação de Escalas

Aplicação desktop completa para monitoramento, comparação e notificação automática de alterações em escalas de trabalho, desenhada para produção e construída em Python moderno com interface gráfica.

## Funcionalidades Principais
1. **Monitoramento de E-mail:** Conecta-se automaticamente à caixa de entrada do Outlook e detecta e-mails contendo anexos de escalas (Excel/CSV) através de palavras-chave.
2. **Motor de Comparação (Diff Engine):** Detecta entradas, saídas e alterações a nível de célula entre a grade anterior e a recém recebida.
3. **Notificação Personalizada:** Dispara e-mails por Microsoft Outlook informando para os funcionários quais foram exatamente as alterações na escala deles.
4. **Dashboard Local:** Mantém banco de dados com todo histórico para consulta.
5. **Interface Gráfica Moderna (GUI):** Interface intuitiva para supervisores criarem os mapeamentos de e-mails, ajustarem configurações e monitorarem os gráficos semanais.

## Estrutura Automática
O sistema opera via **thread independente** a cada `X` minutos. Todo o core dataflow acontece no `core/schedule_processor.py`. O histórico total está persistido num banco leve `SQLite`.

## Instalação e Execução

### Pré-requisitos
- Python 3.10+ instalado no Windows.
- Microsoft Outlook instalado e configurado no mesmo usuário logado.

### Passos
1. Abra um terminal na raiz do projeto.
2. Instale as dependências executando:
   ```bash
   pip install -r requirements.txt
   ```
3. Execute a aplicação inicial:
   ```bash
   python main.py
   ```

## Configuração Inicial recomendada via GUI
1. Vá até a aba "Configurações".
2. Preencha o "Mapeamento de E-mails" no formato `Nome Exato da Planilha=email@empresa.com`.
3. Garanta que o "Modo Teste" está assinalado como `1` durante os testes (para não encher caixas com SPAM).
4. Sincronize (manualmente via botão) ou aguarde o ciclo do relógio!

---
_Desenvolvido usando Pandas, CustomTkinter, Matplotlib e pywin32._
