# Correções Realizadas - Suporte a Formatos .xlsm e Validações

## 📝 Resumo da Correção

**Problema relatado:** O fluxo automático de e-mails não estava processando corretamente as grades recebidas em anexos com o formato `.xlsm` (planilhas do Excel habilitadas para macro).

### Causa Raiz Diagnosticada:
1. `OutlookMonitor` (em `core/outlook_monitor.py`): O filtro de anexos utilizava apenas `endswith(('.xlsx', '.csv', '.xls'))`, descartando anexos `.xlsm`.
2. `MainWindow` (em `gui/main_window.py`): O filtro de seleção de arquivos na GUI para *Carregar Escala Base Inicial* não exibia arquivos `.xlsm` e `.xls`.
3. Verificações de extensão `.csv` usavam `endswith('.csv')` de forma *case-sensitive*, o que poderia ignorar extensoes maiúsculas.

---

## 🛠️ Modificações Aplicadas

1. **`core/outlook_monitor.py`**:
   - Atualizado filtro para incluir `.xlsm`:
     ```python
     if filename.lower().endswith(('.xlsx', '.xlsm', '.csv', '.xls')):
     ```

2. **`gui/main_window.py`**:
   - Atualizados tipos de arquivos permitidos no diálogo `filedialog.askopenfilename`:
     ```python
     filetypes=[("Planilhas e CSV", "*.xlsx *.xlsm *.xls *.csv")]
     ```
   - Corrigida verificação de extensão para `filepath.lower().endswith('.csv')`.

3. **`core/schedule_processor.py`**:
   - Ajustada verificação para `att_path.lower().endswith('.csv')`.

---

## 🧪 Validação dos Testes

- Executado o script `test_carregar_grade.py` contendo o arquivo `GRADE DE JULHO - 14ª VERSÃO.xlsm`.
- O leitor `openpyxl` processou com êxito **8.941 linhas** da grade `.xlsm`.
