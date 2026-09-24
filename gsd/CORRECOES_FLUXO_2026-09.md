# Correções de Fluxo — Setembro/2026

Auditoria do ciclo completo (Outlook → diff → e-mail) com reprodução sobre os
arquivos reais do repositório. 20 defeitos corrigidos, 5 deles críticos.

---

## 🔴 Críticos

### 1. Datas viravam epoch no banco
`database.save_new_grade()` gravava `row.to_json()`, que converte `Timestamp`
em epoch (`1782864000000`). Na releitura o valor voltava como inteiro e nunca
mais casava com a planilha nova. Como `data` compõe a chave de comparação, o
diff não alinhava nenhuma linha.

**Reprodução:** `Check_Pre_Envio_Gerado.xlsx` comparado com ele mesmo →
172 alterações (86 SAIDA + 86 ENTRADA), 42 funcionários notificados.

**Correção:** `core/columns.py` normaliza datas para `dd/mm/aaaa` e horários
para `HH:MM` *antes* da gravação. Migração `schema_version = 2` renormaliza o
histórico existente (testada em 9.640 itens reais, 2s, com backup automático).

### 2. Escala real confundida com grade de TV
`carregar_grade()` detectava "blocos lado a lado" com `"DATA" in header`, e a
coluna `Data_raw` da escala contava como segundo bloco. O achatamento fatiava a
planilha a partir da 1ª coluna `DATA`, descartando `Nome`. Com isso a proteção
de `schedule_processor` (que procurava `NOME`) nunca disparava, e a escala caía
no ramo de cruzamento — onde `df_new = df_old.copy()` **descartava o arquivo
recebido inteiro**.

**Correção:** comparação exata (`header.strip() == "DATA"`) e nova função
`detectar_tipo_planilha()`, que classifica pelo cabeçalho bruto antes de
qualquer transformação. Valida os 5 arquivos reais do repositório.

### 3. `evento_col` resolvia para coluna inexistente
O fallback era a string literal `'Evento/Programa'`, que jamais poderia existir
(as colunas são minúsculas na releitura). Na escala real a coluna é
`Atividade/Descrição`. Resultado: evento sempre `None` → nenhum match → todas
as 86 linhas marcadas `CANCELADO`.

**Correção:** resolução canônica única em `core/columns.py`. E o cruzamento
agora **aborta** se a coluna de evento não for reconhecida, em vez de cancelar.

### 4. Coluna fantasma por acento
A escala tem `Início`; o código escrevia em `inicio`. `df.at[...]` cria a
coluna quando ela não existe, então os horários atualizados e os `CANCELADO`
iam para uma coluna paralela que o diff ignorava.

**Correção:** `Início` e `inicio` passam a ser a mesma coluna canônica.

### 5. E-mail marcado como processado antes do envio
`source_email_id` era gravado antes da comparação e do envio. Qualquer exceção
adiante caía no `except` genérico e, no ciclo seguinte, `is_email_processed()`
retornava `True` — notificação perdida em definitivo, sem retry.

**Correção:** a grade nasce com `status = 'pending'` e só vira `'done'` depois
das notificações. `is_email_processed()` e `get_last_grade_df()` só enxergam
`'done'`; restos de tentativas falhas são descartados na próxima rodada.

---

## 🟠 Alto

| # | Defeito | Correção |
| :-- | :--- | :--- |
| 6 | `update_config` usava `UPDATE ... WHERE key`, descartando em silêncio chaves fora do `DEFAULT_CONFIG` — "Termos Imunes" e "Última Verificação" nunca salvavam | Upsert (`ON CONFLICT DO UPDATE`) + chaves acrescentadas ao `DEFAULT_CONFIG` |
| 7 | `UPDATE changes SET email_enviado=1 WHERE grade_id=?` marcava tudo, inclusive quem não tinha e-mail | Marcação por alteração, a partir do retorno de `send_notification()` |
| 8 | `KeyError: 'funcionario'` derrubava o lote inteiro de e-mails | Coluna canônica garantida + validação antes de salvar a grade |
| 9 | Fuzzy match em 0.85 entregava escala à pessoa errada (Marcos/Márcio = 0.94) | Match exato sem acento resolve o caso real; fuzzy em 0.92 exigindo 1º e último nome iguais, sempre com log |
| 10 | Agendador e botão manual rodavam concorrentes → e-mails duplicados | `threading.Lock` não-bloqueante em `process_cycle()` |

---

## 🟡 Médio

| # | Defeito | Correção |
| :-- | :--- | :--- |
| 11 | `dates_match` comparava `[:5]`: `'2026-07-01'` casava com `'2026-12-25'` | Comparação por data normalizada |
| 12 | Intervalo lido só no boot; valor inválido matava a thread em silêncio | Releitura a cada volta + parsing tolerante + laço protegido |
| 13 | Anexos rebaixados a cada ciclo, pasta sem limpeza | Filtro de já-processado antes do download + retenção de 30 dias |
| 14 | Cada clique em Sincronizar deixava um laço de refresh extra para sempre | Cadeia única com `after_cancel` |
| 15 | `data_deteccao` em UTC (`CURRENT_TIMESTAMP`) e resto em hora local | Hora local explícita |
| 16 | Remover linha do mapeamento não desfazia o vínculo | `update_employee_emails(substituir=True)` |
| 17 | `_entry_refs` era atributo de classe compartilhado | Atributo de instância |
| 18 | Base sempre "a última grade", sem noção de período | Trava `max_funcionarios_por_ciclo` (padrão 25) bloqueia disparo em massa |
| 19 | Histórico recarregava a cada 5s, perdendo o scroll | Redesenho só quando os dados mudam; intervalo para 10s |
| 20 | `except: pass` nu escondia erros de UI | `logger.exception` |

**Menores:** bloco de 9 linhas duplicado em `email_sender`; saudação fixa em
"boa tarde"; `emails_teste/` relativo ao CWD; `cols_to_compare` ignorava
colunas removidas; chave do diff sensível à ordem das linhas; conexões SQLite
sem `close()`.

---

## ✅ Validação

`tests/test_regressao.py` — 26 testes, um por defeito, rodando sobre os
arquivos reais do repositório:

```bash
python -m unittest discover -s tests -v
```

Homologação de ponta a ponta sem tocar no Outlook:

```bash
python run_test.py [arquivo_recebido.xlsx]
```

`run_test.py` foi reescrito para chamar o `ScheduleProcessor` real — antes ele
reimplementava o fluxo inline, e era por isso que o cruzamento quebrado
"passava" no teste.

### Antes × depois, com os arquivos reais

| Cenário | Antes | Depois |
| :--- | :--- | :--- |
| Escala idêntica reprocessada | 172 alterações, 42 e-mails | 0 alterações |
| Escala com 3 mudanças reais | 0 detectadas (arquivo descartado) | exatamente 3 |
| Grade Combate sobre escala de julho | 83 eventos cancelados por engano | cruzamento abortado (8% de match) |
| Migração do banco real | — | 9.640 itens, 2s, backup automático |
