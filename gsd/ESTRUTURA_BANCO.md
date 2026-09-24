# Estrutura do Banco de Dados SQLite

O sistema armazena as grades importadas, histórico de alterações, configurações e mapeamento de funcionários no banco de dados SQLite (`data/scheduler.db`).

---

## 🗄️ Tabelas

### 1. `config`
Armazena parâmetros de execução (ex.: palavra-chave do assunto do e-mail, pasta do Outlook, modo de teste, assuntos de notificação).

### 2. `grades`
Registra o cabeçalho de cada importação realizada.
- `id`: Identificador sequencial
- `import_date`: Data e hora da importação
- `filename`: Nome do arquivo da planilha baixada
- `source_email_id`: ID único do e-mail no Outlook (`EntryID`)

### 3. `grade_items`
Contém as linhas individuais da escala extraída.
- `grade_id`: FK apontando para `grades.id`
- `funcionario`: Nome do funcionário escalado
- `data`: Data da escala
- `horario`: Horário do evento/turno
- `evento`: Descrição do programa ou evento
- `raw_data`: String JSON contendo todas as colunas da planilha original

### 4. `changes`
Histórico de diferenças encontradas entre versões da escala.
- `grade_id`: FK da nova grade
- `funcionario`: Funcionário afetado
- `data_escala`: Data do evento modificado
- `tipo`: Tipo de alteração (`ENTRADA`, `SAIDA`, `ALTERACAO`)
- `campo`: Campo alterado (ex.: horario, evento)
- `valor_antigo`: Valor anterior
- `valor_novo`: Novo valor
- `data_deteccao`: Data/hora do processamento
- `email_enviado`: Flag `0` (não enviado) ou `1` (enviado)

### 5. `funcionarios`
Mapeamento de e-mails para cada funcionário.
- `nome`: Nome completo do funcionário (PK)
- `email`: Endereço de e-mail para envio de notificações
