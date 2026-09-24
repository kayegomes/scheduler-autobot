# GSD - Documentation & System Overview
## Scheduler AutoBot - Gestão e Monitoramento de Escalas

Este diretório `gsd/` é dedicado à documentação técnica, especificação de arquitetura e registro de correções do sistema **Scheduler AutoBot**.

---

## 📌 Índice de Documentos

1. [Visão Geral e Arquitetura](file:///c:/Users/ligomes/Downloads/atualizacao_solution/gsd/README.md)
2. [Fluxo de Processamento de Escalas e Grade TV](file:///c:/Users/ligomes/Downloads/atualizacao_solution/gsd/FLUXO_ESCALAS.md)
3. [Estrutura do Banco de Dados](file:///c:/Users/ligomes/Downloads/atualizacao_solution/gsd/ESTRUTURA_BANCO.md)
4. [Histórico de Correções e Suporte a Formatos (.xlsm)](file:///c:/Users/ligomes/Downloads/atualizacao_solution/gsd/CORRECOES_REALIZADAS.md)

---

## 🚀 Como Executar o Sistema

```bash
# Iniciar a Interface Gráfica (CustomTkinter)
python main.py
```

### Executar Testes de Validação
```bash
# Testar leitura de grades de TV (.xlsx, .xlsm)
python test_carregar_grade.py

# Simular recebimento de e-mail e processamento completo
python run_test.py
```
