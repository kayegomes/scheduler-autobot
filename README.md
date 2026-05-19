# Scheduler AutoBot 🤖🗓️

O **Scheduler AutoBot** é um aplicativo desktop projetado para automatizar o processo de detecção de alterações nas escalas de trabalho e notificação via e-mail dos funcionários afetados. Ele monitora diretamente o seu Microsoft Outlook, processa Grades de TV e envia e-mails lindamente formatados e coloridos avisando sobre as mudanças.

## Principais Funcionalidades

- **Monitoramento Automático do Outlook**: Lê e-mails em segundo plano utilizando a biblioteca `pywin32`.
- **Inteligência de Cruzamento (Match)**: Capacidade de receber "Grades da TV" (como Sportv) e cruzar de forma inteligente com a Escala de Trabalho original, detectando eventos que mudaram de horário ou caíram (foram cancelados).
- **Fuzzy Matching de E-mails**: Sistema resiliente para mapear funcionários. Se um nome vier escrito um pouco diferente da base (ex: sem acento), o sistema é capaz de identificá-lo através do *Fuzzy Matching* (com 85% de tolerância).
- **Notificações Ricas em HTML**: E-mails contendo a tabela completa do funcionário, destacando *em amarelo* apenas as linhas e colunas que sofreram alguma modificação.
- **Painel Gerencial (Dashboard)**: Acompanhamento visual via gráficos das alterações nos últimos 7 dias, funcionários mais impactados e métricas gerais do sistema.
- **Exportação do Histórico**: Ferramenta na aba "Histórico" para salvar todos os relatórios e logs de alterações no formato `.xlsx`.

---

## 🚀 Instalação e Configuração

### Pré-requisitos
- Python 3.9+
- Microsoft Outlook instalado e configurado na máquina (`win32com.client`)

### Passos de Instalação
1. Clone este repositório ou baixe os arquivos para o seu computador.
2. Abra o terminal na pasta do projeto e instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
3. Inicie a interface gráfica (Desktop App):
   ```bash
   python main.py
   ```

---

## 🛠️ Como Utilizar no Dia a Dia

Ao abrir o aplicativo `main.py`, você verá o **Dashboard**, o **Histórico** e as **Configurações**.

### 1. Carregamento da Escala Base (Primeiro Uso)
Para o sistema funcionar, ele precisa saber qual é a escala atual dos seus funcionários.
- Clique no botão verde **"Carregar Escala Base Inicial"** no canto superior direito.
- Selecione o arquivo `.xlsx` da equipe atual.
- *Nota: Isso precisa ser feito toda vez que o mês virar ou que você tiver uma grande reestruturação global da escala que substitua as planilhas anteriores.*

### 2. Configurações
Vá na aba **Configurações** para calibrar a automação:
- **Palavra-chave do Assunto**: (Ex: *Nova Escala* ou *Grade Atualizada*) - Ele filtrará os e-mails não-lidos apenas com esses termos.
- **Pasta do Outlook**: Normalmente *Caixa de Entrada*, mas você pode criar regras e apontar para subpastas.
- **Termos Imunes (Exceções)**: Termos que **NÃO** podem ser considerados como "Eventos de TV" que caíram. Exemplo: `VIAGEM, FOLGA, OFF, REUNIAO, FERIAS`. Se o funcionário estava de FOLGA na escala antiga, e a palavra "FOLGA" não está na Grade de TV, o sistema respeitará e *não* cancelará essa folga.
- **Mapeamento de E-mails**: Preencha linha a linha associando o nome que vem na escala com o e-mail da empresa.
  Exemplo: `Valesca Maranhão=valesca@globo.com`

### 3. Modo de Teste e Sincronização
Ao clicar em **"Sincronizar Agora"** (ou automaticamente a cada X minutos), o sistema varre o e-mail em busca da grade nova.
- Se o **Modo Teste** estiver ligado nas configurações, o sistema **NÃO ENVIARÁ E-MAILS DE VERDADE**. Ele irá processar todas as diferenças e criará uma pasta local chamada `emails_teste/` contendo arquivos HTML individuais para você visualizar no seu navegador exatamente como o e-mail ficaria.
- Quando terminar a homologação, desative o Modo Teste, salve as configurações e o envio para os funcionários começará.

---

## 📂 Arquitetura do Projeto

* `main.py` e `gui/`: Contêm toda a interface gráfica Desktop (CustomTkinter).
* `core/schedule_processor.py`: Orquestrador que junta o e-mail baixado com a inteligência do banco de dados e dispara as etapas.
* `core/match_eventos.py`: Motor avançado que converte uma Grade de TV densa (com várias colunas lado a lado) num formato que se encaixe na escala dos funcionários.
* `core/diff_engine.py`: A inteligência que compara a versão A com a versão B e levanta exatamente qual célula (horário, descrição, status) foi modificada.
* `core/email_sender.py`: Manipulador COM do Outlook para gerar e enviar e-mails em HTML.
* `core/database.py`: Gerencia as tabelas no SQLite local (`data/scheduler.db`).
