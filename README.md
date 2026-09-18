# Automacao Python de Conferencia de Servicos

Automacao em Python para conciliar servicos entre tres bases operacionais, identificar inconsistencias e gerar um relatorio Excel e um dashboard HTML interativo com apoio para auditoria e tratamento de divergencias.

## Sumario

- [Visao geral](#visao-geral)
- [Como funciona](#como-funciona)
- [Tecnologias](#tecnologias)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Pre-requisitos](#pre-requisitos)
- [Configuracao](#configuracao)
- [Execucao local](#execucao-local)
- [Execucao com Docker](#execucao-com-docker)
- [Dashboard interativo](#dashboard-interativo)
- [Servidor local do dashboard](#servidor-local-do-dashboard)
- [Saidas geradas](#saidas-geradas)
- [Variaveis de ambiente](#variaveis-de-ambiente)
- [Regras de comparacao](#regras-de-comparacao)
- [Solucao de problemas](#solucao-de-problemas)

## Visao geral

O projeto compara dados de servicos presentes em tres fontes:

- Lista Servicos Ativos CS (DICI/XCONN)
- planilha de Servicos Ativos (IP e Transporte)
- base de Sites

Durante o processamento, a automacao:

- carrega abas especificas de cada arquivo Excel
- normaliza nomes de servicos para reduzir divergencias de escrita
- elimina linhas vazias, colunas duplicadas, servicos fora do escopo e blocos de servicos ja desativados
- consolida bases IP e Transporte
- identifica servicos ausentes em uma ou mais fontes
- compara velocidade informada no DICI com a capacidade registrada em Ativos
- cruza com um arquivo compartilhado de "ultimo TAS" por servico, preenchido pela equipe
- gera um relatorio Excel com abas separadas por tipo de inconsistencia
- gera um dashboard HTML interativo (busca, filtros, ordenacao) com o resultado
- envia um e-mail com o relatorio em anexo e o dashboard embutido no corpo, quando o envio estiver habilitado

## Como funciona

1. Valida as variaveis de ambiente e os arquivos de entrada.
2. Le as abas configuradas nos arquivos Excel.
3. Limpa e padroniza colunas e nomes dos servicos.
4. Corta a base de Ativos antes do primeiro bloco de servicos desativados (mensal/acumulado) e remove servicos fora do escopo configurado.
5. Remove duplicidades para a comparacao principal e preserva duplicados em abas dedicadas do relatorio.
6. Monta uma base consolidada com a presenca de cada servico nas tres fontes.
7. Classifica cada servico como `OK` ou lista as divergencias encontradas (nao encontrado, capacidade divergente, tecnologia nao classificada, etc.).
8. Cruza o resultado com o arquivo compartilhado de "ultimo TAS" por servico.
9. Exporta o resultado para um arquivo `.xlsx` e para um dashboard `.html`.
10. Envia o relatorio por e-mail, se configurado (anexo em `.xlsx` + dashboard embutido no corpo em HTML).

## Tecnologias

- Python 3.12
- pandas
- openpyxl
- python-dotenv
- SMTP (Office 365 ou servidor compativel)
- servidor HTTP embutido (`http.server`, biblioteca padrao do Python) para publicar o dashboard na rede local
- Docker e Docker Compose para execucao conteinerizada

## Estrutura do projeto

```text
.
|-- main.py
|-- servidor_dashboard.py
|-- iniciar_servidor_dashboard.bat
|-- assets/
|   |--> logo_eletronet_azul.jpg
|   `--> logo_eletronet_branco.jpg
|-- tests/
|   |--> teste_comparacao.py
|-- requirements.txt
|-- Dockerfile
|-- docker-compose.yml
`-- README.md
```

Arquivos gerados pela execucao (ignorados no git, veja `.gitignore`): `Relatorio_Divergencias_*.xlsx`, `Dashboard_Divergencias_*.html`, `dashboard_site/`, `TAS_Servicos.csv`, `automacao_servicos.log`, `servidor_dashboard.log`.

## Pre-requisitos

- Python 3.12 ou superior
- Acesso aos arquivos Excel usados como fonte
- Credenciais SMTP validas para envio de e-mail
- Docker Desktop, se voce quiser executar em container
- (Opcional) Porta liberada no Firewall do Windows, se quiser que outras maquinas da rede acessem o dashboard pelo `servidor_dashboard.py`

## Configuracao

Crie um arquivo `.env` na raiz do projeto com os valores adequados ao seu ambiente.

Exemplo:

```env
EMAIL_REMETENTE=seu_email@empresa.com
EMAIL_SENHA=sua_senha_ou_token
EMAIL_DESTINATARIO=destinatario@empresa.com
SMTP_SERVIDOR=smtp.office365.com
SMTP_PORTA=587

ARQ_DICI=C:\caminho\para\Lista_Servicos_Ativos_CS.xlsx
ARQ_ATIVOS=C:\caminho\para\Servicos_Ativos.xlsx
ARQ_SITES=C:\caminho\para\Servicos_x_Sites.xlsx

(Arquivo compartilhado onde a equipe registra o ultimo TAS de cada servico;
aponte para uma pasta de rede para que todos leiam/editem o mesmo arquivo)
ARQ_TAS=C:\caminho\para\TAS_Servicos.csv

(Para acessar as abas das planilhas excel)
SHEET_DICI = vw_Servicos_Ativos_CS
SHEET_ATIVOS_IP= "mês"_"ano"_IP
SHEET_ATIVOS_TRANS="mês"_"ano"_Transp
SHEET_SITES=Planilha1

(Linha para acessar as informações)
HEADER_ATIVOS=2
HEADER_SITES=2

(Porta usada pelo servidor_dashboard.py; opcional, padrao 8000)
SITE_PORTA=8000

RELATORIO_SAIDA=Relatorio_Divergencias.xlsx
ENVIAR_EMAIL=True
```

## Execucao local

Crie e ative um ambiente virtual:

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
```

Instale as dependencias:

```powershell
pip install -r requirements.txt
```

Execute a automacao:

```powershell
python .\main.py
```

## Execucao com Docker

Construa a imagem:

```powershell
docker compose build
```

Execute o container:

```powershell
docker compose up
```

Observacao: para funcionar em container, os caminhos definidos em `ARQ_DICI`, `ARQ_ATIVOS` e `ARQ_SITES` precisam ser acessiveis dentro do ambiente Docker. Se esses arquivos estiverem apenas em pastas locais do Windows ou unidades mapeadas, talvez seja necessario adaptar volumes e caminhos antes do uso.

## Dashboard interativo

Alem do relatorio Excel, o `main.py` gera um dashboard HTML autocontido (sem depender de internet/CDN) com:

- cartoes de resumo (total analisado, OK, precisam sincronizar)
- busca por nome de servico
- filtro por status (divergente/todos/OK) e por tipo
- botoes de atalho para cada setor (Todos / So IP / So Transporte)
- barra de busca e filtros fixa no topo ao rolar a pagina
- colunas ordenaveis por clique no cabecalho (1º clique = crescente, 2º = decrescente, 3º = volta ao padrao), com seta `▲`/`▼` indicando o estado
- coluna "Ultimo TAS", preenchida a partir do arquivo compartilhado configurado em `ARQ_TAS`
- logo da Eletronet no cabecalho

O arquivo gerado (`Dashboard_Divergencias_YYYYMMDD.html`) abre automaticamente no navegador padrao ao final da execucao (controlado por `ABRIR_DASHBOARD_AUTOMATICO` em `main.py`). O mesmo conteudo tambem e embutido diretamente no corpo do e-mail quando o envio esta habilitado — atencao que alguns clientes de e-mail (Outlook, por exemplo) podem nao executar o JavaScript, entao a busca/ordenacao interativa so funciona abrindo o arquivo `.html` num navegador.

## Servidor local do dashboard

Para deixar o dashboard acessivel por um link fixo enquanto o computador estiver ligado, use `servidor_dashboard.py`. Ele serve **apenas** a pasta `dashboard_site/` (nunca o projeto inteiro, para nao expor o `.env` nem as planilhas de origem); o `main.py` atualiza o conteudo dessa pasta a cada execucao.

Para subir o servidor sem nenhuma janela visivel (roda em segundo plano com `pythonw.exe`):

```powershell
.\iniciar_servidor_dashboard.bat
```

Os links de acesso (local, rede e nome da maquina) ficam registrados em `servidor_dashboard.log` toda vez que o servidor sobe — procure a linha `Na rede local:`. Como o servidor roda sem console, para encerrar use o Gerenciador de Tarefas e finalize o processo `python.exe`/`pythonw.exe`.

Para iniciar automaticamente a cada login do Windows, crie um atalho do `iniciar_servidor_dashboard.bat` na pasta de Inicializacao (`Win+R` → `shell:startup`).

Pontos de atencao:

- por padrao o link so funciona no proprio computador (`http://localhost:8000/`); para outras maquinas da rede acessarem, e preciso liberar a porta (`SITE_PORTA`, padrao `8000`) no Firewall do Windows — combine com o time de TI antes de abrir a porta
- o IP mostrado como "Na rede local" pode mudar se a rede usa DHCP; o link "Pelo nome do PC" tende a ser mais estavel, se a rede da empresa resolver nomes de maquina

## Saidas geradas

O script gera:

- um arquivo Excel de relatorio, por padrao no formato `Relatorio_Divergencias_YYYYMMDD.xlsx`, com as abas `Divergencias`, `Divergencia_Detalhada` e `Resumo`
- um dashboard HTML, por padrao no formato `Dashboard_Divergencias_YYYYMMDD.html`, e uma copia de nome fixo em `dashboard_site/index.html`
- um arquivo de log chamado `automacao_servicos.log`
- um e-mail com o relatorio em anexo e o dashboard embutido no corpo, quando `ENVIAR_EMAIL=True`

As principais colunas do resultado incluem:

- `Servico`, `Tipo` (IP/TRANSPORTE/OUTROS), `Status`
- `Gb_XCONN`, `Gb_Ativos_IP_Total`, `Gb_Ativos_Transporte_Total`
- `Qtd_Portas_IP`, `Detalhe_Portas_IP`, `Qtd_Portas_Trans`, `Detalhe_Portas_Trans`
- `Validado_Via_Soma_Portas`
- `Ultimo_TAS`

## Variaveis de ambiente

| Variavel | Obrigatoria | Descricao |
| --- | --- | --- |
| `EMAIL_REMETENTE` | Sim | Conta usada no envio do e-mail. |
| `EMAIL_SENHA` | Sim, se `ENVIAR_EMAIL=True` | Senha ou token SMTP da conta remetente. |
| `EMAIL_DESTINATARIO` | Sim | Destinatario do relatorio. |
| `SMTP_SERVIDOR` | Nao | Servidor SMTP. Padrao: `smtp.office365.com`. |
| `SMTP_PORTA` | Nao | Porta SMTP. Padrao: `587`. |
| `ARQ_DICI` | Sim | Caminho completo do arquivo DICI. |
| `ARQ_ATIVOS` | Sim | Caminho completo do arquivo de Servicos Ativos. |
| `ARQ_SITES` | Sim | Caminho completo do arquivo de Sites. |
| `ARQ_TAS` | Nao | Caminho do CSV compartilhado com o ultimo TAS por servico. Padrao: `TAS_Servicos.csv` (local). Se o arquivo nao existir, e criado vazio automaticamente. |
| `SITE_PORTA` | Nao | Porta usada pelo `servidor_dashboard.py`. Padrao: `8000`. |
| `RELATORIO_SAIDA` | Nao | Nome ou caminho do arquivo Excel de saida. |
| `SHEET_DICI` | Nao | Aba do arquivo Lista_Servicos_Ativos. |
| `SHEET_ATIVOS_IP` | Nao | Aba IP do arquivo de Ativos. |
| `SHEET_ATIVOS_TRANS` | Nao | Aba Transporte do arquivo de Ativos. |
| `SHEET_SITES` | Nao | Aba da base de Sites. |
| `HEADER_ATIVOS` | Nao | Linha de cabecalho usada na leitura de Ativos. |
| `HEADER_SITES` | Nao | Linha de cabecalho usada na leitura de Sites. |
| `ENVIAR_EMAIL` | Nao | Define se o envio automatico sera habilitado. |

## Regras de comparacao

O processamento segue estas regras principais:

- os nomes dos servicos sao normalizados para caixa alta, sem acentos e sem espacos
- a base de Ativos e filtrada ate antes do primeiro bloco com a palavra `DESATIV` (cobre os blocos de "desativacao mensal" e "acumulado"), para nao contar servico ja desativado como ativo
- servicos cujo nome contem alguma das substrings definidas em `FILTROS_SUBSTRING_IGNORAR` (em `main.py`), ou com nome menor que `FILTRO_TAMANHO_MINIMO` caracteres, sao descartados da analise
- a tecnologia informada no DICI e classificada em `IP`, `TRANSPORTE` ou `OUTROS` pela funcao `classificar_tecnologia` em `main.py`; tecnologias fora dessa lista caem em `OUTROS` e geram divergencia
- a velocidade do DICI e convertida para gigabits por segundo dividindo o valor por `1000`
- a capacidade de Ativos e comparada com a velocidade do DICI com arredondamento de duas casas decimais; quando um servico contratado esta dividido em varias portas/canais, a soma do grupo tambem e validada
- servicos duplicados sao separados em abas especificas e a comparacao principal usa apenas a primeira ocorrencia de cada servico

## Solucao de problemas

Se a automacao falhar, confira primeiro o arquivo `automacao_servicos.log`. Para problemas no servidor local do dashboard, confira `servidor_dashboard.log`.

Erros comuns:

- `FileNotFoundError`: o caminho de um dos arquivos Excel esta incorreto ou inacessivel
- `ValueError` ao ler aba: o nome da planilha configurada nao existe no arquivo informado
- `KeyError`: os nomes das colunas esperadas nao estao presentes na planilha
- falha no envio SMTP: credenciais, porta, servidor ou politica de autenticacao podem estar incorretos
- dashboard nao acessivel por outra maquina: normalmente e o Firewall do Windows bloqueando a porta configurada em `SITE_PORTA`; teste primeiro `http://localhost:8000/` no proprio computador para isolar o problema

## Boas praticas de uso

- nao versione o arquivo `.env`
- teste primeiro com `ENVIAR_EMAIL=False` para validar a geracao do relatorio sem disparar e-mails
- mantenha os nomes das abas e colunas alinhados com o formato real das planilhas de origem
- valide o acesso a unidades de rede, OneDrive e pastas sincronizadas antes de agendar a execucao
