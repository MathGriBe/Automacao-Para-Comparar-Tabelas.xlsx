# Automacao Python de Conferencia de Servicos

Automacao em Python para conciliar servicos entre tres bases operacionais, identificar inconsistencias e gerar um relatorio Excel com apoio para auditoria e tratamento de divergencias.

## Sumario

- [Visao geral](#visao-geral)
- [Como funciona](#como-funciona)
- [Tecnologias](#tecnologias)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Pre-requisitos](#pre-requisitos)
- [Configuracao](#configuracao)
- [Execucao local](#execucao-local)
- [Execucao com Docker](#execucao-com-docker)
- [Saidas geradas](#saidas-geradas)
- [Variaveis de ambiente](#variaveis-de-ambiente)
- [Regras de comparacao](#regras-de-comparacao)
- [Solucao de problemas](#solucao-de-problemas)

## Visao geral

O projeto compara dados de servicos presentes em tres fontes:

- DICI
- planilha de Servicos Ativos
- base de Sites

Durante o processamento, a automacao:

- carrega abas especificas de cada arquivo Excel
- normaliza nomes de servicos para reduzir divergencias de escrita
- elimina linhas vazias, colunas duplicadas e registros fora do bloco ativo
- consolida bases IP e Transporte
- identifica servicos ausentes em uma ou mais fontes
- compara velocidade informada no DICI com a capacidade registrada em Ativos
- gera um relatorio Excel com abas separadas por tipo de inconsistencia
- envia o relatorio por e-mail quando houver divergencias e o envio estiver habilitado

## Como funciona

1. Valida as variaveis de ambiente e os arquivos de entrada.
2. Le as abas configuradas nos arquivos Excel.
3. Limpa e padroniza colunas e nomes dos servicos.
4. Remove duplicidades para a comparacao principal e preserva duplicados em abas dedicadas do relatorio.
5. Monta uma base consolidada com a presenca de cada servico nas tres fontes.
6. Classifica cada servico como `OK`, `Capacidade divergente` ou `Nao consta em: ...`.
7. Exporta o resultado para um arquivo `.xlsx`.
8. Envia o arquivo por e-mail, se configurado.

## Tecnologias

- Python 3.12
- pandas
- openpyxl
- python-dotenv
- SMTP (Office 365 ou servidor compativel)
- Docker e Docker Compose para execucao conteinerizada

## Estrutura do projeto

```text
.
|-- automacao-comparacao.py
|-- teste_comparacao.py
|-- requirements.txt
|-- Dockerfile
|-- docker-compose.yml
|-- tests/
`-- README.md
```

## Pre-requisitos

- Python 3.12 ou superior
- Acesso aos arquivos Excel usados como fonte
- Credenciais SMTP validas para envio de e-mail
- Docker Desktop, se voce quiser executar em container

## Configuracao

Crie um arquivo `.env` na raiz do projeto com os valores adequados ao seu ambiente.

Exemplo:

```env
EMAIL_REMETENTE=seu_email@empresa.com
EMAIL_SENHA=sua_senha_ou_token
EMAIL_DESTINATARIO=destinatario@empresa.com
SMTP_SERVIDOR=smtp.office365.com
SMTP_PORTA=587

ARQ_DICI=C:\caminho\para\Relatorio_Dici_BD.xlsx
ARQ_ATIVOS=C:\caminho\para\Servicos_Ativos.xlsx
ARQ_SITES=C:\caminho\para\Servicos_x_Sites.xlsx

SHEET_DICI_IP=Relatorio_IP
SHEET_DICI_TRANS=Relatorio_Trans
SHEET_ATIVOS_IP=Mar_26_IP
SHEET_ATIVOS_TRANS=Mar_26_Transp
SHEET_SITES=Planilha1

HEADER_ATIVOS=2
HEADER_SITES=2

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
python .\automacao-comparacao.py
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

## Saidas geradas

O script gera:

- um arquivo Excel de relatorio, por padrao no formato `Relatorio_Divergencias_YYYYMMDD.xlsx`
- um arquivo de log chamado `automacao_servicos.log`
- um e-mail com o relatorio em anexo quando `ENVIAR_EMAIL=True` e existirem divergencias

As abas do relatorio incluem:

- `Divergencias`
- `Capacidade_Divergente`
- `Nao_no_Sites`
- `Nao_no_DICI`
- `Nao_no_Ativos`
- `Duplicados_Sites`
- `Duplicados_DICI`
- `Duplicados_Ativos`

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
| `RELATORIO_SAIDA` | Nao | Nome ou caminho do arquivo Excel de saida. |
| `SHEET_DICI_IP` | Nao | Aba IP do arquivo DICI. |
| `SHEET_DICI_TRANS` | Nao | Aba Transporte do arquivo DICI. |
| `SHEET_ATIVOS_IP` | Nao | Aba IP do arquivo de Ativos. |
| `SHEET_ATIVOS_TRANS` | Nao | Aba Transporte do arquivo de Ativos. |
| `SHEET_SITES` | Nao | Aba da base de Sites. |
| `HEADER_ATIVOS` | Nao | Linha de cabecalho usada na leitura de Ativos. |
| `HEADER_SITES` | Nao | Linha de cabecalho usada na leitura de Sites. |
| `ENVIAR_EMAIL` | Nao | Define se o envio automatico sera habilitado. |

## Regras de comparacao

O processamento segue estas regras principais:

- os nomes dos servicos sao normalizados para caixa alta, sem acentos e sem espacos
- a base de Ativos e filtrada ate antes do primeiro bloco com a palavra `DESATIVA`
- a velocidade do DICI e convertida para gigabits por segundo dividindo o valor por `1000`
- a capacidade de Ativos e comparada com a velocidade do DICI com arredondamento de duas casas decimais
- servicos duplicados sao separados em abas especificas e a comparacao principal usa apenas a primeira ocorrencia de cada servico

## Solucao de problemas

Se a automacao falhar, confira primeiro o arquivo `automacao_servicos.log`.

Erros comuns:

- `FileNotFoundError`: o caminho de um dos arquivos Excel esta incorreto ou inacessivel
- `ValueError` ao ler aba: o nome da planilha configurada nao existe no arquivo informado
- `KeyError`: os nomes das colunas esperadas nao estao presentes na planilha
- falha no envio SMTP: credenciais, porta, servidor ou politica de autenticacao podem estar incorretos

## Boas praticas de uso

- nao versione o arquivo `.env`
- teste primeiro com `ENVIAR_EMAIL=False` para validar a geracao do relatorio sem disparar e-mails
- mantenha os nomes das abas e colunas alinhados com o formato real das planilhas de origem
- valide o acesso a unidades de rede, OneDrive e pastas sincronizadas antes de agendar a execucao