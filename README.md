# Automacao Python de Conferencia de Servicos

Automacao em Python que concilia os servicos ativos entre as bases operacionais da Eletronet (XCONN/DICI, Servicos Ativos IP e Transporte, e Sites), aponta as divergencias e gera um relatorio Excel, um dashboard HTML interativo e, opcionalmente, um e-mail de resumo. Roda sozinha toda semana pelo Agendador de Tarefas do Windows.

## Sumario

- [Visao geral](#visao-geral)
- [Fluxo da automacao](#fluxo-da-automacao)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Instalacao](#instalacao)
- [Configuracao (.env)](#configuracao-env)
- [Execucao manual](#execucao-manual)
- [Execucao agendada (semanal)](#execucao-agendada-semanal)
- [Saidas geradas](#saidas-geradas)
- [Regras de comparacao](#regras-de-comparacao)
- [Dashboard interativo](#dashboard-interativo)
- [Compartilhando o dashboard na rede](#compartilhando-o-dashboard-na-rede)
- [E-mail de resumo](#e-mail-de-resumo)
- [Manutencao mensal](#manutencao-mensal)
- [Solucao de problemas](#solucao-de-problemas)

## Visao geral

A automacao compara os servicos presentes em quatro fontes:

| Fonte | O que e | Configurada em |
| --- | --- | --- |
| XCONN / DICI | Lista de Servicos Ativos CS (base de referencia) | `ARQ_DICI`, `SHEET_DICI` |
| Servicos Ativos - IP | Aba IP da planilha de Servicos Ativos | `ARQ_ATIVOS`, `SHEET_ATIVOS_IP` |
| Servicos Ativos - Transporte | Aba Transporte da mesma planilha | `ARQ_ATIVOS`, `SHEET_ATIVOS_TRANS` |
| Sites | Base Servicos x Sites | `ARQ_SITES`, `SHEET_SITES` |

Alem disso, le um CSV compartilhado (`ARQ_TAS`) onde a equipe registra o ultimo TAS de cada servico, que aparece no relatorio e no dashboard.

Cada servico recebe um **Status**: `OK` ou a lista das divergencias encontradas, separadas por `; ` (ex.: `Não encontrado em Ativos (Transporte); Não encontrado em Sites`).

## Fluxo da automacao

Tudo acontece no `main.py`, na funcao `main()`:

1. **`preparar_bases()`**: le as abas configuradas, padroniza nomes de colunas e de servicos (caixa alta, sem acento, sem espaco), descarta nomes fora do formato `Trecho/Cliente/Tecnologia[/Porta]`, corta a planilha de Ativos antes do primeiro bloco `DESATIV` e agrupa portas/canais de um mesmo servico.
2. **`comparar_bases()`**: monta uma base unica com a presenca de cada servico em cada fonte, remove servicos fora do escopo, compara capacidades e monta o `Status` de cada linha. Depois cruza com o CSV de ultimo TAS e ordena as divergencias primeiro.
3. **`gerar_relatorio()`**: grava o Excel `Relatorio_Divergencias_YYYYMMDD.xlsx`.
4. **`gerar_dashboard_html()`**: sobrescreve `dashboard_site/index.html` (a pagina servida na rede), sem criar arquivos novos a cada execucao.
5. **`enviar_email_com_relatorio()`**: envia o e-mail de resumo com o Excel anexado, se `ENVIAR_EMAIL=true`.

Tudo e registrado em `automacao_servicos.log`.

## Estrutura do projeto

```text
.
|-- main.py                         # automacao completa (leitura, comparacao, Excel, dashboard, e-mail)
|-- servidor_dashboard.py           # servidor HTTP que publica dashboard_site/ na rede local
|-- iniciar_servidor_dashboard.bat  # sobe o servidor em segundo plano (sem janela)
|-- run_weekly.bat                  # roda o main.py com o Python do .venv (execucao manual)
|-- assets/                         # logos usados no dashboard e no e-mail
|-- tests/
|-- requirements.txt
|-- Dockerfile / docker-compose.yml
|-- .env                            # configuracao local (NAO versionado)
`-- README.md
```

Arquivos gerados pela execucao (ignorados no git): `Relatorio_Divergencias_*.xlsx`, `dashboard_site/`, `TAS_Servicos.csv`, `automacao_servicos.log`, `servidor_dashboard.log`.

## Instalacao

Requisitos: Windows, Python 3.12+ e acesso de leitura as planilhas de origem (pasta de rede e OneDrive/SharePoint sincronizado).

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Depois crie o `.env` (proxima secao).

## Configuracao (.env)

Crie um arquivo `.env` na raiz do projeto. Ele nao vai pro git porque contem a senha do e-mail.

```env
# --- E-mail ---
ENVIAR_EMAIL=false
EMAIL_REMETENTE=seu_email@eletronet.com
EMAIL_SENHA=sua_senha_ou_token
EMAIL_DESTINATARIO=destinatario@eletronet.com
SMTP_SERVIDOR=smtp.office365.com
SMTP_PORTA=587

# --- Arquivos de origem ---
ARQ_DICI=C:\caminho\Lista_Servicos_Ativos_CS.xlsx
ARQ_ATIVOS=\\servidor\pasta\Servicos Ativos - Eletronet - 2026.xlsx
ARQ_SITES=C:\caminho\Servicos x sites.xlsx
ARQ_TAS=\\servidor\pasta\TAS_Servicos.csv

# --- Abas e linha de cabecalho ---
SHEET_DICI=vw_Servicos_Ativos_CS
SHEET_ATIVOS_IP=Nov_26_IP
SHEET_ATIVOS_TRANS=Nov_26_Transp
SHEET_SITES=Planilha1
HEADER_ATIVOS=2
HEADER_SITES=2

# --- Dashboard na rede ---
DASHBOARD_URL=http://ENU-6576:8000/
SITE_PORTA=8000
```

| Variavel | Obrigatoria | Descricao |
| --- | --- | --- |
| `ENVIAR_EMAIL` | Nao | `true` envia o e-mail; qualquer outro valor (ou ausente) nao envia. Padrao: `false`. |
| `EMAIL_REMETENTE` | Se enviar e-mail | Conta usada no envio. |
| `EMAIL_SENHA` | Se enviar e-mail | Senha ou token SMTP da conta remetente. |
| `EMAIL_DESTINATARIO` | Se enviar e-mail | Destinatario(s) do relatorio. |
| `SMTP_SERVIDOR` | Nao | Padrao: `smtp.office365.com`. |
| `SMTP_PORTA` | Nao | Padrao: `587`. |
| `ARQ_DICI` | Sim | Caminho do arquivo XCONN/DICI. |
| `ARQ_ATIVOS` | Sim | Caminho da planilha de Servicos Ativos (IP e Transporte). |
| `ARQ_SITES` | Sim | Caminho da base de Sites. |
| `ARQ_TAS` | Nao | CSV compartilhado com o ultimo TAS por servico. Padrao: `TAS_Servicos.csv` local; se nao existir, e criado vazio. Aponte para uma pasta de rede para a equipe editar o mesmo arquivo. |
| `SHEET_DICI` | Nao | Aba do XCONN/DICI. Padrao: `vw_Servicos_Ativos_CS`. |
| `SHEET_ATIVOS_IP` | Nao | Aba IP de Ativos (muda todo mes, ver [Manutencao mensal](#manutencao-mensal)). |
| `SHEET_ATIVOS_TRANS` | Nao | Aba Transporte de Ativos (muda todo mes). |
| `SHEET_SITES` | Nao | Aba da base de Sites. Padrao: `Planilha1`. |
| `HEADER_ATIVOS` | Nao | Linha do cabecalho em Ativos (0 = primeira linha). Padrao: `2`. |
| `HEADER_SITES` | Nao | Linha do cabecalho em Sites. Padrao: `2`. |
| `DASHBOARD_URL` | Nao | Link exibido como botao "Abrir dashboard interativo" no e-mail. Vazio = sem botao. |
| `SITE_PORTA` | Nao | Porta do `servidor_dashboard.py`. Padrao: `8000`. |

Configuracoes que ficam direto no `main.py` (mudam pouco):

| Constante | Para que serve |
| --- | --- |
| `FILTROS_SUBSTRING_IGNORAR` | Servicos cujo nome contem algum desses textos sao ignorados (ex.: `ESCH`, `CFO`, `155M`). |
| `FILTRO_TAMANHO_MINIMO` | Servicos com nome menor que isso sao ignorados. |
| `ABRIR_DASHBOARD_AUTOMATICO` | `True` abre o dashboard no navegador ao fim da execucao. Deixe `False` na execucao agendada. |
| `STATUS_NAO_ENCONTRADO_SITES` | Texto da divergencia de Sites, usado no Status e no filtro do dashboard. |

## Execucao manual

Com o `.venv` ativado:

```powershell
python .\main.py
```

Ou de duplo clique no `run_weekly.bat`, que usa o Python do `.venv` e mostra o progresso no console.

> Dica: para testar sem disparar e-mail, deixe `ENVIAR_EMAIL=false` no `.env`.

## Execucao agendada (semanal)

A execucao automatica e feita pela tarefa **"Comparação semanal"** do Agendador de Tarefas do Windows:

| Aba | Configuracao |
| --- | --- |
| Geral | "Executar somente quando o usuario estiver conectado" (precisa da sessao para acessar OneDrive e rede). Configurar para: Windows 10. |
| Disparadores | **Semanal**, segunda-feira, 08:00. **Nao** usar "Ao fazer logon", senao roda toda vez que o PC liga. |
| Acoes | Programa: `<projeto>\.venv\Scripts\pythonw.exe` (com aspas). Argumentos: `"<projeto>\main.py"`. Iniciar em: `<projeto>` (sem aspas). |
| Condicoes | Desmarcar "Iniciar somente se estiver na energia CA" e "Parar se alternar para bateria". |
| Configuracoes | Marcar "Executar assim que possivel apos uma inicializacao agendada perdida" (se o PC estiver desligado na segunda as 8h, roda uma vez quando ligar). Interromper apos 2 horas. |

O `pythonw.exe` roda sem janela de console; o progresso fica no `automacao_servicos.log`.

Para conferir e testar pelo cmd:

```cmd
chcp 65001
schtasks /query /tn "Comparação semanal" /v /fo list
schtasks /run /tn "Comparação semanal"
powershell Get-Content automacao_servicos.log -Tail 20 -Wait
```

Na consulta, "Tipo de Agendamento" deve ser semanal e "Hora da proxima execucao" a proxima segunda. Se aparecer "Somente sob demanda", a tarefa esta sem disparador. "Ultimo resultado" `0` = sucesso; `267009` = em execucao; `267011` = ainda nao rodou.

## Saidas geradas

**Excel** `Relatorio_Divergencias_YYYYMMDD.xlsx`:

| Aba | Conteudo |
| --- | --- |
| `Divergencias` | Servicos com divergencia: servico, cliente, tipo, motivo e ultimo TAS. |
| `Divergencia_Detalhada` | Os mesmos servicos com todas as colunas de apoio (capacidades, portas, presenca em cada base). |
| `Clientes` | Ranking de clientes por quantidade de servicos divergentes. |
| `Resumo` | Totais da execucao. |

Principais colunas: `Servico`, `Cliente`, `Tipo` (IP/TRANSPORTE/OUTROS), `Status`, `Gb_XCONN`, `Gb_Ativos_IP_Total`, `Gb_Ativos_Transporte_Total`, `Qtd_Portas_IP`, `Detalhe_Portas_IP`, `Qtd_Portas_Trans`, `Detalhe_Portas_Trans`, `Validado_Via_Soma_Portas`, `Ultimo_TAS`.

**Dashboard** `dashboard_site/index.html` (sobrescrito a cada execucao).

**Log** `automacao_servicos.log`.

**E-mail** de resumo, quando `ENVIAR_EMAIL=true`.

## Regras de comparacao

**Escopo**

- nomes de servico sao normalizados (caixa alta, sem acentos, sem espacos) antes de comparar
- so entram designacoes no formato `Trecho/Cliente/Tecnologia[/Porta]`; nomes fora disso (ex.: `ANEL1`, `ROTA3`) sao ignorados e listados no log
- a planilha de Ativos e lida so ate o primeiro bloco com `DESATIV` (desativacao mensal / acumulado)
- servicos com algum texto de `FILTROS_SUBSTRING_IGNORAR` ou nome menor que `FILTRO_TAMANHO_MINIMO` sao descartados
- servico duplicado no XCONN: vale o primeiro registro; porta repetida exata em Ativos e removida

**Tipo** (funcao `classificar_tecnologia`)

- `TRANSPORTE`: MetroEthernet, Espectro
- `IP`: IP Flat, IP no PIX, IP no POP, IP Conteudo e variacoes Burstable/Brasil
- qualquer outra tecnologia vira `OUTROS` (e gera divergencia)

Servico que nao esta no XCONN nao tem `NmTecnologia`, entao o tipo e deduzido nesta ordem:

1. pela aba de Ativos onde ele aparece: so na aba IP vira `IP`, so na aba Transporte vira `TRANSPORTE`
2. se nao aparece em nenhuma aba (ou aparece nas duas), pela sigla da designacao (funcao `classificar_por_sigla`): `GETH` vira `TRANSPORTE`; `IPxx` (IPFL, IPOB, IPOP, IPTB, IPTT) vira `IP`
3. sigla fora desses padroes (ex.: `POP`) fica com Tipo vazio

**Divergencias possiveis no Status** (podem aparecer varias juntas, separadas por `; `)

| Divergencia | Quando acontece |
| --- | --- |
| `Não encontrado no XCONN` | Servico existe em outra base mas nao no XCONN. |
| `Tecnologia não classificada (OUTROS)` | Tecnologia do XCONN fora das listas IP/Transporte. |
| `Não encontrado em Ativos (IP)` | Servico IP sem registro na aba IP de Ativos. |
| `Não encontrado em Ativos (Transporte)` | Servico de Transporte sem registro na aba Transporte. |
| `Encontrado indevidamente em Ativos (IP/Transporte)` | Servico aparece na aba de Ativos do outro tipo. |
| `Presente em Ativos mas sem cadastro no XCONN` | Esta em Ativos, mas nao no XCONN. |
| `Capacidade divergente entre XCONN e Ativos` | Velocidade do XCONN (convertida para Gb, /1000) diferente da capacidade em Ativos, com 2 casas decimais. Servicos divididos em varias portas sao validados pela soma (`Validado_Via_Soma_Portas`). |
| `Não encontrado em Sites` | Servico sem registro na base de Sites. |

## Dashboard interativo

HTML autocontido (funciona offline, sem CDN) com:

- cards de resumo (total analisado, OK, precisam sincronizar)
- botoes de setor: Todos os setores / So IP / So Transporte
- filtro de status: **So divergentes** (padrao) / Todos / So OK / **So nao encontrados em Sites**
- filtro de tipo, filtro de cliente e busca por servico ou cliente
- contador "X de Y servico(s) exibido(s)", que reflete todos os filtros combinados
- ranking de clientes com mais divergencias (clique para filtrar a tabela)
- colunas ordenaveis (1o clique crescente, 2o decrescente, 3o volta ao padrao)
- coluna "Ultimo TAS" vinda do `ARQ_TAS`

O filtro "So nao encontrados em Sites" verifica se o Status **contem** esse texto, entao tambem pega servicos com outras divergencias juntas. Para obter a mesma lista em Python (exportar para Excel, por exemplo):

```python
from main import preparar_bases, comparar_bases, filtrar_nao_encontrados_em_sites

resultado = comparar_bases(*preparar_bases())
filtrar_nao_encontrados_em_sites(resultado).to_excel("Nao_encontrados_Sites.xlsx", index=False)
```

## Compartilhando o dashboard na rede

O `servidor_dashboard.py` publica **apenas** a pasta `dashboard_site/` (nunca o projeto inteiro, para nao expor o `.env` nem as planilhas). Cada execucao do `main.py` atualiza essa pasta, entao o link e sempre o mesmo e mostra o resultado mais recente.

Para o pessoal da rede acessar, tres coisas precisam estar ok:

1. **Servidor rodando.** Suba com `iniciar_servidor_dashboard.bat` (sem janela). Para subir sozinho a cada logon, crie uma tarefa no Agendador com disparador "Ao fazer logon" executando `<projeto>\.venv\Scripts\pythonw.exe "<projeto>\servidor_dashboard.py"`. Essa tarefa pode rodar a cada logon, porque so publica a pagina e nao gera nem envia relatorio.
2. **Porta liberada no Firewall** (PowerShell como administrador; se for bloqueado, pedir a TI):
   ```powershell
   New-NetFirewallRule -DisplayName "Dashboard Divergencias (8000)" -Direction Inbound -Protocol TCP -LocalPort 8000 -Action Allow -Profile Domain,Private
   ```
3. **`DASHBOARD_URL` no `.env`**, para o botao aparecer no e-mail. Prefira o nome do PC (`http://ENU-6576:8000/`) ao IP, que pode mudar no Wi-Fi/DHCP.

Os links de acesso ficam no `servidor_dashboard.log` (linha `Na rede local:`). Para parar o servidor: Gerenciador de Tarefas > `pythonw.exe` > Finalizar tarefa.

Limitacao: o link so funciona enquanto este PC estiver ligado e o servidor rodando.

## E-mail de resumo

Controlado por `ENVIAR_EMAIL` no `.env` (`true` envia, `false` nao envia). O corpo e um HTML proprio (o Outlook nao executa JavaScript, entao nao e o dashboard): cards com total/OK/divergentes, divergencias por setor, os 10 clientes mais desincronizados, os principais motivos e o botao "Abrir dashboard interativo" (se `DASHBOARD_URL` estiver preenchido). O Excel completo vai anexado e o assunto informa quantos servicos precisam ser sincronizados.

## Manutencao mensal

As abas de Servicos Ativos mudam de nome todo mes (ex.: `Out_26_IP` -> `Nov_26_IP`). Quando a nova aba for criada, atualize no `.env`:

```env
SHEET_ATIVOS_IP=Nov_26_IP
SHEET_ATIVOS_TRANS=Nov_26_Transp
```

Se esquecer, a automacao compara com a aba do mes anterior (ou falha com `ValueError` se ela nao existir mais).

## Solucao de problemas

Comece sempre pelo `automacao_servicos.log` (ou `servidor_dashboard.log` para o servidor).

| Sintoma | Causa provavel |
| --- | --- |
| Relatorio gerado toda vez que liga o PC | Existe tarefa com disparador "Ao fazer logon" rodando o `main.py`. Remova o disparador ou a tarefa. |
| Tarefa semanal nao roda | Tarefa sem disparador ("Somente sob demanda") ou caminho sem aspas. Veja [Execucao agendada](#execucao-agendada-semanal). |
| Janela de cmd aberta durante a execucao | Tarefa chamando `run_weekly.bat`/`python.exe` em vez de `pythonw.exe`. |
| E-mail enviado mesmo com envio desligado | Confira se `ENVIAR_EMAIL` esta no `.env` e se o `main.py` le do `.env` (nao deve haver valor fixo no codigo). |
| `FileNotFoundError` | Caminho de planilha errado ou rede/OneDrive indisponivel. |
| `ValueError` ao ler aba | Nome de aba no `.env` nao existe (veja [Manutencao mensal](#manutencao-mensal)). |
| `KeyError` | Colunas da planilha mudaram de nome ou `HEADER_*` aponta para a linha errada. |
| Falha no SMTP | Credenciais, porta ou politica de autenticacao da conta. |
| Dashboard nao abre em outra maquina | Servidor parado ou porta bloqueada no Firewall. Teste `http://localhost:8000/` no proprio PC para isolar. |
