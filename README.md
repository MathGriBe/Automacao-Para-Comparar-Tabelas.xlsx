📊 Automação de Conciliação de Serviços

📌 Visão Geral
Script em Python responsável pela comparação entre Serviços Ativos (SA), Serviços x Sites (SS) e Dici (Base do XCONN):

DICI
Serviços Ativos
Sites (SharePoint sincronizado via OneDrive)

O processo realiza:
- Validação de existência dos serviços nas bases
- Comparação de capacidade (DICI x Ativos)

⚠ Identificação de divergências
- Geração de relatório Excel com múltiplas abas 
- Envio automático de e-mail (quando há divergências)
- Registro de execução em log

🛠 Tecnologias Utilizadas
- Python 3.12
- pandas
- openpyxl
- python-dotenv
- SMTP (Office365)


Estrutura do Projeto:
automacao-servicos/
│
├── automacao_comparacao.py
├── requirements.txt
├── README.md
├── .gitignore
└── .env (não versionado)

⚙️ Configuração do Ambiente
1️⃣ Criar ambiente virtual
py -3.12 -m venv .venv
2️⃣ Ativar ambiente (Windows)
.venv\Scripts\activate
3️⃣ Instalar dependências
pip install -r requirements.txt


⚠ O arquivo .env não deve ser versionado.


▶ Execução
python automacao_comparacao.py

📊 Saída Gerada
O relatório Excel contém as seguintes abas:
- Divergencias
- So_no_Sites
- So_no_DICI
- So_no_Ativos

Também são gerados:
Arquivo de log (automacao_servicos.log)
Envio automático de e-mail quando houver divergências