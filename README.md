# 📊 Automação de Comparação de Dados

## 📌 Visão Geral
Este projeto consiste em um script em Python responsável por comparar dados de clientes entre duas bases (ou dois lados da mesma planilha), validando inconsistências e gerando relatórios automáticos.

Projeto utilizado para análise:
**Analise-Banda_CLI-PILOTO**

---

## ⚙️ O que o processo realiza?

- 🔍 Validação de dados:
  - Cliente
  - Circuito
  - Sigla

- ⚠️ Identificação de divergências:
  - Dados diferentes entre as bases
  - Registros não encontrados

- 📊 Geração de relatório em Excel:
  - Base completa
  - Divergências
  - Registros válidos (OK)

- 📝 Registro de logs (execução e erros)

---

## 🛠️ Tecnologias utilizadas

- Python 3.x
- pandas
- openpyxl
- python-dotenv

Dependências completas:

et_xmlfile==2.0.0
numpy==2.4.3
openpyxl==3.1.5
pandas==3.0.1
python-dateutil==2.9.0.post0
python-dotenv==1.2.2
six==1.17.0
tzdata==2025.3


---

## 📁 Estrutura do Projeto


projeto/
│
├── main.py
├── .env
├── requirements.txt
├── README.md
└── resultado_comparacao.xlsx (gerado automaticamente)


---

## 🔐 Variáveis de Ambiente

Crie um arquivo `.env` na raiz do projeto:


ARQ_EXCEL=caminho/do/arquivo.xlsx
SHEET_NAME=nome_da_aba
RELATORIO_SAIDA=resultado_comparacao.xlsx


### ⚠️ Importante
O arquivo `.env` **não deve ser versionado**.

Adicione ao `.gitignore`:


.env
.venv


---

## ▶️ Como executar

1. Ativar o ambiente virtual:


.venv\Scripts\activate


2. Instalar dependências:


pip install -r requirements.txt


3. Executar o script:


python main.py


---

## 📊 Saída gerada

O script gera um arquivo Excel contendo:

- 📄 Base_Comparacao → dados completos
- ⚠️ Divergencias → inconsistências encontradas
- ✅ OK → registros válidos
- 🔍 Outros (dependendo do cenário):
  - Não encontrados
  - Duplicados

---

## 🧠 Regras de comparação

- Comparação baseada na coluna **Circuito**
- Validação dos campos:
  - Cliente
  - Sigla

---

## 🚀 Possíveis melhorias futuras

- Integração com SharePoint
- Execução automatizada (agendador / Docker)
- Interface gráfica
- Validação de múltiplas tabelas
- Alertas por e-mail

---

## 👨‍💻 Autor

Projeto desenvolvido para automação e análise de dados internos.