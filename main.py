import logging
import os
import smtplib
import unicodedata
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
from dotenv import load_dotenv

# =============================
# CONFIGURAÇÕES
# =============================

load_dotenv()

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "")
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO", "")

SMTP_SERVIDOR = os.getenv("SMTP_SERVIDOR", "smtp.office365.com")
SMTP_PORTA = int(os.getenv("SMTP_PORTA", "587"))

ARQ_DICI = os.getenv("ARQ_DICI", r"Z:\CORe\Automacoes\Lista_Servicos_Ativos_CS.xlsx")
ARQ_ATIVOS = os.getenv("ARQ_ATIVOS", r"Z:\Customer Services\Serviços Ativos\Serviços Ativos - Eletronet - 2026.xlsx")
ARQ_SITES = os.getenv("ARQ_SITES", r"C:\Users\mbevilaqua\Eletronet SA\Guilherme Luis Dias De Oliveira - Controle de Implantação\Serviços x sites_base Junho_24 1.xlsx")

RELATORIO_SAIDA = f"Relatorio_Divergencias_{datetime.now().strftime('%Y%m%d')}.xlsx"

LOG_ARQUIVO = "automacao_servicos.log"

SHEET_DICI = 'vw_Servicos_Ativos_CS'
SHEET_ATIVOS_IP = 'Abr_26_IP'
SHEET_ATIVOS_TRANS = 'Abr_26_Transp'
SHEET_SITES = 'Planilha1'

HEADER_ATIVOS = 2
HEADER_SITES = 2

ENVIAR_EMAIL = False

# =============================
# LOG
# =============================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_ARQUIVO, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

logger = logging.getLogger(__name__)

# =============================
# FUNÇÕES AUXILIARES
# =============================

def limpar_nomes_colunas(df):
    df.columns = df.columns.astype(str).str.strip()
    return df


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ASCII", "ignore").decode("ASCII")
    valor = valor.replace(" ", "")
    return valor


def limpar_texto(col):
    return col.apply(normalizar_texto)


def classificar_tecnologia(tecnologia):
    transporte = ["MetroEthernet", "Espectro"]
    ip = [
        "IP Flat", "IP Flat Burstable",
        "IP no PIX", "IP no PIX Burstable",
        "IP no POP", "IP no POP Brasil", "IP no POP Burstable"
    ]

    if tecnologia in transporte:
        return "TRANSPORTE"
    elif tecnologia in ip:
        return "IP"
    return "OUTROS"


def ler_arquivo_excel(caminho, sheet_name, header=None):
    if header is None:
        return pd.read_excel(caminho, sheet_name=sheet_name)
    return pd.read_excel(caminho, sheet_name=sheet_name, header=header)


# =============================
# PROCESSAMENTO
# =============================

def preparar_bases():
    logger.info("Lendo arquivos...")

    dici = ler_arquivo_excel(ARQ_DICI, SHEET_DICI)
    ativos_ip = ler_arquivo_excel(ARQ_ATIVOS, SHEET_ATIVOS_IP, HEADER_ATIVOS)
    ativos_trans = ler_arquivo_excel(ARQ_ATIVOS, SHEET_ATIVOS_TRANS, HEADER_ATIVOS)
    sites = ler_arquivo_excel(ARQ_SITES, SHEET_SITES, HEADER_SITES)

    # Limpeza
    dici = limpar_nomes_colunas(dici)
    ativos_ip = limpar_nomes_colunas(ativos_ip)
    ativos_trans = limpar_nomes_colunas(ativos_trans)
    sites = limpar_nomes_colunas(sites)

    # =============================
    # DICI (BASE PRINCIPAL)
    # =============================

    dici = dici[["NmServico", "NmTecnologia", "ValorMB"]].copy()

    dici["Servico"] = limpar_texto(dici["NmServico"])
    dici["Tipo"] = dici["NmTecnologia"].apply(classificar_tecnologia)
    dici["Velocidade_Gb"] = pd.to_numeric(dici["ValorMB"], errors="coerce") / 1000

    # =============================
    # ATIVOS
    # =============================

    ativos_ip = ativos_ip[["Serviços", "Capacidade (GB)"]].copy()
    ativos_trans = ativos_trans[["Serviços", "Capacidade (GB)"]].copy()

    ativos_ip["Servico"] = limpar_texto(ativos_ip["Serviços"])
    ativos_trans["Servico"] = limpar_texto(ativos_trans["Serviços"])

    ativos_ip["Capacidade"] = pd.to_numeric(ativos_ip["Capacidade (GB)"], errors="coerce")
    ativos_trans["Capacidade"] = pd.to_numeric(ativos_trans["Capacidade (GB)"], errors="coerce")

    # =============================
    # SITES
    # =============================

    sites = sites[["Nome Serviço"]].copy()
    sites["Servico"] = limpar_texto(sites["Nome Serviço"])

    return dici, ativos_ip, ativos_trans, sites


def comparar(dici, ativos_ip, ativos_trans, sites):
    resultados = []

    for _, row in dici.iterrows():
        servico = row["Servico"]
        tipo = row["Tipo"]
        gb = row["Velocidade_Gb"]

        if tipo == "TRANSPORTE":
            base = ativos_trans
        elif tipo == "IP":
            base = ativos_ip
        else:
            continue

        match = base[base["Servico"] == servico]

        if match.empty:
            status = "Não encontrado no Ativos"
            capacidade = None
        else:
            capacidade = match.iloc[0]["Capacidade"]

            if round(gb, 2) != round(capacidade, 2):
                status = "Capacidade divergente"
            else:
                status = "OK"

        existe_site = not sites[sites["Servico"] == servico].empty

        resultados.append({
            "Servico": row["NmServico"],
            "Tipo": tipo,
            "Gb_DICI": gb,
            "Gb_Ativos": capacidade,
            "Existe_Sites": existe_site,
            "Status": status
        })

    return pd.DataFrame(resultados)


def gerar_relatorio(df):
    df.to_excel(RELATORIO_SAIDA, index=False)
    logger.info(f"Relatório gerado: {RELATORIO_SAIDA}")


# =============================
# MAIN
# =============================

def main():
    logger.info("Iniciando...")

    dici, ativos_ip, ativos_trans, sites = preparar_bases()

    resultado = comparar(dici, ativos_ip, ativos_trans, sites)

    gerar_relatorio(resultado)

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()