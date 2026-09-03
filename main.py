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

ARQ_DICI = os.path.normpath(os.getenv("ARQ_DICI", ""))
ARQ_ATIVOS = os.path.normpath(os.getenv("ARQ_ATIVOS", ""))
ARQ_SITES = os.path.normpath(os.getenv("ARQ_SITES", ""))

RELATORIO_SAIDA = f"Relatorio_Divergencias_{datetime.now().strftime('%Y%m%d')}.xlsx"

LOG_ARQUIVO = "automacao_servicos.log"

SHEET_DICI = os.getenv("SHEET_DICI", "vw_Servicos_Ativos_CS")
SHEET_ATIVOS_IP = os.getenv("SHEET_ATIVOS_IP", "Abr_26_IP")
SHEET_ATIVOS_TRANS = os.getenv("SHEET_ATIVOS_TRANS", "Abr_26_Transp")
SHEET_SITES = os.getenv("SHEET_SITES", "Planilha1")

HEADER_ATIVOS = int(os.getenv("HEADER_ATIVOS", "2"))
HEADER_SITES = int(os.getenv("HEADER_SITES", "2"))

ENVIAR_EMAIL = False

# Serviços que devem ser desconsiderados da análise
FILTRO_SUBSTRING_IGNORAR = "ESCH"
FILTRO_TAMANHO_MINIMO = 15

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
    # DICI
    # =============================

    dici = dici[["NmServico", "NmTecnologia", "ValorMB"]].copy()

    dici["Servico"] = limpar_texto(dici["NmServico"])
    dici["Tipo"] = dici["NmTecnologia"].apply(classificar_tecnologia)
    dici["Velocidade_Gb"] = pd.to_numeric(dici["ValorMB"], errors="coerce") / 1000

    # Se houver serviço duplicado no DICI, mantém o primeiro registro
    dici = dici.drop_duplicates(subset="Servico", keep="first")

    # =============================
    # ATIVOS
    # =============================

    ativos_ip = ativos_ip[["Serviços", "Capacidade (GB)"]].copy()
    ativos_trans = ativos_trans[["Serviços", "Capacidade (GB)"]].copy()

    ativos_ip["Servico"] = limpar_texto(ativos_ip["Serviços"])
    ativos_trans["Servico"] = limpar_texto(ativos_trans["Serviços"])

    ativos_ip["Capacidade_IP"] = pd.to_numeric(ativos_ip["Capacidade (GB)"], errors="coerce")
    ativos_trans["Capacidade_Trans"] = pd.to_numeric(ativos_trans["Capacidade (GB)"], errors="coerce")

    ativos_ip = ativos_ip.drop_duplicates(subset="Servico", keep="first")
    ativos_trans = ativos_trans.drop_duplicates(subset="Servico", keep="first")

    # =============================
    # SITES
    # =============================

    sites = sites[["Nome Serviço"]].copy()
    sites["Servico"] = limpar_texto(sites["Nome Serviço"])
    sites = sites.drop_duplicates(subset="Servico", keep="first")

    return dici, ativos_ip, ativos_trans, sites


def comparar_bases(dici, ativos_ip, ativos_trans, sites):
    """
    Compara TODAS as bases entre si (não apenas o DICI como referência).
    Faz um outer join por 'Servico' e avalia, para cada serviço encontrado
    em QUALQUER uma das bases, se ele existe/é consistente nas demais.
    """

    # --- Base DICI reduzida para o merge ---
    dici_m = dici[["Servico", "NmServico", "Tipo", "Velocidade_Gb"]].rename(
        columns={"NmServico": "Nome_XCONN"}
    )

    ativos_ip_m = ativos_ip[["Servico", "Capacidade_IP"]]
    ativos_trans_m = ativos_trans[["Servico", "Capacidade_Trans"]]
    sites_m = sites[["Servico"]].assign(Existe_Sites=True)

    # --- Outer join de todas as bases pelo código normalizado do serviço ---
    base = dici_m.merge(ativos_ip_m, on="Servico", how="outer")
    base = base.merge(ativos_trans_m, on="Servico", how="outer")
    base = base.merge(sites_m, on="Servico", how="outer")

    # --- Flags de existência em cada base ---
    base["Existe_XCONN"] = base["Nome_XCONN"].notna()
    base["Existe_Ativos_IP"] = base["Capacidade_IP"].notna()
    base["Existe_Ativos_Transporte"] = base["Capacidade_Trans"].notna()
    base["Existe_Sites"] = base["Existe_Sites"].fillna(False)

    # --- Nome de exibição: usa o nome do XCONN quando existir, senão o código normalizado ---
    base["Servico_Exibicao"] = base["Nome_XCONN"].fillna(base["Servico"])

    # --- Filtro: desconsidera serviços com "ESCH" no nome ou com nome muito curto ---
    nome_upper = base["Servico_Exibicao"].astype(str).str.upper()
    mask_ignorar_esch = nome_upper.str.contains(FILTRO_SUBSTRING_IGNORAR, na=False)
    mask_ignorar_curto = base["Servico_Exibicao"].astype(str).str.len() < FILTRO_TAMANHO_MINIMO
    qtd_ignorados = int((mask_ignorar_esch | mask_ignorar_curto).sum())
    if qtd_ignorados:
        logger.info(
            f"Ignorando {qtd_ignorados} serviço(s) por conter '{FILTRO_SUBSTRING_IGNORAR}' "
            f"ou ter menos de {FILTRO_TAMANHO_MINIMO} caracteres."
        )
    base = base[~(mask_ignorar_esch | mask_ignorar_curto)].copy()

    # --- Capacidade de Ativos correspondente ao tipo declarado no XCONN ---
    def capacidade_ativos_correspondente(row):
        if row["Tipo"] == "IP":
            return row["Capacidade_IP"]
        elif row["Tipo"] == "TRANSPORTE":
            return row["Capacidade_Trans"]
        return pd.NA

    base["Gb_Ativos_Correspondente"] = base.apply(capacidade_ativos_correspondente, axis=1)

    # --- Monta a lista de divergências linha a linha ---
    def montar_status(row):
        divergencias = []

        if not row["Existe_XCONN"]:
            divergencias.append("Não encontrado no XCONN")

        if row["Existe_XCONN"] and row["Tipo"] == "OUTROS":
            divergencias.append("Tecnologia não classificada (OUTROS)")

        if row["Existe_XCONN"] and row["Tipo"] == "IP" and not row["Existe_Ativos_IP"]:
            divergencias.append("Não encontrado em Ativos (IP)")

        if row["Existe_XCONN"] and row["Tipo"] == "TRANSPORTE" and not row["Existe_Ativos_Transporte"]:
            divergencias.append("Não encontrado em Ativos (Transporte)")

        # Serviço aparecendo na planilha de Ativos "errada" para o tipo dele
        if row["Tipo"] == "IP" and row["Existe_Ativos_Transporte"]:
            divergencias.append("Encontrado indevidamente em Ativos (Transporte)")
        if row["Tipo"] == "TRANSPORTE" and row["Existe_Ativos_IP"]:
            divergencias.append("Encontrado indevidamente em Ativos (IP)")

        # Serviço presente em Ativos mas ausente do XCONN (não dá para saber IP/Transporte esperado)
        if not row["Existe_XCONN"] and (row["Existe_Ativos_IP"] or row["Existe_Ativos_Transporte"]):
            divergencias.append("Presente em Ativos mas sem cadastro no XCONN")

        # Divergência de capacidade entre XCONN e Ativos
        if pd.notna(row["Velocidade_Gb"]) and pd.notna(row["Gb_Ativos_Correspondente"]):
            if round(row["Velocidade_Gb"], 2) != round(row["Gb_Ativos_Correspondente"], 2):
                divergencias.append("Capacidade divergente entre XCONN e Ativos")

        if not row["Existe_Sites"]:
            divergencias.append("Não encontrado em Sites")

        if not divergencias:
            return "OK"
        return "; ".join(divergencias)

    base["Status"] = base.apply(montar_status, axis=1)

    resultado = base[[
        "Servico_Exibicao",
        "Tipo",
        "Existe_XCONN",
        "Existe_Ativos_IP",
        "Existe_Ativos_Transporte",
        "Existe_Sites",
        "Velocidade_Gb",
        "Capacidade_IP",
        "Capacidade_Trans",
        "Status",
    ]].rename(columns={
        "Servico_Exibicao": "Servico",
        "Velocidade_Gb": "Gb_XCONN",
        "Capacidade_IP": "Gb_Ativos_IP",
        "Capacidade_Trans": "Gb_Ativos_Transporte",
    })

    # Ordena colocando as divergências primeiro
    resultado = resultado.sort_values(
        by="Status", key=lambda s: (s == "OK")
    ).reset_index(drop=True)

    return resultado


def gerar_relatorio(df):
    total = len(df)
    divergentes_mask = df["Status"] != "OK"
    divergentes = int(divergentes_mask.sum())

    resumo = pd.DataFrame({
        "Métrica": ["Total de serviços analisados", "Serviços OK", "Serviços com divergência"],
        "Valor": [total, total - divergentes, divergentes],
    })

    # A planilha de Divergencias exibe apenas os serviços que não estão OK
    df_divergentes = df[divergentes_mask].reset_index(drop=True)

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        df_divergentes.to_excel(writer, sheet_name="Divergencias", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

    logger.info(f"Relatório gerado: {RELATORIO_SAIDA} ({divergentes}/{total} divergências)")


def enviar_email_com_relatorio(arquivo_anexo):
    if not ENVIAR_EMAIL:
        logger.info("Envio de e-mail desabilitado nas configurações.")
        return

    logger.info("Preparando envio de e-mail...")

    msg = EmailMessage()
    msg["Subject"] = f"Relatório de Divergências - {datetime.now().strftime('%d/%m/%Y')}"
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = EMAIL_DESTINATARIO
    msg.set_content("Olá,\n\nSegue em anexo o relatório de divergências gerado pela automação.")

    try:
        with open(arquivo_anexo, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=arquivo_anexo
            )

        with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as server:
            server.starttls()
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.send_message(msg)

        logger.info("E-mail enviado com sucesso!")
    except Exception as e:
        logger.error(f"Erro ao enviar e-mail: {e}")

# =============================
# MAIN
# =============================

def main():
    logger.info("Iniciando...")

    dici, ativos_ip, ativos_trans, sites = preparar_bases()
    resultado = comparar_bases(dici, ativos_ip, ativos_trans, sites)
    gerar_relatorio(resultado)

    enviar_email_com_relatorio(RELATORIO_SAIDA)

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()