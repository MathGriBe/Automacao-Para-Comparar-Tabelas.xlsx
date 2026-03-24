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

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "matheus.bevilaqua@eletronet.com")
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO", "matheusgrisostomo0@outlook.com")

SMTP_SERVIDOR = os.getenv("SMTP_SERVIDOR", "smtp.office365.com")
SMTP_PORTA = int(os.getenv("SMTP_PORTA", "587"))

ARQ_ATIVOS = os.getenv(
    "ARQ_ATIVOS",
    r"\\gjas-fileserver\Global\Customer Services\Serviços Ativos\Serviços Ativos - Eletronet - 2026.xlsx",
)
ARQ_SITES = os.getenv(
    "ARQ_SITES",
    r"\\gjas-fileserver\Global\CORe\Automacoes\Guilherme Luis Dias De Oliveira - Controle de Implantação\Serviços x sites_base Junho_24 1.xlsx",
)

RELATORIO_SAIDA = os.getenv(
    "RELATORIO_SAIDA",
    f"Relatorio_Divergencias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
)

LOG_ARQUIVO = "automacao_servicos.log"

SHEET_ATIVOS_IP = os.getenv("SHEET_ATIVOS_IP", "Mar_26_IP")
SHEET_ATIVOS_TRANS = os.getenv("SHEET_ATIVOS_TRANS", "Mar_26_Transp")
SHEET_SITES = os.getenv("SHEET_SITES", "Planilha1")

HEADER_ATIVOS = int(os.getenv("HEADER_ATIVOS", "2"))
HEADER_SITES = int(os.getenv("HEADER_SITES", "2"))

ENVIAR_EMAIL = os.getenv("ENVIAR_EMAIL", "False").strip().lower() == "true"

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

def limpar_nomes_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, df.columns != ""]
    df = df.loc[:, df.columns != "nan"]
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def normalizar_texto(valor: object) -> str:
    if pd.isna(valor):
        return ""

    valor = str(valor).strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ASCII", "ignore").decode("ASCII")
    valor = valor.replace(" ", "")
    return valor


def limpar_texto(col: pd.Series) -> pd.Series:
    return col.apply(normalizar_texto)


def filtrar_bloco_ativo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Mantém apenas o bloco superior da aba, cortando a partir da
    primeira linha que contenha 'DESATIVA'.
    """
    df = df.copy().reset_index(drop=True)
    df = df.dropna(how="all").reset_index(drop=True)

    mascara_desativacao = df.astype(str).apply(
        lambda row: row.str.contains("DESATIVA", case=False, na=False).any(),
        axis=1,
    )

    indices = df.index[mascara_desativacao].tolist()

    if indices:
        corte = indices[0]
        df = df.iloc[:corte].copy()
        logger.info("Bloco de desativação encontrado. Corte aplicado na linha interna %s.", corte)
    else:
        logger.info("Nenhum bloco de desativação encontrado na aba.")

    if "Serviços" in df.columns:
        df = df[df["Serviços"].notna()].copy()
        df = df[df["Serviços"].astype(str).str.strip() != ""].copy()
        df = df[df["Serviços"].astype(str).str.strip().str.upper() != "SERVIÇOS"].copy()
        df = df[~df["Serviços"].astype(str).str.contains("DESATIVA", case=False, na=False)].copy()

    if "Clientes" in df.columns:
        df = df[df["Clientes"].notna()].copy()
        df = df[df["Clientes"].astype(str).str.strip() != ""].copy()
        df = df[df["Clientes"].astype(str).str.strip().str.upper() != "CLIENTES"].copy()

    return df.reset_index(drop=True)


def validar_variaveis_ambiente() -> None:
    if not EMAIL_REMETENTE:
        raise ValueError("EMAIL_REMETENTE não configurado.")
    if ENVIAR_EMAIL and not EMAIL_SENHA:
        raise ValueError("EMAIL_SENHA não configurada para envio de e-mail.")
    if not EMAIL_DESTINATARIO:
        raise ValueError("EMAIL_DESTINATARIO não configurado.")


def validar_arquivos() -> None:
    arquivos = {
        "ARQ_ATIVOS": ARQ_ATIVOS,
        "ARQ_SITES": ARQ_SITES,
    }

    for nome, caminho in arquivos.items():
        if not os.path.exists(caminho):
            raise FileNotFoundError(f"{nome} não encontrado: {caminho}")


def listar_abas_excel(caminho: str) -> None:
    try:
        xls = pd.ExcelFile(caminho)
        logger.info("Abas disponíveis em '%s': %s", caminho, xls.sheet_names)
    except Exception as exc:
        raise RuntimeError(f"Erro ao listar abas do arquivo '{caminho}': {exc}") from exc


def ler_arquivo_excel(caminho: str, sheet_name: str, header: int | None = None) -> pd.DataFrame:
    try:
        if header is None:
            return pd.read_excel(caminho, sheet_name=sheet_name)
        return pd.read_excel(caminho, sheet_name=sheet_name, header=header)
    except ValueError as exc:
        try:
            abas_disponiveis = pd.ExcelFile(caminho).sheet_names
        except Exception:
            abas_disponiveis = []

        raise ValueError(
            f"Aba '{sheet_name}' não encontrada no arquivo '{caminho}'. "
            f"Abas disponíveis: {abas_disponiveis}"
        ) from exc
    except Exception as exc:
        raise RuntimeError(f"Erro ao ler arquivo '{caminho}' / aba '{sheet_name}': {exc}") from exc


def carregar_dados() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    logger.info("Iniciando leitura dos arquivos...")

    listar_abas_excel(ARQ_ATIVOS)
    listar_abas_excel(ARQ_SITES)

    ativos_ip = ler_arquivo_excel(ARQ_ATIVOS, SHEET_ATIVOS_IP, HEADER_ATIVOS)
    ativos_trans = ler_arquivo_excel(ARQ_ATIVOS, SHEET_ATIVOS_TRANS, HEADER_ATIVOS)
    sites = ler_arquivo_excel(ARQ_SITES, SHEET_SITES, HEADER_SITES)

    logger.info("Arquivos carregados com sucesso.")
    return ativos_ip, ativos_trans, sites


def preparar_bases(
    ativos_ip: pd.DataFrame,
    ativos_trans: pd.DataFrame,
    sites: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    ativos_ip = limpar_nomes_colunas(ativos_ip)
    ativos_trans = limpar_nomes_colunas(ativos_trans)
    sites = limpar_nomes_colunas(sites)

    logger.info("Colunas Ativos IP: %s", list(ativos_ip.columns))
    logger.info("Colunas Ativos Trans: %s", list(ativos_trans.columns))
    logger.info("Colunas Sites: %s", list(sites.columns))

    ativos_ip = filtrar_bloco_ativo(ativos_ip)
    ativos_trans = filtrar_bloco_ativo(ativos_trans)

    logger.info("Qtd Ativos IP após filtro: %s", len(ativos_ip))
    logger.info("Qtd Ativos Trans após filtro: %s", len(ativos_trans))

    ativos = pd.concat([ativos_ip, ativos_trans], ignore_index=True)

    ativos = limpar_nomes_colunas(ativos)
    sites = limpar_nomes_colunas(sites)

    logger.info("Qtd Ativos consolidado: %s", len(ativos))
    logger.info("Qtd Sites consolidado: %s", len(sites))

    ativos = ativos[["Serviços", "Capacidade (GB)"]].copy()
    ativos = ativos.rename(columns={"Capacidade (GB)": "Capacidade"})
    sites = sites[["Nome Serviço"]].copy()

    ativos["Servico_Ativos"] = ativos["Serviços"]
    sites["Servico_Sites"] = sites["Nome Serviço"]

    ativos["Servico"] = limpar_texto(ativos["Serviços"])
    sites["Servico"] = limpar_texto(sites["Nome Serviço"])

    ativos = ativos[ativos["Servico"] != ""].copy()
    sites = sites[sites["Servico"] != ""].copy()

    ativos["Capacidade"] = pd.to_numeric(ativos["Capacidade"], errors="coerce")

    dup_sites = sites[sites["Servico"].duplicated(keep=False)].copy().sort_values("Servico")
    dup_ativos = ativos[ativos["Servico"].duplicated(keep=False)].copy().sort_values("Servico")

    logger.info("Duplicados em Sites: %s", len(dup_sites))
    logger.info("Duplicados em Ativos: %s", len(dup_ativos))

    if not dup_sites.empty:
        logger.info("Top duplicados Sites:\n%s", dup_sites.head(20).to_string())

    if not dup_ativos.empty:
        logger.info("Top duplicados Ativos:\n%s", dup_ativos.head(20).to_string())

    sites_unicos = sites.drop_duplicates(subset=["Servico"], keep="first").copy()
    ativos_unicos = ativos.drop_duplicates(subset=["Servico"], keep="first").copy()

    return (
        sites_unicos,
        ativos_unicos,
        dup_sites,
        dup_ativos,
    )


def comparar_bases(
    sites_unicos: pd.DataFrame,
    ativos_unicos: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    todos_servicos = pd.DataFrame(
        {
            "Servico": sorted(
                set(sites_unicos["Servico"]) | set(ativos_unicos["Servico"])
            )
        }
    )

    base = (
        todos_servicos
        .merge(
            ativos_unicos[["Servico", "Servico_Ativos", "Capacidade"]],
            on="Servico",
            how="left",
        )
        .merge(
            sites_unicos[["Servico", "Servico_Sites"]],
            on="Servico",
            how="left",
        )
    )

    base["Existe_Ativos"] = base["Servico_Ativos"].notna()
    base["Existe_Sites"] = base["Servico_Sites"].notna()

    def validar(row: pd.Series) -> str:
        if row["Existe_Ativos"] and row["Existe_Sites"]:
            return "OK"
        if row["Existe_Ativos"] and not row["Existe_Sites"]:
            return "Não consta em: Sites"
        if row["Existe_Sites"] and not row["Existe_Ativos"]:
            return "Não consta em: Ativos"
        return "Inconsistência"

    base["Status"] = base.apply(validar, axis=1)

    divergencias = base[base["Status"] != "OK"].copy()
    capacidade_divergente = base.iloc[0:0].copy()
    nao_no_sites = base[~base["Existe_Sites"]].copy()
    nao_no_ativos = base[~base["Existe_Ativos"]].copy()

    logger.info("Qtd base final de comparação: %s", len(base))
    logger.info("Qtd divergências: %s", len(divergencias))
    logger.info("Qtd não encontrados em Sites: %s", len(nao_no_sites))
    logger.info("Qtd não encontrados em Ativos: %s", len(nao_no_ativos))

    return divergencias, capacidade_divergente, nao_no_sites, nao_no_ativos


def gerar_relatorio(
    divergencias: pd.DataFrame,
    capacidade_divergente: pd.DataFrame,
    nao_no_sites: pd.DataFrame,
    nao_no_ativos: pd.DataFrame,
    dup_sites: pd.DataFrame,
    dup_ativos: pd.DataFrame,
) -> None:
    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        divergencias.to_excel(writer, sheet_name="Divergencias", index=False)
        capacidade_divergente.to_excel(writer, sheet_name="Capacidade_Divergente", index=False)
        nao_no_sites.to_excel(writer, sheet_name="Nao_no_Sites", index=False)
        nao_no_ativos.to_excel(writer, sheet_name="Nao_no_Ativos", index=False)
        dup_sites.to_excel(writer, sheet_name="Duplicados_Sites", index=False)
        dup_ativos.to_excel(writer, sheet_name="Duplicados_Ativos", index=False)

    logger.info("Relatório gerado com sucesso: %s", RELATORIO_SAIDA)


def enviar_email(arquivo: str) -> None:
    msg = EmailMessage()
    msg["Subject"] = "Relatório Automático de Divergências - Serviços"
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = EMAIL_DESTINATARIO

    msg.set_content(
        """Olá,

Segue relatório automático de divergências.

Att,
Automação Python
"""
    )

    with open(arquivo, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=os.path.basename(arquivo),
        )

    with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_REMETENTE, EMAIL_SENHA)
        smtp.send_message(msg)

    logger.info("E-mail enviado com sucesso.")


def main() -> None:
    logger.info("===== INÍCIO DA AUTOMAÇÃO =====")

    try:
        validar_variaveis_ambiente()
        validar_arquivos()

        logger.info("Arquivo ATIVOS em uso: %s", ARQ_ATIVOS)
        logger.info("Arquivo SITES em uso: %s", ARQ_SITES)
        logger.info(
            "Última modificação ATIVOS: %s",
            datetime.fromtimestamp(os.path.getmtime(ARQ_ATIVOS)),
        )
        logger.info(
            "Última modificação SITES: %s",
            datetime.fromtimestamp(os.path.getmtime(ARQ_SITES)),
        )

        ativos_ip, ativos_trans, sites = carregar_dados()

        (
            sites_unicos,
            ativos_unicos,
            dup_sites,
            dup_ativos,
        ) = preparar_bases(ativos_ip, ativos_trans, sites)

        (
            divergencias,
            capacidade_divergente,
            nao_no_sites,
            nao_no_ativos,
        ) = comparar_bases(sites_unicos, ativos_unicos)

        gerar_relatorio(
            divergencias,
            capacidade_divergente,
            nao_no_sites,
            nao_no_ativos,
            dup_sites,
            dup_ativos,
        )

        if ENVIAR_EMAIL and not divergencias.empty:
            enviar_email(RELATORIO_SAIDA)
        elif divergencias.empty:
            logger.info("Nenhuma divergência encontrada. E-mail não enviado.")
        else:
            logger.info("Envio de e-mail desabilitado para teste.")

    except FileNotFoundError as exc:
        logger.error("Arquivo não encontrado: %s", exc)
    except KeyError as exc:
        logger.error("Coluna não encontrada: %s", exc)
        logger.error("Confira os nomes das colunas exibidos no log.")
    except ValueError as exc:
        logger.error("Erro de validação: %s", exc)
    except RuntimeError as exc:
        logger.error("Erro de leitura/processamento: %s", exc)
    except Exception as exc:
        logger.exception("Erro inesperado durante a automação: %s", exc)
    finally:
        logger.info("===== FIM DA AUTOMAÇÃO =====")


if __name__ == "__main__":
    main()