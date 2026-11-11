import base64
import html as html_lib
import json
import logging
import os
import re
import smtplib
import unicodedata
import webbrowser
from datetime import datetime
from email.message import EmailMessage
from email.utils import make_msgid

import pandas as pd
from dotenv import load_dotenv
from openpyxl.styles import Font, PatternFill

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

# Arquivo compartilhado (idealmente numa pasta de rede) onde a equipe
# registra o último TAS de cada serviço. Aponte ARQ_TAS no .env para o
# caminho de rede; se não configurado, usa um arquivo local.
ARQ_TAS = os.getenv("ARQ_TAS", "TAS_Servicos.csv")

RELATORIO_SAIDA = f"Relatorio_Divergencias_{datetime.now().strftime('%Y%m%d')}.xlsx"
DASHBOARD_SAIDA = f"Dashboard_Divergencias_{datetime.now().strftime('%Y%m%d')}.html"
ABRIR_DASHBOARD_AUTOMATICO = False

# Pasta servida pelo servidor_dashboard.py. Contém SÓ o dashboard (nunca o
# projeto inteiro, pra não expor o .env e as planilhas na rede). O nome do
# arquivo é fixo para o link de acesso não mudar a cada execução.
SITE_DIR = "dashboard_site"
SITE_ARQUIVO = os.path.join(SITE_DIR, "index.html")

LOGO_ARQUIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_eletronet_azul.jpg")
# No e-mail o logo fica numa faixa branca (o JPG não tem transparência),
# então usa a versão de fundo branco.
LOGO_EMAIL_ARQUIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo_eletronet_branco.jpg")

# Link do dashboard (servidor_dashboard.py) exibido como botão no e-mail.
# Ex.: http://NOME-DO-PC:8000/ . Se vazio, o botão não aparece.
DASHBOARD_URL = os.getenv("DASHBOARD_URL", "")

# Quantos itens listar no corpo do e-mail (a lista completa vai no anexo)
EMAIL_QTD_CLIENTES = 10
EMAIL_QTD_MOTIVOS = 6

LOG_ARQUIVO = "automacao_servicos.log"

SHEET_DICI = os.getenv("SHEET_DICI", "vw_Servicos_Ativos_CS")
SHEET_ATIVOS_IP = os.getenv("SHEET_ATIVOS_IP", "Set_26_IP")
SHEET_ATIVOS_TRANS = os.getenv("SHEET_ATIVOS_TRANS", "Set_26_Transp")
SHEET_SITES = os.getenv("SHEET_SITES", "Planilha1")

HEADER_ATIVOS = int(os.getenv("HEADER_ATIVOS", "2"))
HEADER_SITES = int(os.getenv("HEADER_SITES", "2"))

ENVIAR_EMAIL = os.getenv("ENVIAR_EMAIL", "false").strip().lower() == "true"

# Texto de divergência usado no Status e no filtro "Só não encontrados em Sites"
STATUS_NAO_ENCONTRADO_SITES = "Não encontrado em Sites"

# Serviços que devem ser desconsiderados da análise (substring contida no nome)
FILTROS_SUBSTRING_IGNORAR = ["ESCH", "CFO", "155M", "RPEE", "622M"]
FILTRO_TAMANHO_MINIMO = 15

# Formato de uma designação de serviço válida:
# Trecho/Cliente/Tecnologia[/Porta] (a porta é opcional, pois o XCONN às vezes
# só tem o nome base). Linhas das bases fora desse formato — nomes de anel,
# rotas, títulos de bloco (ex.: ANEL7LITORALSUPERIOR, SERRADAMESA-CAMACARI) —
# não são serviços e são descartadas antes da comparação. Designações com
# complemento entre parênteses (ex.: ASS-GJA/A01/GETH/002 (VC4-2)) só marcam
# uma rota na planilha e também são descartadas.
PADRAO_DESIGNACAO_SERVICO = re.compile(r"^[^/()]+/[^/()]+/[^/()]+(/[^/()]+)?$")

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


def manter_so_designacoes_validas(df, coluna, nome_base):
    """Descarta as linhas cujo nome não segue PADRAO_DESIGNACAO_SERVICO."""
    mask_valido = df[coluna].str.match(PADRAO_DESIGNACAO_SERVICO)
    invalidos = df.loc[~mask_valido & (df[coluna] != ""), coluna].unique()
    if len(invalidos):
        logger.info(
            f"{nome_base}: ignorando {len(invalidos)} nome(s) fora do formato "
            f"Trecho/Cliente/Tecnologia/Porta: {', '.join(invalidos)}"
        )
    return df[mask_valido].copy()


def extrair_servico_base(servico):
    """
    Remove o sufixo de porta/canal do final do código do serviço
    (ex.: 'MUZMBO/ASR/IPFL/001' -> 'MUZMBO/ASR/IPFL'), permitindo
    agrupar todas as portas de um mesmo serviço contratado antes
    de comparar a capacidade total com o XCONN.
    """
    return re.sub(r"/\d+$", "", servico)


def classificar_tecnologia(tecnologia):
    transporte = ["MetroEthernet", "Espectro"]
    ip = [
        "IP Flat", "IP Flat Burstable",
        "IP no PIX", "IP no PIX Burstable",
        "IP no POP", "IP no POP Brasil", "IP no POP Burstable",
        "IP Conteudo",
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


def cortar_antes_do_marcador(df, marcador):
    """
    As abas de Ativos trazem, abaixo da tabela de serviços ativos, blocos de
    'DESATIVAÇÃO (mensal)' e '(acumulado 2025)' com as mesmas colunas
    (Serviços, Capacidade (GB)...). Sem esse corte, esses serviços já
    desativados entrariam na comparação como se ainda estivessem ativos.
    Corta a tabela na primeira linha em que QUALQUER coluna contém o
    marcador, mantendo só as linhas de serviços realmente ativos.
    """
    mask = df.astype(str).apply(lambda col: col.str.contains(marcador, case=False, na=False))
    linhas_marcadas = df.index[mask.any(axis=1)]
    if linhas_marcadas.empty:
        return df
    primeira_linha = linhas_marcadas[0]
    logger.info(
        f"Cortando tabela na linha {primeira_linha} (marcador '{marcador}' encontrado) "
        f"para não incluir serviços desativados."
    )
    return df.loc[:primeira_linha].iloc[:-1].copy()


def carregar_logo_base64(caminho):
    """Lê o logo e devolve como data URI, pra embutir no HTML sem depender de arquivo externo."""
    try:
        with open(caminho, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        return f"data:image/jpeg;base64,{b64}"
    except Exception as e:
        logger.warning(f"Não foi possível carregar o logo ({caminho}): {e}")
        return ""


def carregar_tas(caminho):
    """
    Lê o arquivo compartilhado (CSV) onde a equipe registra o último TAS de
    cada serviço (colunas "Servico" e "Ultimo_TAS"). Se ainda não existir,
    cria um vazio com o cabeçalho esperado para a equipe começar a preencher.
    """
    colunas = ["Servico", "Ultimo_TAS"]

    if not os.path.exists(caminho):
        try:
            pd.DataFrame(columns=colunas).to_csv(caminho, index=False, encoding="utf-8-sig")
            logger.info(f"Arquivo de TAS não encontrado, criado vazio em: {caminho}")
        except Exception as e:
            logger.warning(f"Não foi possível criar o arquivo de TAS ({caminho}): {e}")
        return pd.DataFrame(columns=["Servico_Chave", "Ultimo_TAS"])

    try:
        tas = pd.read_csv(caminho, dtype=str, encoding="utf-8-sig")
    except Exception as e:
        logger.warning(f"Não foi possível ler o arquivo de TAS ({caminho}): {e}")
        return pd.DataFrame(columns=["Servico_Chave", "Ultimo_TAS"])

    if "Servico" not in tas.columns or "Ultimo_TAS" not in tas.columns:
        logger.warning(f"Arquivo de TAS ({caminho}) sem as colunas esperadas {colunas}.")
        return pd.DataFrame(columns=["Servico_Chave", "Ultimo_TAS"])

    tas = tas[colunas].dropna(subset=["Servico"])
    tas["Servico_Chave"] = tas["Servico"].apply(normalizar_texto)
    tas = tas.drop_duplicates(subset="Servico_Chave", keep="last")
    return tas[["Servico_Chave", "Ultimo_TAS"]]


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

    if "nmempresa" not in dici.columns:
        dici["nmempresa"] = pd.NA
    dici = dici[["NmServico", "nmempresa", "NmTecnologia", "ValorMB"]].copy()

    dici["Servico"] = limpar_texto(dici["NmServico"])
    dici = manter_so_designacoes_validas(dici, "Servico", "XCONN")
    dici["Tipo"] = dici["NmTecnologia"].apply(classificar_tecnologia)
    dici["Velocidade_Gb"] = pd.to_numeric(dici["ValorMB"], errors="coerce") / 1000

    # Se houver serviço duplicado no DICI, mantém o primeiro registro
    dici = dici.drop_duplicates(subset="Servico", keep="first")

    # =============================
    # ATIVOS
    # =============================

    # Remove os blocos de "DESATIVAÇÃO (mensal)" / "(acumulado 2025)" que
    # aparecem abaixo da tabela principal nas duas abas — senão serviços já
    # desativados entram na comparação como se estivessem ativos.
    ativos_ip = cortar_antes_do_marcador(ativos_ip, "DESATIV")
    ativos_trans = cortar_antes_do_marcador(ativos_trans, "DESATIV")

    for df_ativos in (ativos_ip, ativos_trans):
        if "Clientes" not in df_ativos.columns:
            df_ativos["Clientes"] = pd.NA
    ativos_ip = ativos_ip[["Serviços", "Clientes", "Capacidade (GB)"]].copy()
    ativos_trans = ativos_trans[["Serviços", "Clientes", "Capacidade (GB)"]].copy()

    # Conjunto de serviços cadastrados no XCONN, usado para decidir a chave de agrupamento
    servicos_dici = set(dici["Servico"])

    def montar_chave_servico(servico_porta):
        """
        Decide a chave usada para casar com o XCONN:
        - Se o nome EXATO (com o número final) já existe como serviço próprio
          no XCONN, usa ele direto (comportamento normal, sem agrupar nada).
        - Senão, tenta o nome BASE (sem o número final) — cobre o caso de um
          serviço contratado dividido em múltiplas portas/canais, onde o
          XCONN só tem o nome base cadastrado (ex.: MUZMBO/ASR/IPFL/001 e
          /002 agrupados sob MUZMBO/ASR/IPFL).
        - Se nenhum dos dois existir no XCONN, mantém o nome exato (vai
          aparecer como "não encontrado no XCONN", igual antes).
        Isso evita agrupar por engano serviços cujo código só termina em
        número por outro motivo (ex.: um ID de circuito).
        """
        base = extrair_servico_base(servico_porta)
        if servico_porta in servicos_dici:
            return servico_porta
        if base in servicos_dici:
            return base
        return servico_porta

    # Nome completo (com porta/canal) e chave de agrupamento
    ativos_ip["Servico_Porta"] = limpar_texto(ativos_ip["Serviços"])
    ativos_trans["Servico_Porta"] = limpar_texto(ativos_trans["Serviços"])
    ativos_ip = manter_so_designacoes_validas(ativos_ip, "Servico_Porta", "Ativos IP")
    ativos_trans = manter_so_designacoes_validas(ativos_trans, "Servico_Porta", "Ativos Transporte")

    ativos_ip["Servico"] = ativos_ip["Servico_Porta"].apply(montar_chave_servico)
    ativos_trans["Servico"] = ativos_trans["Servico_Porta"].apply(montar_chave_servico)

    ativos_ip["Capacidade_IP"] = pd.to_numeric(ativos_ip["Capacidade (GB)"], errors="coerce")
    ativos_trans["Capacidade_Trans"] = pd.to_numeric(ativos_trans["Capacidade (GB)"], errors="coerce")

    # Remove porta duplicada exata (mesma porta aparecendo 2x na planilha)
    ativos_ip = ativos_ip.drop_duplicates(subset="Servico_Porta", keep="first")
    ativos_trans = ativos_trans.drop_duplicates(subset="Servico_Porta", keep="first")

    # Detalhe por porta, para auditoria (ex.: "MUZMBO/ASR/IPFL/001=60 | MUZMBO/ASR/IPFL/002=40")
    ativos_ip["Detalhe_Porta_IP"] = (
        ativos_ip["Servico_Porta"] + "=" + ativos_ip["Capacidade_IP"].round(2).astype(str)
    )
    ativos_trans["Detalhe_Porta_Trans"] = (
        ativos_trans["Servico_Porta"] + "=" + ativos_trans["Capacidade_Trans"].round(2).astype(str)
    )

    # Agrupa pela chave definida acima, somando a capacidade quando há mais de uma porta
    ativos_ip = ativos_ip.groupby("Servico", as_index=False).agg(
        Capacidade_IP=("Capacidade_IP", "sum"),
        Qtd_Portas_IP=("Servico_Porta", "count"),
        Detalhe_Portas_IP=("Detalhe_Porta_IP", lambda s: " | ".join(s)),
        Cliente_IP=("Clientes", "first"),
    )
    ativos_trans = ativos_trans.groupby("Servico", as_index=False).agg(
        Capacidade_Trans=("Capacidade_Trans", "sum"),
        Qtd_Portas_Trans=("Servico_Porta", "count"),
        Detalhe_Portas_Trans=("Detalhe_Porta_Trans", lambda s: " | ".join(s)),
        Cliente_Trans=("Clientes", "first"),
    )

    # =============================
    # SITES
    # =============================

    sites = sites[["Nome Serviço"]].copy()
    sites["Servico"] = limpar_texto(sites["Nome Serviço"])
    sites = manter_so_designacoes_validas(sites, "Servico", "Sites")
    sites = sites.drop_duplicates(subset="Servico", keep="first")

    return dici, ativos_ip, ativos_trans, sites


def comparar_bases(dici, ativos_ip, ativos_trans, sites):
    """
    Compara TODAS as bases entre si (não apenas o DICI como referência).
    Faz um outer join por 'Servico' e avalia, para cada serviço encontrado
    em QUALQUER uma das bases, se ele existe/é consistente nas demais.
    """

    # --- Base DICI reduzida para o merge ---
    dici_m = dici[["Servico", "NmServico", "nmempresa", "Tipo", "Velocidade_Gb"]].rename(
        columns={"NmServico": "Nome_XCONN", "nmempresa": "Cliente_XCONN"}
    )

    ativos_ip_m = ativos_ip[["Servico", "Capacidade_IP", "Qtd_Portas_IP", "Detalhe_Portas_IP", "Cliente_IP"]]
    ativos_trans_m = ativos_trans[["Servico", "Capacidade_Trans", "Qtd_Portas_Trans", "Detalhe_Portas_Trans", "Cliente_Trans"]]
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

    # --- Cliente: nome da empresa no XCONN; senão o nome da planilha de Ativos;
    # senão o código do cliente na designação (Trecho/CLIENTE/Tecnologia/Porta) ---
    def limpar_cliente(serie):
        serie = serie.astype("string").str.strip()
        return serie.mask(serie == "")

    codigo_cliente = base["Servico"].str.split("/").str[1]
    base["Cliente"] = (
        limpar_cliente(base["Cliente_XCONN"])
        .fillna(limpar_cliente(base["Cliente_IP"]))
        .fillna(limpar_cliente(base["Cliente_Trans"]))
        .fillna(codigo_cliente)
        .fillna("")
    )
    # O mesmo cliente pode vir com grafias diferentes entre as bases
    # (ex.: "MAXXNET" x "Maxxnet"): unifica pela grafia mais frequente.
    chave_cliente = base["Cliente"].apply(normalizar_texto)
    grafia_preferida = (
        base.assign(Chave_Cliente=chave_cliente)
        .groupby("Chave_Cliente")["Cliente"]
        .agg(lambda s: s.value_counts().index[0])
    )
    base["Cliente"] = chave_cliente.map(grafia_preferida)

    # --- Filtro: desconsidera serviços com algum dos nomes de FILTROS_SUBSTRING_IGNORAR ou nome muito curto ---
    nome_upper = base["Servico_Exibicao"].astype(str).str.upper()
    mask_ignorar_substring = pd.Series(False, index=base.index)
    for substring in FILTROS_SUBSTRING_IGNORAR:
        mask_ignorar_substring |= nome_upper.str.contains(substring.upper(), na=False)
    mask_ignorar_curto = base["Servico_Exibicao"].astype(str).str.len() < FILTRO_TAMANHO_MINIMO
    qtd_ignorados = int((mask_ignorar_substring | mask_ignorar_curto).sum())
    if qtd_ignorados:
        logger.info(
            f"Ignorando {qtd_ignorados} serviço(s) por conter {FILTROS_SUBSTRING_IGNORAR} "
            f"ou ter menos de {FILTRO_TAMANHO_MINIMO} caracteres."
        )
    base = base[~(mask_ignorar_substring | mask_ignorar_curto)].copy()

    # --- Capacidade de Ativos correspondente ao tipo declarado no XCONN ---
    def capacidade_ativos_correspondente(row):
        if row["Tipo"] == "IP":
            return row["Capacidade_IP"]
        elif row["Tipo"] == "TRANSPORTE":
            return row["Capacidade_Trans"]
        return pd.NA

    base["Gb_Ativos_Correspondente"] = base.apply(capacidade_ativos_correspondente, axis=1)

    # --- Sinaliza quando a capacidade veio da soma de mais de uma porta/canal ---
    def validado_via_soma_portas(row):
        if row["Tipo"] == "IP" and pd.notna(row.get("Qtd_Portas_IP")) and row["Qtd_Portas_IP"] > 1:
            return True
        if row["Tipo"] == "TRANSPORTE" and pd.notna(row.get("Qtd_Portas_Trans")) and row["Qtd_Portas_Trans"] > 1:
            return True
        return False

    base["Validado_Via_Soma_Portas"] = base.apply(validado_via_soma_portas, axis=1)

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
            divergencias.append(STATUS_NAO_ENCONTRADO_SITES)

        if not divergencias:
            return "OK"
        return "; ".join(divergencias)

    base["Status"] = base.apply(montar_status, axis=1)

    # --- Revalida por SOMA DO GRUPO: cobre o caso em que cada porta já tem seu
    # próprio registro no XCONN (então a comparação porta a porta acusa
    # divergência em cada uma), mas a SOMA de todas as portas do mesmo serviço
    # base é igual nas duas pontas (ex.: XCONN tem 10 + 2,5 = 12,5 e Ativos tem
    # 6,25 + 6,25 = 12,5 — o total contratado bate, só a distribuição entre
    # portas é diferente).
    base["Servico_Base_Agrupamento"] = base["Servico"].apply(extrair_servico_base)

    def revalidar_grupo_por_soma(indices):
        grupo = base.loc[indices]
        if len(grupo) < 2:
            return

        tipos_grupo = grupo["Tipo"].dropna().unique()
        if len(tipos_grupo) != 1:
            return
        tipo = tipos_grupo[0]

        soma_xconn = grupo["Velocidade_Gb"].sum(min_count=1)
        if tipo == "IP":
            soma_ativos = grupo["Capacidade_IP"].sum(min_count=1)
        elif tipo == "TRANSPORTE":
            soma_ativos = grupo["Capacidade_Trans"].sum(min_count=1)
        else:
            return

        if pd.isna(soma_xconn) or pd.isna(soma_ativos):
            return
        if round(soma_xconn, 2) != round(soma_ativos, 2):
            return

        # Soma do grupo bate: remove a divergência de capacidade de cada linha
        for idx in indices:
            partes = [
                p for p in base.at[idx, "Status"].split("; ")
                if p != "Capacidade divergente entre XCONN e Ativos"
            ]
            base.at[idx, "Status"] = "; ".join(partes) if partes else "OK"
            base.at[idx, "Validado_Via_Soma_Portas"] = True

    for _, indices in base.groupby("Servico_Base_Agrupamento").groups.items():
        revalidar_grupo_por_soma(indices)

    resultado = base[[
        "Servico_Exibicao",
        "Cliente",
        "Tipo",
        "Existe_XCONN",
        "Existe_Ativos_IP",
        "Existe_Ativos_Transporte",
        "Existe_Sites",
        "Velocidade_Gb",
        "Capacidade_IP",
        "Qtd_Portas_IP",
        "Detalhe_Portas_IP",
        "Capacidade_Trans",
        "Qtd_Portas_Trans",
        "Detalhe_Portas_Trans",
        "Validado_Via_Soma_Portas",
        "Status",
    ]].rename(columns={
        "Servico_Exibicao": "Servico",
        "Velocidade_Gb": "Gb_XCONN",
        "Capacidade_IP": "Gb_Ativos_IP_Total",
        "Capacidade_Trans": "Gb_Ativos_Transporte_Total",
    })

    # --- Preenche o último TAS registrado pela equipe (arquivo compartilhado) ---
    tas = carregar_tas(ARQ_TAS)
    resultado["Servico_Chave"] = resultado["Servico"].apply(normalizar_texto)
    resultado = resultado.merge(tas, on="Servico_Chave", how="left").drop(columns="Servico_Chave")
    resultado["Ultimo_TAS"] = resultado["Ultimo_TAS"].fillna("")

    # Ordena colocando as divergências primeiro
    resultado = resultado.sort_values(
        by="Status", key=lambda s: (s == "OK")
    ).reset_index(drop=True)

    return resultado


def filtrar_nao_encontrados_em_sites(df):
    """Devolve só os serviços cujo Status contém "Não encontrado em Sites"
    (sozinho ou junto com outras divergências separadas por "; ")."""
    mask = df["Status"].str.contains(STATUS_NAO_ENCONTRADO_SITES, regex=False, na=False)
    return df[mask].copy()


COR_FUNDO_CABECALHO = "1F497D"
COR_FONTE_CABECALHO = "F79646"


def estilizar_cabecalho(worksheet):
    """Aplica fundo azul (#1F497D) e letra laranja (#F79646) na primeira linha da aba."""
    preenchimento = PatternFill(
        start_color=COR_FUNDO_CABECALHO, end_color=COR_FUNDO_CABECALHO, fill_type="solid"
    )
    fonte = Font(color=COR_FONTE_CABECALHO, bold=True)
    for celula in worksheet[1]:
        celula.fill = preenchimento
        celula.font = fonte


def ranking_clientes(df):
    """Clientes ordenados pela quantidade de serviços desincronizados (maior primeiro)."""
    agrupado = (
        df.assign(Divergente=df["Status"] != "OK")
        .groupby("Cliente", as_index=False)
        .agg(Servicos_Divergentes=("Divergente", "sum"), Total_Servicos=("Servico", "count"))
    )
    agrupado["Percentual_Divergente"] = (
        agrupado["Servicos_Divergentes"] / agrupado["Total_Servicos"] * 100
    ).round(1)
    return agrupado.sort_values(
        by=["Servicos_Divergentes", "Percentual_Divergente", "Cliente"],
        ascending=[False, False, True],
    ).reset_index(drop=True)


def gerar_relatorio(df):
    total = len(df)
    divergentes_mask = df["Status"] != "OK"
    divergentes = int(divergentes_mask.sum())

    resumo = pd.DataFrame({
        "Métrica": ["Total de serviços analisados", "Serviços OK", "Serviços com divergência"],
        "Valor": [total, total - divergentes, divergentes],
    })

    # Divergencias: resumida, só com o essencial (serviço, tipo, motivo, último TAS)
    df_divergentes = (
        df[divergentes_mask][["Servico", "Cliente", "Tipo", "Status", "Ultimo_TAS"]]
        .rename(columns={"Status": "Motivo_Divergencia"})
        .reset_index(drop=True)
    )

    # Divergencia_Detalhada: mesmas linhas divergentes, mas com todas as colunas
    df_divergentes_detalhado = df[divergentes_mask].reset_index(drop=True)

    df_clientes = ranking_clientes(df)

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        df_divergentes.to_excel(writer, sheet_name="Divergencias", index=False)
        df_divergentes_detalhado.to_excel(writer, sheet_name="Divergencia_Detalhada", index=False)
        df_clientes.to_excel(writer, sheet_name="Clientes", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        for nome_aba in ["Divergencias", "Divergencia_Detalhada", "Clientes", "Resumo"]:
            estilizar_cabecalho(writer.sheets[nome_aba])

    logger.info(f"Relatório gerado: {RELATORIO_SAIDA} ({divergentes}/{total} divergências)")


def gerar_dashboard_html(df):
    """
    Gera um dashboard HTML autocontido (sem depender de internet/CDN) com
    busca, filtro por Tipo/Status e cores, pra visualizar o que precisa
    sincronizar sem precisar abrir o Excel.
    """
    total = len(df)
    divergentes = int((df["Status"] != "OK").sum())
    ok = total - divergentes

    colunas_exibicao = ["Servico", "Cliente", "Tipo", "Status", "Gb_XCONN"]
    if "Gb_Ativos_IP_Total" in df.columns:
        colunas_exibicao.append("Gb_Ativos_IP_Total")
    if "Gb_Ativos_Transporte_Total" in df.columns:
        colunas_exibicao.append("Gb_Ativos_Transporte_Total")
    colunas_exibicao.append("Validado_Via_Soma_Portas")
    if "Ultimo_TAS" in df.columns:
        colunas_exibicao.append("Ultimo_TAS")

    registros = df[colunas_exibicao].fillna("").to_dict(orient="records")
    dados_json = json.dumps(registros, ensure_ascii=False)

    # IP e TRANSPORTE sempre aparecem no dropdown/botões de setor, mesmo que
    # não haja nenhum serviço desse tipo na execução atual.
    tipos_padrao = ["IP", "TRANSPORTE"]
    tipos_extras = [t for t in sorted(df["Tipo"].dropna().unique().tolist()) if t not in tipos_padrao]
    tipos_disponiveis = tipos_padrao + tipos_extras
    opcoes_tipo = "".join(f'<option value="{t}">{t}</option>' for t in tipos_disponiveis)

    logo_src = carregar_logo_base64(LOGO_ARQUIVO)
    logo_html = f'<img src="{logo_src}" class="logo" alt="Eletronet">' if logo_src else ""

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Dashboard de Divergências - {datetime.now().strftime('%d/%m/%Y')}</title>
<style>
    :root {{
        --azul: #1F497D;
        --laranja: #F79646;
        --verde: #2e7d32;
        --vermelho: #c62828;
        --fundo: #f4f6f8;
    }}
    * {{ box-sizing: border-box; }}
    body {{
        font-family: -apple-system, Segoe UI, Arial, sans-serif;
        background: var(--fundo);
        margin: 0;
        padding: 24px;
        color: #222;
    }}
    .cabecalho {{
        display: flex; justify-content: space-between; align-items: flex-start;
        gap: 16px; flex-wrap: wrap;
    }}
    .logo {{ height: 75px; width: 150px; }}
    h1 {{ color: var(--azul); margin-bottom: 4px; }}
    .subtitulo {{ color: #666; margin-top: 0; margin-bottom: 20px; }}
    .cards {{ display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }}
    .card {{
        background: white; border-radius: 8px; padding: 16px 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.15); min-width: 140px;
    }}
    .card .numero {{ font-size: 28px; font-weight: bold; }}
    .card .rotulo {{ color: #666; font-size: 13px; }}
    .card.total .numero {{ color: var(--azul); }}
    .card.ok .numero {{ color: var(--verde); }}
    .card.divergente .numero {{ color: var(--vermelho); }}
    .barra-fixa {{
        position: sticky; top: 0; z-index: 10;
        background: var(--fundo); padding: 12px 0 4px;
        box-shadow: 0 4px 6px -6px rgba(0,0,0,0.3);
    }}
    .setores {{ display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }}
    .btn-setor {{
        padding: 8px 16px; border: 1px solid var(--azul); border-radius: 6px;
        background: white; color: var(--azul); font-size: 13px; font-weight: 600;
        cursor: pointer;
    }}
    .btn-setor:hover {{ background: #eef2f7; }}
    .btn-setor.ativo {{ background: var(--azul); color: white; }}
    .controles {{
        display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center;
    }}
    .controles input, .controles select {{
        padding: 8px 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px;
    }}
    .controles input[type="text"] {{ flex: 1; min-width: 220px; }}
    table {{
        width: 100%; border-collapse: collapse; background: white;
        border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.15);
    }}
    thead th {{
        background: var(--azul); color: var(--laranja); text-align: left;
        padding: 10px 12px; font-size: 13px; cursor: pointer; user-select: none;
        white-space: nowrap; border-right: 1px solid rgba(255,255,255,0.25);
    }}
    thead th:last-child {{ border-right: none; }}
    thead th:hover {{ opacity: 0.9; }}
    .seta-ordenacao {{ display: inline-block; width: 12px; font-size: 11px; opacity: 0.9; }}
    tbody td {{
        padding: 9px 12px; font-size: 13px;
        border-top: 1px dashed #ccc; border-right: 1px solid #e0e0e0;
    }}
    tbody td:last-child {{ border-right: none; }}
    tbody tr.status-ok {{ background: #f1f8f1; }}
    tbody tr.status-divergente {{ background: #fdf1f0; }}
    .badge {{
        display: inline-block; padding: 2px 8px; border-radius: 10px;
        font-size: 12px; font-weight: 600;
    }}
    .badge-ok {{ background: #c8e6c9; color: var(--verde); }}
    .badge-divergente {{ background: #ffcdd2; color: var(--vermelho); }}
    .painel-clientes {{
        background: white; border-radius: 8px; padding: 12px 16px; margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.15);
    }}
    .painel-clientes summary {{
        cursor: pointer; font-weight: 600; color: var(--azul); font-size: 15px;
    }}
    .painel-clientes .dica {{ color: #666; font-size: 12px; margin: 6px 0 10px; }}
    .lista-clientes {{ max-height: 320px; overflow-y: auto; }}
    .item-cliente {{
        display: grid; grid-template-columns: minmax(120px, 260px) 1fr auto;
        gap: 12px; align-items: center; padding: 5px 6px; border-radius: 4px;
        cursor: pointer; font-size: 13px;
    }}
    .item-cliente:hover {{ background: #eef2f7; }}
    .item-cliente.ativo {{ background: #dce6f2; font-weight: 600; }}
    .item-cliente .nome {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .item-cliente .barra {{ background: #eee; border-radius: 3px; height: 10px; overflow: hidden; }}
    .item-cliente .barra div {{ background: var(--vermelho); height: 100%; }}
    .item-cliente .numeros {{ color: #555; white-space: nowrap; font-variant-numeric: tabular-nums; }}
    .contagem {{ color: #666; font-size: 13px; margin-top: 8px; }}
    .sem-resultado {{ text-align: center; padding: 30px; color: #888; }}
</style>
</head>
<body>

<div class="cabecalho">
    <div>
        <h1>Dashboard de Divergências</h1>
        <p class="subtitulo">Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
    </div>
    {logo_html}
</div>

<div class="cards">
    <div class="card total"><div class="numero">{total}</div><div class="rotulo">Total analisado</div></div>
    <div class="card ok"><div class="numero">{ok}</div><div class="rotulo">OK</div></div>
    <div class="card divergente"><div class="numero">{divergentes}</div><div class="rotulo">Precisam sincronizar</div></div>
</div>

<details class="painel-clientes" open>
    <summary>Clientes mais desincronizados</summary>
    <p class="dica">Ordenado pela quantidade de serviços com divergência (respeita o filtro de setor). Clique num cliente para filtrar a tabela; clique de novo para limpar.</p>
    <div class="lista-clientes" id="listaClientes"></div>
</details>

<div class="barra-fixa">
    <div class="setores">
        <button type="button" class="btn-setor" data-tipo="todos">Todos os setores</button>
        <button type="button" class="btn-setor" data-tipo="IP">Só IP</button>
        <button type="button" class="btn-setor" data-tipo="TRANSPORTE">Só Transporte</button>
    </div>

    <div class="controles">
        <input type="text" id="busca" placeholder="Buscar por nome do serviço...">
        <select id="filtroStatus">
            <option value="divergente">Só divergentes</option>
            <option value="todos">Todos</option>
            <option value="ok">Só OK</option>
            <option value="nao_sites">Só não encontrados em Sites</option>
        </select>
        <select id="filtroTipo">
            <option value="todos">Todos os tipos</option>
            {opcoes_tipo}
        </select>
        <select id="filtroCliente">
            <option value="todos">Todos os clientes</option>
        </select>
    </div>
</div>

<table id="tabela">
    <thead>
        <tr>
            <th data-col="Servico">Serviço<span class="seta-ordenacao"></span></th>
            <th data-col="Cliente">Cliente<span class="seta-ordenacao"></span></th>
            <th data-col="Tipo">Tipo<span class="seta-ordenacao"></span></th>
            <th data-col="Status">Status<span class="seta-ordenacao"></span></th>
            <th data-col="Gb_XCONN">Gb XCONN<span class="seta-ordenacao"></span></th>
            <th data-col="Gb_Ativos_IP_Total">Gb Ativos IP<span class="seta-ordenacao"></span></th>
            <th data-col="Gb_Ativos_Transporte_Total">Gb Ativos Transp.<span class="seta-ordenacao"></span></th>
            <th data-col="Validado_Via_Soma_Portas">Validado por soma?<span class="seta-ordenacao"></span></th>
            <th data-col="Ultimo_TAS">Último TAS<span class="seta-ordenacao"></span></th>
        </tr>
    </thead>
    <tbody></tbody>
</table>
<div class="sem-resultado" id="semResultado" style="display:none;">Nenhum serviço encontrado com esse filtro.</div>
<p class="contagem" id="contagem"></p>

<script>
    const dados = {dados_json};
    const STATUS_NAO_SITES = {json.dumps(STATUS_NAO_ENCONTRADO_SITES, ensure_ascii=False)};
    let ordenacaoAtual = {{ coluna: null, asc: true }};

    function escaparHtml(texto) {{
        return String(texto).replace(/[&<>"']/g, function(c) {{
            return {{ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }}[c];
        }});
    }}

    // Clientes ordenados por nº de serviços desincronizados (maior primeiro),
    // considerando só o setor selecionado.
    function calcularRankingClientes() {{
        const filtroTipo = document.getElementById('filtroTipo').value;
        const porCliente = {{}};
        dados.forEach(function(linha) {{
            if (filtroTipo !== 'todos' && linha.Tipo !== filtroTipo) return;
            const nome = linha.Cliente || '(sem cliente)';
            if (!porCliente[nome]) porCliente[nome] = {{ nome: nome, divergentes: 0, total: 0 }};
            porCliente[nome].total++;
            if (linha.Status !== 'OK') porCliente[nome].divergentes++;
        }});
        return Object.values(porCliente).sort(function(a, b) {{
            return (b.divergentes - a.divergentes)
                || (b.divergentes / b.total - a.divergentes / a.total)
                || a.nome.localeCompare(b.nome);
        }});
    }}

    function renderizarClientes() {{
        const ranking = calcularRankingClientes();
        const selecionado = document.getElementById('filtroCliente').value;

        // Dropdown: mesma ordem do ranking, com a contagem de divergências
        const select = document.getElementById('filtroCliente');
        select.innerHTML = '<option value="todos">Todos os clientes</option>' + ranking.map(function(c) {{
            return '<option value="' + escaparHtml(c.nome) + '">' + escaparHtml(c.nome) +
                ' (' + c.divergentes + ' desinc.)</option>';
        }}).join('');
        select.value = ranking.some(function(c) {{ return c.nome === selecionado; }}) ? selecionado : 'todos';

        // Painel: só quem tem algo pra sincronizar
        const comDivergencia = ranking.filter(function(c) {{ return c.divergentes > 0; }});
        const maximo = comDivergencia.length ? comDivergencia[0].divergentes : 1;
        const lista = document.getElementById('listaClientes');
        lista.innerHTML = comDivergencia.length ? '' : '<p class="dica">Nenhum cliente com divergência.</p>';
        comDivergencia.forEach(function(c) {{
            const item = document.createElement('div');
            item.className = 'item-cliente' + (select.value === c.nome ? ' ativo' : '');
            item.title = c.nome;
            item.innerHTML =
                '<span class="nome">' + escaparHtml(c.nome) + '</span>' +
                '<div class="barra"><div style="width:' + (c.divergentes / maximo * 100) + '%"></div></div>' +
                '<span class="numeros">' + c.divergentes + ' de ' + c.total +
                ' (' + Math.round(c.divergentes / c.total * 100) + '%)</span>';
            item.addEventListener('click', function() {{
                select.value = select.value === c.nome ? 'todos' : c.nome;
                renderizarClientes();
                renderizar();
            }});
            lista.appendChild(item);
        }});
    }}

    function renderizar() {{
        const busca = document.getElementById('busca').value.trim().toUpperCase();
        const filtroStatus = document.getElementById('filtroStatus').value;
        const filtroTipo = document.getElementById('filtroTipo').value;
        const filtroCliente = document.getElementById('filtroCliente').value;

        let linhas = dados.filter(function(linha) {{
            const ehOk = linha.Status === 'OK';
            if (filtroStatus === 'divergente' && ehOk) return false;
            if (filtroStatus === 'ok' && !ehOk) return false;
            if (filtroStatus === 'nao_sites' && !String(linha.Status).includes(STATUS_NAO_SITES)) return false;
            if (filtroTipo !== 'todos' && linha.Tipo !== filtroTipo) return false;
            if (filtroCliente !== 'todos' && (linha.Cliente || '(sem cliente)') !== filtroCliente) return false;
            if (busca && !String(linha.Servico).toUpperCase().includes(busca)
                && !String(linha.Cliente).toUpperCase().includes(busca)) return false;
            return true;
        }});

        if (ordenacaoAtual.coluna) {{
            const col = ordenacaoAtual.coluna;
            linhas = linhas.slice().sort(function(a, b) {{
                const va = a[col], vb = b[col];
                if (va < vb) return ordenacaoAtual.asc ? -1 : 1;
                if (va > vb) return ordenacaoAtual.asc ? 1 : -1;
                return 0;
            }});
        }}

        const tbody = document.querySelector('#tabela tbody');
        tbody.innerHTML = '';
        document.getElementById('semResultado').style.display = linhas.length ? 'none' : 'block';

        linhas.forEach(function(linha) {{
            const ehOk = linha.Status === 'OK';
            const tr = document.createElement('tr');
            tr.className = ehOk ? 'status-ok' : 'status-divergente';
            const badge = ehOk
                ? '<span class="badge badge-ok">OK</span>'
                : '<span class="badge badge-divergente">' + linha.Status + '</span>';
            tr.innerHTML =
                '<td>' + linha.Servico + '</td>' +
                '<td>' + escaparHtml(linha.Cliente || '') + '</td>' +
                '<td>' + linha.Tipo + '</td>' +
                '<td>' + badge + '</td>' +
                '<td>' + linha.Gb_XCONN + '</td>' +
                '<td>' + (linha.Gb_Ativos_IP_Total !== undefined ? linha.Gb_Ativos_IP_Total : '') + '</td>' +
                '<td>' + (linha.Gb_Ativos_Transporte_Total !== undefined ? linha.Gb_Ativos_Transporte_Total : '') + '</td>' +
                '<td>' + (linha.Validado_Via_Soma_Portas === true ? 'Sim' : '') + '</td>' +
                '<td>' + (linha.Ultimo_TAS || '') + '</td>';
            tbody.appendChild(tr);
        }});

        document.getElementById('contagem').textContent = linhas.length + ' de ' + dados.length + ' serviço(s) exibido(s)';
    }}

    function atualizarBotoesSetor() {{
        const atual = document.getElementById('filtroTipo').value;
        document.querySelectorAll('.btn-setor').forEach(function(btn) {{
            btn.classList.toggle('ativo', btn.getAttribute('data-tipo') === atual);
        }});
    }}

    document.querySelectorAll('.btn-setor').forEach(function(btn) {{
        btn.addEventListener('click', function() {{
            document.getElementById('filtroTipo').value = btn.getAttribute('data-tipo');
            atualizarBotoesSetor();
            renderizarClientes();
            renderizar();
        }});
    }});

    function atualizarSetasOrdenacao() {{
        document.querySelectorAll('#tabela thead th').forEach(function(th) {{
            const seta = th.querySelector('.seta-ordenacao');
            const col = th.getAttribute('data-col');
            if (ordenacaoAtual.coluna !== col) {{
                seta.textContent = '';
            }} else {{
                seta.textContent = ordenacaoAtual.asc ? '▲' : '▼';
            }}
        }});
    }}

    document.querySelectorAll('#tabela thead th').forEach(function(th) {{
        th.addEventListener('click', function() {{
            const col = th.getAttribute('data-col');
            if (ordenacaoAtual.coluna !== col) {{
                // 1º clique na coluna: ordena crescente
                ordenacaoAtual = {{ coluna: col, asc: true }};
            }} else if (ordenacaoAtual.asc) {{
                // 2º clique: ordena decrescente
                ordenacaoAtual.asc = false;
            }} else {{
                // 3º clique: volta pra ordem padrão (sem ordenação, só filtro/busca)
                ordenacaoAtual = {{ coluna: null, asc: true }};
            }}
            atualizarSetasOrdenacao();
            renderizar();
        }});
    }});

    document.getElementById('busca').addEventListener('input', renderizar);
    document.getElementById('filtroStatus').addEventListener('change', renderizar);
    document.getElementById('filtroTipo').addEventListener('change', function() {{
        atualizarBotoesSetor();
        renderizarClientes();
        renderizar();
    }});
    document.getElementById('filtroCliente').addEventListener('change', function() {{
        renderizarClientes();
        renderizar();
    }});

    atualizarBotoesSetor();
    atualizarSetasOrdenacao();
    renderizarClientes();
    renderizar();
</script>
</body>
</html>"""

    with open(DASHBOARD_SAIDA, "w", encoding="utf-8") as f:
        f.write(html)

    logger.info(f"Dashboard gerado: {DASHBOARD_SAIDA}")

    # Cópia com nome fixo, servida pelo servidor_dashboard.py: assim o link
    # de acesso continua o mesmo a cada execução, sempre com o conteúdo mais
    # recente.
    os.makedirs(SITE_DIR, exist_ok=True)
    with open(SITE_ARQUIVO, "w", encoding="utf-8") as f:
        f.write(html)
    logger.info(f"Cópia do dashboard atualizada em: {SITE_ARQUIVO}")

    if ABRIR_DASHBOARD_AUTOMATICO:
        try:
            webbrowser.open(f"file://{os.path.abspath(DASHBOARD_SAIDA)}")
        except Exception as e:
            logger.warning(f"Não foi possível abrir o dashboard automaticamente: {e}")


def resumo_para_email(df):
    """Números usados no corpo do e-mail (HTML e texto puro)."""
    divergentes = df[df["Status"] != "OK"]
    motivos = (
        divergentes["Status"].str.split("; ").explode().value_counts()
    )
    setores = []
    for tipo, rotulo in [("IP", "IP"), ("TRANSPORTE", "Transporte")]:
        do_tipo = df[df["Tipo"] == tipo]
        setores.append((rotulo, int((do_tipo["Status"] != "OK").sum()), len(do_tipo)))
    ranking = ranking_clientes(df)
    return {
        "total": len(df),
        "divergentes": len(divergentes),
        "ok": len(df) - len(divergentes),
        "motivos": list(motivos.items())[:EMAIL_QTD_MOTIVOS],
        "setores": setores,
        "clientes": ranking[ranking["Servicos_Divergentes"] > 0].head(EMAIL_QTD_CLIENTES),
        "qtd_clientes_afetados": int((ranking["Servicos_Divergentes"] > 0).sum()),
    }


def gerar_email_texto(resumo):
    """Versão em texto puro, para clientes de e-mail que não exibem HTML."""
    linhas = [
        "Olá,",
        "",
        f"Resultado da comparação de bases de {datetime.now().strftime('%d/%m/%Y')}:",
        f"  - Serviços analisados: {resumo['total']}",
        f"  - OK: {resumo['ok']}",
        f"  - Precisam sincronizar: {resumo['divergentes']}",
        "",
        "Clientes mais desincronizados:",
    ]
    for _, c in resumo["clientes"].iterrows():
        linhas.append(f"  - {c['Cliente']}: {c['Servicos_Divergentes']} de {c['Total_Servicos']}")
    if DASHBOARD_URL:
        linhas += ["", f"Dashboard interativo: {DASHBOARD_URL}"]
    linhas += ["", "O relatório completo segue em anexo (Excel)."]
    return "\n".join(linhas)


def gerar_email_html(resumo, logo_cid=None):
    """
    Corpo do e-mail em HTML. Feito com tabelas e estilos inline (sem
    JavaScript, grid ou <style> complexo) porque é o que o Outlook respeita.
    """
    esc = html_lib.escape
    azul, laranja = f"#{COR_FUNDO_CABECALHO}", f"#{COR_FONTE_CABECALHO}"
    verde, vermelho, cinza = "#2e7d32", "#c62828", "#666666"
    fonte = "font-family:Segoe UI,Arial,sans-serif;"
    total = resumo["total"] or 1
    pct_div = round(resumo["divergentes"] / total * 100, 1)

    logo = (
        f'<img src="cid:{logo_cid}" width="110" height="56" alt="Eletronet" '
        f'style="display:block;border:0;">' if logo_cid else ""
    )

    def card(numero, rotulo, cor, detalhe=""):
        return (
            f'<td width="33%" valign="top" style="padding:0 6px;">'
            f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
            f'style="background:#ffffff;border:1px solid #e3e7ec;border-top:4px solid {cor};">'
            f'<tr><td style="padding:14px 16px;{fonte}">'
            f'<div style="font-size:28px;font-weight:bold;color:{cor};">{numero}</div>'
            f'<div style="font-size:13px;color:{cinza};">{rotulo}</div>'
            f'<div style="font-size:12px;color:{cinza};">{detalhe}&nbsp;</div>'
            f'</td></tr></table></td>'
        )

    def titulo_secao(texto):
        return (
            f'<tr><td style="padding:24px 24px 8px;{fonte}font-size:16px;font-weight:bold;'
            f'color:{azul};">{texto}</td></tr>'
        )

    def barra(valor, maximo, cor):
        largura = max(int(round(valor / (maximo or 1) * 100)), 2)
        return (
            f'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" '
            f'style="background:#eeeeee;"><tr>'
            f'<td width="{largura}%" height="10" style="background:{cor};font-size:0;line-height:0;">&nbsp;</td>'
            + (f'<td style="font-size:0;line-height:0;">&nbsp;</td>' if largura < 100 else "")
            + "</tr></table>"
        )

    setores_html = " &nbsp;·&nbsp; ".join(
        f"{rotulo}: <b>{div}</b> de {tot}" for rotulo, div, tot in resumo["setores"]
    )

    # Clientes mais desincronizados
    linhas_clientes = ""
    max_cli = int(resumo["clientes"]["Servicos_Divergentes"].max()) if len(resumo["clientes"]) else 1
    for i, (_, c) in enumerate(resumo["clientes"].iterrows()):
        fundo = "#ffffff" if i % 2 == 0 else "#f7f9fb"
        linhas_clientes += (
            f'<tr style="background:{fundo};">'
            f'<td style="padding:7px 10px;{fonte}font-size:13px;color:#222;">{esc(str(c["Cliente"]))}</td>'
            f'<td width="35%" style="padding:7px 10px;">{barra(c["Servicos_Divergentes"], max_cli, vermelho)}</td>'
            f'<td align="right" style="padding:7px 10px;{fonte}font-size:13px;color:#444;white-space:nowrap;">'
            f'<b>{int(c["Servicos_Divergentes"])}</b> de {int(c["Total_Servicos"])}</td></tr>'
        )
    if not linhas_clientes:
        linhas_clientes = (
            f'<tr><td style="padding:10px;{fonte}font-size:13px;color:{verde};">'
            f'Nenhum cliente com divergência. 🎉</td></tr>'
        )

    # Principais motivos de divergência
    linhas_motivos = ""
    max_mot = resumo["motivos"][0][1] if resumo["motivos"] else 1
    for i, (motivo, qtd) in enumerate(resumo["motivos"]):
        fundo = "#ffffff" if i % 2 == 0 else "#f7f9fb"
        linhas_motivos += (
            f'<tr style="background:{fundo};">'
            f'<td style="padding:7px 10px;{fonte}font-size:13px;color:#222;">{esc(motivo)}</td>'
            f'<td width="30%" style="padding:7px 10px;">{barra(qtd, max_mot, laranja)}</td>'
            f'<td align="right" style="padding:7px 10px;{fonte}font-size:13px;color:#444;"><b>{qtd}</b></td></tr>'
        )

    botao = ""
    if DASHBOARD_URL:
        botao = (
            f'<tr><td align="center" style="padding:24px;">'
            f'<table role="presentation" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="background:{azul};border-radius:6px;">'
            f'<a href="{esc(DASHBOARD_URL)}" style="display:inline-block;padding:12px 28px;{fonte}'
            f'font-size:14px;font-weight:bold;color:#ffffff;text-decoration:none;">'
            f'Abrir dashboard interativo</a></td></tr></table>'
            f'<div style="{fonte}font-size:11px;color:{cinza};padding-top:6px;">'
            f'Funciona na rede interna da Eletronet.</div></td></tr>'
        )

    if resumo["divergentes"]:
        frase_resumo = (
            f'<b style="color:{vermelho};">{resumo["divergentes"]} serviço(s)</b> de '
            f'{resumo["qtd_clientes_afetados"]} cliente(s) precisam ser sincronizados.'
        )
    else:
        frase_resumo = f'<b style="color:{verde};">todas as bases estão sincronizadas.</b>'

    tabela = (
        'role="presentation" width="100%" cellpadding="0" cellspacing="0" '
        'style="border:1px solid #e3e7ec;border-collapse:collapse;"'
    )

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f6f8;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6f8;">
<tr><td align="center" style="padding:24px 12px;">
<table role="presentation" width="640" cellpadding="0" cellspacing="0" style="width:100%;max-width:640px;background:#ffffff;border:1px solid #e3e7ec;">

{f'<tr><td style="padding:16px 24px 12px;">{logo}</td></tr>' if logo else ''}
<tr><td style="background:{azul};padding:18px 24px;border-bottom:4px solid {laranja};{fonte}">
    <div style="font-size:20px;font-weight:bold;color:#ffffff;">Relatório de Divergências</div>
    <div style="font-size:13px;color:{laranja};padding-top:4px;">
        Comparação XCONN × Ativos × Sites · {datetime.now().strftime('%d/%m/%Y')}</div>
</td></tr>

<tr><td style="padding:22px 24px 6px;{fonte}font-size:14px;color:#333;line-height:1.5;">
    Olá,<br>
    Segue o resultado da comparação desta semana: {frase_resumo}
</td></tr>

<tr><td style="padding:12px 18px 0;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
        {card(resumo['total'], 'Serviços analisados', azul)}
        {card(resumo['ok'], 'OK', verde)}
        {card(resumo['divergentes'], 'Precisam sincronizar', vermelho, f'{pct_div}% do total')}
    </tr></table>
</td></tr>
<tr><td style="padding:10px 24px 0;{fonte}font-size:12px;color:{cinza};">Divergências por setor: {setores_html}</td></tr>

{titulo_secao('Clientes mais desincronizados')}
<tr><td style="padding:0 24px;"><table {tabela}>{linhas_clientes}</table></td></tr>

{titulo_secao('Principais motivos') if linhas_motivos else ''}
{f'<tr><td style="padding:0 24px;"><table {tabela}>{linhas_motivos}</table></td></tr>' if linhas_motivos else ''}

{botao or '<tr><td style="padding:12px;"></td></tr>'}

<tr><td style="padding:0 24px 22px;{fonte}font-size:13px;color:#333;">
    📎 A lista completa, com o detalhamento de cada serviço, está na planilha em anexo.
</td></tr>

<tr><td style="background:#f4f6f8;padding:14px 24px;{fonte}font-size:11px;color:#888;border-top:1px solid #e3e7ec;">
    E-mail gerado automaticamente pela automação de comparação de bases. Não é necessário responder.
</td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""


def enviar_email_com_relatorio(resultado, arquivo_anexo):
    if not ENVIAR_EMAIL:
        logger.info("Envio de e-mail desabilitado nas configurações.")
        return

    logger.info("Preparando envio de e-mail...")

    resumo = resumo_para_email(resultado)

    msg = EmailMessage()
    msg["Subject"] = (
        f"Relatório de Divergências - {datetime.now().strftime('%d/%m/%Y')} "
        f"({resumo['divergentes']} para sincronizar)"
    )
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = EMAIL_DESTINATARIO
    msg.set_content(gerar_email_texto(resumo))

    # Logo vai como imagem inline (cid:), que o Outlook exibe; data URI ele bloqueia.
    logo_bytes = None
    try:
        with open(LOGO_EMAIL_ARQUIVO, "rb") as f:
            logo_bytes = f.read()
    except Exception as e:
        logger.warning(f"Não foi possível carregar o logo do e-mail ({LOGO_EMAIL_ARQUIVO}): {e}")

    logo_cid = make_msgid(domain="eletronet.local")[1:-1] if logo_bytes else None
    msg.add_alternative(gerar_email_html(resumo, logo_cid), subtype="html")
    if logo_bytes:
        msg.get_payload()[1].add_related(
            logo_bytes, maintype="image", subtype="jpeg", cid=f"<{logo_cid}>"
        )

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
    gerar_dashboard_html(resultado)

    enviar_email_com_relatorio(resultado, RELATORIO_SAIDA)

    logger.info("Finalizado.")


if __name__ == "__main__":
    main()