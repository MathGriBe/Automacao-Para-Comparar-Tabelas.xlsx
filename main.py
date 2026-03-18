import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

ARQ_EXCEL = os.getenv("ARQ_EXCEL")
SHEET_NAME = os.getenv("SHEET_NAME", "Planilha1")
RELATORIO_SAIDA = os.getenv("RELATORIO_SAIDA", "resultado_comparacao.xlsx")


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def preparar_lado_esquerdo(df):
    esq = df.iloc[:, [0, 1, 2]].copy()
    esq.columns = ["Cliente", "Circuito", "SIGLA"]

    for col in ["Cliente", "Circuito", "SIGLA"]:
        esq[col] = esq[col].apply(normalizar_texto)

    esq = esq[esq["Circuito"] != ""].copy()
    return esq


def preparar_lado_direito(df):
    dir_ = df.iloc[:, [5, 6, 7]].copy()
    dir_.columns = ["Cliente", "Circuito", "SIGLA"]

    for col in ["Cliente", "Circuito", "SIGLA"]:
        dir_[col] = dir_[col].apply(normalizar_texto)

    dir_ = dir_[dir_["Circuito"] != ""].copy()
    return dir_


def validar(row):
    if pd.isna(row["Circuito_DIR"]):
        return "Não encontrado no lado direito"

    erros = []

    if row["Cliente_ESQ"] != row["Cliente_DIR"]:
        erros.append("Cliente")

    if row["SIGLA_ESQ"] != row["SIGLA_DIR"]:
        erros.append("SIGLA")

    if erros:
        return "Divergência em: " + ", ".join(erros)

    return "OK"


def main():
    print("===== INÍCIO =====")
    print("Arquivo:", ARQ_EXCEL)
    print("Aba:", SHEET_NAME)

    if not ARQ_EXCEL:
        raise ValueError("ARQ_EXCEL não foi definido no .env")

    if not os.path.exists(ARQ_EXCEL):
        raise FileNotFoundError(f"Arquivo não encontrado: {ARQ_EXCEL}")

    df = pd.read_excel(ARQ_EXCEL, sheet_name=SHEET_NAME)

    esquerda = preparar_lado_esquerdo(df)
    direita = preparar_lado_direito(df)

    dup_esq = esquerda[esquerda["Circuito"].duplicated(keep=False)].copy()
    dup_dir = direita[direita["Circuito"].duplicated(keep=False)].copy()

    esquerda_unica = esquerda.drop_duplicates(subset=["Circuito"], keep="first").copy()
    direita_unica = direita.drop_duplicates(subset=["Circuito"], keep="first").copy()

    base = esquerda_unica.merge(
        direita_unica,
        on="Circuito",
        how="left",
        suffixes=("_ESQ", "_DIR")
    )

    base["Circuito_DIR"] = base["Circuito"].where(base["Cliente_DIR"].notna())

    base["Status"] = base.apply(validar, axis=1)

    divergencias = base[base["Status"] != "OK"].copy()
    ok = base[base["Status"] == "OK"].copy()

    so_no_direito = direita_unica[~direita_unica["Circuito"].isin(esquerda_unica["Circuito"])].copy()

    with pd.ExcelWriter(RELATORIO_SAIDA, engine="openpyxl") as writer:
        base.to_excel(writer, sheet_name="Base_Comparacao", index=False)
        divergencias.to_excel(writer, sheet_name="Divergencias", index=False)
        ok.to_excel(writer, sheet_name="OK", index=False)
        so_no_direito.to_excel(writer, sheet_name="So_no_Direito", index=False)
        dup_esq.to_excel(writer, sheet_name="Duplicados_Esquerda", index=False)
        dup_dir.to_excel(writer, sheet_name="Duplicados_Direita", index=False)

    print(f"Relatório gerado: {RELATORIO_SAIDA}")
    print("===== FIM =====")


if __name__ == "__main__":
    main()