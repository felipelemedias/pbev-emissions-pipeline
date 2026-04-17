"""
Extrai dados da tabela PBEV 2026 (INMETRO) usando pdfplumber e salva em CSV.

Uso:
    pip install pdfplumber pandas
    python extract_pbev.py
"""

import sys
import pandas as pd
import pdfplumber
from pathlib import Path

PDF_PATH = Path("data/pbve/Tabela PBEV 2026_20_JAN-REV04.pdf")
CSV_PATH = Path("data/pbev_2026.csv")

L8_LIMITS = {"co_mg_km": 700.0, "nmog_nox_mg_km": 60.0}

# Índices confirmados inspecionando linha do SPIN 1.8L (CO=800 no [11])
COL_IDX = {
    "categoria":        0,
    "marca":            1,
    "modelo":           2,
    "versao":           3,
    "motor":            4,
    "tipo_propulsao":   5,
    "transmissao":      6,
    "combustivel":      9,
    "nmog_nox_mg_km":  10,
    "co_mg_km":        11,
    "nota_verde_l8":   13,
    "co2_fossil_g_km": 15,
    "consumo_mj_km":   23,
    "classificacao_pbe": 25,
}

HEADER_KEYWORDS = {
    "marca", "modelo", "categoria", "versão", "versao", "motor",
    "combustível", "combustivel", "emissões", "emissoes", "transmissão",
}


def to_float(val) -> float | None:
    if val is None:
        return None
    s = str(val).strip().replace(",", ".")
    if s in ("\\", "-", "", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def is_data_row(row: list) -> bool:
    if len(row) < 12:
        return False
    marca = str(row[COL_IDX["marca"]] or "").strip()
    modelo = str(row[COL_IDX["modelo"]] or "").strip()
    if not marca or not modelo:
        return False
    if marca.lower() in HEADER_KEYWORDS or modelo.lower() in HEADER_KEYWORDS:
        return False
    # Linha de dado válida tem pelo menos marca e modelo preenchidos
    return True


def extract_rows() -> list[dict]:
    records = []
    with pdfplumber.open(PDF_PATH) as pdf:
        print(f"PDF aberto: {len(pdf.pages)} páginas")
        for i, page in enumerate(pdf.pages, 1):
            count = 0
            for table in page.extract_tables():
                for row in table:
                    if not is_data_row(row):
                        continue
                    record = {}
                    for col_name, idx in COL_IDX.items():
                        val = row[idx] if idx < len(row) else None
                        record[col_name] = str(val).strip() if val is not None else None
                    records.append(record)
                    count += 1
            print(f"  Página {i}: {count} veículos")

    return records


def clean(df: pd.DataFrame) -> pd.DataFrame:
    numeric = ["nmog_nox_mg_km", "co_mg_km", "co2_fossil_g_km", "consumo_mj_km"]
    for col in numeric:
        df[col] = df[col].apply(to_float)

    text = ["categoria", "marca", "modelo", "versao", "motor",
            "tipo_propulsao", "transmissao", "combustivel",
            "nota_verde_l8", "classificacao_pbe"]
    for col in text:
        df[col] = df[col].replace({"None": pd.NA, "\\": pd.NA, "-": pd.NA})

    return df.drop_duplicates().reset_index(drop=True)


def compute_compliance(df: pd.DataFrame) -> pd.DataFrame:
    for col, limit in L8_LIMITS.items():
        excesso = col.split("_")[0] + "_excesso_pct"
        df[excesso] = (df[col] - limit).clip(lower=0) / limit * 100
        df[excesso] = df[excesso].round(2)

    df["viola_l8"] = (
        (df["co_mg_km"] > L8_LIMITS["co_mg_km"]) |
        (df["nmog_nox_mg_km"] > L8_LIMITS["nmog_nox_mg_km"])
    )
    return df


def main() -> None:
    if not PDF_PATH.exists():
        print(f"PDF não encontrado: {PDF_PATH}")
        sys.exit(1)

    rows = extract_rows()
    if not rows:
        print("Nenhum veículo extraído.")
        sys.exit(1)

    df = pd.DataFrame(rows)
    df = clean(df)
    df = compute_compliance(df)
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")

    print(f"\nSalvo em {CSV_PATH}")
    print(f"Total: {len(df)} modelos/versões | {df['marca'].nunique()} marcas")
    print(f"Violam L8 (CO ou NMOG+NOx): {df['viola_l8'].sum()}")
    print(f"\nExemplo de violação:")
    print(df[df["viola_l8"]].head(3)[["marca", "modelo", "versao", "co_mg_km", "nmog_nox_mg_km"]].to_string())


if __name__ == "__main__":
    main()
