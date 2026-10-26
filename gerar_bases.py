"""Gera as bases do dashboard SmartAgri Analytics (atividade_pratica_4).

Saida: Producao_Agro e Manutencao_Agro (.xlsx e .csv)
"""

from datetime import date, timedelta
from pathlib import Path
import random

import pandas as pd

BASE = Path(__file__).resolve().parent
SEED = 42

# ID_Talhao, Cultura, Area_ha, Produtividade_base_t_ha, 1a data de colheita
TALHOES = [
    ("Talhão 01", "Soja", 92.4, 3.62, date(2026, 2, 10)),
    ("Talhão 02", "Soja", 76.0, 3.41, date(2026, 2, 14)),
    ("Talhão 03", "Soja", 118.5, 3.78, date(2026, 2, 22)),
    ("Talhão 04", "Milho", 64.2, 9.35, date(2026, 6, 24)),
    ("Talhão 05", "Milho", 88.7, 8.94, date(2026, 7, 1)),
    ("Talhão 06", "Milho", 45.0, 10.12, date(2026, 7, 9)),
    ("Talhão 07", "Algodão", 53.6, 4.48, date(2026, 9, 16)),
    ("Talhão 08", "Algodão", 71.3, 4.72, date(2026, 9, 28)),
]

LOTE_PESOS = [0.27, 0.24, 0.26, 0.23]
LOTE_INTERVALO = [0, 8, 16, 25]

# ID_Equipamento: (custo por hora de parada R$, horas base por evento)
EQUIPAMENTOS = {
    "Trator 01": 320.0,
    "Trator 02": 320.0,
    "Colheitadeira A": 450.0,
    "Colheitadeira B": 450.0,
    "Pulverizador B": 380.0,
    "Semeadora C": 300.0,
    "Distribuidor de Insumos D": 340.0,
    "Grade Niveladora E": 260.0,
}

# (equipamento, data, cultura atendida, horas de operacao)
MANUTENCOES = [
    ("Semeadora C", date(2026, 3, 5), "Milho", 168.0),
    ("Trator 02", date(2026, 3, 12), "Milho", 142.5),
    ("Colheitadeira A", date(2026, 2, 18), "Soja", 214.0),
    ("Colheitadeira B", date(2026, 2, 26), "Soja", 198.5),
    ("Trator 01", date(2026, 2, 20), "Soja", 176.0),
    ("Trator 02", date(2026, 3, 3), "Soja", 154.0),
    ("Distribuidor de Insumos D", date(2026, 3, 18), "Soja", 132.5),
    ("Colheitadeira A", date(2026, 3, 8), "Soja", 96.0),
    ("Trator 01", date(2026, 4, 15), "Soja", 188.0),
    ("Grade Niveladora E", date(2026, 4, 22), "Soja", 146.5),
    ("Trator 02", date(2026, 5, 6), "Soja", 162.0),
    ("Grade Niveladora E", date(2026, 5, 14), "Milho", 138.0),
    ("Semeadora C", date(2026, 4, 8), "Algodão", 152.5),
    ("Pulverizador B", date(2026, 5, 20), "Algodão", 174.0),
    ("Trator 01", date(2026, 5, 27), "Algodão", 128.5),
    ("Pulverizador B", date(2026, 6, 11), "Algodão", 186.0),
    ("Semeadora C", date(2026, 6, 19), "Milho", 143.5),
    ("Trator 02", date(2026, 6, 26), "Milho", 171.0),
    ("Colheitadeira A", date(2026, 7, 3), "Milho", 205.5),
    ("Distribuidor de Insumos D", date(2026, 7, 10), "Milho", 158.0),
    ("Trator 01", date(2026, 7, 17), "Milho", 182.5),
    ("Colheitadeira A", date(2026, 7, 24), "Milho", 88.0),
    ("Pulverizador B", date(2026, 8, 5), "Algodão", 192.0),
    ("Trator 02", date(2026, 8, 12), "Algodão", 166.5),
    ("Distribuidor de Insumos D", date(2026, 8, 20), "Algodão", 149.0),
    ("Pulverizador B", date(2026, 9, 2), "Algodão", 124.5),
    ("Colheitadeira B", date(2026, 9, 18), "Algodão", 212.0),
    ("Trator 02", date(2026, 9, 24), "Algodão", 178.5),
    ("Colheitadeira B", date(2026, 10, 2), "Algodão", 94.0),
    ("Distribuidor de Insumos D", date(2026, 10, 7), "Algodão", 136.0),
    ("Semeadora C", date(2026, 9, 30), "Soja", 164.0),
    ("Trator 01", date(2026, 10, 6), "Soja", 194.5),
    ("Pulverizador B", date(2026, 10, 13), "Soja", 157.0),
    ("Grade Niveladora E", date(2026, 10, 9), "Soja", 141.5),
    ("Semeadora C", date(2026, 10, 16), "Soja", 72.0),
    ("Trator 02", date(2026, 10, 20), "Soja", 108.5),
]


def gerar_producao(rng: random.Random) -> pd.DataFrame:
    linhas = []
    for id_talhao, cultura, area, prod_base, primeira_data in TALHOES:
        pesos = LOTE_PESOS[:]
        rng.shuffle(pesos)
        for i, peso in enumerate(pesos):
            prod = prod_base * rng.uniform(0.93, 1.07)
            area_lote = area * peso
            volume = prod * area_lote
            linhas.append(
                {
                    "ID_Talhao": id_talhao,
                    "Cultura": cultura,
                    "Produtividade_Sectares": round(prod, 2),
                    "Volume_Colhido_Ton": round(volume, 1),
                    "Data_Colheita": primeira_data + timedelta(days=LOTE_INTERVALO[i]),
                }
            )
    df = pd.DataFrame(linhas)
    df["Data_Colheita"] = pd.to_datetime(df["Data_Colheita"])
    return df.sort_values("Data_Colheita", kind="stable").reset_index(drop=True)


def gerar_manutencao(rng: random.Random) -> pd.DataFrame:
    linhas = []
    for equipamento, data, cultura, horas in MANUTENCOES:
        taxa = EQUIPAMENTOS[equipamento]
        parada = rng.uniform(1.8, 14.0)
        if horas > 200:
            parada *= 1.15
        parada = round(min(parada, horas * 0.16), 1)
        custo = round(parada * taxa + rng.uniform(150.0, 900.0), 2)
        linhas.append(
            {
                "ID_Equipamento": equipamento,
                "Horas_Operacao": round(horas, 1),
                "Tempo_Parada_Manutencao": parada,
                "Custo_Manutencao": custo,
                "Cultura": cultura,
                "Data_Manutencao": data,
            }
        )
    df = pd.DataFrame(linhas)
    df["Data_Manutencao"] = pd.to_datetime(df["Data_Manutencao"])
    return df.sort_values("Data_Manutencao", kind="stable").reset_index(drop=True)


def validar(prod: pd.DataFrame, manut: pd.DataFrame) -> None:
    assert not prod.isna().any().any(), "nulo em Producao_Agro"
    assert not manut.isna().any().any(), "nulo em Manutencao_Agro"
    assert (prod["Volume_Colhido_Ton"] > 0).all(), "volume <= 0"
    assert (prod["Produtividade_Sectares"] > 0).all(), "produtividade <= 0"
    assert (manut["Tempo_Parada_Manutencao"] > 0).all(), "parada <= 0"
    assert (manut["Tempo_Parada_Manutencao"] < manut["Horas_Operacao"]).all()
    assert (manut["Custo_Manutencao"] > 0).all(), "custo <= 0"
    assert set(prod["Cultura"]) == set(manut["Cultura"]), "culturas divergentes"


def salvar(df: pd.DataFrame, nome: str) -> None:
    df.to_csv(BASE / f"{nome}.csv", index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(BASE / f"{nome}.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nome, startrow=1)
        ws = writer.sheets[nome]
        ws["A1"] = f"SmartAgri Analytics - {nome}"
        for col in df.columns:
            idx = df.columns.get_loc(col) + 1
            maior = max(len(str(col)), df[col].astype(str).str.len().max())
            ws.column_dimensions[ws.cell(row=2, column=idx).column_letter].width = min(maior + 2, 40)
        for campo, formato in (("Data_Colheita", "DD/MM/YYYY"), ("Data_Manutencao", "DD/MM/YYYY")):
            if campo in df.columns:
                i = df.columns.get_loc(campo) + 1
                for r in range(3, 3 + len(df)):
                    ws.cell(row=r, column=i).number_format = formato


def main() -> None:
    rng = random.Random(SEED)
    prod = gerar_producao(rng)
    manut = gerar_manutencao(rng)
    validar(prod, manut)
    salvar(prod, "Producao_Agro")
    salvar(manut, "Manutencao_Agro")
    print(f"Producao_Agro: {len(prod)} linhas | {prod['ID_Talhao'].nunique()} talhoes | "
          f"{prod['Volume_Colhido_Ton'].sum():,.1f} ton")
    print(f"Manutencao_Agro: {len(manut)} linhas | {manut['ID_Equipamento'].nunique()} equipamentos | "
          f"{manut['Custo_Manutencao'].sum():,.2f} BRL | {manut['Tempo_Parada_Manutencao'].sum():,.1f} h parada")
    print(f"Produtividade media: {prod['Produtividade_Sectares'].mean():.2f} t/ha")


if __name__ == "__main__":
    main()
