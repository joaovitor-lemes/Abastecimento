import os
import glob
from datetime import datetime
import pandas as pd

# Tolerâncias por modelo
TOLERANCIAS = {
    "VW 24.280 Mec": {"km": 300, "horimetro": 48},
    "Atego 2426": {"km": 400, "horimetro": 48},
    "Atego 1719": {"km": 400, "horimetro": 48},
    "TECTOR 24-280": {"km": 300, "horimetro": 48},
    "Tector 24-320": {"km": 200, "horimetro": 48},
    "Atego 2433": {"km": 400, "horimetro": 48},
    "Atego 2429": {"km": 400, "horimetro": 48},
    "VW 17.260 Mec.": {"km": 300, "horimetro": 48},
    "VW 18.260": {"km": 400, "horimetro": 48},
    "VW 24.260 Aut.": {"km": 400, "horimetro": 48},
    "VW 24.260": {"km": 400, "horimetro": 48},
    "VW 26.260": {"km": 400, "horimetro": 48},
    "636D": {"km": 400, "horimetro": 48},
    "416.0": {"km": 400, "horimetro": 48},
    "VW 11.180": {"km": 400, "horimetro": 48},
    "VW 17.210": {"km": 400, "horimetro": 48}
}

TOLERANCIA_PADRAO = {"km": 300, "horimetro": 10}

def encontrar_arquivo_txt():
    arquivos_txt = glob.glob("*.txt")
    if not arquivos_txt:
        print("Nenhum arquivo .txt encontrado no diretório.")
        return None
    print(f"📄 Arquivo encontrado: {arquivos_txt[0]}")
    return arquivos_txt[0]

def analisar_abastecimentos(caminho_arquivo):
    df = pd.read_csv(caminho_arquivo, sep=";", encoding="utf-8")

    df["Data/Hora"] = pd.to_datetime(df["Data/Hora"], errors="coerce", dayfirst=True)
    df = df.dropna(subset=["Data/Hora"])
    df.sort_values(by=["Placa", "Data/Hora"], inplace=True)
    df_ordenado = df.copy()

    erros = []
    total_linhas = len(df)

    for placa, grupo in df.groupby("Placa"):
        grupo = grupo.reset_index(drop=True)

        for i in range(1, len(grupo)):
            prev = grupo.loc[i - 1]
            atual = grupo.loc[i]

            km_atual = atual["Odometro"]
            km_anterior = prev["Odometro"]
            hr_atual = atual["Horimetro"]
            hr_anterior = prev["Horimetro"]
            modelo = atual["Modelo"]

            tolerancia = TOLERANCIAS.get(modelo, TOLERANCIA_PADRAO)
            tol_km = tolerancia["km"]
            tol_hr = tolerancia["horimetro"]

            if km_atual < km_anterior:
                erros.append({**atual.to_dict(), "Erro": "KM abaixo do anterior"})
            elif km_atual - km_anterior > tol_km:
                erros.append({**atual.to_dict(), "Erro": f"KM acima da tolerância de {tol_km} km para o modelo {modelo}"})

            if hr_atual < hr_anterior:
                erros.append({**atual.to_dict(), "Erro": "Horímetro abaixo do anterior"})
            elif hr_atual - hr_anterior > tol_hr:
                erros.append({**atual.to_dict(), "Erro": f"Horímetro acima da tolerância de {tol_hr} h para o modelo {modelo}"})

    nome_arquivo = "analise_abastecimentos.xlsx"
    with pd.ExcelWriter(nome_arquivo, engine="openpyxl") as writer:
        df_ordenado.to_excel(writer, sheet_name="Ordenado", index=False)
        if erros:
            df_erros = pd.DataFrame(erros)
            df_erros.to_excel(writer, sheet_name="Erros", index=False)
            print(f"\n Arquivo '{nome_arquivo}' gerado com aba de erros ({len(erros)} linhas).")
        else:
            print(f"\n Arquivo '{nome_arquivo}' gerado sem erros detectados.")

    print(f"\n Total de linhas processadas: {total_linhas}")
    print(f" Total de linhas com erro: {len(erros)}")

# Executa o script
arquivo = encontrar_arquivo_txt()
if arquivo:
    analisar_abastecimentos(arquivo)

