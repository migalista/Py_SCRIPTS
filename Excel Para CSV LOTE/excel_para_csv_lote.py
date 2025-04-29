# _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _  _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _

import os
import pandas as pd

# Mensagem de autoria
print("Esse script foi criado por Miguel Mariano Cabrera\n")

# Obtem o caminho da pasta onde o script está
pasta_atual = os.path.dirname(os.path.abspath(__file__))

# Lista todos os arquivos da pasta
arquivos = os.listdir(pasta_atual)

# Filtra apenas os arquivos .xlsx (ignora temporários e .csvs)
arquivos_excel = [arquivo for arquivo in arquivos if arquivo.endswith(".xlsx") and not arquivo.startswith('~$')]

# Verifica se há arquivos para processar
if not arquivos_excel:
    print("Nenhum arquivo .xlsx encontrado na pasta.")
else:
    for arquivo in arquivos_excel:
        caminho_excel = os.path.join(pasta_atual, arquivo)
        nome_base = os.path.splitext(arquivo)[0]
        caminho_csv = os.path.join(pasta_atual, nome_base + ".csv")

        try:
            df = pd.read_excel(caminho_excel, engine="openpyxl")
            df.to_csv(caminho_csv, index=False)
            print(f"Convertido: {arquivo} → {nome_base}.csv")
        except Exception as e:
            print(f"Erro ao converter {arquivo}: {e}")
