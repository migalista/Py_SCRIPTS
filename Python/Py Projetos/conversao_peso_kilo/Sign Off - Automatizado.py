import pandas as pd
import os
from datetime import datetime

# Configurações
NOME_ARQUIVO = 'Sign Off.xlsx'
CAMINHO_ALTERNATIVO = r'C:\Caminho\Alternativo\Sign Off.xlsx'  # <-- edite se necessário
COLUNA_IDH = 6     # Coluna G (índice 6)
COLUNA_UM = 7      # Coluna H (índice 7)
COLUNA_INICIO_VALORES = 8   # Coluna I (índice 8)
COLUNA_FIM_VALORES = 80     # Coluna CB (índice 80)
ABA_PADRAO = 'Standart Unit'
ABA_CONVERSAO = 'Conversão'

def localizar_arquivo():
    if os.path.exists(NOME_ARQUIVO):
        print(f"✔ Arquivo encontrado no diretório atual.")
        return NOME_ARQUIVO
    else:
        print("❌ Arquivo não encontrado no diretório atual.")
        if os.path.exists(CAMINHO_ALTERNATIVO):
            print(f"✔ Utilizando caminho alternativo: {CAMINHO_ALTERNATIVO}")
            return CAMINHO_ALTERNATIVO
        else:
            raise FileNotFoundError("Arquivo não encontrado nem no diretório atual nem no caminho alternativo.")

try:
    caminho_arquivo = localizar_arquivo()
    df_standart = pd.read_excel(caminho_arquivo, sheet_name=ABA_PADRAO)
    df_conversao = pd.read_excel(caminho_arquivo, sheet_name=ABA_CONVERSAO)

    # Verificar se as colunas A (IDH) e B (Peso Kg) existem na aba Conversão
    if df_conversao.shape[1] < 2:
        raise ValueError("A aba 'Conversão' precisa conter pelo menos duas colunas: IDH (coluna A) e Peso Kg (coluna B).")

    # Criar um dicionário de conversão a partir da aba Conversão
    conversao_dict = dict(zip(df_conversao.iloc[:, 0], df_conversao.iloc[:, 1]))

    # Filtrar apenas linhas com unidade diferente de "KG"
    df_filtrado = df_standart[df_standart.iloc[:, COLUNA_UM].str.upper() != "KG"]

    for index, row in df_filtrado.iterrows():
        idh = row[COLUNA_IDH]
        fator = conversao_dict.get(idh, None)

        if fator is None:
            continue  # pular se não encontrar o fator

        for col in range(COLUNA_INICIO_VALORES, COLUNA_FIM_VALORES + 1):
            valor = df_standart.iat[index, col]
            if pd.notnull(valor):
                df_standart.iat[index, col] = valor * fator

    # Criar novo nome com data
    data_execucao = datetime.now().strftime('%Y-%m-%d')
    novo_nome = f"Sign Off - {data_execucao}.xlsx"
    df_standart.to_excel(novo_nome, index=False, sheet_name=ABA_PADRAO)
    print(f"✔ Arquivo salvo com o nome: {novo_nome}")

except Exception as e:
    print(f"\n🚫 ERRO: {e}")
    print("O script encontrou um problema durante a execução.")
