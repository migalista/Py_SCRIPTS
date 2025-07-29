# ========================================================
# SCRIPT: Atualizar dados de um arquivo Excel para CSV - -
# VERSÃO INTERATIVA COM INSTRUÇÕES DETALHADAS - - - - - -
# CRIADOR: MIGUEL MARIANO CABRERA - - - - - - - - - - - -
# ========================================================

input(str("Script Criado por Miguel Mariano Cabrera \nPressione Enter!"))
print('\n\n\n')


import pandas as pd
import os

print("==== ATUALIZAÇÃO DE DADOS DO EXCEL PARA CSV ====\n")
# ========================================================
# PASSO 1: Solicita ao usuário os caminhos dos arquivos
# ========================================================

print("Informe o caminho COMPLETO do arquivo Excel (.xlsx) que você deseja converter.")
print("Exemplo: C:\\Users\\seu_usuario\\Documents\\arquivo.xlsx")
excel_path = input("Digite o caminho do arquivo Excel: ").strip('"')

print("\nInforme o caminho COMPLETO onde o arquivo CSV deve ser salvo.")
print("Exemplo: C:\\Users\\seu_usuario\\Documents\\saida.csv")
csv_path = input("Digite o caminho de saída do arquivo CSV: ").strip('"')

# ========================================================
# PASSO 2: Verifica se o arquivo Excel existe
# ========================================================
if not os.path.exists(excel_path):
    print("\nERRO: O arquivo Excel informado não foi encontrado.")
    print("Verifique se o caminho está correto, com extensão .xlsx, e tente novamente.")
    exit()

# ========================================================
# PASSO 3: Remove o CSV antigo, se existir
# ========================================================
if os.path.exists(csv_path):
    os.remove(csv_path)
    print("Arquivo CSV antigo encontrado e removido.")

# ========================================================
# PASSO 4: Lê os dados do Excel
# ========================================================
try:
    df = pd.read_excel(excel_path)
    print("Arquivo Excel carregado com sucesso.")
except Exception as e:
    print("\nERRO ao tentar ler o arquivo Excel. Verifique se ele está fechado e se é um .xlsx válido.")
    print("Detalhes do erro:", e)
    exit()

# ========================================================
# PASSO 5: Salva o novo CSV
# ========================================================
try:
    df.to_csv(csv_path, index=False)
    print(f"\nNovo arquivo CSV gerado com sucesso em:\n{csv_path}")
except Exception as e:
    print("\nERRO ao salvar o arquivo CSV.")
    print("Detalhes do erro:", e)
    exit()

print("\nProcesso finalizado com sucesso.")