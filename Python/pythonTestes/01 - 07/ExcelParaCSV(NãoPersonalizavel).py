# ========================================================
# SCRIPT: Atualizar dados de um arquivo Excel para CSV para adições ou atualizações de base no Power Bi
# AUTOR: Miguel Mariano Cabrera
# OBJETIVO: , eLer um Excelxcluir o CSV antigo (se existir),
#           e gerar um novo CSV atualizado para uso no Power BI.
# ========================================================

import pandas as pd
import os

# ========================================================
# REGRAS PARA FUNCIONAR EM QUALQUER COMPUTADOR:
# 
# 1. É necessário ter o Python instalado (recomenda-se versão 3.10+).
# 2. Instalar a biblioteca pandas: pip install pandas
# 3. Instalar a biblioteca openpyxl (necessária para ler .xlsx):
#    pip install openpyxl
#
# 4. Os caminhos dos arquivos abaixo DEVEM existir no computador.
#    - Mude os caminhos se o local dos arquivos for diferente.
# ========================================================

# Caminho do arquivo Excel de entrada
excel_path = r"C:\Pasta Main\Excel\testes\Teste1.xlsx"

# Caminho do arquivo CSV de saída
csv_path = r"C:\Pasta Main\Excel\testes\saida_teste1.csv"

# ========================================================
# PASSO 1: Verifica se o CSV já existe. Se sim, apaga.
# Isso evita que dados antigos fiquem salvos por engano.
# ========================================================
if os.path.exists(csv_path):
    os.remove(csv_path)

# ========================================================
# PASSO 2: Lê os dados do Excel (.xlsx) atualizado
# ========================================================
df = pd.read_excel(excel_path)

# ========================================================
# PASSO 3: Salva os dados em um novo arquivo CSV
# index=False -> não salva o número da linha no CSV
# ========================================================
df.to_csv(csv_path, index=False)

# ========================================================
# PASSO 4: Mensagem final
# ========================================================
print("CSV atualizado com sucesso!")
