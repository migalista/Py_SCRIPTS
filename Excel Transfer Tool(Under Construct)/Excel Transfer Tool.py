import pandas as pd
import os
import shutil
import string
print("===Esse script possui 145 linhas de programação===")
input(" ==== Esse Script foi criado por Miguel Mariano Cabrera ==== \n clique enter para continuar!")

input("\n É nescessario ponutar algumas coisas antes do uso do script\n\n 1° Os arquivos em Excel não podem estar abertos\n 2° O script não adiciona valores a arquivos Excel em branco!\n 3° Verifique todos os dados que você ira fornecer para o script!\n 4° Para facilitar o uso, COPIE E COLE em um BLOCO DE NOTAS as informações solicitadas no script (Exemplo Caminho raiz colunas e linhas que serão utilizadas)\n Aproveite!")


# Funções auxiliares
def mostrar_caminho_e_confirmar(caminho):
    print(f"\nVocê informou o caminho: {caminho}")
    confirmacao = input("Está correto? (s/n): ").strip().lower()
    return confirmacao == 's'

def salvar_como_csv(df, caminho_csv):
    df.to_csv(caminho_csv, index=False)
    print(f"Arquivo salvo como CSV em: {caminho_csv}")

def mover_arquivo(origem, destino):
    shutil.move(origem, destino)
    print(f"Arquivo movido para: {destino}")

def arte_sucesso():
    print("""
                                          sucesso
    """)

def arte_erro():
    print("""
                                         error
    """)

try:
    print("Exemplos de caminho raiz: C:/Pasta/Arquivo.xlsx ou C:/Users/Usuario/Documentos/Arquivo.xlsx\n")

    caminho_origem = input("Digite o caminho completo do ARQUIVO DE ORIGEM (.xlsx): ").strip()
    if not mostrar_caminho_e_confirmar(caminho_origem):
        raise Exception("Caminho da ORIGEM incorreto. Encerrando.")

    caminho_destino = input("Digite o caminho completo do ARQUIVO DE DESTINO (.xlsx): ").strip()
    if not mostrar_caminho_e_confirmar(caminho_destino):
        raise Exception("Caminho do DESTINO incorreto. Encerrando.")

    df_origem = pd.read_excel(caminho_origem)

    linhas, colunas = df_origem.shape
    print(f"\nA planilha de ORIGEM possui {linhas} linhas e {colunas} colunas.")
    print("\nCOLUNAS DISPONÍVEIS:")
    colunas_dict = {}
    for idx, col in enumerate(df_origem.columns):
        letra_coluna = string.ascii_uppercase[idx]
        colunas_dict[letra_coluna] = idx
        print(f"{letra_coluna}: {col}")

    print("""
Você pode usar as LETRAS acima para escolher as colunas (ex: da coluna A até E).
As LINHAS devem ser informadas em NÚMEROS (ex: da linha 1 até a 100).
""")

    modo_operacao = input("Você deseja SOBRESCREVER, ADICIONAR ou fazer AMBAS as operações? (sobrescrever/adicionar/ambos): ").strip().lower()
    opcao_transferencia = input("Deseja copiar a planilha INTEIRA ou transferir COLUNAS e LINHAS específicas? (inteira/parcial): ").strip().lower()

    if opcao_transferencia == 'inteira':
        df_corte = df_origem
    elif opcao_transferencia == 'parcial':
        multiplos_colagens = input("Deseja realizar MÚLTIPLAS COLAGENS em locais diferentes? (s/n): ").strip().lower() == 's'
        df_destino = pd.read_excel(caminho_destino)

        while True:
            linha_inicio = int(input(f"Informe a LINHA INICIAL da ORIGEM (1 a {linhas}): ")) - 1
            linha_fim = int(input(f"Informe a LINHA FINAL da ORIGEM (até {linhas}): "))
            coluna_inicio_letra = input("Informe a LETRA da COLUNA INICIAL da ORIGEM (ex: A): ").strip().upper()
            coluna_fim_letra = input("Informe a LETRA da COLUNA FINAL da ORIGEM (ex: E): ").strip().upper()

            if coluna_inicio_letra not in colunas_dict or coluna_fim_letra not in colunas_dict:
                raise Exception("COLUNAS inválidas. Encerrando.")

            coluna_inicio = colunas_dict[coluna_inicio_letra]
            coluna_fim = colunas_dict[coluna_fim_letra] + 1
            df_recorte = df_origem.iloc[linha_inicio:linha_fim, coluna_inicio:coluna_fim]

            if modo_operacao == 'adicionar':
                linha_destino = df_destino.dropna(how='all').shape[0]
            else:
                linha_destino = int(input("Informe a LINHA INICIAL de DESTINO (ex: 1): ")) - 1

            coluna_destino_letra = input("Informe a LETRA da COLUNA INICIAL de DESTINO (ex: A): ").strip().upper()
            coluna_destino = colunas_dict.get(coluna_destino_letra, None)
            if coluna_destino is None:
                raise Exception("COLUNA de DESTINO inválida. Encerrando.")

            linhas_necessarias = linha_destino + df_recorte.shape[0]
            while len(df_destino) < linhas_necessarias:
                df_destino.loc[len(df_destino)] = [None] * df_destino.shape[1]

            colunas_necessarias = coluna_destino + df_recorte.shape[1]
            if df_destino.shape[1] < colunas_necessarias:
                for i in range(colunas_necessarias - df_destino.shape[1]):
                    df_destino[f'NovaColuna_{i}'] = None

            for i in range(df_recorte.shape[0]):
                for j in range(df_recorte.shape[1]):
                    df_destino.iat[linha_destino + i, coluna_destino + j] = df_recorte.iat[i, j]

            if not multiplos_colagens:
                break
            continuar = input("Deseja fazer OUTRA COLAGEM? (s/n): ").strip().lower()
            if continuar != 's':
                break

        df_corte = df_destino
    else:
        raise Exception("Opção inválida. Encerrando.")

    with pd.ExcelWriter(caminho_destino, engine='openpyxl', mode='w') as writer:
        df_corte.to_excel(writer, index=False)

    print("Arquivo atualizado com sucesso!")

    salvar_csv = input("Deseja SALVAR o arquivo também como CSV? (s/n): ").strip().lower()
    if salvar_csv == 's':
        caminho_csv = caminho_destino.replace('.xlsx', '.csv')
        salvar_como_csv(df_corte, caminho_csv)

    acao_final = input("Deseja RENOMEAR, MOVER, ou AMBOS o arquivo? (renomear/mover/ambos/nao): ").strip().lower()

    if acao_final in ['renomear', 'ambos']:
        novo_nome = input("Digite o NOVO NOME para o arquivo (sem caminho, exemplo: novo_arquivo.xlsx): ").strip()
        novo_caminho = os.path.join(os.path.dirname(caminho_destino), novo_nome)
        os.rename(caminho_destino, novo_caminho)
        caminho_destino = novo_caminho
        print(f"Arquivo renomeado para: {novo_nome}")

    if acao_final in ['mover', 'ambos']:
        nova_pasta = input("Digite o NOVO CAMINHO da pasta onde deseja mover o arquivo: ").strip()
        novo_caminho_final = os.path.join(nova_pasta, os.path.basename(caminho_destino))
        mover_arquivo(caminho_destino, novo_caminho_final)

    arte_sucesso()

except Exception as e:
    print(f"Erro: {e}")
    arte_erro()