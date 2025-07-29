import pandas as pd
import os
import shutil
import string

def exibir_e_confirmar_caminho(caminho):
    """Exibe um caminho para o usuário e pede confirmação."""
    print(f"\nVocê informou o caminho: {caminho}")
    confirmacao = input("Está correto? (s/n): ").strip().lower()
    return confirmacao == 's'

def salvar_dataframe_como_csv(df, caminho_csv):
    """Salva um DataFrame em um arquivo CSV."""
    try:
        df.to_csv(caminho_csv, index=False)
        print(f"Arquivo salvo como CSV em: {caminho_csv}")
    except Exception as e:
        print(f"Erro ao salvar como CSV: {e}")

def mover_arquivo(origem, destino):
    """Move um arquivo da origem para o destino."""
    try:
        shutil.move(origem, destino)
        print(f"Arquivo movido para: {destino}")
    except Exception as e:
        print(f"Erro ao mover arquivo: {e}")

def imprimir_arte_ascii(tipo_arte):
    """Imprime arte ASCII para sucesso ou erro."""
    if tipo_arte == "sucesso":
        print("\n" + "="*40)
        print("                 SUCESSO                ")
        print("="*40 + "\n")
    elif tipo_arte == "erro":
        print("\n" + "!"*40)
        print("                   ERRO                 ")
        print("!"*40 + "\n")

def obter_letras_colunas_excel(num_colunas):
    """Gera um mapeamento das letras das colunas do Excel para seus índices inteiros."""
    letras_colunas = {}
    for i in range(num_colunas):
        if i < 26:
            letras_colunas[string.ascii_uppercase[i]] = i
        else:
            # Para colunas além de Z (ex: AA, AB, etc.)
            primeiro_caractere_codigo = (i // 26) - 1
            segundo_caractere_codigo = i % 26
            letras_colunas[string.ascii_uppercase[primeiro_caractere_codigo] + string.ascii_uppercase[segundo_caractere_codigo]] = i
    return letras_colunas

def principal():
    """Função principal para executar o script de manipulação de Excel."""
    print("="*40)
    print("      Script de Manipulação de Planilhas      ")
    print("="*40)
    print("Criado por Miguel Mariano Cabrera")
    input("Pressione Enter para continuar...")

    print("\n" + "-"*40)
    print("       ATENÇÃO: Instruções Importantes!      ")
    print("-"*40)
    print("1. Os arquivos Excel não podem estar abertos.")
    print("2. Este script não adiciona valores a arquivos Excel em branco.")
    print("3. Verifique todos os dados fornecidos com atenção.")
    print("4. Para facilitar, copie e cole as informações solicitadas (caminhos, colunas, linhas) em um bloco de notas.")
    input("Aproveite! Pressione Enter para começar...")

    try:
        print("\nExemplos de caminho raiz: C:/Pasta/Arquivo.xlsx ou C:/Users/Usuario/Documentos/Arquivo.xlsx\n")

        # Obter e validar caminho de origem
        caminho_origem = input("Digite o caminho completo do ARQUIVO DE ORIGEM (.xlsx): ").strip()
        if not exibir_e_confirmar_caminho(caminho_origem):
            raise ValueError("Caminho da ORIGEM incorreto. Encerrando.")
        if not os.path.exists(caminho_origem):
            raise FileNotFoundError(f"Arquivo de origem não encontrado: {caminho_origem}")

        # Obter e validar caminho de destino
        caminho_destino = input("Digite o caminho completo do ARQUIVO DE DESTINO (.xlsx): ").strip()
        if not exibir_e_confirmar_caminho(caminho_destino):
            raise ValueError("Caminho do DESTINO incorreto. Encerrando.")

        # Carregar DataFrame de origem
        try:
            df_origem = pd.read_excel(caminho_origem)
        except Exception as e:
            raise Exception(f"Erro ao carregar o arquivo de origem: {e}")

        linhas_origem, colunas_origem = df_origem.shape
        print(f"\nA planilha de ORIGEM possui {linhas_origem} linhas e {colunas_origem} colunas.")

        mapa_colunas = obter_letras_colunas_excel(colunas_origem)
        print("\nCOLUNAS DISPONÍVEIS NA ORIGEM:")
        for letra_coluna, indice in mapa_colunas.items():
            if indice < len(df_origem.columns): # Garante que não exceda os limites para df_origem.columns
                print(f"{letra_coluna}: {df_origem.columns[indice]}")
            else:
                break # Para se exceder as colunas reais disponíveis

        print("\nVocê pode usar as LETRAS acima para escolher as colunas (ex: da coluna A até E).")
        print("As LINHAS devem ser informadas em NÚMEROS (ex: da linha 1 até a 100).")

        opcoes_modo = {'sobrescrever', 'adicionar', 'ambos'}
        modo_operacao = input("Você deseja SOBRESCREVER, ADICIONAR ou fazer AMBAS as operações? (sobrescrever/adicionar/ambos): ").strip().lower()
        if modo_operacao not in opcoes_modo:
            raise ValueError("Opção de operação inválida. Escolha 'sobrescrever', 'adicionar' ou 'ambos'.")

        opcoes_transferencia = {'inteira', 'parcial'}
        opcao_transferencia = input("Deseja copiar a planilha INTEIRA ou transferir COLUNAS e LINHAS específicas? (inteira/parcial): ").strip().lower()
        if opcao_transferencia not in opcoes_transferencia:
            raise ValueError("Opção de transferência inválida. Escolha 'inteira' ou 'parcial'.")

        df_final = None # Este será o DataFrame a ser gravado no destino

        if opcao_transferencia == 'inteira':
            df_final = df_origem.copy() # Usa .copy() para evitar modificar o df_origem original
        elif opcao_transferencia == 'parcial':
            df_destino = None
            if os.path.exists(caminho_destino):
                try:
                    df_destino = pd.read_excel(caminho_destino)
                except Exception as e:
                    raise Exception(f"Erro ao carregar o arquivo de destino: {e}")
            else:
                print(f"Arquivo de destino '{caminho_destino}' não encontrado. Criando um novo DataFrame para o destino.")
                df_destino = pd.DataFrame() # Começa com um DataFrame vazio se o destino não existir

            multiplas_colagens = input("Deseja realizar MÚLTIPLAS COLAGENS em locais diferentes? (s/n): ").strip().lower() == 's'

            while True:
                try:
                    linha_inicio_origem = int(input(f"Informe a LINHA INICIAL da ORIGEM (1 a {linhas_origem}): ")) - 1
                    linha_fim_origem = int(input(f"Informe a LINHA FINAL da ORIGEM (até {linhas_origem}): "))

                    if not (0 <= linha_inicio_origem < linha_fim_origem <= linhas_origem):
                        raise ValueError("Intervalo de linhas da ORIGEM inválido.")

                    letra_coluna_inicio_origem = input("Informe a LETRA da COLUNA INICIAL da ORIGEM (ex: A): ").strip().upper()
                    letra_coluna_fim_origem = input("Informe a LETRA da COLUNA FINAL da ORIGEM (ex: E): ").strip().upper()

                    coluna_inicio_origem = mapa_colunas.get(letra_coluna_inicio_origem)
                    coluna_fim_origem = mapa_colunas.get(letra_coluna_fim_origem)

                    if coluna_inicio_origem is None or coluna_fim_origem is None or coluna_inicio_origem > coluna_fim_origem:
                        raise ValueError("Intervalo de COLUNAS da ORIGEM inválido.")

                    df_recorte = df_origem.iloc[linha_inicio_origem:linha_fim_origem, coluna_inicio_origem:coluna_fim_origem + 1]

                    if modo_operacao == 'adicionar':
                        # Encontra a primeira linha vazia no destino
                        linha_destino = df_destino.shape[0] if df_destino.empty else df_destino.dropna(how='all').shape[0]
                    else:
                        linha_destino = int(input("Informe a LINHA INICIAL de DESTINO (ex: 1): ")) - 1
                        if linha_destino < 0:
                            raise ValueError("A linha de destino não pode ser negativa.")

                    letra_coluna_destino = input("Informe a LETRA da COLUNA INICIAL de DESTINO (ex: A): ").strip().upper()
                    coluna_destino = mapa_colunas.get(letra_coluna_destino)
                    if coluna_destino is None:
                        raise ValueError("COLUNA de DESTINO inválida.")

                    # Garante que o DataFrame de destino tenha linhas e colunas suficientes
                    linhas_necessarias = linha_destino + df_recorte.shape[0]
                    if df_destino.shape[0] < linhas_necessarias:
                        # Preenche com linhas vazias
                        for _ in range(linhas_necessarias - df_destino.shape[0]):
                            df_destino.loc[len(df_destino)] = [None] * df_destino.shape[1]

                    colunas_necessarias = coluna_destino + df_recorte.shape[1]
                    if df_destino.shape[1] < colunas_necessarias:
                        # Adiciona novas colunas se necessário, nomeando-as 'NovaColuna_X'
                        for i in range(colunas_necessarias - df_destino.shape[1]):
                            df_destino[f'NovaColuna_{df_destino.shape[1]}'] = None

                    # Cola os dados
                    for idx_linha in range(df_recorte.shape[0]):
                        for idx_coluna in range(df_recorte.shape[1]):
                            df_destino.iat[linha_destino + idx_linha, coluna_destino + idx_coluna] = df_recorte.iat[idx_linha, idx_coluna]

                    if not multiplas_colagens:
                        break
                    
                    continuar = input("Deseja fazer OUTRA COLAGEM? (s/n): ").strip().lower()
                    if continuar != 's':
                        break

                except ValueError as ve:
                    print(f"Erro de entrada: {ve}. Por favor, tente novamente com valores válidos.")
                except Exception as e:
                    print(f"Erro inesperado durante a colagem: {e}")
                    break # Sai do loop em erros inesperados

            df_final = df_destino

        # Salvar o arquivo Excel atualizado
        if df_final is not None:
            try:
                # Se 'sobrescrever' ou 'ambos', escrevemos no modo 'w'.
                # Se 'adicionar' com parcial, df_final já contém as adições.
                with pd.ExcelWriter(caminho_destino, engine='openpyxl', mode='w') as writer:
                    df_final.to_excel(writer, index=False)
                print("Arquivo atualizado com sucesso!")
            except Exception as e:
                raise Exception(f"Erro ao salvar o arquivo Excel: {e}")
        else:
            print("Nenhuma operação de transferência foi realizada, o arquivo de destino não será salvo.")

        # Salvamento opcional em CSV
        salvar_csv = input("Deseja SALVAR o arquivo também como CSV? (s/n): ").strip().lower()
        if salvar_csv == 's':
            csv_path = caminho_destino.replace('.xlsx', '.csv')
            salvar_dataframe_como_csv(df_final, csv_path)

        # Ações finais: renomear/mover
        opcoes_acao_final = {'renomear', 'mover', 'ambos', 'nao'}
        acao_final = input("Deseja RENOMEAR, MOVER, ou AMBOS o arquivo? (renomear/mover/ambos/nao): ").strip().lower()
        if acao_final not in opcoes_acao_final:
            print("Opção de ação final inválida. Nenhuma ação será realizada.")

        if acao_final in ['renomear', 'ambos']:
            try:
                novo_nome = input("Digite o NOVO NOME para o arquivo (sem caminho, exemplo: novo_arquivo.xlsx): ").strip()
                if not novo_nome.endswith('.xlsx'):
                    novo_nome += '.xlsx' # Garante que mantenha a extensão .xlsx
                
                # Constrói o novo caminho no mesmo diretório do destino atual
                novo_caminho = os.path.join(os.path.dirname(caminho_destino), novo_nome)
                
                if os.path.exists(novo_caminho) and novo_caminho != caminho_destino:
                    sobrescrever = input(f"O arquivo '{os.path.basename(novo_caminho)}' já existe. Deseja sobrescrevê-lo? (s/n): ").strip().lower()
                    if sobrescrever != 's':
                        print("Renomeação cancelada.")
                    else:
                        os.remove(novo_caminho) # Remove o arquivo existente antes de renomear
                        os.rename(caminho_destino, novo_caminho)
                        caminho_destino = novo_caminho # Atualiza caminho_destino para operações subsequentes
                        print(f"Arquivo renomeado para: {novo_nome}")
                else:
                    os.rename(caminho_destino, novo_caminho)
                    caminho_destino = novo_caminho # Atualiza caminho_destino para operações subsequentes
                    print(f"Arquivo renomeado para: {novo_nome}")

            except OSError as oe:
                print(f"Erro ao renomear o arquivo: {oe}. Certifique-se de que o arquivo não está em uso e o nome é válido.")
            except Exception as e:
                print(f"Erro inesperado durante a renomeação: {e}")

        if acao_final in ['mover', 'ambos']:
            try:
                nova_pasta = input("Digite o NOVO CAMINHO da pasta onde deseja mover o arquivo: ").strip()
                if not os.path.isdir(nova_pasta):
                    os.makedirs(nova_pasta, exist_ok=True) # Cria o diretório se ele não existir
                    print(f"Pasta '{nova_pasta}' criada.")

                novo_caminho_final = os.path.join(nova_pasta, os.path.basename(caminho_destino))
                mover_arquivo(caminho_destino, novo_caminho_final)
                caminho_destino = novo_caminho_final # Atualiza para confirmação final
            except OSError as oe:
                print(f"Erro ao mover o arquivo: {oe}. Verifique se o caminho de destino é válido e você tem permissão.")
            except Exception as e:
                print(f"Erro inesperado durante a movimentação: {e}")

        imprimir_arte_ascii("sucesso")

    except (ValueError, FileNotFoundError, Exception) as e:
        print(f"Erro: {e}")
        imprimir_arte_ascii("erro")

if __name__ == "__main__":
    principal()