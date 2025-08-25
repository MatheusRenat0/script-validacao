import pandas as pd
import numpy as np

# --- 1. CONFIGURAÇÃO ---
# Altere os nomes dos arquivos e das colunas conforme necessário.

# Caminho para os arquivos Excel
caminho_planilha1 = 'planilha1.xlsx'  # <-- COLOQUE O NOME DO SEU PRIMEIRO ARQUIVO AQUI
caminho_planilha2 = 'planilha2.xlsx'  # <-- COLOQUE O NOME DO SEU SEGUNDO ARQUIVO AQUI

# Mapeamento das colunas que serão comparadas
# A chave (à esquerda) é o nome padrão que usaremos no script.
# O valor (à direita) é o nome EXATO da coluna no seu arquivo Excel.
mapa_colunas_planilha1 = {
    'RE': 'RE do funcionario', # <-- Altere 'RE do funcionario' se o nome for diferente
    'Nome': 'nome do funcionario',
    'IMEI': 'Imei do aparelho',
    'Linha': 'Linha',
    'ICCID': 'ICCID'
}

mapa_colunas_planilha2 = {
    'RE': 'RE', # <-- Exemplo: na planilha 2, a coluna se chama apenas 'RE'
    'Nome': 'Funcionário',
    'IMEI': 'IMEI',
    'Linha': 'Número da Linha',
    'ICCID': 'ICCID do Chip'
}

# Coluna que será usada como identificador único para cruzar as informações.
coluna_chave = 'RE'

# --- FIM DA CONFIGURAÇÃO ---


def carregar_e_padronizar(caminho_arquivo, mapa_colunas):
    """Carrega uma planilha, seleciona e renomeia as colunas de interesse."""
    try:
        # Define o tipo de todas as colunas como string ao carregar para evitar problemas
        # de formatação (ex: RE 123 lido como 123.0, ou ICCID perdendo zeros à esquerda)
        df = pd.read_excel(caminho_arquivo, dtype=str)
        
        colunas_reais = list(mapa_colunas.values())
        df = df[colunas_reais]

        mapa_inverso = {v: k for k, v in mapa_colunas.items()}
        df = df.rename(columns=mapa_inverso)

        df = df.dropna(subset=[coluna_chave])
        return df

    except FileNotFoundError:
        print(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        return None
    except KeyError as e:
        print(f"ERRO: A coluna {e} não foi encontrada no arquivo '{caminho_arquivo}'. Verifique o mapeamento.")
        return None

# --- 2. PROCESSAMENTO ---

print("Iniciando a comparação...")
# Configuração para que o pandas mostre todas as colunas no terminal
pd.set_option('display.max.columns', None)
pd.set_option('display.width', 1000)

df1 = carregar_e_padronizar(caminho_planilha1, mapa_colunas_planilha1)
df2 = carregar_e_padronizar(caminho_planilha2, mapa_colunas_planilha2)

if df1 is None or df2 is None:
    print("Processo interrompido devido a erro no carregamento dos arquivos.")
else:
    df_comparacao = pd.merge(
        df1,
        df2,
        on=coluna_chave,
        how='outer',
        suffixes=('_p1', '_p2')
    )

    # --- 3. ANÁLISE E CLASSIFICAÇÃO ---
    
    colunas_comparacao = [col for col in mapa_colunas_planilha1.keys() if col != coluna_chave]

    # Categoria 1: Apenas na Planilha 1
    apenas_em_p1 = df_comparacao[df_comparacao['Nome_p2'].isnull()].copy()
    apenas_em_p1 = apenas_em_p1[[coluna_chave] + [f"{col}_p1" for col in colunas_comparacao]]
    apenas_em_p1.columns = [coluna_chave] + colunas_comparacao

    # Categoria 2: Apenas na Planilha 2
    apenas_em_p2 = df_comparacao[df_comparacao['Nome_p1'].isnull()].copy()
    apenas_em_p2 = apenas_em_p2[[coluna_chave] + [f"{col}_p2" for col in colunas_comparacao]]
    apenas_em_p2.columns = [coluna_chave] + colunas_comparacao
    
    # Categoria 3 e 4: Itens em comum
    comuns = df_comparacao.dropna(subset=['Nome_p1', 'Nome_p2']).copy()
    
    if not comuns.empty:
        for col in colunas_comparacao:
            comuns[f'divergencia_{col}'] = np.where(comuns[f'{col}_p1'] != comuns[f'{col}_p2'], 'Sim', 'Não')
        
        comuns['tem_divergencia'] = np.where(
            (comuns[[f'divergencia_{col}' for col in colunas_comparacao]] == 'Sim').any(axis=1), 'Sim', 'Não'
        )
        
        comuns_com_divergencia = comuns[comuns['tem_divergencia'] == 'Sim']
        comuns_sem_divergencia = comuns[comuns['tem_divergencia'] == 'Não']
    else:
        # Se não há itens em comum, cria DataFrames vazios para evitar erros
        comuns_com_divergencia = pd.DataFrame()
        comuns_sem_divergencia = pd.DataFrame()

    # --- 4. EXIBIÇÃO DOS RESULTADOS NO TERMINAL ---

    print("\n" + "="*80)
    print("||" + " RELATÓRIO DE COMPARAÇÃO ".center(76) + "||")
    print("="*80)

    print(f"\n>>> Itens em COMUM COM DIVERGÊNCIAS ({len(comuns_com_divergencia)} registros):")
    if not comuns_com_divergencia.empty:
        print(comuns_com_divergencia)
    else:
        print("Nenhum item encontrado nesta categoria.")

    print(f"\n>>> Itens ENCONTRADOS APENAS na Planilha 1 ('{caminho_planilha1}') ({len(apenas_em_p1)} registros):")
    if not apenas_em_p1.empty:
        print(apenas_em_p1)
    else:
        print("Nenhum item encontrado nesta categoria.")

    print(f"\n>>> Itens ENCONTRADOS APENAS na Planilha 2 ('{caminho_planilha2}') ({len(apenas_em_p2)} registros):")
    if not apenas_em_p2.empty:
        print(apenas_em_p2)
    else:
        print("Nenhum item encontrado nesta categoria.")

    print(f"\n>>> Itens IDÊNTICOS em ambas as planilhas ({len(comuns_sem_divergencia)} registros):")
    if not comuns_sem_divergencia.empty:
        # Para os idênticos, não precisamos mostrar as colunas repetidas de p1 e p2
        colunas_para_mostrar = [coluna_chave] + colunas_comparacao
        print(comuns_sem_divergencia[colunas_para_mostrar])
    else:
        print("Nenhum item encontrado nesta categoria.")

    print("\n" + "="*80)
    print("Comparação finalizada.")
    print("="*80)