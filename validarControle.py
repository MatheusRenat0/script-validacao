import pandas as pd
import os

# --- Configuração dos Arquivos ---
# Altere os nomes dos arquivos se necessário
arquivo_master = 'planilha_master.xlsx'
arquivo_validar = 'planilha_para_validar.xlsx'
arquivo_relatorio = 'relatorio_de_validacao.xlsx'

# Coluna que será usada como chave para encontrar o funcionário
# IMPORTANTE: Use um identificador único como 'Matrícula' ou 'CPF' se possível
coluna_chave = 'Nome do Funcionário'
# ---------------------------------

def validar_planilhas():
    """
    Função principal que carrega as planilhas, trata células mescladas,
    compara os dados e gera um relatório com as divergências.
    """
    # 1. VERIFICAÇÃO INICIAL
    # Verifica se os arquivos de entrada existem na pasta
    if not os.path.exists(arquivo_master) or not os.path.exists(arquivo_validar):
        print(f"Erro: Verifique se os arquivos '{arquivo_master}' e '{arquivo_validar}' estão na mesma pasta do script.")
        return

    print("Carregando planilhas...")
    # 2. CARREGAMENTO E LIMPEZA DOS DADOS
    # Carrega as planilhas
    df_master_raw = pd.read_excel(arquivo_master)
    df_validar_raw = pd.read_excel(arquivo_validar)

    # TRATAMENTO PARA CÉLULAS MESCLADAS:
    # O método 'ffill' (forward fill) preenche os valores vazios (NaN)
    # com o último valor válido encontrado na mesma coluna.
    df_master_raw.ffill(inplace=True)
    df_validar_raw.ffill(inplace=True)

    # Converte todas as colunas para o tipo 'string' para evitar erros de formatação
    # e preenche quaisquer células restantes vazias com um texto vazio.
    df_master = df_master_raw.astype(str).fillna('')
    df_validar = df_validar_raw.astype(str).fillna('')
    
    # Lista para armazenar os erros encontrados
    erros = []

    print("Iniciando a validação de cada linha...")

    # 3. COMPARAÇÃO LINHA A LINHA
    # Itera sobre cada linha da planilha que queremos validar
    for index, row_validar in df_validar.iterrows():
        chave_valor = row_validar[coluna_chave]

        # Procura pelo mesmo funcionário (usando a coluna_chave) na planilha master
        master_info = df_master[df_master[coluna_chave] == chave_valor]

        # Se o funcionário não foi encontrado na planilha master
        if master_info.empty:
            erros.append({
                coluna_chave: chave_valor,
                'Campo com Erro': 'Funcionário',
                'Valor Esperado (Master)': 'N/A',
                'Valor Encontrado (Validação)': chave_valor,
                'Descrição do Erro': f'Funcionário "{chave_valor}" não encontrado na planilha master.'
            })
            continue # Pula para o próximo funcionário

        # Se encontrou o funcionário, faz a comparação campo a campo
        master_info = master_info.iloc[0] # Pega a primeira ocorrência, caso haja duplicados

        # Comparações (usamos .strip() para remover espaços em branco no início/fim)
        if row_validar['IMEI do Aparelho'].strip() != master_info['IMEI do Aparelho'].strip():
            erros.append({
                coluna_chave: chave_valor,
                'Campo com Erro': 'IMEI do Aparelho',
                'Valor Esperado (Master)': master_info['IMEI do Aparelho'],
                'Valor Encontrado (Validação)': row_validar['IMEI do Aparelho'],
                'Descrição do Erro': 'O IMEI do aparelho não corresponde ao registro master.'
            })

        if row_validar['Número da Linha'].strip() != master_info['Número da Linha'].strip():
            erros.append({
                coluna_chave: chave_valor,
                'Campo com Erro': 'Número da Linha',
                'Valor Esperado (Master)': master_info['Número da Linha'],
                'Valor Encontrado (Validação)': row_validar['Número da Linha'],
                'Descrição do Erro': 'O número da linha não corresponde ao registro master.'
            })

        if row_validar['ICCID do Chip'].strip() != master_info['ICCID do Chip'].strip():
            erros.append({
                coluna_chave: chave_valor,
                'Campo com Erro': 'ICCID do Chip',
                'Valor Esperado (Master)': master_info['ICCID do Chip'],
                'Valor Encontrado (Validação)': row_validar['ICCID do Chip'],
                'Descrição do Erro': 'O ICCID do chip (SIM Card) não corresponde ao registro master.'
            })

    # 4. GERAÇÃO DO RELATÓRIO FINAL
    if not erros:
        print("\n--- Validação Concluída ---")
        print("✅ Nenhuma divergência encontrada. As planilhas estão consistentes!")
    else:
        print(f"\n--- Validação Concluída ---")
        print(f"⚠️ Foram encontradas {len(erros)} divergências.")
        # Cria um DataFrame do pandas com a lista de erros
        df_erros = pd.DataFrame(erros)
        
        # Salva o DataFrame em um novo arquivo Excel
        df_erros.to_excel(arquivo_relatorio, index=False)
        print(f"📄 Um relatório detalhado foi salvo no arquivo: '{arquivo_relatorio}'")

# Executa a função principal quando o script é chamado
if __name__ == "__main__":
    validar_planilhas()