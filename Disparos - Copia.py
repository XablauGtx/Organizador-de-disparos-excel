import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle

# Defina o caminho para a pasta "Nova pasta" na área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
folder_path = os.path.join(desktop_path, "Nova pasta")

# Verifique se a pasta existe
if not os.path.exists(folder_path):
    raise FileNotFoundError(f"A pasta 'Nova pasta' não foi encontrada no Desktop.")

# Liste todos os arquivos Excel na pasta
planilhas = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".xlsx")]

# Lista para armazenar os DataFrames
dfs = []

# Função para converter valores para moeda
def format_to_currency(value):
    try:
        # Verifica se o valor já é numérico
        if isinstance(value, (int, float)):
            # Formatar como moeda, usando a vírgula para separar centavos
            return f"R${value:,.2f}"
        
        # Caso não seja numérico, tenta converter para float
        value = str(value).replace(' ', '')  # Remove espaços
        value = float(value)  # Converte para float

        # Retorna o valor formatado como moeda
        return f"R${value:,.2f}"
    except (ValueError, TypeError):
        # Retorna o valor original caso não consiga converter
        return value  # Retorna o valor original caso não consiga converter

# Ler cada planilha e armazenar os dados em um DataFrame
for arquivo in planilhas:
    print(f"Lendo arquivo: {arquivo}")
    df = pd.read_excel(arquivo)
    
    # Verifique os nomes das colunas no arquivo Excel
    print(f"Colunas no arquivo {arquivo}: {df.columns.tolist()}")  # Exibe os nomes das colunas
    
    # Agora, reorganize as colunas de acordo com seus nomes reais 
    try:
        # Exemplo de reorganização (ajuste conforme necessário)
        df_reorganizado = df[['id', 'prop1', 'prop2', 'prop3', 'prop4', 'prop5']]
        df_reorganizado = df_reorganizado[['id', 'prop1', 'prop2', 'prop3', 'prop5', 'prop4']]
        df_reorganizado.columns = ['recipients',  'var1', 'var2', 'var3', 'var4', 'var5']

        # Adicionando as colunas em branco para 'template_name' e 'code' entre 'recipients' e 'var1'
        df_reorganizado.insert(1, 'template_name', 'boletoscoonecta')
        df_reorganizado.insert(2, 'code', 'pt_BR')
        df_reorganizado['type'] = 'OFICIAL'
        # Substituir os pontos por vírgulas e converter para o formato de moeda na coluna 'var3'
        df_reorganizado['var3'] = df_reorganizado['var3'].apply(format_to_currency)
        # Remover o valor 'https://short.hinova.com.br/' da coluna 'var5'
        df_reorganizado['var5'] = df_reorganizado['var5'].str.replace('https://short.hinova.com.br/v2/', '', regex=False)

        


        dfs.append(df_reorganizado)
    except KeyError as e:
        print(f"Erro ao acessar as colunas: {e}")

# Combinar todas as planilhas em uma única
df_combinado = pd.concat(dfs, ignore_index=True)

# Solicitar ao usuário um nome para o arquivo de saída
nome_arquivo = input("Digite o nome do arquivo de saída (sem extensão): ").strip()

# Garantir que o nome não esteja vazio
while not nome_arquivo:
    print("O nome do arquivo não pode ser vazio.")
    nome_arquivo = input("Digite o nome do arquivo de saída (sem extensão): ").strip()

# Adicionar a extensão .csv ao nome do arquivo
nome_arquivo += ".csv"

# Criar o caminho completo para salvar o arquivo na área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
arquivo_saida = os.path.join(desktop_path, nome_arquivo)

# Salvar o DataFrame combinado no arquivo especificado
df_combinado.to_csv(arquivo_saida, index=False, sep=';', encoding='utf-8')

print(f"Dados combinados foram salvos em: {arquivo_saida}")
