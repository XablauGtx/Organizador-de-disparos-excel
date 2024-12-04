import pandas as pd
import tkinter as tk
from tkinter import simpledialog

def dividir_planilha_arquivo(input_file, output_dir, linhas_por_arquivo=1000, separador=";"):
    # Inicializar a interface do Tkinter
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    # Perguntar o nome base para os arquivos de saída
    nome_base = simpledialog.askstring("Entrada", "Digite o nome base para os arquivos de saída:")
    
    if not nome_base:
        print("Nome base não fornecido. Operação cancelada.")
        return
    
    # Carregar a planilha no formato CSV
    df = pd.read_csv(input_file, sep=separador)
    
    # Garantir que as colunas necessárias estão presentes
    colunas_desejadas = ['recipients', 'template_name', 'code', 'var1', 'var2', 'var3', 'var4', 'var5', 'type']
    df = df[colunas_desejadas]  # Filtrar apenas as colunas desejadas
    
    # Converter a coluna 'recipients' para string, para evitar o problema do .0
    df['recipients'] = df['recipients'].astype(str).str.replace('.0', '', regex=False)
    
    # Calcular o número total de arquivos necessários
    total_linhas = len(df)
    total_arquivos = (total_linhas // linhas_por_arquivo) + (1 if total_linhas % linhas_por_arquivo != 0 else 0)
    
    # Dividir a planilha em arquivos menores
    for i in range(total_arquivos):
        start_row = i * linhas_por_arquivo
        end_row = start_row + linhas_por_arquivo
        df_chunk = df[start_row:end_row]
        
        # Criar o nome do arquivo de saída
        output_file = f"{output_dir}\\{nome_base}{i+1}.csv"
        df_chunk.to_csv(output_file, index=False, sep=separador)  # Usando o mesmo separador do arquivo de entrada
        print(f"Arquivo gerado: {output_file}")

# Exemplo de uso
input_file = r"C:\\Users\\gustavo.barbosa\\Desktop\\Boletos Gerais.csv"
output_dir = r"C:\\Users\\gustavo.barbosa\\Desktop\\Nova pasta\\Nova pasta 3"
dividir_planilha_arquivo(input_file, output_dir)






