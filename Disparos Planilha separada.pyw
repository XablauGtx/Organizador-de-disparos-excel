import pandas as pd
import os
from openpyxl import load_workbook
import tkinter as tk
from tkinter import simpledialog, messagebox
from PIL import Image, ImageTk 
import tkinter as tk
from tkinter import messagebox
import tkinter.font as tkFont
import subprocess

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
        # Converte para string e remove espaços
        value = str(value).replace(' ', '')
        # Converte para float
        value = float(value)
        # Formata como moeda
        return f"R${value:,.2f}"
    except ValueError:
        return value  # Retorna o valor original caso não consiga converter

# Ler cada planilha e armazenar os dados em um DataFrame
for arquivo in planilhas:
    print(f"Lendo arquivo: {arquivo}")
    df = pd.read_excel(arquivo)
    
    try:
        # Reorganizar colunas (ajuste conforme necessário)
        df_reorganizado = df[['id', 'prop1', 'prop2', 'prop3', 'prop4', 'prop5']]
        df_reorganizado = df_reorganizado[['id', 'prop1', 'prop2', 'prop3', 'prop5', 'prop4']]
        df_reorganizado.columns = ['recipients', 'var1', 'var2', 'var3', 'var4', 'var5']

        # Adicionar colunas extras
        df_reorganizado.insert(1, 'template_name', 'boletoscoonecta')
        df_reorganizado.insert(2, 'code', 'pt_BR')
        df_reorganizado['type'] = 'OFICIAL'

        # Converter valores para moeda na coluna 'var3'
        df_reorganizado['var3'] = df_reorganizado['var3'].apply(format_to_currency)

        # Remover o valor específico da coluna 'var5'
        df_reorganizado['var5'] = df_reorganizado['var5'].str.replace('https://short.hinova.com.br/v2/', '', regex=False)
        

        dfs.append(df_reorganizado)
    except KeyError as e:
        print(f"Erro ao acessar as colunas: {e}")

# Combinar todas as planilhas em uma única
df_combinado = pd.concat(dfs, ignore_index=True)
###################################################################################################################################################
# Função para salvar o arquivo com uma janela personalizada
def salvar_arquivo_personalizado(df_combinado):
    # Criar a janela principal personalizada
    janela = tk.Tk()
    janela.title("Salvar Arquivo")
    janela.geometry("400x400")  # Ajuste o tamanho da janela conforme necessário
    janela.resizable(False, False)
    janela.configure(bg="#0000ff")

    # Criando uma fonte personalizada com tkinter.font
    fonte_personalizada = tkFont.Font(family="Blantic", size=15, weight="normal")  # Fonte Arial, tamanho 14, negrito

    # Carregar e exibir o logo
    try:
        logo_path = 'C:\\Users\\gustavo.barbosa\\Pictures\\Lenovo\\Coonecta.png'  # Caminho para a sua imagem PNG com fundo transparente
        imagem = Image.open(logo_path).convert("RGBA")  # Garantir que a imagem seja carregada com canal alfa (transparente)
        imagem = imagem.resize((150, 150))  # Ajuste o tamanho da logo
        logo = ImageTk.PhotoImage(imagem)
        label_logo = tk.Label(janela, image=logo, bg="#0000ff")  # Fundo da janela será a cor #f4f4f9
        label_logo.image = logo  # Referência para evitar que a imagem seja descartada
        label_logo.pack(pady=0)  # Adiciona o logo com algum espaçamento
    except Exception as e:
        print(f"Erro ao carregar logo: {e}")

    # Função que será chamada ao clicar no botão "Salvar"
    def salvar():
        nome_arquivo = entrada_nome.get().strip()  # Pega o texto do campo de entrada
        if not nome_arquivo:  # Verifica se o nome do arquivo foi fornecido
            messagebox.showerror("Erro", "O nome do arquivo não pode estar vazio.")
            return

        # Adicionar a extensão e criar o caminho do arquivo
        arquivo_saida = os.path.join(desktop_path, f"{nome_arquivo}.csv")
        try:
            df_combinado.to_csv(arquivo_saida, index=False, sep=';', encoding='utf-8')
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{arquivo_saida}")
            janela.destroy()  # Fechar a janela após salvar
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar o arquivo:\n{e}")

    # Rótulo para a instrução
    tk.Label(
        janela, 
        text="Digite o nome do arquivo:", 
        font=fonte_personalizada, 
        bg="#0000ff", 
        fg="#ffe7e5"
    ).pack(pady=10)

    # Campo de entrada para o nome do arquivo
    entrada_nome = tk.Entry(janela, font=("Arial", 12), width=30)
    entrada_nome.pack(pady=5)

    # Botão "Salvar"
    salvar_btn = tk.Button(
        janela, 
        text="Salvar", 
        command=salvar,  # Aqui associamos o comando à função salvar
        font=fonte_personalizada, 
        bg="#fffffe", 
        fg="black", 
        activebackground="#0000ff", 
        activeforeground="white", 
        relief="flat", 
        width=10
    )
    salvar_btn.pack(pady=20)

    # Manter a janela aberta até o usuário interagir
    janela.mainloop()
######################################################################################################################################
    # Adicionar extensão e criar o caminho completo para salvar
    arquivo_saida = os.path.join(desktop_path,)
    
# Chamar a função para exibir a janela personalizada
salvar_arquivo_personalizado(df_combinado)

# Caminho do outro arquivo Python
arquivo_python = 'C:\\Users\\gustavo.barbosa\\Desktop\\Nova pasta\\Nova pasta 2\\Separador de planilha.py'

# Executando o outro arquivo
subprocess.run(['python', arquivo_python])