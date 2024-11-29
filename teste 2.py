import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk  # Importando o Pillow para carregar e exibir imagens

# Função para salvar o arquivo com uma janela personalizada
def salvar_arquivo_personalizado(df_combinado):
    # Criar a janela principal personalizada
    janela = tk.Tk()
    janela.title("Salvar Arquivo")
    janela.geometry("400x300")  # Ajuste o tamanho da janela conforme necessário
    janela.resizable(False, False)
    janela.configure(bg="#f4f4f9")

    # Carregar e exibir o logo
    try:
        logo_path = 'logo.png'  # Caminho para a sua imagem
        imagem = Image.open(logo_path)
        imagem = imagem.resize((100, 100))  # Ajuste o tamanho do logo
        logo = ImageTk.PhotoImage(imagem)
        label_logo = tk.Label(janela, image=logo, bg="#f4f4f9")
        label_logo.image = logo  # Referência para evitar que a imagem seja descartada
        label_logo.pack(pady=10)  # Adiciona o logo com algum espaçamento
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
        font=("Arial", 14), 
        bg="#f4f4f9", 
        fg="#333"
    ).pack(pady=10)

    # Campo de entrada para o nome do arquivo
    entrada_nome = tk.Entry(janela, font=("Arial", 12), width=30)
    entrada_nome.pack(pady=5)

    # Botão "Salvar"
    salvar_btn = tk.Button(
        janela, 
        text="Salvar", 
        command=salvar,  # Aqui associamos o comando à função salvar
        font=("Arial", 12), 
        bg="#4CAF50", 
        fg="white", 
        activebackground="#45a049", 
        activeforeground="white", 
        relief="flat", 
        width=10
    )
    salvar_btn.pack(pady=20)

    # Manter a janela aberta até o usuário interagir
    janela.mainloop()

# Função para formatar valores como moeda (já incluída no seu código original)
def format_to_currency(value):
    try:
        # Converte para string e remove espaços
        value = str(value).replace(' ', '')
        
        # Tenta converter para float
        value = float(value)
        
        # Formatar como moeda, usando a vírgula para separar centavos
        return f"R${value:,.2f}"
    except ValueError:
        return value  # Retorna o valor original caso não consiga converter

# Defina o caminho para a área de trabalho
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# Exemplo de DataFrame para salvar (com valores de exemplo)
df_combinado = pd.DataFrame({
    "id": [1, 2, 3],
    "prop1": ["A", "B", "C"],
    "prop2": ["X", "Y", "Z"],
    "prop3": [1000.50, 2000.75, 3000.99],
    "prop4": ["A1", "B2", "C3"],
    "prop5": ["https://short.hinova.com.br/", "https://short.hinova.com.br/", "https://short.hinova.com.br/"]
})

# Reorganize as colunas conforme necessário
df_reorganizado = df_combinado[['id', 'prop1', 'prop2', 'prop3', 'prop4', 'prop5']]
df_reorganizado = df_reorganizado[['id', 'prop1', 'prop2', 'prop3', 'prop5', 'prop4']]
df_reorganizado.columns = ['recipients', 'var1', 'var2', 'var3', 'var4', 'var5']

# Adicionando as colunas em branco para 'template_name' e 'code'
df_reorganizado.insert(1, 'template_name', 'boletoscoonecta')
df_reorganizado.insert(2, 'code', 'pt_BR')
df_reorganizado['type'] = 'OFICIAL'

# Substituir os pontos por vírgulas e converter para o formato de moeda na coluna 'var3'
df_reorganizado['var3'] = df_reorganizado['var3'].apply(format_to_currency)

# Remover o link da coluna 'var5'
df_reorganizado['var5'] = df_reorganizado['var5'].str.replace("https://short.hinova.com.br/", "", regex=False)

# Combinar todos os dados para a planilha final
df_combinado = df_reorganizado

# Chamar a função para exibir a janela personalizada
salvar_arquivo_personalizado(df_combinado)



