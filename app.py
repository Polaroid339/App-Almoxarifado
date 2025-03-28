import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import pandas as pd
from pandastable import Table, TableModel
import csv
from datetime import datetime


arquivos = {
    "estoque": "Planilhas/Estoque.csv",
    "entrada": "Planilhas/Entrada.csv",
    "saida": "Planilhas/Saida.csv"
}

# Funções

def criar_planilhas():
    colunas = {
        "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"],
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA"]
    }
    os.makedirs("Planilhas", exist_ok=True)

    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            df = pd.DataFrame(columns=colunas[nome])
            df.to_csv(arquivo, index=False, encoding="utf-8")
            
            
def obter_proximo_codigo():
    try:
        with open(arquivos["estoque"], "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))
            if len(reader) > 1:
                return int(reader[-1][0]) + 1
            else:
                return 3
    except FileNotFoundError:
        return 3

def pesquisar_tabela():
    query = pesquisar_entry.get().strip().lower()
    if query:
        df_filtered = df[df.apply(lambda row: row.astype(
            str).str.lower().str.contains(query).any(), axis=1)]
    else:
        df_filtered = df

    pandas_table.model.df = df_filtered
    pandas_table.redraw()

def limpar_tabela(): 
    pesquisar_entry.delete(0, tk.END)
    pandas_table.model.df = df
    pandas_table.redraw()
        
def cadastrar_estoque():
    codigo = obter_proximo_codigo()

    descricao = desc_entry.get().strip().upper()
    if not descricao:
        messagebox.showerror("Erro", "Descrição não pode ser vazia.")
        return

    quantidade = quantidade_entry.get().strip()
    if not quantidade.isdigit():
        messagebox.showerror("Erro", "Quantidade deve ser um número inteiro.")
        return
    quantidade = int(quantidade)

    valor_un = valor_entry.get().strip()
    try:
        valor_un = float(valor_un)
    except ValueError:
        messagebox.showerror("Erro", "Valor unitário deve ser um número válido.")
        return

    localizacao = localizacao_entry.get().strip().upper()
    if not localizacao:
        messagebox.showerror("Erro", "Localização não pode ser vazia.")
        return

    data = datetime.now().strftime("%H:%M %d/%m/%Y")

    confirmacao = messagebox.askyesno(
        "Confirmação", f"Você deseja cadastrar o produto com código {codigo} e descrição {descricao}?"
    )
    if confirmacao:
        valor_total = quantidade * valor_un

        try:
            with open(arquivos["estoque"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([int(codigo), descricao, valor_un, valor_total, quantidade, data, localizacao])
            messagebox.showinfo("Sucesso", f"Produto cadastrado com sucesso! \n{descricao} Código: {codigo}")

            desc_entry.delete(0, tk.END)
            quantidade_entry.delete(0, tk.END)
            valor_entry.delete(0, tk.END)
            localizacao_entry.delete(0, tk.END)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o produto: {e}")

def salvar_mudancas():
    pandas_table.model.df = pandas_table.model.df.copy()
    updated_df = pandas_table.model.df
    updated_df.to_csv(os.path.join("Planilhas", "Estoque.csv"), index=False)
    pandas_table.redraw()
    messagebox.showinfo("Sucesso", "Alterações salvas com sucesso!")
    
def atualizar_tabela():
    global df
    try:
        df = pd.read_csv(os.path.join("Planilhas", "Estoque.csv"), encoding="utf-8")
        pandas_table.updateModel(TableModel(df))
        pandas_table.redraw()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar a tabela: {e}")



# Configuração da janela principal

main = tk.Tk()
main.config(bg="#C1BABA")
main.title("Almoxarifado")
main.geometry("1100x600")   
main.resizable(False, False)
main.iconbitmap("favicon.ico")

criar_planilhas()

# Criando o sistema de abas

notebook = ttk.Notebook(main)
notebook.pack(expand=True, fill="both")



# Aba Estoque

estoque_tab = ttk.Frame(notebook)
notebook.add(estoque_tab, text="Estoque")

df = pd.read_csv(os.path.join("Planilhas", "Estoque.csv"), encoding="utf-8")

pandas_table_table_frame = tk.Frame(master=estoque_tab)
pandas_table_table_frame.place(x=20, y=20, width=1057, height=467)
pandas_table = Table(parent=pandas_table_table_frame, dataframe=df)
pandas_table.show()

pesquisar_entry = tk.Entry(master=estoque_tab)
pesquisar_entry.config(bg="#fff", fg="#000")
pesquisar_entry.place(x=20, y=517, width=371, height=43)

limpar_button = tk.Button(master=estoque_tab, text="Limpar", command=limpar_tabela)
limpar_button.config(bg="#EF7E65", fg="#000")
limpar_button.place(x=391, y=517, width=70, height=43)

pesquisar_button = tk.Button(master=estoque_tab, text="Buscar", command=pesquisar_tabela)
pesquisar_button.config(bg="#67F5A5", fg="#000")
pesquisar_button.place(x=461, y=517, width=70, height=43)

save_button = tk.Button(master=estoque_tab, text="Salvar Alterações", command=salvar_mudancas)
save_button.config(bg="#54befc", fg="#000")
save_button.place(x=941, y=517, width=120, height=43)

refresh_button = tk.Button(master=estoque_tab, text="Atualizar", command=atualizar_tabela)
refresh_button.config(bg="#54befc", fg="#000")
refresh_button.place(x=851, y=517, width=80, height=43)



# Aba Cadastro

cadastro_tab = ttk.Frame(notebook)
notebook.add(cadastro_tab, text="Cadastro")

desc_label = tk.Label(master=cadastro_tab, text="Descrição", font=("Arial", 12))
desc_label.place(x=20, y=20)
desc_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
desc_entry.config(bg="#fff", fg="#000")
desc_entry.place(x=20, y=40, width=371, height=40)

quantidade_label = tk.Label(master=cadastro_tab, text="Quantidade", font=("Arial", 12))
quantidade_label.place(x=20, y=90)
quantidade_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
quantidade_entry.config(bg="#fff", fg="#000")
quantidade_entry.place(x=20, y=110, width=371, height=40)

valor_label = tk.Label(master=cadastro_tab, text="Valor Unitário", font=("Arial", 12))
valor_label.place(x=20, y=160)
valor_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
valor_entry.config(bg="#fff", fg="#000")
valor_entry.place(x=20, y=180, width=371, height=40)

localizacao_label = tk.Label(master=cadastro_tab, text="Localização", font=("Arial", 12))
localizacao_label.place(x=20, y=230)
localizacao_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
localizacao_entry.config(bg="#fff", fg="#000")
localizacao_entry.place(x=20, y=250, width=371, height=40)

cadastro_button = tk.Button(master=cadastro_tab, text="Cadastrar", command=cadastrar_estoque)
cadastro_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
cadastro_button.place(x=20, y=320, width=371, height=40)

main.mainloop()
