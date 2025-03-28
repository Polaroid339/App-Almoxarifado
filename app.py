import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import os
import pandas as pd
from pandastable import Table, TableModel


def criar_planilhas():
    colunas = {
        "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"],
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA"]
    }

    # Cria o diretório "Planilhas" caso não exista
    os.makedirs("Planilhas", exist_ok=True)

    # Definindo os nomes dos arquivos
    arquivos = {
        "estoque": "Planilhas/Estoque.csv",
        "entrada": "Planilhas/Entrada.csv",
        "saida": "Planilhas/Saida.csv"
    }

    # Criação dos arquivos CSV utilizando pandas
    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            # Cria um DataFrame vazio com as colunas especificadas
            df = pd.DataFrame(columns=colunas[nome])

            # Salva o DataFrame no arquivo CSV
            df.to_csv(arquivo, index=False, encoding="utf-8")

# Função para pesquisar na tabela


def search_table():
    query = pesquisar_entry.get().strip().lower()
    if query:
        df_filtered = df[df.apply(lambda row: row.astype(
            str).str.lower().str.contains(query).any(), axis=1)]
    else:
        df_filtered = df  # Mostra tudo se não houver consulta

    pandas_table.model.df = df_filtered  # Atualiza o DataFrame da tabela
    pandas_table.redraw()  # Atualiza a exibição


def verificar_valor(tipo_valor):
    # Obtém o valor inserido no campo de texto
    valor = entry.get() # type: ignore

    try:
        if tipo_valor == "int":
            # Verifica se o valor é um número inteiro
            valor_int = int(valor)
        elif tipo_valor == "float":
            # Verifica se o valor é um número flutuante
            valor_float = float(valor)
        else:
            raise ValueError("Tipo de valor desconhecido.")

        # Se a conversão for bem-sucedida, exibe a mensagem de sucesso
        messagebox.showinfo(
            "Sucesso", f"Valor '{valor}' aceito como {tipo_valor}!")

    except ValueError:
        # Se não for do tipo esperado, exibe o popup de erro
        messagebox.showerror(
            "Erro", f"Por favor, insira um valor válido para o tipo {tipo_valor}.")

# Função para salvar as alterações


def save_changes():
    # Garante que as edições sejam capturadas
    pandas_table.model.df = pandas_table.model.df.copy()
    updated_df = pandas_table.model.df  # Obtém os dados editados na tabela
    updated_df.to_csv(os.path.join("Planilhas", "Estoque.csv"), index=False)
    pandas_table.redraw()


# Configuração da Janela principal
main = tk.Tk()
main.config(bg="#C1BABA")
main.title("Sistema de Abas")
main.geometry("1100x600")

criar_planilhas()

# Criando o Notebook (sistema de abas)
notebook = ttk.Notebook(main)
notebook.pack(expand=True, fill="both")

# Criando a aba "Estoque"
estoque_tab = ttk.Frame(notebook)
notebook.add(estoque_tab, text="Estoque")

# Carregar CSV em um DataFrame
df = pd.read_csv(os.path.join("Planilhas", "Estoque.csv"))

# Frame da tabela
pandas_table_table_frame = tk.Frame(master=estoque_tab)
pandas_table_table_frame.place(x=20, y=20, width=1057, height=467)
pandas_table = Table(parent=pandas_table_table_frame, dataframe=df)
pandas_table.show()

# Campo de entrada para pesquisa
pesquisar_entry = tk.Entry(master=estoque_tab)
pesquisar_entry.config(bg="#fff", fg="#000")
pesquisar_entry.place(x=20, y=517, width=371, height=43)

# Botão para pesquisar
pesquisar_button = tk.Button(master=estoque_tab, text="Buscar", command=search_table)
pesquisar_button.config(bg="#E4E2E2", fg="#000")
pesquisar_button.place(x=391, y=517, width=80, height=43)

# Botão para salvar alterações
save_button = tk.Button(
    master=estoque_tab, text="Salvar Alterações", command=save_changes)
save_button.config(bg="#E4E2E2", fg="#000")
save_button.place(x=941, y=517, width=120, height=43)

cadastro_tab = ttk.Frame(notebook)
notebook.add(cadastro_tab, text="Cadastro")

main.mainloop()
