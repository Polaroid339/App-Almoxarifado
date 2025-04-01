import os
import csv
import time
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
from pandastable import Table, TableModel


arquivos = {
    "estoque": "Planilhas/Estoque.csv",
    "entrada": "Planilhas/Entrada.csv",
    "saida": "Planilhas/Saida.csv"
}



# Funções

def criar_planilhas():
    """
    Cria os arquivos CSV necessários para o funcionamento do sistema, caso não existam.
    """
    
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
    """
    Obtém o próximo código disponível para um novo produto.
    """
    
    try:
        with open(arquivos["estoque"], "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))
            if len(reader) > 1:
                return int(float(reader[-1][0])) + 1
            else:
                return 3
    except (FileNotFoundError, ValueError):
        return 3

    
def atualizar_estoque(codigo, nova_quantidade):
    """
    Atualiza a quantidade de um produto no estoque.
    """
    
    with open(arquivos["estoque"], "r", encoding="utf-8") as f:
        produtos = list(csv.reader(f))

    for produto in produtos:
        if produto[0] == codigo:
            try:
                produto[4] = str(nova_quantidade)
                produto[3] = str(float(produto[2]) * int(nova_quantidade))
            except ValueError:
                messagebox.showerror("Erro", "Erro ao atualizar o estoque. Verifique os valores numéricos.")
                return

    with open(arquivos["estoque"], "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(produtos)


def buscar_produto(codigo):
    """
    Busca um produto no estoque pelo código.
    """
    
    try:
        with open(arquivos["estoque"], "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader)
            for row in reader:
                if row[0] == codigo:
                    return row
    except FileNotFoundError:
        pass
    return None


def pesquisar_tabela():
    """
    Filtra a tabela com base na entrada do usuário.
    """
    
    query = pesquisar_entry.get().strip().lower()
    if query:
        df_filtered = df[df.apply(lambda row: row.astype(
            str).str.lower().str.contains(query).any(), axis=1)]
    else:
        df_filtered = df

    pandas_table.model.df = df_filtered
    pandas_table.redraw()


def limpar_tabela(): 
    """
    Limpa a tabela de pesquisa e exibe todos os produtos.
    """
    
    pesquisar_entry.delete(0, tk.END)
    pandas_table.model.df = df
    pandas_table.redraw()

        
def cadastrar_estoque():
    """
    Cadastra um novo produto no estoque.
    """
    
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
            return

            
def registrar_entrada():
    """
    Registra a entrada de um produto no estoque.
    """
    
    codigo = codigo_entry.get().strip()
    if not codigo:
        messagebox.showerror("Erro", "O código do produto não pode ser vazio.")
        return

    produto = buscar_produto(codigo)
    if not produto:
        messagebox.showerror("Erro", "Código do produto não encontrado.")
        return

    quantidade_adicionada = quantidade_entrada_entry.get().strip()
    if not quantidade_adicionada.isdigit():
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro.")
        return
    quantidade_adicionada = int(quantidade_adicionada)

    try:
        nova_quantidade = int(produto[4]) + quantidade_adicionada
    except ValueError:
        messagebox.showerror("Erro", "Erro ao calcular a nova quantidade. Verifique os valores no estoque.")
        return

    nova_quantidade = int(produto[4]) + quantidade_adicionada
    data = datetime.now().strftime("%H:%M %d/%m/%Y")

    with open(arquivos["estoque"], "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        linhas = list(reader)

    valor_un = None
    for linha in linhas:
        if linha[0] == codigo:
            valor_un = float(linha[2])
            break

    if valor_un is None:
        messagebox.showerror("Erro", "Código do produto não encontrado no estoque!")
        return

    valor_total = valor_un * quantidade_adicionada

    confirmacao = messagebox.askyesno(
        "Confirmação",
        f"Você deseja registrar a entrada nesse produto?\n\n"
        f"Código: {codigo}\n"
        f"Descrição: {produto[1]}\n"
        f"Quantidade a adicionar: {quantidade_adicionada}\n"
        f"Valor Unitário: R$ {valor_un:.2f}\n"
        f"Valor Total: R$ {valor_total:.2f}"
    )

    if confirmacao:
        with open(arquivos["entrada"], "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([codigo, produto[1], quantidade_adicionada, valor_un, valor_total, data])

        atualizar_estoque(codigo, nova_quantidade)
        messagebox.showinfo(
            "Sucesso",
            f"Entrada registrada e estoque atualizado!\n"
            f"{produto[1]} | Quantidade: {nova_quantidade}"
        )
        
        codigo_entry.delete(0, tk.END)
        quantidade_entrada_entry.delete(0, tk.END)
        

def registrar_saida():
    """
    Registra a saída de um produto do estoque.
    """
    
    codigo = codigo_saida_entry.get().strip()
    if not codigo:
        messagebox.showerror("Erro", "O código do produto não pode ser vazio.")
        return

    produto = buscar_produto(codigo)
    if not produto:
        messagebox.showerror("Erro", "Código do produto não encontrado.")
        return

    solicitante = solicitante_entry.get().strip().upper()
    if not solicitante:
        messagebox.showerror("Erro", "O nome do solicitante não pode ser vazio.")
        return

    quantidade_retirada = quantidade_saida_entry.get().strip()
    if not quantidade_retirada.isdigit() or int(quantidade_retirada) <= 0:
        messagebox.showerror("Erro", "A quantidade deve ser um número inteiro maior que zero.")
        return

    quantidade_retirada = int(quantidade_retirada)

    if quantidade_retirada > int(produto[4]):
        messagebox.showerror("Erro", "Quantidade insuficiente no estoque!")
        return

    nova_quantidade = int(produto[4]) - quantidade_retirada
    data = datetime.now().strftime("%H:%M %d/%m/%Y")

    confirmacao = messagebox.askyesno(
        "Confirmação",
        f"Você deseja registrar a saída deste produto?\n\n"
        f"Código: {codigo}\n"
        f"Descrição: {produto[1]}\n"
        f"Quantidade a retirar: {quantidade_retirada}\n"
        f"Solicitante: {solicitante}\n"
        f"Quantidade restante: {nova_quantidade}"
    )

    if confirmacao:
        with open(arquivos["saida"], "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([codigo, produto[1], quantidade_retirada, solicitante, data])

        atualizar_estoque(codigo, nova_quantidade)
        messagebox.showinfo(
            "Sucesso",
            f"Saída registrada e estoque atualizado!\n"
            f"{produto[1]} | Quantidade restante: {nova_quantidade}"
        )

        codigo_saida_entry.delete(0, tk.END)
        solicitante_entry.delete(0, tk.END)
        quantidade_saida_entry.delete(0, tk.END)


def exportar_conteudo():
    """
    Exporta o conteúdo das planilhas para um arquivo Excel e gera um relatório de produtos esgotados.
    """
    
    pasta_saida = "Relatorios"
    os.makedirs(pasta_saida, exist_ok=True)
    caminho_excel = os.path.join(pasta_saida, "Relatorio_Almoxarifado.xlsx")
    caminho_txt = os.path.join(pasta_saida, "Produtos_Esgotados.txt")

    try:
        with pd.ExcelWriter(caminho_excel) as writer:
            for nome, arquivo in arquivos.items():
                try:
                    df = pd.read_csv(arquivo, encoding="utf-8")
                    df.to_excel(writer, sheet_name=nome.capitalize(), index=False)
                except FileNotFoundError:
                    messagebox.showwarning("Aviso", f"Arquivo {arquivo} não encontrado. Ignorando...")

        df_estoque = pd.read_csv(arquivos["estoque"], encoding="utf-8")
        if "QUANTIDADE" not in df_estoque.columns:
            messagebox.showerror("Erro", "Coluna 'QUANTIDADE' não encontrada no estoque.")
            return

        produtos_esgotados = df_estoque[df_estoque["QUANTIDADE"] == 0]

        with open(caminho_txt, "w", encoding="utf-8") as f:
            f.write("Relatório de Produtos Esgotados\n")
            f.write("-" * 40 + "\n")
            for _, row in produtos_esgotados.iterrows():
                f.write(f"Código: {row['CODIGO']} | Descrição: {row['DESCRICAO']}\n")
            f.write("-" * 40 + "\n")

        messagebox.showinfo("Sucesso", f"Relatórios exportados com sucesso!\n\n"
                                       f"Excel: {caminho_excel}\n"
                                       f"Produtos Esgotados: {caminho_txt}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao exportar relatórios: {e}")  


tabela_atual = "estoque"


def trocar_tabela(nome_tabela):
    """
    Troca a tabela exibida na interface gráfica.
    """
    
    global tabela_atual, df

    if nome_tabela not in arquivos:
        messagebox.showerror("Erro", f"Tabela {nome_tabela} não encontrada.")
        return

    tabela_atual = nome_tabela

    try:
        df = pd.read_csv(arquivos[nome_tabela], encoding="utf-8")
        pandas_table.updateModel(TableModel(df))
        pandas_table.redraw()

        atualizar_cores_botoes()

        messagebox.showinfo("Tabela Atualizada", f"Agora exibindo a tabela {nome_tabela.capitalize()}")
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo da tabela {nome_tabela} não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a tabela {nome_tabela}: {e}")


def atualizar_cores_botoes():
    """
    Atualiza as cores dos botões de tabela com base na tabela atual.
    """
    
    tabela_estoque_button.config(bg="#C1BABA", fg="#000")
    tabela_entrada_button.config(bg="#C1BABA", fg="#000")
    tabela_saida_button.config(bg="#C1BABA", fg="#000")

    if tabela_atual == "estoque":
        tabela_estoque_button.config(bg="#54befc", fg="#000")
    elif tabela_atual == "entrada":
        tabela_entrada_button.config(bg="#54befc", fg="#000")
    elif tabela_atual == "saida":
        tabela_saida_button.config(bg="#54befc", fg="#000")


def salvar_mudancas():
    """
    Salva as alterações feitas na tabela atual no arquivo CSV correspondente.
    """
    
    try:
        df_original = pd.read_csv(arquivos[tabela_atual], encoding="utf-8")

        df_atualizado = pandas_table.model.df.copy()

        for index, row in df_atualizado.iterrows():
            if not row.equals(df_original.loc[index]):
                df_original.loc[index] = row

        if os.path.exists(arquivos[tabela_atual]):
            shutil.copy(arquivos[tabela_atual], arquivos[tabela_atual].replace(".csv", "_backup.csv"))

        df_original.to_csv(arquivos[tabela_atual], index=False, encoding="utf-8")

        pandas_table.updateModel(TableModel(df_original))
        pandas_table.redraw()

        messagebox.showinfo("Sucesso", f"Alterações na tabela {tabela_atual.capitalize()} salvas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar alterações na tabela {tabela_atual}: {e}")
        

def criar_backup_periodico():
    """
    Cria backups periódicos dos arquivos de dados e remove backups com mais de 3 dias.
    """
    
    pasta_backup = "Backups"
    os.makedirs(pasta_backup, exist_ok=True)

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    try:
        for nome, arquivo in arquivos.items():
            if os.path.exists(arquivo):
                nome_backup = f"{nome}_{timestamp}.csv"
                caminho_backup = os.path.join(pasta_backup, nome_backup)
                shutil.copy(arquivo, caminho_backup)

        print(f"Backup criado com sucesso em {timestamp}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar backup: {e}")

    try:
        agora = time.time()
        for arquivo in os.listdir(pasta_backup):
            caminho_arquivo = os.path.join(pasta_backup, arquivo)
            if os.path.isfile(caminho_arquivo):
                tempo_modificacao = os.path.getmtime(caminho_arquivo)
                if (agora - tempo_modificacao) > (3 * 24 * 60 * 60):
                    os.remove(caminho_arquivo)
                    print(f"Backup antigo removido: {arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao remover backups antigos: {e}")

    main.after(10800000, criar_backup_periodico)


def atualizar_tabela():
    """
    Atualiza a tabela atual com os dados mais recentes do arquivo CSV correspondente.
    """
    
    global df
    try:
        df = pd.read_csv(arquivos[tabela_atual], encoding="utf-8")

        if os.path.exists("./Planilhas/Estoque.csv"):
            shutil.copy("./Planilhas/Estoque.csv", "./Planilhas/Estoque_backup.csv")

        if tabela_atual == "estoque":
            df["VALOR TOTAL"] = df["VALOR UN"] * df["QUANTIDADE"]
            df.to_csv(arquivos["estoque"], index=False, encoding="utf-8")         

        pandas_table.updateModel(TableModel(df))
        pandas_table.redraw()

    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo da tabela {tabela_atual} não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar a tabela {tabela_atual}: {e}")



# Configuração da janela principal

main = tk.Tk()
main.config(bg="#C1BABA")
main.title("Almoxarifado")
main.geometry("1100x600")   
main.resizable(False, False)

criar_planilhas()

# Criando o sistema de abas

notebook = ttk.Notebook(main)
notebook.pack(expand=True, fill="both")



# Aba Estoque

estoque_tab = ttk.Frame(notebook)
notebook.add(estoque_tab, text="Estoque")

df = pd.read_csv(os.path.join("Planilhas", "Estoque.csv"), encoding="utf-8")

pandas_table_table_frame = tk.Frame(master=estoque_tab)
pandas_table_table_frame.place(x=20, y=20, width=1057, height=483)
pandas_table = Table(parent=pandas_table_table_frame, dataframe=df)
pandas_table.show()

pesquisar_entry = tk.Entry(master=estoque_tab)
pesquisar_entry.config(bg="#fff", fg="#000", borderwidth=3)
pesquisar_entry.place(x=20, y=517, width=295, height=43)

limpar_button = tk.Button(master=estoque_tab, text="Limpar", command=limpar_tabela)
limpar_button.config(bg="#EF7E65", fg="#000")
limpar_button.place(x=320, y=517, width=70, height=43)

pesquisar_button = tk.Button(master=estoque_tab, text="Buscar", command=pesquisar_tabela)
pesquisar_button.config(bg="#67F5A5", fg="#000")
pesquisar_button.place(x=390, y=517, width=70, height=43)

save_button = tk.Button(master=estoque_tab, text="Salvar Alterações", command=salvar_mudancas)
save_button.config(bg="#54befc", fg="#000")
save_button.place(x=941, y=517, width=120, height=43)

refresh_button = tk.Button(master=estoque_tab, text="Atualizar", command=atualizar_tabela)
refresh_button.config(bg="#54befc", fg="#000")
refresh_button.place(x=851, y=517, width=80, height=43)

exportar_button = tk.Button(master=estoque_tab, text="Exportar", command=exportar_conteudo)
exportar_button.config(bg="#FFFF00", fg="#000")
exportar_button.place(x=490, y=517, width=80, height=43)

tabela_estoque_button = tk.Button(master=estoque_tab, text="Estoque", command=lambda: trocar_tabela("estoque"))
tabela_estoque_button.config(bg="#C1BABA", fg="#000")
tabela_estoque_button.place(x=607, y=517, width=70, height=43)

tabela_entrada_button = tk.Button(master=estoque_tab, text="Entrada", command=lambda: trocar_tabela("entrada"))
tabela_entrada_button.config(bg="#C1BABA", fg="#000")
tabela_entrada_button.place(x=677, y=517, width=70, height=43)

tabela_saida_button = tk.Button(master=estoque_tab, text="Saída", command=lambda: trocar_tabela("saida"))
tabela_saida_button.config(bg="#C1BABA", fg="#000")
tabela_saida_button.place(x=747, y=517, width=70, height=43)



# Aba Cadastro

cadastro_tab = ttk.Frame(notebook)
notebook.add(cadastro_tab, text="Cadastro")

titulo_cadastro_label = tk.Label(master=cadastro_tab, text="Cadastro de Produto", font=("Arial", 16))
titulo_cadastro_label.place(x=20, y=20)

desc_label = tk.Label(master=cadastro_tab, text="Descrição", font=("Arial", 12))
desc_label.place(x=20, y=60)
desc_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
desc_entry.config(bg="#fff", fg="#000")
desc_entry.place(x=20, y=90, width=371, height=30)

quantidade_label = tk.Label(master=cadastro_tab, text="Quantidade", font=("Arial", 12))
quantidade_label.place(x=20, y=130)
quantidade_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
quantidade_entry.config(bg="#fff", fg="#000")
quantidade_entry.place(x=20, y=160, width=371, height=30)

valor_label = tk.Label(master=cadastro_tab, text="Valor Unitário", font=("Arial", 12))
valor_label.place(x=20, y=200)
valor_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
valor_entry.config(bg="#fff", fg="#000")
valor_entry.place(x=20, y=230, width=371, height=30)

localizacao_label = tk.Label(master=cadastro_tab, text="Localização", font=("Arial", 12))
localizacao_label.place(x=20, y=270)
localizacao_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
localizacao_entry.config(bg="#fff", fg="#000")
localizacao_entry.place(x=20, y=300, width=371, height=30)

cadastro_button = tk.Button(master=cadastro_tab, text="Cadastrar", command=cadastrar_estoque)
cadastro_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
cadastro_button.place(x=20, y=350, width=371, height=40)



# Aba Movimentação

movimentacao_tab = ttk.Frame(notebook)
notebook.add(movimentacao_tab, text="Movimentação")


# Entrada

titulo_entrada_label = tk.Label(master=movimentacao_tab, text="Registrar Entrada", font=("Arial", 16))
titulo_entrada_label.place(x=80, y=20)

codigo_label = tk.Label(master=movimentacao_tab, text="Código", font=("Arial", 12))
codigo_label.place(x=80, y=60)
codigo_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
codigo_entry.config(bg="#fff", fg="#000")
codigo_entry.place(x=80, y=90, width=371, height=30)

quantidade_entrada_label = tk.Label(master=movimentacao_tab, text="Quantidade", font=("Arial", 12))
quantidade_entrada_label.place(x=80, y=130)
quantidade_entrada_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
quantidade_entrada_entry.config(bg="#fff", fg="#000")
quantidade_entrada_entry.place(x=80, y=160, width=371, height=30)

entrada_button = tk.Button(master=movimentacao_tab, text="Registrar Entrada", command=registrar_entrada)
entrada_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
entrada_button.place(x=80, y=210, width=371, height=40)

separator = ttk.Separator(movimentacao_tab, orient="vertical")
separator.place(x=550, y=20, height=530)


# Saída

titulo_saida_label = tk.Label(master=movimentacao_tab, text="Registrar Saída", font=("Arial", 16))
titulo_saida_label.place(x=645, y=20)

codigo_saida_label = tk.Label(master=movimentacao_tab, text="Código", font=("Arial", 12))
codigo_saida_label.place(x=645, y=60)
codigo_saida_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
codigo_saida_entry.config(bg="#fff", fg="#000")
codigo_saida_entry.place(x=645, y=90, width=371, height=30)

solicitante_label = tk.Label(master=movimentacao_tab, text="Solicitante", font=("Arial", 12))
solicitante_label.place(x=645, y=130)
solicitante_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
solicitante_entry.config(bg="#fff", fg="#000")
solicitante_entry.place(x=645, y=160, width=371, height=30)

quantidade_saida_label = tk.Label(master=movimentacao_tab, text="Quantidade", font=("Arial", 12))
quantidade_saida_label.place(x=645, y=200)
quantidade_saida_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
quantidade_saida_entry.config(bg="#fff", fg="#000")
quantidade_saida_entry.place(x=645, y=230, width=371, height=30)

saida_button = tk.Button(master=movimentacao_tab, text="Registrar Saída", command=registrar_saida)
saida_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
saida_button.place(x=645, y=280, width=371, height=40)

criar_backup_periodico()

main.mainloop()

# python -m PyInstaller --onefile --name=Almoxarifado --windowed --icon=favicon.ico --add-data "Planilhas;Planilhas" main.py
