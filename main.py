import os
import csv
import time
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from usuarios import usuarios
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
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA", "ID"],
        "epis": ["CA", "DESCRICAO", "QUANTIDADE"]
    }
    os.makedirs("Planilhas", exist_ok=True)

    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            df = pd.DataFrame(columns=colunas[nome])
            df.to_csv(arquivo, index=False, encoding="utf-8")

    if not os.path.exists("Planilhas/Epis.csv"):
        df_epis = pd.DataFrame(columns=["CA", "DESCRICAO", "QUANTIDADE"])
        df_epis.to_csv("Planilhas/Epis.csv", index=False, encoding="utf-8")
            
            
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


def pesquisar_tabela(event=None):
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
    try:
        quantidade = float(quantidade.replace(",", "."))

    except ValueError:
        messagebox.showerror("Erro", "Quantidade deve ser um número válido.")
        return

    valor_un = valor_entry.get().strip()
    try:
        valor_un = float(valor_un.replace(",", "."))

    except ValueError:
        messagebox.showerror("Erro", "Valor unitário deve ser um número válido.")
        return

    localizacao = localizacao_entry.get().strip().upper()

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
    try:
        quantidade_adicionada = float(quantidade_adicionada.replace(",", "."))
        if quantidade_adicionada <= 0:
            raise ValueError("Quantidade deve ser maior que zero.")
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número válido e maior que zero.")
        return

    try:
        nova_quantidade = float(produto[4]) + quantidade_adicionada
    except ValueError:
        messagebox.showerror("Erro", "Erro ao calcular a nova quantidade. Verifique os valores no estoque.")
        return

    data = datetime.now().strftime("%H:%M %d/%m/%Y")

    valor_un = float(produto[2])
    valor_total = valor_un * quantidade_adicionada

    confirmacao = messagebox.askyesno(
        "Confirmação",
        f"Você deseja registrar a entrada neste produto?\n\n"
        f"Código: {codigo}\n"
        f"Descrição: {produto[1]}\n"
        f"Quantidade a adicionar: {quantidade_adicionada}\n"
        f"Valor Unitário: R$ {valor_un:.2f}\n"
        f"Valor Total: R$ {valor_total:.2f}"
    )

    if confirmacao:
        with open(arquivos["entrada"], "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([codigo, produto[1], quantidade_adicionada, valor_un, valor_total, data, operador_logado_id])

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
    try:
        quantidade_retirada = float(quantidade_retirada.replace(",", "."))
        if quantidade_retirada <= 0:
            raise ValueError("Quantidade deve ser maior que zero.")
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número válido e maior que zero.")
        return

    if quantidade_retirada > float(produto[4]):
        messagebox.showerror("Erro", "Quantidade insuficiente no estoque!")
        return

    nova_quantidade = float(produto[4]) - quantidade_retirada
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
            writer.writerow([codigo, produto[1], quantidade_retirada, solicitante, data, operador_logado_id])

        atualizar_estoque(codigo, nova_quantidade)
        messagebox.showinfo(
            "Sucesso",
            f"Saída registrada e estoque atualizado!\n"
            f"{produto[1]} | Quantidade restante: {nova_quantidade}"
        )

        codigo_saida_entry.delete(0, tk.END)
        solicitante_entry.delete(0, tk.END)
        quantidade_saida_entry.delete(0, tk.END)


def registrar_epi():
    """
    Registra um novo EPI no arquivo Epis.csv ou atualiza a quantidade de um EPI existente.
    """
    ca = ca_entry.get().strip().upper()
    descricao = descricao_epi_entry.get().strip().upper()
    quantidade = quantidade_epi_entry.get().strip()

    if not (ca or descricao):
        messagebox.showerror("Erro", "Você deve preencher pelo menos o CA ou a Descrição.")
        return

    if not quantidade:
        messagebox.showerror("Erro", "A Quantidade deve ser preenchida.")
        return

    try:
        quantidade = float(quantidade)
        if quantidade <= 0:
            raise ValueError("Quantidade deve ser maior que zero.")
    except ValueError:
        messagebox.showerror("Erro", "A Quantidade deve ser um número válido e maior que zero.")
        return

    try:
        df_epis = pd.read_csv("Planilhas/Epis.csv", encoding="utf-8", dtype={"CA": str})
        df_epis["CA"] = df_epis["CA"].fillna("").astype(str).str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].fillna("").astype(str).str.strip().str.upper()

        if ca:
            epi_existente_ca = df_epis[df_epis["CA"] == ca]
            if not epi_existente_ca.empty:
                adicionar_quantidade = messagebox.askyesno(
                    "EPI Já Existente",
                    f"Já existe um EPI com este CA.\n"
                    f"CA: {epi_existente_ca.iloc[0]['CA']}\n"
                    f"Descrição: {epi_existente_ca.iloc[0]['DESCRICAO']}\n"
                    f"Quantidade Atual: {epi_existente_ca.iloc[0]['QUANTIDADE']}\n\n"
                    f"Deseja adicionar {quantidade} à quantidade existente?"
                )
                if adicionar_quantidade:
                    nova_quantidade = float(epi_existente_ca.iloc[0]["QUANTIDADE"]) + quantidade
                    df_epis.loc[epi_existente_ca.index, "QUANTIDADE"] = nova_quantidade
                    df_epis.to_csv("Planilhas/Epis.csv", index=False, encoding="utf-8")
                    atualizar_tabela_epis()
                    messagebox.showinfo("Sucesso", f"Quantidade atualizada com sucesso!\nCA: {ca}, Nova Quantidade: {nova_quantidade}")
                else:
                    messagebox.showinfo("Operação Cancelada", "A quantidade não foi alterada.")
                return

        if descricao:
            epi_existente_desc = df_epis[df_epis["DESCRICAO"] == descricao]
            if not epi_existente_desc.empty:
                adicionar_quantidade = messagebox.askyesno(
                    "EPI Já Existente",
                    f"Já existe um EPI com esta descrição.\n"
                    f"CA: {epi_existente_desc.iloc[0]['CA']}\n"
                    f"Descrição: {epi_existente_desc.iloc[0]['DESCRICAO']}\n"
                    f"Quantidade Atual: {epi_existente_desc.iloc[0]['QUANTIDADE']}\n\n"
                    f"Deseja adicionar {quantidade} à quantidade existente?"
                )
                if adicionar_quantidade:
                    nova_quantidade = float(epi_existente_desc.iloc[0]["QUANTIDADE"]) + quantidade
                    df_epis.loc[epi_existente_desc.index, "QUANTIDADE"] = nova_quantidade
                    df_epis.to_csv("Planilhas/Epis.csv", index=False, encoding="utf-8")
                    atualizar_tabela_epis()
                    messagebox.showinfo("Sucesso", f"Quantidade atualizada com sucesso!\nDescrição: {descricao}, Nova Quantidade: {nova_quantidade}")
                else:
                    messagebox.showinfo("Operação Cancelada", "A quantidade não foi alterada.")
                return

        confirmacao = messagebox.askyesno(
            "Confirmação",
            f"Você deseja registrar este EPI?\n\n"
            f"CA: {ca if ca else 'Sem CA'}\n"
            f"Descrição: {descricao if descricao else 'Sem Descrição'}\n"
            f"Quantidade: {quantidade}"
        )
        if not confirmacao:
            messagebox.showinfo("Operação Cancelada", "O registro do EPI foi cancelado.")
            return

        with open("Planilhas/Epis.csv", "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([ca, descricao, quantidade])

        messagebox.showinfo("Sucesso", f"EPI registrado com sucesso!\nDescrição: {descricao}, Quantidade: {quantidade}")

        ca_entry.delete(0, tk.END)
        descricao_epi_entry.delete(0, tk.END)
        quantidade_epi_entry.delete(0, tk.END)

        atualizar_tabela_epis()

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao registrar EPI: {e}")
   
     
def registrar_retirada():
    """
    Registra a retirada de um EPI por um colaborador.
    """
    colaborador = colaborador_entry.get().strip().upper()
    identificador = ca_retirada_entry.get().strip().upper()
    quantidade_retirada = quantidade_retirada_entry.get().strip()

    if not colaborador or not identificador:
        messagebox.showerror("Erro", "Colaborador e CA/Descrição devem ser preenchidos.")
        return
        
    try:
        quantidade_retirada = float(quantidade_retirada)
        if quantidade_retirada <= 0:
            raise ValueError("Quantidade deve ser maior que zero.")
    except ValueError:
        messagebox.showerror("Erro", "A quantidade deve ser um número válido e maior que zero.")
        return

    try:
        df_epis = pd.read_csv("Planilhas/Epis.csv", encoding="utf-8", dtype={"CA": str})
        df_epis["CA"] = df_epis["CA"].fillna("").astype(str).str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].fillna("").astype(str).str.strip().str.upper()

        epi = df_epis[(df_epis["CA"].str.upper() == identificador) | 
                      (df_epis["DESCRICAO"].str.upper() == identificador)]

        if epi.empty:
            messagebox.showerror("Erro", f"O EPI com CA ou Descrição '{identificador}' não foi encontrado.")
            return

        descricao = epi.iloc[0]["DESCRICAO"]
        quantidade_disponivel = int(epi.iloc[0]["QUANTIDADE"])

        if quantidade_retirada > quantidade_disponivel:
            messagebox.showerror("Erro", f"Quantidade insuficiente no estoque para o EPI '{descricao}'.")
            return

        pasta_colaborador = os.path.join("Colaboradores", colaborador)
        if not os.path.exists(pasta_colaborador):
            confirmacao_pasta = messagebox.askyesno(
                "Colaborador Não Encontrado",
                f"A pasta para o colaborador '{colaborador}' não foi encontrada. Deseja criá-la?"
            )
            if confirmacao_pasta:
                os.makedirs(pasta_colaborador, exist_ok=True)
            else:
                messagebox.showinfo("Operação Cancelada", "A retirada foi cancelada.")
                return

        confirmacao_retirada = messagebox.askyesno(
            "Confirmação",
            f"Você deseja registrar a retirada deste EPI?\n\n"
            f"Colaborador: {colaborador}\n"
            f"Descrição: {descricao}\n"
            f"Quantidade a retirar: {quantidade_retirada}\n"
            f"Quantidade restante: {quantidade_disponivel - quantidade_retirada}"
        )

        if not confirmacao_retirada:
            messagebox.showinfo("Operação Cancelada", "A retirada foi cancelada.")
            return

        df_epis.loc[(df_epis["CA"] == identificador) | (df_epis["DESCRICAO"] == identificador), "QUANTIDADE"] = quantidade_disponivel - quantidade_retirada
        df_epis.to_csv("Planilhas/Epis.csv", index=False, encoding="utf-8")
        atualizar_tabela_epis()

        nome_arquivo = f"{colaborador}_{datetime.now().strftime('%Y_%m')}.csv"
        caminho_arquivo = os.path.join(pasta_colaborador, nome_arquivo)

        if not os.path.exists(caminho_arquivo):
            with open(caminho_arquivo, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["CA", "DESCRICAO", "QTD RETIRADA", "DATA"])

        data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(caminho_arquivo, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([epi.iloc[0]["CA"], descricao, quantidade_retirada, data])

        messagebox.showinfo("Sucesso", f"Retirada registrada para o colaborador {colaborador}.\n"
                                       f"Descrição: {descricao}, Quantidade: {quantidade_retirada}")

        colaborador_entry.delete(0, tk.END)
        ca_retirada_entry.delete(0, tk.END)
        quantidade_retirada_entry.delete(0, tk.END)

    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo Epis.csv não encontrado.")
        return
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao acessar o arquivo Epis.csv: {e}")
        return    


def atualizar_tabela_epis():
    """
    Atualiza a tabela de EPIs com os dados mais recentes do arquivo Epis.csv.
    """
    global df_epis
    try:
        df_epis = pd.read_csv("Planilhas/Epis.csv", encoding="utf-8")
        epis_table.updateModel(TableModel(df_epis))
        epis_table.redraw()
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo Epis.csv não encontrado.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar a tabela de EPIs: {e}")
    

def exportar_conteudo():
    """
    Exporta o conteúdo das planilhas para um arquivo Excel e gera um relatório de produtos esgotados.
    """
    pasta_saida = "Relatorios"
    os.makedirs(pasta_saida, exist_ok=True)
    
    data_atual = datetime.now().strftime("%d-%m-%Y")
    caminho_excel = os.path.join(pasta_saida, f"Relatorio Almoxarifado {data_atual}.xlsx")
    caminho_txt = os.path.join(pasta_saida, f"Produtos Esgotados {data_atual}.txt")

    try:
        with pd.ExcelWriter(caminho_excel) as writer:
            arquivos_com_epis = {**arquivos, "epis": "Planilhas/Epis.csv"}
            for nome, arquivo in arquivos_com_epis.items():
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
            f.write(f"Relatório de Produtos Esgotados - {data_atual}\n")
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

    arquivo_ultimo_backup = os.path.join(pasta_backup, "ultimo_backup.txt")

    agora = time.time()
    if os.path.exists(arquivo_ultimo_backup):
        with open(arquivo_ultimo_backup, "r", encoding="utf-8") as f:
            ultimo_backup = float(f.read().strip())
        if (agora - ultimo_backup) < 3 * 60 * 60:
            main.after(10800000, criar_backup_periodico)
            return

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    try:
        arquivos_com_epis = {**arquivos, "epis": "Planilhas/Epis.csv"}
        for nome, arquivo in arquivos_com_epis.items():
            if os.path.exists(arquivo):
                nome_backup = f"{nome}_{timestamp}.csv"
                caminho_backup = os.path.join(pasta_backup, nome_backup)
                shutil.copy(arquivo, caminho_backup)

        with open(arquivo_ultimo_backup, "w", encoding="utf-8") as f:
            f.write(str(agora))

        print(f"Backup criado com sucesso em {timestamp}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar backup: {e}")

    try:
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
    
    
def corrigir_planilhas():
    """
    Corrige as planilhas de entrada e saída, preenchendo colunas vazias com '1'.
    """
    planilhas = ["entrada", "saida"]
    colunas_esperadas = {
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA", "ID"]
    }

    for planilha in planilhas:
        caminho = arquivos[planilha]
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                linhas = list(csv.reader(f))

            if len(linhas) > 0 and len(linhas[0]) != len(colunas_esperadas[planilha]):
                linhas[0] = colunas_esperadas[planilha]

            linhas_corrigidas = []
            for linha in linhas:
                if len(linha) < len(colunas_esperadas[planilha]):
                    linha.extend(["1"] * (len(colunas_esperadas[planilha]) - len(linha)))
                elif len(linha) > len(colunas_esperadas[planilha]):
                    linha = linha[:len(colunas_esperadas[planilha])]
                linhas_corrigidas.append(linha)

            with open(caminho, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(linhas_corrigidas)

        except FileNotFoundError:
            print(f"Arquivo {caminho} não encontrado. Ignorando...")
        except Exception as e:
            print(f"Erro ao corrigir {caminho}: {e}")


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


def validar_login():
    """
    Função chamada ao clicar no botão de login.
    """
    global operador_logado_id

    usuario = usuario_entry.get().strip()
    senha = senha_entry.get().strip()

    if usuario in usuarios and senha == usuarios[usuario]["senha"]:
        operador_logado_id = usuarios[usuario]["id"]
        messagebox.showinfo("Sucesso", f"Login realizado com sucesso!\nOperador ID: {operador_logado_id}")
        root.destroy()
    else:
        messagebox.showerror("Erro", "Usuário ou senha inválidos!")


def fechar_login():
    """
    Função chamada ao fechar a janela de login.
    Encerra o programa completamente.
    """
    if messagebox.askyesno("Confirmação", "Deseja realmente sair?"):
        root.destroy()
        os._exit(0)
    

def fechar_aplicacao():
    """
    Função chamada ao fechar a janela principal.
    Encerra o programa completamente.
    """
    if messagebox.askyesno("Confirmação", "Deseja realmente sair?"):
        main.destroy()
        os._exit(0)


def focar_proximo(event):
    """
    Move o foco para o próximo widget quando a tecla Enter é pressionada.
    """
    event.widget.tk_focusNext().focus()
    return


# Configuração da janela de login

root = tk.Tk()
root.title("Almoxarifado")
root.geometry("300x230")
root.resizable(False, False)

root.protocol("WM_DELETE_WINDOW", fechar_login)

titulo_label = tk.Label(root, text="Login", font=("Arial", 16))
titulo_label.pack(pady=10)

usuario_label = tk.Label(root, text="Usuário:", font=("Arial", 12))
usuario_label.pack()
usuario_entry = tk.Entry(root, font=("Arial", 12))
usuario_entry.bind("<Return>", focar_proximo)
usuario_entry.pack(pady=5)

senha_label = tk.Label(root, text="Senha:", font=("Arial", 12))
senha_label.pack()
senha_entry = tk.Entry(root, font=("Arial", 12), show="*")
senha_entry.bind("<Return>", lambda event: validar_login())
senha_entry.pack(pady=5)

login_button = tk.Button(root, text="Entrar", font=("Arial", 12), bg="#67F5A5", command=validar_login)
login_button.pack(pady=10)

root.mainloop()



# Configuração da janela principal

main = tk.Tk()
main.config(bg="#C1BABA")
main.title("Almoxarifado")
main.geometry("1100x600")   
main.resizable(False, False)

main.protocol("WM_DELETE_WINDOW", fechar_aplicacao)

criar_planilhas()
criar_backup_periodico()
corrigir_planilhas()


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
pesquisar_entry.bind("<KeyRelease>", pesquisar_tabela)
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
exportar_button.place(x=494, y=517, width=80, height=43)

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
desc_entry.bind("<Return>", focar_proximo)
desc_entry.config(bg="#fff", fg="#000")
desc_entry.place(x=20, y=90, width=371, height=30)

quantidade_label = tk.Label(master=cadastro_tab, text="Quantidade", font=("Arial", 12))
quantidade_label.place(x=20, y=130)
quantidade_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
quantidade_entry.bind("<Return>", focar_proximo)
quantidade_entry.config(bg="#fff", fg="#000")
quantidade_entry.place(x=20, y=160, width=371, height=30)

valor_label = tk.Label(master=cadastro_tab, text="Valor Unitário", font=("Arial", 12))
valor_label.place(x=20, y=200)
valor_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
valor_entry.bind("<Return>", focar_proximo) 
valor_entry.config(bg="#fff", fg="#000")
valor_entry.place(x=20, y=230, width=371, height=30)

localizacao_label = tk.Label(master=cadastro_tab, text="Localização", font=("Arial", 12))
localizacao_label.place(x=20, y=270)
localizacao_entry = tk.Entry(master=cadastro_tab, font=("Arial", 12))
localizacao_entry.bind("<Return>", lambda event: cadastrar_estoque())
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
codigo_entry.bind("<Return>", focar_proximo)
codigo_entry.config(bg="#fff", fg="#000")
codigo_entry.place(x=80, y=90, width=371, height=30)

quantidade_entrada_label = tk.Label(master=movimentacao_tab, text="Quantidade", font=("Arial", 12))
quantidade_entrada_label.place(x=80, y=130)
quantidade_entrada_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
quantidade_entrada_entry.bind("<Return>", lambda event: registrar_entrada())
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
codigo_saida_entry.bind("<Return>", focar_proximo)
codigo_saida_entry.config(bg="#fff", fg="#000")
codigo_saida_entry.place(x=645, y=90, width=371, height=30)

solicitante_label = tk.Label(master=movimentacao_tab, text="Solicitante", font=("Arial", 12))
solicitante_label.place(x=645, y=130)
solicitante_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
solicitante_entry.bind("<Return>", focar_proximo)
solicitante_entry.config(bg="#fff", fg="#000")
solicitante_entry.place(x=645, y=160, width=371, height=30)

quantidade_saida_label = tk.Label(master=movimentacao_tab, text="Quantidade", font=("Arial", 12))
quantidade_saida_label.place(x=645, y=200)
quantidade_saida_entry = tk.Entry(master=movimentacao_tab, font=("Arial", 12))
quantidade_saida_entry.bind("<Return>", lambda event: registrar_saida())
quantidade_saida_entry.config(bg="#fff", fg="#000")
quantidade_saida_entry.place(x=645, y=230, width=371, height=30)

saida_button = tk.Button(master=movimentacao_tab, text="Registrar Saída", command=registrar_saida)
saida_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
saida_button.place(x=645, y=280, width=371, height=40)



# Aba EPIs

epis_tab = ttk.Frame(notebook)
notebook.add(epis_tab, text="EPIs")

df_epis = pd.read_csv("Planilhas/Epis.csv", encoding="utf-8")

epis_table_frame = tk.Frame(master=epis_tab)
epis_table_frame.place(x=20, y=20, width=650, height=530)

epis_table = Table(parent=epis_table_frame, dataframe=df_epis)
epis_table.show()

ca_label = tk.Label(master=epis_tab, text="CA", font=("Arial", 12))
ca_label.place(x=700, y=20)
ca_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
ca_entry.bind("<Return>", focar_proximo)
ca_entry.place(x=700, y=50, width=300, height=30)

descricao_epi_label = tk.Label(master=epis_tab, text="Descrição", font=("Arial", 12))
descricao_epi_label.place(x=700, y=90)
descricao_epi_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
descricao_epi_entry.bind("<Return>", focar_proximo)
descricao_epi_entry.place(x=700, y=120, width=300, height=30)

quantidade_epi_label = tk.Label(master=epis_tab, text="Quantidade", font=("Arial", 12))
quantidade_epi_label.place(x=700, y=160)
quantidade_epi_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
quantidade_epi_entry.bind("<Return>", lambda event: registrar_epi())    
quantidade_epi_entry.place(x=700, y=190, width=300, height=30)

registrar_epi_button = tk.Button(master=epis_tab, text="Registrar EPI", command=lambda: registrar_epi())
registrar_epi_button.config(bg="#67F5A5", fg="#000", font=("Arial", 12))
registrar_epi_button.place(x=700, y=230, width=300, height=40)

ca_retirada_label = tk.Label(master=epis_tab, text="CA ou Descrição", font=("Arial", 12))
ca_retirada_label.place(x=700, y=300)
ca_retirada_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
ca_retirada_entry.bind("<Return>", focar_proximo)
ca_retirada_entry.place(x=700, y=330, width=300, height=30)

quantidade_retirada_label = tk.Label(master=epis_tab, text="Quantidade", font=("Arial", 12))
quantidade_retirada_label.place(x=700, y=370)
quantidade_retirada_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
quantidade_retirada_entry.bind("<Return>", focar_proximo)
quantidade_retirada_entry.place(x=700, y=400, width=300, height=30)

colaborador_label = tk.Label(master=epis_tab, text="Colaborador", font=("Arial", 12))
colaborador_label.place(x=700, y=440)
colaborador_entry = tk.Entry(master=epis_tab, font=("Arial", 12))
colaborador_entry.bind("<Return>", lambda event: registrar_retirada())
colaborador_entry.place(x=700, y=470, width=300, height=30)

registrar_retirada_button = tk.Button(master=epis_tab, text="Registrar Retirada", command=lambda: registrar_retirada())
registrar_retirada_button.config(bg="#54befc", fg="#000", font=("Arial", 12))
registrar_retirada_button.place(x=700, y=510, width=300, height=40)


main.mainloop()
