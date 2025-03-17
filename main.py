import csv
import os
import pandas as pd
from datetime import datetime

# Definição dos arquivos CSV
arquivos = {
    "estoque": "Estoque.csv",
    "entrada": "Entrada.csv",
    "saida": "Saida.csv"
}

# Função para criar os arquivos CSV caso não existam
def criar_planilhas():
    colunas = {
        "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"],
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA"]
    }
    
    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            with open(arquivo, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(colunas[nome])
    print("\nPlanilhas criadas/verificadas com sucesso!")

# Função para exibir relatório completo do estoque
def exibir_relatorio():
    os.system('cls')
    try:
        df = pd.read_csv(arquivos["estoque"], encoding="utf-8")
        print("\nRelatório de Estoque:")
        print(df.to_string(index=False))
    except FileNotFoundError:
        print("Arquivo de estoque não encontrado.")

# Função para exportar as planilhas para Excel
def exportar_para_excel():
    pasta_saida = "Relatorios"
    os.makedirs(pasta_saida, exist_ok=True)
    caminho_arquivo = os.path.join(pasta_saida, "Relatorio_Almoxarifado.xlsx")
    
    with pd.ExcelWriter(caminho_arquivo) as writer:
        for nome, arquivo in arquivos.items():
            try:
                df = pd.read_csv(arquivo, encoding="utf-8")
                df.to_excel(writer, sheet_name=nome.capitalize(), index=False)
            except FileNotFoundError:
                print(f"Arquivo {arquivo} não encontrado.")
    os.system('cls')
    print(f"Relatórios exportados para {caminho_arquivo}")

# Função para obter o próximo código disponível
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

# Função para buscar detalhes de um produto pelo código
def buscar_produto(codigo):
    try:
        with open(arquivos["estoque"], "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            next(reader)  # Pular cabeçalho
            for row in reader:
                if row[0] == codigo:
                    return row  # Retorna a linha completa do produto
    except FileNotFoundError:
        pass
    return None

# Função para atualizar o estoque
def atualizar_estoque(codigo, nova_quantidade):
    with open(arquivos["estoque"], "r", encoding="utf-8") as f:
        produtos = list(csv.reader(f))
    
    for produto in produtos:
        if produto[0] == codigo:
            produto[4] = str(nova_quantidade)
            produto[3] = str(float(produto[2]) * int(nova_quantidade))
    
    with open(arquivos["estoque"], "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(produtos)

# Função para cadastrar um item no estoque
def cadastrar_estoque():
    codigo = obter_proximo_codigo()
    descricao = input("Descrição do produto: ").upper()
    valor_un = float(input("Valor unitário (R$): "))
    quantidade = int(input("Quantidade: "))
    localizacao = input("Localização do produto: ").upper()
    data = datetime.now().strftime("%H:%M %d/%m/%Y")
    valor_total = quantidade * valor_un
    
    with open(arquivos["estoque"], "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([codigo, descricao, valor_un, valor_total, quantidade, data, localizacao])
    os.system('cls')
    print(f"Produto cadastrado no estoque com código {codigo}!")

# Função para registrar entrada de produto
def registrar_entrada():
    codigo = input("Código do produto: ")
    produto = buscar_produto(codigo)
    
    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = input("Deseja registrar entrada neste produto? (S/N): ").strip().upper()
        if confirmacao == "S":
            quantidade_adicionada = int(input("Quantidade adicionada: "))
            nova_quantidade = int(produto[4]) + quantidade_adicionada
            data = datetime.now().strftime("%H:%M %d/%m/%Y")
            
            with open(arquivos["entrada"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([codigo, produto[1], quantidade_adicionada, produto[2], float(produto[2]) * quantidade_adicionada, data])
            
            atualizar_estoque(codigo, nova_quantidade)
            os.system('cls')
            print("Entrada registrada e estoque atualizado!")
        else:
            os.system('cls')
            print("Operação cancelada.")
    else:
        os.system('cls')
        print("Código do produto não encontrado.")

# Função para registrar saída de produto
def registrar_saida():
    codigo = input("Código do produto: ")
    produto = buscar_produto(codigo)
    
    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = input("Deseja dar saída neste produto? (S/N): ").strip().upper()
        if confirmacao == "S":
            quantidade_retirada = int(input("Quantidade retirada: "))
            if quantidade_retirada > int(produto[4]):
                os.system('cls')
                print("Quantidade insuficiente no estoque!")
                return
            solicitante = input("Nome do solicitante: ").upper()
            nova_quantidade = int(produto[4]) - quantidade_retirada
            data = datetime.now().strftime("%H:%M %d/%m/%Y")
            
            with open(arquivos["saida"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([codigo, produto[1], quantidade_retirada, solicitante, data])
            
            atualizar_estoque(codigo, nova_quantidade)
            os.system('cls')
            print("Saída registrada e estoque atualizado!")
        else:
            os.system('cls')
            print("Operação cancelada.")
    else:
        os.system('cls')
        print("Código do produto não encontrado.")

# Menu de opções
def menu():
    criar_planilhas()
    while True:
        print("\n1 - Cadastrar produto no estoque")
        print("2 - Registrar entrada de produto")
        print("3 - Registrar saída de produto")
        print("4 - Exibir relatório de estoque")
        print("5 - Exportar planilhas para Excel")
        print("6 - Sair\n")
        opcao = input("Escolha uma opção: ")
        
        if opcao == "1":
            cadastrar_estoque()
        elif opcao == "2":
            registrar_entrada()
        elif opcao == "3":
            registrar_saida()
        elif opcao == "4":
            exibir_relatorio()
        elif opcao == "5":
            exportar_para_excel()
        elif opcao == "6":
            print("Saindo...")
            break
        else:
            os.system('cls')
            print("Opção inválida!")

# Execução do programa
if __name__ == "__main__":
    menu()
