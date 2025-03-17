import csv
import os
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
        "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA"],
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA"]
    }
    
    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            with open(arquivo, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(colunas[nome])
    print("Planilhas criadas/verificadas com sucesso!")

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
    descricao = input("Descrição do produto: ")
    valor_un = float(input("Valor unitário (R$): "))
    quantidade = int(input("Quantidade: "))
    data = datetime.now().strftime("%d/%m/%Y")
    valor_total = quantidade * valor_un
    
    with open(arquivos["estoque"], "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([codigo, descricao, valor_un, valor_total, quantidade, data])
    print(f"Produto cadastrado no estoque com código {codigo}!")

# Função para registrar entrada de produto
def registrar_entrada():
    codigo = input("Código do produto: ")
    produto = buscar_produto(codigo)
    
    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = input("Deseja registrar entrada neste produto? (S/N): ").strip().lower()
        if confirmacao == "s":
            quantidade_adicionada = int(input("Quantidade adicionada: "))
            nova_quantidade = int(produto[4]) + quantidade_adicionada
            data = datetime.now().strftime("%d/%m/%Y")
            
            with open(arquivos["entrada"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([codigo, produto[1], quantidade_adicionada, produto[2], float(produto[2]) * quantidade_adicionada, data])
            
            atualizar_estoque(codigo, nova_quantidade)
            print("Entrada registrada e estoque atualizado!")
        else:
            print("Operação cancelada.")
    else:
        print("Código do produto não encontrado.")

# Função para registrar saída de produto
def registrar_saida():
    codigo = input("Código do produto: ")
    produto = buscar_produto(codigo)
    
    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = input("Deseja dar saída neste produto? (S/N): ").strip().lower()
        if confirmacao == "s":
            quantidade_retirada = int(input("Quantidade retirada: "))
            if quantidade_retirada > int(produto[4]):
                print("Quantidade insuficiente no estoque!")
                return
            solicitante = input("Nome do solicitante: ")
            nova_quantidade = int(produto[4]) - quantidade_retirada
            data = datetime.now().strftime("%d/%m/%Y")
            
            with open(arquivos["saida"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([codigo, produto[1], quantidade_retirada, solicitante, data])
            
            atualizar_estoque(codigo, nova_quantidade)
            print("Saída registrada e estoque atualizado!")
        else:
            print("Operação cancelada.")
    else:
        print("Código do produto não encontrado.")

# Menu de opções
def menu():
    criar_planilhas()
    while True:
        print("\n1 - Cadastrar produto no estoque")
        print("2 - Registrar entrada de produto")
        print("3 - Registrar saída de produto")
        print("4 - Sair")
        opcao = input("Escolha uma opção: ")
        
        if opcao == "1":
            cadastrar_estoque()
        elif opcao == "2":
            registrar_entrada()
        elif opcao == "3":
            registrar_saida()
        elif opcao == "4":
            print("Saindo...")
            break
        else:
            print("Opção inválida!")

# Execução do programa
if __name__ == "__main__":
    menu()