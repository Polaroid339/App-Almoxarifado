import csv
import os
import pandas as pd
from datetime import datetime
from tabulate import tabulate
from formatter import logger
import shutil

"""
Gestão de Almoxarifado
Sistema para controle de estoque e movimentações de produtos
"""

# Definição dos arquivos CSV
arquivos = {
    "estoque": "./Planilhas/Estoque.csv",
    "entrada": "./Planilhas/Entrada.csv",
    "saida": "./Planilhas/Saida.csv"
}


# Função para criar os arquivos CSV caso não existam

def criar_planilhas():
    colunas = {
        "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"],
        "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA"],
        "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA"]
    }

    os.makedirs("Planilhas", exist_ok=True)

    for nome, arquivo in arquivos.items():
        if not os.path.exists(arquivo):
            with open(arquivo, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(colunas[nome])


# Funções para entrada de dados

def inputm(mensagem):
    entrada = input(mensagem).strip().upper()
    if entrada == "MENU":
        os.system('cls')
        menu()
        return None
    return entrada


def entrada_inteiro(mensagem):
    while True:
        try:
            return int(inputm(mensagem))
        except ValueError:
            logger.warning("Erro! Digite um número inteiro válido.")


def entrada_float(mensagem):
    while True:
        try:
            return float(inputm(mensagem))
        except ValueError:
            logger.warning("Erro! Digite um número decimal válido.")


# Função para exibir relatório completo do estoque

def exibir_relatorio():
    os.system('cls')
    try:
        df = pd.read_csv(arquivos["estoque"], encoding="utf-8")

        if df.empty:
            logger.warning("\nO estoque está vazio!\n")
            return

        total_produtos = len(df)
        itens_por_pagina = 20
        total_paginas = ((total_produtos - 1) // itens_por_pagina) + 1
        pagina_atual = 0

        while True:
            os.system('cls')
            inicio = pagina_atual * itens_por_pagina
            fim = inicio + itens_por_pagina
            pagina = df.iloc[inicio:fim]

            logger.info("\nRelatório de Estoque\n")
            print(tabulate(pagina, headers='keys',
                  tablefmt='fancy_grid', showindex=False))

            logger.debug(f"\nPágina {pagina_atual + 1} de {total_paginas}")
            print("\n1 - Próxima página")
            print("2 - Página anterior")
            print("3 - Ir para uma página específica")
            print("4 - Registrar entrada de produto")
            print("5 - Registrar saída de produto")
            print("6 - Voltar ao menu")

            opcao = inputm("\n> Escolha uma opção: ")

            if opcao == "1" and fim < total_produtos:
                pagina_atual += 1
            elif opcao == "2" and pagina_atual > 0:
                pagina_atual -= 1
            elif opcao == "3":
                num_pagina = inputm(
                    f"\n> Digite um número de página (1-{total_paginas}): ")
                if num_pagina.isdigit():
                    num_pagina = int(num_pagina) - 1
                    if 0 <= num_pagina < total_paginas:
                        pagina_atual = num_pagina
                    else:
                        logger.warning("\nPágina inválida! Tente novamente.")
                else:
                    logger.warning("\nEntrada inválida! Digite um número.")
            elif opcao == "4":
                registrar_entrada()
            elif opcao == "5":
                registrar_saida()
            elif opcao == "6":
                os.system('cls')
                break
            else:
                logger.warning("\nOpção inválida! Tente novamente.")
    except FileNotFoundError:
        os.system('cls')
        logger.warning("\nArquivo de estoque não encontrado.\n")


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
    logger.info(f"Relatórios exportados para {caminho_arquivo}")


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
    descricao = inputm("> Descrição do produto: ").upper()
    quantidade = entrada_inteiro("> Quantidade: ")
    valor_un = entrada_float("> Valor unitário (R$): ")
    localizacao = inputm("> Localização do produto: ").upper()
    data = datetime.now().strftime("%H:%M %d/%m/%Y")

    confirmacao = inputm(
        "> Deseja cadastrar este produto? (S/N): ").strip().upper()
    if confirmacao == "S":
        valor_total = quantidade * valor_un

        with open(arquivos["estoque"], "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([codigo, descricao, valor_un,
                            valor_total, quantidade, data, localizacao])
        os.system('cls')
        logger.info(f"Produto cadastrado no estoque com código {codigo}!")
    else:
        os.system('cls')
        logger.warning("Operação cancelada.")


# Função para registrar entrada de produto

def registrar_entrada():
    codigo = inputm("> Código do produto: ")
    print()
    produto = buscar_produto(codigo)

    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = inputm(
            "> Deseja registrar entrada neste produto? (S/N): ").strip().upper()
        if confirmacao == "S":
            quantidade_adicionada = entrada_inteiro(
                "> Quantidade adicionada: ")
            nova_quantidade = int(produto[4]) + quantidade_adicionada
            data = datetime.now().strftime("%H:%M %d/%m/%Y")

            with open(arquivos["entrada"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow([codigo, produto[1], quantidade_adicionada, produto[2], float(
                    produto[2]) * quantidade_adicionada, data])

            atualizar_estoque(codigo, nova_quantidade)
            os.system('cls')
            logger.info("Entrada registrada e estoque atualizado!")
        else:
            os.system('cls')
            logger.warning("Operação cancelada.")
    else:
        os.system('cls')
        logger.warning("Código do produto não encontrado.")


# Função para registrar saída de produto

def registrar_saida():
    codigo = inputm("> Código do produto: ")
    print()
    produto = buscar_produto(codigo)

    if produto:
        print(f"Produto encontrado: {produto[1]}")
        confirmacao = inputm(
            "> Deseja dar saída neste produto? (S/N): ").strip().upper()
        if confirmacao == "S":
            quantidade_retirada = entrada_inteiro("> Quantidade retirada: ")
            if quantidade_retirada > int(produto[4]):
                os.system('cls')
                logger.warning("Quantidade insuficiente no estoque!")
                return
            solicitante = inputm("> Nome do solicitante: ").upper()
            nova_quantidade = int(produto[4]) - quantidade_retirada
            data = datetime.now().strftime("%H:%M %d/%m/%Y")

            with open(arquivos["saida"], "a", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(
                    [codigo, produto[1], quantidade_retirada, solicitante, data])

            atualizar_estoque(codigo, nova_quantidade)
            os.system('cls')
            logger.info("Saída registrada e estoque atualizado!")
        else:
            os.system('cls')
            logger.warning("Operação cancelada.")
    else:
        os.system('cls')
        logger.warning("Código do produto não encontrado.")


# Função para editar um produto no estoque

def editar_produto():
    codigo = inputm("> Código do produto a ser editado: ")
    print()
    produto = buscar_produto(codigo)

    if produto:
        print(f"Produto encontrado: {produto}\n")
        nova_descricao = inputm(
            f"> Nova descrição ({produto[1]}): ").upper() or produto[1]
        novo_valor = float(inputm(
            f"> Novo valor unitário ({produto[2]}): ")) or produto[2]
        nova_localizacao = inputm(
            f"> Nova localização ({produto[6]}): ").upper() or produto[6]

        with open(arquivos["estoque"], "r", encoding="utf-8") as f:
            produtos = list(csv.reader(f))

        for linha in produtos:
            if linha[0] == codigo:
                linha[1], linha[2], linha[6] = nova_descricao, str(
                    novo_valor), nova_localizacao

        with open(arquivos["estoque"], "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(produtos)

        os.system('cls')
        logger.info("Produto atualizado com sucesso!")
    else:
        os.system('cls')
        logger.warning("Código do produto não encontrado.")


# Função para pesquisar um produto no estoque

def pesquisar_produto():
    nome_busca = inputm(
        "> Pesquisar produto (COD; DES; LOC; DAT): ").strip().upper()

    try:
        df = pd.read_csv(arquivos["estoque"], encoding="utf-8")
        resultado = df[
            df["CODIGO"].astype(str).str.contains(nome_busca, na=False, case=False) |
            df["DESCRICAO"].str.contains(nome_busca, na=False, case=False) |
            df["LOCALIZACAO"].str.contains(nome_busca, na=False, case=False) |
            df["DATA"].str.contains(nome_busca, na=False, case=False)
        ]
        if not resultado.empty:
            total_produtos = len(resultado)
            itens_por_pagina = 20
            total_paginas = ((total_produtos - 1) // itens_por_pagina) + 1
            pagina_atual = 0

            while True:
                os.system('cls')
                inicio = pagina_atual * itens_por_pagina
                fim = inicio + itens_por_pagina
                pagina = resultado.iloc[inicio:fim]

                logger.info("\nProdutos encontrados:\n")
                print(tabulate(pagina, headers='keys',
                      tablefmt='fancy_grid', showindex=False))

                logger.debug(f"\nPágina {pagina_atual + 1} de {total_paginas}")
                print("\n1 - Próxima página")
                print("2 - Página anterior")
                print("3 - Ir para uma página específica")
                print("4 - Registrar entrada de produto")
                print("5 - Registrar saída de produto")
                print("6 - Pesquisar outro produto")
                print("7 - Voltar ao menu")

                opcao = inputm("\n> Escolha uma opção: ")

                if opcao == "1" and fim < total_produtos:
                    pagina_atual += 1
                elif opcao == "2" and pagina_atual > 0:
                    pagina_atual -= 1
                elif opcao == "3":
                    num_pagina = inputm(
                        f"\n> Digite um número de página (1-{total_paginas}): ")
                    if num_pagina.isdigit():
                        num_pagina = int(num_pagina) - 1
                        if 0 <= num_pagina < total_paginas:
                            pagina_atual = num_pagina
                        else:
                            logger.warning(
                                "\nPágina inválida! Tente novamente.")
                    else:
                        logger.warning("\nEntrada inválida! Digite um número.")

                elif opcao == "4":
                    registrar_entrada()
                elif opcao == "5":
                    registrar_saida()
                elif opcao == "6":
                    os.system('cls')
                    pesquisar_produto()
                elif opcao == "7":
                    os.system('cls')
                    break
                else:
                    logger.warning("\nOpção inválida! Tente novamente.")
        else:
            os.system('cls')
            logger.warning("Nenhum produto encontrado com esse nome.")
    except FileNotFoundError:
        os.system('cls')
        logger.warning("Arquivo de estoque não encontrado.")


# Função para excluir um produto do estoque

def excluir_produto():
    codigo = inputm("> Código do produto a ser excluído: ").strip()
    print()
    produto = buscar_produto(codigo)

    if produto:
        print(f"Produto encontrado: {produto[1]}")
        logger.error(
            "Excluir produtos não é recomendado, pois desordena a ordem dos códigos.")
        confirmacao = inputm(
            "> Tem certeza que deseja excluir este produto? (S/N): ").strip().upper()

        if confirmacao == "S":
            with open(arquivos["estoque"], "r", encoding="utf-8") as f:
                produtos = list(csv.reader(f))

            # Filtrar para manter apenas os produtos diferentes do código informado
            produtos = [linha for linha in produtos if linha[0] != codigo]

            with open(arquivos["estoque"], "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerows(produtos)

            os.system('cls')
            logger.info(f"Produto {produto[1]} excluído com sucesso!")
        else:
            os.system('cls')
            logger.warning("Operação cancelada.")
    else:
        os.system('cls')
        logger.warning("Código do produto não encontrado.")

    # Função para exibir relatório de produtos esgotados e exportá-los para um arquivo .txt


def itens_esgotados():
    os.system('cls')
    try:
        # Carregar o estoque atual
        df = pd.read_csv(arquivos["estoque"], encoding="utf-8")

        if df.empty:
            logger.warning("\nO estoque está vazio!\n")
            return

        # Filtrar produtos com quantidade igual ou menor que zero
        faltantes = df[df["QUANTIDADE"].astype(int) <= 0]

        if faltantes.empty:
            logger.info("\nNão há produtos esgotados.\n")
            return

        # Exibir o relatório na tela
        logger.info("\nProdutos esgotados do estoque:\n")
        print(tabulate(faltantes[["CODIGO", "DESCRICAO"]],
              headers='keys', tablefmt='fancy_grid', showindex=False))
        print()
        confirmacao = inputm(
            "> Deseja exportar os itens esgotados? (S/N): ").strip().upper()
        if confirmacao == "S":

            # Criar a pasta Relatorios caso não exista
            os.makedirs("Relatorios", exist_ok=True)

            # Exportar para arquivo .txt
            caminho_arquivo = os.path.join(
                "Relatorios", "Produtos_Esgotados.txt")
            with open(caminho_arquivo, "w", encoding="utf-8") as f:
                f.write("Relatório de produtos esgotados\n")
                f.write("-" * 40 + "\n")
                for index, row in faltantes.iterrows():
                    f.write(
                        f"Código: {row['CODIGO']} | Descrição: {row['DESCRICAO']}\n")
                f.write("-" * 40 + "\n")

            os.system('cls')
            logger.info(
                f"Relatório de produtos exgotados exportado para {caminho_arquivo}")
        else:
            os.system('cls')
            logger.warning("Operação cancelada.")

    except FileNotFoundError:
        os.system('cls')
        logger.warning("\nArquivo de estoque não encontrado.\n")


# Menu de opções

def menu():
    criar_planilhas()
    if os.path.exists("./Planilhas/Estoque.csv"):
        shutil.copy("./Planilhas/Estoque.csv", "./Planilhas/Estoque_backup.csv")

    df = pd.read_csv("./Planilhas/Estoque.csv")
    df['VALOR TOTAL'] = df['VALOR UN'] * df['QUANTIDADE']
    df.to_csv("./Planilhas/Estoque.csv", index=False)

    while True:
        print("\033[1;36;40m+----------------------------------------------------------------+\033[0;37;40m")
        print("\033[1;36;40m|                   [ GESTÃO DE ALMOXARIFADO ]                   |\033[0;37;40m")
        print("\033[1;36;40m+----------------------------------------------------------------+\033[0;37;40m")
        print("\033[1;36;40m|                                                                |\033[0;37;40m")
        print("\033[1;36;40m|                         MENU PRINCIPAL                         |\033[0;37;40m")
        print("\033[1;36;40m|                                                                |\033[0;37;40m")
        print("\033[1;36;40m|               [1]\033[0;37;40m..Cadastrar produto no estoque                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [2]\033[0;37;40m..Registrar entrada de produto                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [3]\033[0;37;40m....Registrar saída de produto                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [4]\033[0;37;40m...Exibir relatório de estoque                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [5]\033[0;37;40m..Pesquisar produto no estoque                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [6]\033[0;37;40m.Exportar planilhas para Excel                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [7]\033[0;37;40m......Exportar itens esgotados                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [8]\033[0;37;40m.....Editar produto no estoque                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               [9]\033[0;37;40m....Excluir produto do estoque                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|               \033[1;31;40m[0]\033[0;37;40m..........................Sair                \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|                                                                |\033[0;37;40m")
        print("\033[1;36;40m|\033[0;37;40m              Digite 'menu' a qualquer momento para             \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|\033[0;37;40m                   voltar ao menu principal.                    \033[1;36;40m|\033[0;37;40m")
        print("\033[1;36;40m|                                                                |\033[0;37;40m")
        print("\033[1;36;40m+----------------------------------------------------------------+\033[0;37;40m\n")
        opcao = inputm("> Escolha uma opção: ")
        print()

        try:
            if opcao == "1":
                cadastrar_estoque()
            elif opcao == "2":
                registrar_entrada()
            elif opcao == "3":
                registrar_saida()
            elif opcao == "4":
                exibir_relatorio()
            elif opcao == "5":
                pesquisar_produto()
            elif opcao == "6":
                exportar_para_excel()
            elif opcao == "7":
                itens_esgotados()
            elif opcao == "8":
                editar_produto()
            elif opcao == "9":
                excluir_produto()
            elif opcao == "0":
                print("Saindo...")
                break
            else:
                os.system('cls')
                print("Opção inválida!")

        except Exception as e:
            os.system('cls')
            logger.critical(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    menu()

# python -m PyInstaller --onefile --name=Almoxarifado main.py
 