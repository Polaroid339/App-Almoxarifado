import flet as ft
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

# Função para carregar os dados do estoque
def carregar_estoque():
    if os.path.exists(arquivos["estoque"]):
        return pd.read_csv(arquivos["estoque"], encoding="utf-8").values.tolist()
    return []

# Função para adicionar um item ao estoque
def adicionar_item(codigo, descricao, valor_un, quantidade, localizacao):
    data = datetime.now().strftime("%H:%M %d/%m/%Y")
    valor_total = float(valor_un) * int(quantidade)
    with open(arquivos["estoque"], "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([codigo, descricao, valor_un, valor_total, quantidade, data, localizacao])

# Função para exportar os dados para Excel
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
                pass
    return caminho_arquivo

# Função principal do app
def main(page: ft.Page):
    page.title = "Gestão de Almoxarifado"
    criar_planilhas()
    
    tabela_estoque = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("Código")),
            ft.DataColumn(ft.Text("Descrição")),
            ft.DataColumn(ft.Text("Valor Unitário")),
            ft.DataColumn(ft.Text("Valor Total")),
            ft.DataColumn(ft.Text("Quantidade")),
            ft.DataColumn(ft.Text("Data")),
            ft.DataColumn(ft.Text("Localização"))
        ],
        rows=[ft.DataRow(cells=[ft.DataCell(ft.Text(str(item))) for item in linha]) for linha in carregar_estoque()]
    )

    # Campos de entrada
    codigo_input = ft.TextField(label="Código")
    descricao_input = ft.TextField(label="Descrição")
    valor_un_input = ft.TextField(label="Valor Unitário")
    quantidade_input = ft.TextField(label="Quantidade")
    localizacao_input = ft.TextField(label="Localização")
    
    # Função para cadastrar produto
    def cadastrar_produto(e):
        adicionar_item(
            codigo_input.value, descricao_input.value.upper(), valor_un_input.value,
            quantidade_input.value, localizacao_input.value.upper()
        )
        tabela_estoque.rows.append(ft.DataRow(cells=[
            ft.DataCell(ft.Text(codigo_input.value)),
            ft.DataCell(ft.Text(descricao_input.value.upper())),
            ft.DataCell(ft.Text(valor_un_input.value)),
            ft.DataCell(ft.Text(str(float(valor_un_input.value) * int(quantidade_input.value)))),
            ft.DataCell(ft.Text(quantidade_input.value)),
            ft.DataCell(ft.Text(datetime.now().strftime("%H:%M %d/%m/%Y"))),
            ft.DataCell(ft.Text(localizacao_input.value.upper()))
        ]))
        page.update()
        page.snack_bar = ft.SnackBar(ft.Text("Produto cadastrado com sucesso!"))
        page.snack_bar.open = True
        page.update()
    
    # Botão de exportação
    def exportar_estoque(e):
        caminho = exportar_para_excel()
        page.snack_bar = ft.SnackBar(ft.Text(f"Relatório exportado para {caminho}"))
        page.snack_bar.open = True
        page.update()
    
    form_cadastro = ft.Column([
        codigo_input, descricao_input, valor_un_input, quantidade_input, localizacao_input,
        ft.ElevatedButton("Cadastrar Produto", on_click=cadastrar_produto),
        ft.ElevatedButton("Exportar para Excel", on_click=exportar_estoque)
    ])
    
    page.add(
        ft.Text("Controle de Almoxarifado", size=24, weight=ft.FontWeight.BOLD),
        tabela_estoque,
        form_cadastro
    )

ft.app(target=main)
