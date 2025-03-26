import flet as ft
import pandas as pd
import os
from datetime import datetime
from gestao import atualizar_estoque

# Caminho dos arquivos CSV
DIR_PLANILHAS = "Planilhas"
ARQUIVO_ESTOQUE = os.path.join(DIR_PLANILHAS, "Estoque.csv")

# Criar diretório e arquivo CSV caso não existam
def inicializar_estoque():
    os.makedirs(DIR_PLANILHAS, exist_ok=True)
    if not os.path.exists(ARQUIVO_ESTOQUE):
        df = pd.DataFrame(columns=["CODIGO", "DESCRICAO", "VALOR UN",
                          "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"])
        df.to_csv(ARQUIVO_ESTOQUE, index=False)

def carregar_estoque():
    if os.path.exists(ARQUIVO_ESTOQUE):
        return pd.read_csv(ARQUIVO_ESTOQUE)
    return pd.DataFrame(columns=["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"])

def main(page: ft.Page):
    
    page.title = "Gestão de Almoxarifado"
    page.scroll = "adaptive"
    page.window.maximized = True
    page.padding = 30
    page.theme_mode = ft.ThemeMode.LIGHT
    #page.bgcolor = ft.Colors.WHITE
    
    inicializar_estoque()

    def atualizar_tabela():
        df = carregar_estoque()
        tabela.rows = [
            ft.DataRow(cells=[ft.DataCell(ft.Text(str(row[col])))
                       for col in df.columns])
            for _, row in df.iterrows()
        ]
        page.update()

    tabela = ft.DataTable(
        columns=[ft.DataColumn(ft.Text(col))
                 for col in carregar_estoque().columns],
        rows=[],
        expand=True,  # Permite que a tabela ocupe o espaço corretamente
        horizontal_lines= ft.BorderSide(1),
        vertical_lines= ft.BorderSide(1)
    )
    atualizar_tabela()

    def cadastrar_produto(e):
        df = carregar_estoque()
        codigo = len(df) + 1
        descricao = descricao_input.value
        valor_un = float(valor_un_input.value)
        quantidade = int(quantidade_input.value)
        valor_total = valor_un * quantidade
        localizacao = localizacao_input.value
        data = datetime.now().strftime("%H:%M %d/%m/%Y")
        novo_produto = pd.DataFrame(
            [[codigo, descricao, valor_un, valor_total, quantidade, data, localizacao]], columns=df.columns)
        df = pd.concat([df, novo_produto], ignore_index=True)
        df.to_csv(ARQUIVO_ESTOQUE, index=False)
        atualizar_tabela()
        atualizar_estoque()   
        
        descricao_input.value = valor_un_input.value = quantidade_input.value = localizacao_input.value = ""
        resultado.value = f"Produto cadastrado com sucesso! Código: {codigo}"
        resultado.color = ft.Colors.GREEN
        page.update()
        
    def pesquisar_produto():
        nome_busca = pesquiar_input.value

        try:
            df = pd.read_csv(ARQUIVO_ESTOQUE)
            resultado = df[
                df["CODIGO"].astype(str).str.contains(nome_busca, na=False, case=False) |
                df["DESCRICAO"].str.contains(nome_busca, na=False, case=False) |
                df["LOCALIZACAO"].str.contains(nome_busca, na=False, case=False) |
                df["DATA"].str.contains(nome_busca, na=False, case=False)
            ]
            tabela.rows = [
                ft.DataRow(cells=[ft.DataCell(ft.Text(str(row[col])))
                           for col in resultado.columns])
                for _, row in resultado.iterrows()
            ]
            page.update()

        except FileNotFoundError:
            os.system('cls')

    descricao_input = ft.TextField(label="Descrição")
    
    valor_un_input = ft.TextField(
        label="Valor Unitário",
        keyboard_type="number",
        prefix_text="R$ ",
        input_filter=ft.InputFilter(allow=True, regex_string=r"^[0-9]*\.?[0-9]*$", replacement_string=""))
    
    quantidade_input = ft.TextField(
        label="Quantidade",
        input_filter=ft.InputFilter(allow=True, regex_string=r"^[0-9]*$", replacement_string=""))
    
    localizacao_input = ft.TextField(label="Localização")
    
    pesquiar_input = ft.TextField(label="Pesquisar")
    
    resultado = ft.Text("")

    tabs = ft.Tabs(
        selected_index=0,
        animation_duration=300,
        tabs=[
            ft.Tab(
                text="Cadastrar",
                content=ft.Column(
                    [
                        ft.Text("Cadastro de Produto", size=20, weight="bold"),
                        descricao_input,
                        valor_un_input,
                        quantidade_input,
                        localizacao_input,
                        ft.ElevatedButton("Cadastrar Produto",
                                          on_click=cadastrar_produto),
                        ft.Divider(),
                        resultado
                    ], spacing=10, horizontal_alignment=ft.CrossAxisAlignment.CENTER       
                ),
            ),
            ft.Tab(
                text="Estoque",
                content=ft.Column([
                    ft.Text("Aqui está a seção de Estoque."),
                    pesquiar_input,
                    ft.ElevatedButton("Pesquisar", on_click=pesquisar_produto),
                    ft.Divider(),
                    ft.Container(content=tabela, expand=True)
                ]),
            ),
            ft.Tab(
                text="Configurações",
                content=ft.Column([
                    ft.Text("Ajuste as configurações do sistema aqui."),
                    ft.Switch(label="Modo escuro")
                ]),
            ),
        ])

    page.add(tabs)

ft.app(target=main)