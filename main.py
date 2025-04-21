import os
import csv
import time
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from usuarios import usuarios  # Assume que usuarios.py está no mesmo diretório
from pandastable import Table, TableModel

# --- Configuração ---
PLANILHAS_DIR = "Planilhas"
BACKUP_DIR = "Backups"
COLABORADORES_DIR = "Colaboradores"

ARQUIVOS = {
    "estoque": os.path.join(PLANILHAS_DIR, "Estoque.csv"),
    "entrada": os.path.join(PLANILHAS_DIR, "Entrada.csv"),
    "saida": os.path.join(PLANILHAS_DIR, "Saida.csv"),
    "epis": os.path.join(PLANILHAS_DIR, "Epis.csv")
}

# --- Classes Auxiliares para Diálogos ---
# ... (O código das classes LookupDialog, EditDialogBase, EditProductDialog, EditEPIDialog permanece o mesmo) ...
class LookupDialog(tk.Toplevel):
    """Um diálogo de busca simples e pesquisável."""
    def __init__(self, parent, title, df, search_cols, return_col):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x400")
        self.transient(parent)
        self.grab_set()

        self.df_full = df.copy()
        self.search_cols = search_cols
        self.return_col = return_col
        self.result = None

        # Frame de Busca
        search_frame = ttk.Frame(self, padding="5")
        search_frame.pack(fill=tk.X)
        ttk.Label(search_frame, text="Buscar:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._filter_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_entry.focus_set()

        # Frame da Listbox
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, exportselection=False)
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<Double-Button-1>", self._select_item)
        self.listbox.bind("<Return>", self._select_item)

        self._populate_listbox(self.df_full)

        # Frame de Botões
        button_frame = ttk.Frame(self, padding="5")
        button_frame.pack(fill=tk.X)
        # Aplicando estilo aos botões do diálogo
        ttk.Button(button_frame, text="Selecionar", command=self._select_item, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.wait_window(self) # Espera até que a janela seja fechada

    def _populate_listbox(self, df):
        self.listbox.delete(0, tk.END)
        # Ajuste o formato de exibição conforme necessário
        for index, row in df.iterrows():
            # Exemplo de exibição: Código - Descrição (ou CA - Descrição para EPIs)
            display_text = f"{row[self.return_col]} - {row.get('DESCRICAO', '')}"
            self.listbox.insert(tk.END, display_text)

    def _filter_list(self, *args):
        query = self.search_var.get().strip().lower()
        df_to_search = self.df_full # Usa o df original para filtrar
        if not query:
            df_filtered = df_to_search
        else:
            try:
                # --- LÓGICA DE BUSCA CORRIGIDA ---
                mask = df_to_search.apply(
                    lambda row: any(query in str(row[col]).lower() # Usa 'in' para verificar substring
                                    for col in self.search_cols if pd.notna(row[col])),
                    axis=1
                )
                # --- FIM DA CORREÇÃO ---
                df_filtered = df_to_search[mask]
            except Exception as e:
                 print(f"Erro durante o filtro de busca: {e}") # Registra o erro
                 df_filtered = df_to_search # Mostra tudo em caso de erro

        self._populate_listbox(df_filtered)

    def _select_item(self, event=None):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            # Recupera os dados armazenados
            list_item_text = self.listbox.get(index)
            # Extrai o identificador (primeira parte antes de ' - ')
            self.result = list_item_text.split(' - ')[0]
            self.destroy()
        else:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um item da lista.", parent=self)


class EditDialogBase(tk.Toplevel):
    """Classe base para diálogos de edição."""
    def __init__(self, parent, title, item_data):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.geometry("400x300") # Ajuste conforme necessário

        self.item_data = item_data # Dicionário com dados originais
        self.updated_data = None # Conterá o resultado se salvo

        self.entries = {}
        self._create_widgets()

        # Frame de Botões
        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)
        # Aplicando estilo aos botões do diálogo
        ttk.Button(button_frame, text="Salvar", command=self._save, style="Success.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _create_widgets(self):
        # A ser implementado pelas subclasses
        pass

    def _add_entry(self, frame, field_key, label_text, row, col, **kwargs):
        """Auxiliar para criar rótulo e entrada."""
        ttk.Label(frame, text=label_text + ":").grid(row=row, column=col*2, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(frame, **kwargs)
        entry.grid(row=row, column=col*2 + 1, sticky=tk.EW, padx=5, pady=2)
        if field_key in self.item_data:
             entry.insert(0, str(self.item_data[field_key]))
        self.entries[field_key] = entry
        frame.columnconfigure(col*2 + 1, weight=1) # Faz a entrada expandir

    def _validate_and_collect(self):
        # A ser implementado pelas subclasses para validação específica
        collected_data = {}
        for key, entry in self.entries.items():
             collected_data[key] = entry.get().strip()
        return collected_data # Retorna os dados coletados

    def _save(self):
        try:
            self.updated_data = self._validate_and_collect()
            if self.updated_data: # validação passou se retornou dados
                self.destroy()
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e), parent=self)


class EditProductDialog(EditDialogBase):
    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._add_entry(main_frame, "CODIGO", "Código", 0, 0, state='readonly') # Somente leitura
        self._add_entry(main_frame, "DESCRICAO", "Descrição", 1, 0)
        self._add_entry(main_frame, "VALOR UN", "Valor Unitário", 2, 0)
        self._add_entry(main_frame, "QUANTIDADE", "Quantidade", 3, 0)
        self._add_entry(main_frame, "LOCALIZACAO", "Localização", 4, 0)
        # DATA e VALOR TOTAL são geralmente calculados, não editados diretamente aqui

    def _validate_and_collect(self):
        data = super()._validate_and_collect()

        # Validação
        if not data["DESCRICAO"]:
            raise ValueError("Descrição não pode ser vazia.")

        try:
            data["VALOR UN"] = float(str(data["VALOR UN"]).replace(",", "."))
            if data["VALOR UN"] < 0: raise ValueError("Valor Unitário não pode ser negativo.")
        except ValueError:
            raise ValueError("Valor Unitário deve ser um número válido.")

        try:
            data["QUANTIDADE"] = float(str(data["QUANTIDADE"]).replace(",", "."))
            if data["QUANTIDADE"] < 0: raise ValueError("Quantidade não pode ser negativa.")
        except ValueError:
            raise ValueError("Quantidade deve ser um número válido.")

        # Calcula o Valor Total com base nos valores atualizados
        data["VALOR TOTAL"] = data["VALOR UN"] * data["QUANTIDADE"]

        return data

class EditEPIDialog(EditDialogBase):
     def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._add_entry(main_frame, "CA", "CA", 0, 0) # Permitir editar CA? Ou tornar somente leitura?
        self._add_entry(main_frame, "DESCRICAO", "Descrição", 1, 0)
        self._add_entry(main_frame, "QUANTIDADE", "Quantidade", 2, 0)

     def _validate_and_collect(self):
        data = super()._validate_and_collect()

        if not (data["CA"] or data["DESCRICAO"]):
             raise ValueError("Pelo menos CA ou Descrição deve ser preenchido.")

        try:
            data["QUANTIDADE"] = float(str(data["QUANTIDADE"]).replace(",", "."))
            if data["QUANTIDADE"] < 0: raise ValueError("Quantidade não pode ser negativa.")
        except ValueError:
            raise ValueError("Quantidade deve ser um número válido.")

        # Garante que CA/Desc estejam em maiúsculas
        data["CA"] = data["CA"].upper()
        data["DESCRICAO"] = data["DESCRICAO"].upper()
        return data

# --- Classe Principal da Aplicação ---

class AlmoxarifadoApp:
    def __init__(self, root, user_id):
        self.root = root
        self.operador_logado_id = user_id
        self.active_table_name = "estoque"
        self.current_table_df = None # Armazena o dataframe para a pandastable ativa

        self.root.title(f"Almoxarifado - Operador: {self.operador_logado_id}")
        self.root.geometry("1150x650") # Um pouco maior para a barra de status
        # self.root.config(bg="#E0E0E0") # Fundo mais claro
        # self.root.resizable(False, False) # Considere permitir redimensionamento

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- ORDEM CORRIGIDA ---
        self._criar_pastas_e_planilhas()
        self._setup_ui() # Configura a UI, incluindo estilos
        self._criar_backup_periodico()
        self._load_and_display_table(self.active_table_name)

        # Agenda verificação periódica de backup (ex: a cada 3 horas)
        self.root.after(10800000, self._schedule_backup) # 10800000 ms = 3 horas

    def _setup_ui(self):
            # Estilo
            style = ttk.Style()
            style.theme_use('clam') # Ou 'alt', 'default', 'classic'

            # --- Definição de Estilos Coloridos ---
            # Estilo Principal (Azul)
            style.configure("Accent.TButton",
                            foreground="white",
                            background="#0078D7")
            style.map("Accent.TButton",
                      background=[('active', '#005A9E')]) # Azul mais escuro ao passar o mouse

            # Estilo de Sucesso (Verde)
            style.configure("Success.TButton",
                            foreground="white",
                            background="#107C10")
            style.map("Success.TButton",
                      background=[('active', '#0A530A')]) # Verde mais escuro

            # Estilo de Aviso/Edição (Laranja/Amarelo)
            style.configure("Edit.TButton",
                            foreground="black",
                            background="#FFB900")
            style.map("Edit.TButton",
                      background=[('active', '#D89D00')]) # Amarelo mais escuro

            # Estilo de Perigo/Exclusão (Vermelho)
            style.configure("Delete.TButton",
                            foreground="black",
                            background="#D83B01")
            style.map("Delete.TButton",
                      background=[('active', '#A42E00')]) # Vermelho mais escuro

            # Estilo Padrão (Cinza - para botões menos importantes)
            style.configure("Secondary.TButton",
                            foreground="black",
                            background="#CCCCCC", # Cinza claro
                            padding=5)
            style.map("Secondary.TButton",
                      background=[('active', '#B3B3B3')]) # Cinza um pouco mais escuro
            # --- Fim da Definição de Estilos ---


            # Frame principal
            main_frame = ttk.Frame(self.root, padding="5")
            main_frame.pack(expand=True, fill="both")

            self.status_var = tk.StringVar()
            # Empacota temporariamente a barra de status no topo para reservar espaço, depois a move
            self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
            # self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # Empacota depois


            # Notebook (Abas)
            self.notebook = ttk.Notebook(main_frame)
            self.notebook.pack(expand=True, fill="both", pady=(0, 5))


            # Cria Abas (Agora podem chamar _update_status com segurança)
            self._create_estoque_tab()
            self._create_cadastro_tab()
            self._create_movimentacao_tab()
            self._create_epis_tab()


            self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # Empacota na posição final desejada
            self._update_status("Pronto.")

    def _create_estoque_tab(self):
        self.estoque_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.estoque_tab, text="Estoque & Tabelas")

        # --- Controles Superiores ---
        controls_frame = ttk.Frame(self.estoque_tab)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        # Pesquisa
        search_frame = ttk.LabelFrame(controls_frame, text="Pesquisar na Tabela Atual", padding="5")
        search_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.pesquisar_entry = ttk.Entry(search_frame, width=40)
        self.pesquisar_entry.bind("<KeyRelease>", self._pesquisar_tabela_event)
        self.pesquisar_entry.pack(side=tk.LEFT, padx=(0, 5))
        # Aplicando estilos
        ttk.Button(search_frame, text="Buscar", command=self._pesquisar_tabela, style="Success.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(search_frame, text="Limpar", command=self._limpar_pesquisa).pack(side=tk.LEFT)

        # Troca de Tabela
        switch_frame = ttk.LabelFrame(controls_frame, text="Visualizar Tabela", padding="5")
        switch_frame.pack(side=tk.LEFT, padx=(0, 10))
        # Botões de visualização usarão o estilo padrão ou Accent quando ativos
        self.btn_view_estoque = ttk.Button(switch_frame, text="Estoque", command=lambda: self._trocar_tabela_view("estoque"))
        self.btn_view_estoque.pack(side=tk.LEFT, padx=2)
        self.btn_view_entrada = ttk.Button(switch_frame, text="Entradas", command=lambda: self._trocar_tabela_view("entrada"))
        self.btn_view_entrada.pack(side=tk.LEFT, padx=2)
        self.btn_view_saida = ttk.Button(switch_frame, text="Saídas", command=lambda: self._trocar_tabela_view("saida"))
        self.btn_view_saida.pack(side=tk.LEFT, padx=2)

        # Frame de Ações (Alinhado à Direita)
        action_frame = ttk.Frame(controls_frame)
        action_frame.pack(side=tk.RIGHT)

        # Frame Editar/Excluir (Agrupado)
        edit_delete_frame = ttk.LabelFrame(action_frame, text="Item Selecionado", padding="5")
        edit_delete_frame.pack(side=tk.LEFT, padx=(0,10))
        # Aplicando estilos
        self.edit_button = ttk.Button(edit_delete_frame, text="Editar", command=self._edit_selected_item, state=tk.DISABLED, style="Edit.TButton")
        self.edit_button.pack(side=tk.LEFT, padx=2)
        self.delete_button = ttk.Button(edit_delete_frame, text="Excluir", command=self._delete_selected_item, state=tk.DISABLED, style="Delete.TButton")
        self.delete_button.pack(side=tk.LEFT, padx=2)


        # Frame de Ações Gerais (Agrupado)
        general_action_frame = ttk.LabelFrame(action_frame, text="Ações", padding="5")
        general_action_frame.pack(side=tk.LEFT, padx=(0, 10))
        # Aplicando estilo
        ttk.Button(general_action_frame, text="Atualizar", command=self._atualizar_tabela_atual, style="Success.TButton").pack(side=tk.LEFT, padx=2)


        # Frame de Exportação (Agrupado)
        export_frame = ttk.LabelFrame(action_frame, text="Relatórios", padding="5")
        export_frame.pack(side=tk.LEFT)
        # Aplicando estilo
        ttk.Button(export_frame, text="Exportar", command=self._exportar_conteudo, style="Accent.TButton").pack(side=tk.LEFT, padx=2)


        # --- Frame da Tabela ---
        self.pandas_table_frame = ttk.Frame(self.estoque_tab)
        self.pandas_table_frame.pack(expand=True, fill="both")

        # Cria tabela dummy inicialmente, será substituída
        self.pandas_table = Table(parent=self.pandas_table_frame)
        self.pandas_table.show()
        # Desabilita edição direta
        self.pandas_table.editable = False
        # Adiciona binding para habilitar/desabilitar botões na mudança de seleção
        self.pandas_table.bind("<<TableSelectChanged>>", self._on_table_select)

        self._atualizar_cores_botoes_view() # Define o estado inicial do botão


    def _create_cadastro_tab(self):
        self.cadastro_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.cadastro_tab, text="Cadastrar Produto")

        container = ttk.Frame(self.cadastro_tab)
        container.pack(anchor=tk.CENTER) # Centraliza os elementos do formulário

        ttk.Label(container, text="Cadastrar Novo Produto no Estoque", font="-weight bold -size 14").grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Usa uma estrutura consistente para rótulos e entradas
        fields = [
            ("DESCRICAO", "Descrição:"),
            ("QUANTIDADE", "Quantidade:"),
            ("VALOR UN", "Valor Unitário:"),
            ("LOCALIZACAO", "Localização:")
        ]

        self.cadastro_entries = {}
        for i, (key, label_text) in enumerate(fields):
            ttk.Label(container, text=label_text, font="-size 12").grid(row=i+1, column=0, sticky=tk.W, padx=5, pady=5)
            entry = ttk.Entry(container, width=40)
            entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
            entry.grid(row=i+1, column=1, sticky=tk.EW, padx=5, pady=5)
            entry.bind("<Return>", self._focar_proximo_cadastro) # Associa a tecla Enter
            self.cadastro_entries[key] = entry

        # Associa a última entrada à função cadastrar no Return
        self.cadastro_entries["LOCALIZACAO"].bind("<Return>", lambda e: self._cadastrar_estoque())

        # Botão
        # Aplicando estilo
        cadastrar_button = ttk.Button(container, text="Cadastrar Produto", command=self._cadastrar_estoque, style="Success.TButton")
        cadastrar_button.grid(row=len(fields)+1, column=0, columnspan=2, pady=(20, 0), sticky=tk.EW)

        # Torna a coluna de entrada expansível
        container.columnconfigure(1, weight=1)

    def _create_movimentacao_tab(self):
        self.movimentacao_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.movimentacao_tab, text="Movimentação Estoque")

        # Layout Principal em Grid
        self.movimentacao_tab.columnconfigure(0, weight=1)
        self.movimentacao_tab.columnconfigure(1, minsize=20) # Espaço separador
        self.movimentacao_tab.columnconfigure(2, weight=1)
        self.movimentacao_tab.rowconfigure(0, weight=1)

        # --- Frame de Entrada ---
        entrada_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Entrada", padding="15")
        entrada_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        entrada_frame.columnconfigure(1, weight=1) # Faz as entradas expandirem

        # Widgets de Entrada
        ttk.Label(entrada_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_codigo_entry = ttk.Entry(entrada_frame)
        self.entrada_codigo_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.entrada_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        self.entrada_codigo_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        # Aplicando estilo
        ttk.Button(entrada_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("entrada"), style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5))


        ttk.Label(entrada_frame, text="Qtd. Entrada:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_qtd_entry = ttk.Entry(entrada_frame)
        self.entrada_qtd_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.entrada_qtd_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.entrada_qtd_entry.bind("<Return>", lambda e: self._registrar_entrada())

        # Aplicando estilo
        entrada_button = ttk.Button(entrada_frame, text="Registrar Entrada", command=self._registrar_entrada, style="Success.TButton")
        entrada_button.grid(row=2, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)

        # --- Separador ---
        sep = ttk.Separator(self.movimentacao_tab, orient=tk.VERTICAL)
        sep.grid(row=0, column=1, sticky="ns", padx=5, pady=5)

        # --- Frame de Saída ---
        saida_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Saída", padding="15")
        saida_frame.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        saida_frame.columnconfigure(1, weight=1) # Faz as entradas expandirem

        # Widgets de Saída
        ttk.Label(saida_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_codigo_entry = ttk.Entry(saida_frame)
        self.saida_codigo_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.saida_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        self.saida_codigo_entry.bind("<Return>", lambda e: self._focar_proximo(e))
        # Aplicando estilo
        ttk.Button(saida_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("saida"), style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5))


        ttk.Label(saida_frame, text="Solicitante:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_solicitante_entry = ttk.Entry(saida_frame)
        self.saida_solicitante_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.saida_solicitante_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_solicitante_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(saida_frame, text="Qtd. Saída:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_qtd_entry = ttk.Entry(saida_frame)
        self.saida_qtd_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.saida_qtd_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_qtd_entry.bind("<Return>", lambda e: self._registrar_saida())

        # Aplicando estilo
        saida_button = ttk.Button(saida_frame, text="Registrar Saída", command=self._registrar_saida, style="Success.TButton") # Usando vermelho para saída
        saida_button.grid(row=3, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)


    def _create_epis_tab(self):
        self.epis_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.epis_tab, text="EPIs")

        # Layout: Tabela à esquerda, Formulários à direita
        self.epis_tab.columnconfigure(0, weight=2) # Área da tabela
        self.epis_tab.columnconfigure(1, weight=1) # Área dos formulários
        self.epis_tab.rowconfigure(0, weight=1)

        # --- Frame da Tabela de EPIs ---
        epis_table_frame = ttk.Frame(self.epis_tab)
        epis_table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # Controles para a Tabela de EPIs
        epi_table_controls = ttk.Frame(epis_table_frame)
        epi_table_controls.pack(fill=tk.X, pady=(0, 5))

        # Aplicando estilos
        ttk.Button(epi_table_controls, text="Atualizar Lista", command=self._atualizar_tabela_epis, style="Success.TButton").pack(side=tk.LEFT, padx=(0, 10))

        self.edit_epi_button = ttk.Button(epi_table_controls, text="Editar EPI Sel.", command=self._edit_selected_epi, state=tk.DISABLED, style="Accent.TButton")
        self.edit_epi_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_epi_button = ttk.Button(epi_table_controls, text="Excluir EPI Sel.", command=self._delete_selected_epi, state=tk.DISABLED, style="Accent.TButton")
        self.delete_epi_button.pack(side=tk.LEFT, padx=(0, 5))


        # Pandastable de EPIs
        self.epis_pandas_frame = ttk.Frame(epis_table_frame)
        self.epis_pandas_frame.pack(expand=True, fill="both")
        self.epis_table = Table(parent=self.epis_pandas_frame, editable=False) # Desabilita edição direta
        self.epis_table.show()
        self._carregar_epis() # Carga inicial
        self.epis_table.bind("<<TableSelectChanged>>", self._on_epi_table_select)

        # --- Frame de Formulários (Direita) ---
        forms_frame = ttk.Frame(self.epis_tab)
        forms_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        forms_frame.rowconfigure(1, weight=1) # Permite que o frame de retirar expanda se necessário


        # --- Frame Registrar EPI ---
        registrar_frame = ttk.LabelFrame(forms_frame, text="Registrar / Adicionar EPI", padding="10")
        registrar_frame.pack(fill=tk.X, pady=(0, 15))
        registrar_frame.columnconfigure(1, weight=1)

        ttk.Label(registrar_frame, text="CA:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_ca_entry = ttk.Entry(registrar_frame)
        self.epi_ca_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.epi_ca_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_ca_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        ttk.Label(registrar_frame, text="Descrição:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_desc_entry = ttk.Entry(registrar_frame)
        self.epi_desc_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.epi_desc_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_desc_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(registrar_frame, text="Quantidade:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_qtd_entry = ttk.Entry(registrar_frame)
        self.epi_qtd_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.epi_qtd_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_qtd_entry.bind("<Return>", lambda e: self._registrar_epi())

        # Aplicando estilo
        registrar_epi_button = ttk.Button(registrar_frame, text="Registrar / Adicionar", command=self._registrar_epi, style="Success.TButton")
        registrar_epi_button.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky=tk.EW)


        # --- Frame Retirar EPI ---
        retirar_frame = ttk.LabelFrame(forms_frame, text="Registrar Retirada de EPI", padding="10")
        retirar_frame.pack(fill=tk.BOTH, expand=True) # Preenche o espaço restante
        retirar_frame.columnconfigure(1, weight=1)


        ttk.Label(retirar_frame, text="CA/Descrição:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_id_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_id_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.retirar_epi_id_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2), pady=5)
        self.retirar_epi_id_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        # Aplicando estilo
        ttk.Button(retirar_frame, text="Buscar", width=8, command=self._show_epi_lookup, style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5), pady=5)


        ttk.Label(retirar_frame, text="Qtd. Retirada:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_qtd_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_qtd_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.retirar_epi_qtd_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_qtd_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(retirar_frame, text="Colaborador:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_colab_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_colab_entry.config(font="-size 12") # Define o tamanho da fonte para as entradas
        self.retirar_epi_colab_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_colab_entry.bind("<Return>", lambda e: self._registrar_retirada())

        # Aplicando estilo
        retirar_epi_button = ttk.Button(retirar_frame, text="Registrar Retirada", command=self._registrar_retirada, style="Success.TButton") # Usando vermelho para retirada
        retirar_epi_button.grid(row=3, column=0, columnspan=3, pady=(10, 5), sticky=tk.EW)


    # --- Manipuladores de Eventos da UI & Auxiliares ---
    # ... (O código dos métodos _update_status, _focar_proximo, _focar_proximo_cadastro permanece o mesmo) ...
    def _update_status(self, message, error=False):
        """Atualiza a barra de status."""
        self.status_var.set(message)
        if error:
            self.status_bar.config(foreground="red")
        else:
            self.status_bar.config(foreground="black")
        # print(message) # Também imprime no console para depuração

    def _focar_proximo(self, event):
        """Move o foco para o próximo widget na tecla Enter."""
        try:
            event.widget.tk_focusNext().focus()
        except Exception: # Ignora se não houver próximo widget
            pass
        return "break" # Impede o comportamento padrão do Enter (como adicionar nova linha)

    def _focar_proximo_cadastro(self, event):
        """Foca o próximo especificamente para campos de cadastro, ordenados."""
        widget = event.widget
        ordered_entries = [
            self.cadastro_entries["DESCRICAO"],
            self.cadastro_entries["QUANTIDADE"],
            self.cadastro_entries["VALOR UN"],
            self.cadastro_entries["LOCALIZACAO"]
        ]
        try:
            current_index = ordered_entries.index(widget)
            if current_index < len(ordered_entries) - 1:
                ordered_entries[current_index + 1].focus()
            else:
                 # Se for a última entrada, talvez focar o botão? Ou chamar o envio? Vamos focar o botão
                 # Assumindo que o botão está acessível, caso contrário, chama _cadastrar_estoque()
                 # Encontra o widget do botão se necessário, ou apenas chama a função
                 # Por simplicidade aqui, chamamos a função se Return for pressionado no último campo
                 if widget == self.cadastro_entries["LOCALIZACAO"]:
                      self._cadastrar_estoque()
                 else: #Não deve acontecer com base na verificação do índice
                      pass

        except ValueError:
             # Widget não está na lista esperada, tenta o foco padrão
            try:
                 event.widget.tk_focusNext().focus()
            except: pass # ignora erros adicionais
        return "break"

    # ... (O código dos métodos _on_table_select, _on_epi_table_select, _pesquisar_tabela_event, _pesquisar_tabela, _limpar_pesquisa permanece o mesmo) ...
    def _on_table_select(self, event=None):
        """Habilita/desabilita botões de editar/excluir com base na seleção da tabela."""
        selected = self.pandas_table.getSelectedRowData()
        # Assumindo seleção única por simplicidade. Ajuste se for necessária seleção múltipla.
        if selected is not None and len(selected) == 1:
            self.edit_button.config(state=tk.NORMAL)
            # Permitir exclusão apenas da visualização de estoque? Ou basear no tipo de tabela?
            if self.active_table_name == "estoque":
                 self.delete_button.config(state=tk.NORMAL)
            else:
                 self.delete_button.config(state=tk.DISABLED)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)


    def _on_epi_table_select(self, event=None):
        """Habilita/desabilita botões de editar/excluir EPI."""
        selected = self.epis_table.getSelectedRowData()
        if selected is not None and len(selected) == 1:
            self.edit_epi_button.config(state=tk.NORMAL)
            self.delete_epi_button.config(state=tk.NORMAL)
        else:
            self.edit_epi_button.config(state=tk.DISABLED)
            self.delete_epi_button.config(state=tk.DISABLED)


    def _pesquisar_tabela_event(self, event=None):
        """Manipulador para liberação de tecla na entrada de pesquisa."""
        # Opcional: Adicione um pequeno atraso aqui, se necessário, para evitar pesquisar a cada tecla
        self._pesquisar_tabela()

    def _pesquisar_tabela(self):
        """Filtra a tabela pandas atualmente exibida."""
        if self.current_table_df is None or self.pandas_table is None:
            return
        query = self.pesquisar_entry.get().strip().lower()

        if query:
            try:
                # Filtra com base nas colunas de string que contêm a consulta
                string_cols = self.current_table_df.select_dtypes(include='object').columns
                mask = self.current_table_df[string_cols].apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)

                # Também tenta converter outras colunas para string e verificar
                other_cols = self.current_table_df.select_dtypes(exclude='object').columns
                mask_others = self.current_table_df[other_cols].astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)

                df_filtered = self.current_table_df[mask | mask_others]

            except Exception as e:
                self._update_status(f"Erro na pesquisa: {e}", error=True)
                df_filtered = self.current_table_df # Mostra tudo em caso de erro
        else:
            df_filtered = self.current_table_df # Mostra tudo se a consulta estiver vazia

        # Atualiza o modelo da tabela com segurança
        try:
            self.pandas_table.updateModel(TableModel(df_filtered))
            self.pandas_table.redraw()
            self._update_status(f"Exibindo {len(df_filtered)} resultados para '{query}'" if query else f"Exibindo todos os {len(df_filtered)} registros.")
        except Exception as e:
            self._update_status(f"Erro ao atualizar tabela: {e}", error=True)

    def _limpar_pesquisa(self):
        """Limpa a entrada de pesquisa e redefine a visualização da tabela."""
        if self.pandas_table is None or self.current_table_df is None:
            return
        self.pesquisar_entry.delete(0, tk.END)
        try:
            self.pandas_table.updateModel(TableModel(self.current_table_df))
            self.pandas_table.redraw()
            self._update_status(f"Exibindo todos os {len(self.current_table_df)} registros.")
            self._on_table_select() # Redefine os estados dos botões
        except Exception as e:
             self._update_status(f"Erro ao limpar pesquisa: {e}", error=True)

    def _atualizar_cores_botoes_view(self):
        """Destaca o botão da tabela atualmente visualizada."""
        buttons = {
            "estoque": self.btn_view_estoque,
            "entrada": self.btn_view_entrada,
            "saida": self.btn_view_saida
        }
        for name, button in buttons.items():
             # Verifica se o botão existe antes de configurar
             # Usando hasattr para segurança, embora devam existir se _create_estoque_tab foi chamado
             if hasattr(self, f"btn_view_{name}"):
                 target_button = getattr(self, f"btn_view_{name}")
                 if name == self.active_table_name:
                      # Aplica o estilo Accent ao botão ativo
                      target_button.config(style="Accent.TButton")
                 else:
                      # Aplica o estilo padrão (ou Secondary se preferir) aos inativos
                      target_button.config(style="TButton") # Ou "Secondary.TButton"


    # --- Carregamento de Dados e Operações de Arquivo ---
    # ... (O código dos métodos _criar_pastas_e_planilhas, _safe_read_csv, _safe_write_csv, _load_and_display_table, _atualizar_tabela_atual, _trocar_tabela_view, _carregar_epis, _atualizar_tabela_epis, _buscar_produto, _atualizar_estoque_produto, _obter_proximo_codigo permanece o mesmo) ...
    def _criar_pastas_e_planilhas(self):
        """Cria diretórios e arquivos CSV necessários se não existirem."""
        os.makedirs(PLANILHAS_DIR, exist_ok=True)
        os.makedirs(BACKUP_DIR, exist_ok=True)
        os.makedirs(COLABORADORES_DIR, exist_ok=True)

        colunas = {
            "estoque": ["CODIGO", "DESCRICAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO"],
            "entrada": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID"],
            "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA", "ID"],
            "epis": ["CA", "DESCRICAO", "QUANTIDADE"]
        }

        for nome, arquivo in ARQUIVOS.items():
            if not os.path.exists(arquivo):
                try:
                    df = pd.DataFrame(columns=colunas.get(nome, []))
                    df.to_csv(arquivo, index=False, encoding="utf-8")
                    print(f"Arquivo criado: {arquivo}")
                except Exception as e:
                    messagebox.showerror("Erro ao Criar Arquivo", f"Não foi possível criar {arquivo}: {e}")
                    # Considere sair se arquivos essenciais não puderem ser criados

    def _safe_read_csv(self, file_path):
        """Lê um arquivo CSV com segurança, retornando um DataFrame vazio em caso de erro."""
        try:
            # Define explicitamente os dtypes para evitar confusão, esp para códigos/CAs
            dtype_map = {'CODIGO': str, 'CA': str} # Adicione outros se necessário
            return pd.read_csv(file_path, encoding="utf-8", dtype=dtype_map)
        except FileNotFoundError:
            self._update_status(f"Aviso: Arquivo não encontrado: {file_path}. Verifique as pastas.", error=True)
            return pd.DataFrame() # Retorna df vazio
        except Exception as e:
            self._update_status(f"Erro ao ler {file_path}: {e}", error=True)
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo {os.path.basename(file_path)}.\nVerifique se ele não está corrompido ou aberto em outro programa.\n\nDetalhes: {e}")
            return pd.DataFrame()

    def _safe_write_csv(self, df, file_path, create_backup=True):
        """Escreve um DataFrame em CSV com segurança, opcionalmente criando um backup."""
        backup_path = None
        try:
            # Backup antes de escrever
            if create_backup and os.path.exists(file_path):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                backup_path = file_path.replace(".csv", f"_backup_{timestamp}.csv")
                shutil.copy2(file_path, backup_path) # copy2 preserva metadados

            df.to_csv(file_path, index=False, encoding="utf-8")
            return True # Indica sucesso
        except Exception as e:
            self._update_status(f"Erro Crítico ao salvar {os.path.basename(file_path)}: {e}", error=True)
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar as alterações em {os.path.basename(file_path)}.\n\nDetalhes: {e}\n\n{'Um backup pode ter sido criado: ' + os.path.basename(backup_path) if backup_path else 'Não foi possível criar backup.'}")
            # Considere opções de recuperação ou avisar o usuário que os dados podem estar inconsistentes
            return False # Indica falha

    def _load_and_display_table(self, table_name):
        """Carrega dados do CSV e atualiza a pandastable principal."""
        file_path = ARQUIVOS.get(table_name)
        if not file_path:
            messagebox.showerror("Erro Interno", f"Nome de tabela inválido: {table_name}")
            return

        self._update_status(f"Carregando tabela '{table_name.capitalize()}'...")
        df = self._safe_read_csv(file_path)

        # Realiza verificações/correções de sanidade se necessário *com cuidado*
        # Exemplo: Garante que colunas numéricas sejam numéricas, preenche NaNs razoavelmente
        if table_name == "estoque":
             numeric_cols = ["VALOR UN", "VALOR TOTAL", "QUANTIDADE"]
             df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
             # Recalcular VALOR TOTAL para consistência? Arriscado se os dados estiverem ruins.
             # df["VALOR TOTAL"] = df["VALOR UN"] * df["QUANTIDADE"]
             # Só salva de volta imediatamente se o cálculo não for destrutivo
             # self._safe_write_csv(df, file_path, create_backup=False) # Evita loops de backup

        # Garante que colunas de string sejam strings
        string_like_cols = ["DESCRICAO", "LOCALIZACAO", "SOLICITANTE", "ID", "CODIGO", "CA"]
        for col in string_like_cols:
             if col in df.columns:
                 df[col] = df[col].astype(str).fillna('') # Preenche NaN com string vazia


        self.current_table_df = df # Armazena o dataframe carregado
        self.active_table_name = table_name

        # Atualiza pandastable
        try:
            # Verifica se a tabela existe antes de usar
            if hasattr(self, 'pandas_table') and self.pandas_table:
                self.pandas_table.updateModel(TableModel(self.current_table_df))
                # Ajusta larguras das colunas se necessário - pode precisar de lógica mais específica
                # self.pandas_table.autoResizeColumns() # Use com cautela em tabelas grandes
                self.pandas_table.redraw()
                self._update_status(f"Tabela '{table_name.capitalize()}' carregada ({len(self.current_table_df)} registros).")
                self._atualizar_cores_botoes_view()
                self._on_table_select() # Atualiza estados dos botões
            else:
                 self._update_status(f"Erro: Tabela principal (pandas_table) não inicializada.", error=True)

        except Exception as e:
            self._update_status(f"Erro ao exibir tabela '{table_name.capitalize()}': {e}", error=True)

    def _atualizar_tabela_atual(self):
        """Recarrega os dados para a visualização da tabela ativa atual."""
        if self.active_table_name:
             # Reaplica o filtro de pesquisa se ativo
            query = self.pesquisar_entry.get().strip()
            self._load_and_display_table(self.active_table_name)
            if query:
                 self.pesquisar_entry.delete(0, tk.END)
                 self.pesquisar_entry.insert(0, query)
                 self._pesquisar_tabela() # Reaplica a pesquisa

    def _trocar_tabela_view(self, nome_tabela):
        """Muda a tabela sendo visualizada na aba 'Estoque & Tabelas'."""
        if nome_tabela in ARQUIVOS:
             # Limpa a pesquisa antes de trocar
             self.pesquisar_entry.delete(0, tk.END)
             self._load_and_display_table(nome_tabela)
        else:
            messagebox.showerror("Erro", f"Configuração de tabela '{nome_tabela}' não encontrada.")


    def _carregar_epis(self):
         """Carrega dados de EPIs na tabela da aba EPIs."""
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         # Tratamento básico de tipos para EPIs
         df_epis["CA"] = df_epis["CA"].astype(str).fillna("")
         df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("")
         df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

         try:
             # Verifica se a tabela existe antes de usar
             if hasattr(self, 'epis_table') and self.epis_table:
                 self.epis_table.updateModel(TableModel(df_epis))
                 self.epis_table.redraw()
                 self._update_status(f"Lista de EPIs atualizada ({len(df_epis)} itens).")
                 self._on_epi_table_select() # Atualiza estado do botão
             else:
                 self._update_status(f"Erro: Tabela de EPIs (epis_table) não inicializada.", error=True)

         except Exception as e:
              self._update_status(f"Erro ao exibir EPIs: {e}", error=True)


    def _atualizar_tabela_epis(self):
         """Método de conveniência para recarregar dados de EPIs."""
         self._carregar_epis()

    def _buscar_produto(self, codigo):
        """Busca um produto no estoque pelo código. Retorna uma Series ou None."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        # Garante que a coluna Codigo seja string para comparação
        if 'CODIGO' in df_estoque.columns:
            df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
            produto = df_estoque[df_estoque['CODIGO'] == str(codigo)]
            if not produto.empty:
                return produto.iloc[0] # Retorna a primeira correspondência como uma Series
        return None

    def _atualizar_estoque_produto(self, codigo, nova_quantidade):
        """Atualiza a quantidade e valor total de um produto no estoque."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        if 'CODIGO' not in df_estoque.columns:
             self._update_status(f"Erro interno: Coluna 'CODIGO' não encontrada no estoque.", error=True)
             return False

        codigo_str = str(codigo)
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str) # Garante consistência

        # Encontra o índice
        idx = df_estoque.index[df_estoque['CODIGO'] == codigo_str].tolist()

        if not idx:
             self._update_status(f"Erro interno: Produto {codigo_str} não encontrado para atualizar.", error=True)
             return False # Produto não encontrado

        try:
            # Atualiza quantidade (lida com possíveis problemas de tipo)
            idx = idx[0] # Usa o primeiro índice se múltiplos existirem de alguma forma

            # Verifica se as colunas existem antes de acessá-las
            if "VALOR UN" not in df_estoque.columns:
                self._update_status(f"Erro: Coluna 'VALOR UN' não encontrada para atualizar produto {codigo_str}.", error=True)
                return False
            if "QUANTIDADE" not in df_estoque.columns:
                 self._update_status(f"Erro: Coluna 'QUANTIDADE' não encontrada para atualizar produto {codigo_str}.", error=True)
                 return False
            if "VALOR TOTAL" not in df_estoque.columns:
                 self._update_status(f"Erro: Coluna 'VALOR TOTAL' não encontrada para atualizar produto {codigo_str}.", error=True)
                 return False
            if "DATA" not in df_estoque.columns:
                 self._update_status(f"Erro: Coluna 'DATA' não encontrada para atualizar produto {codigo_str}.", error=True)
                 return False


            current_valor_un = pd.to_numeric(df_estoque.loc[idx, "VALOR UN"], errors='coerce')
            if pd.isna(current_valor_un):
                current_valor_un = 0 # Padrão se a conversão falhar
            nova_quantidade = pd.to_numeric(nova_quantidade, errors='coerce')
            if pd.isna(nova_quantidade) or nova_quantidade < 0:
                 raise ValueError("Nova quantidade inválida.")


            df_estoque.loc[idx, "QUANTIDADE"] = nova_quantidade
            df_estoque.loc[idx, "VALOR TOTAL"] = current_valor_un * nova_quantidade
            df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y") # Atualiza timestamp

            # Salva as alterações
            if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                 # Se a tabela de estoque estiver sendo exibida, atualiza-a
                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual()
                 return True
            else:
                 return False # Salvar falhou

        except (ValueError, KeyError, IndexError) as e:
             self._update_status(f"Erro ao atualizar estoque para {codigo_str}: {e}", error=True)
             messagebox.showerror("Erro de Atualização", f"Não foi possível atualizar o produto {codigo_str}.\nVerifique os dados na planilha.\n\nDetalhes: {e}")
             return False

    def _obter_proximo_codigo(self):
        """Obtém o próximo código disponível para um novo produto."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        if df_estoque.empty or 'CODIGO' not in df_estoque.columns:
            return "1"  # Começa de 1 se vazio ou sem coluna codigo

        # Converte para numérico com segurança, encontra o máximo, lida com códigos não numéricos
        codes_numeric = pd.to_numeric(df_estoque['CODIGO'], errors='coerce')
        max_code = codes_numeric.max()

        if pd.isna(max_code):
            # Se todos os códigos não forem numéricos ou a conversão falhar, começa de 1
            return "1"
        else:
            # Garante que estamos retornando um inteiro + 1 como string
            try:
                return str(int(max_code) + 1)
            except ValueError: # Caso max_code seja float com decimal (improvável, mas seguro)
                return str(int(round(max_code)) + 1)


    # --- Métodos de Lógica Principal (Cadastro, Movimentação, EPIs) ---
    # ... (O código dos métodos _cadastrar_estoque, _registrar_entrada, _registrar_saida, _registrar_epi, _registrar_retirada permanece o mesmo, mas os botões dentro deles já terão os estilos aplicados) ...
    def _cadastrar_estoque(self):
        """Cadastra um novo produto no estoque."""
        desc = self.cadastro_entries["DESCRICAO"].get().strip().upper()
        qtd_str = self.cadastro_entries["QUANTIDADE"].get().strip().replace(",",".")
        val_un_str = self.cadastro_entries["VALOR UN"].get().strip().replace(",",".")
        loc = self.cadastro_entries["LOCALIZACAO"].get().strip().upper()

        # Validação
        if not desc:
             messagebox.showerror("Erro de Validação", "Descrição não pode ser vazia.")
             self.cadastro_entries["DESCRICAO"].focus_set()
             return
        try:
            quantidade = float(qtd_str)
            if quantidade < 0: raise ValueError("Quantidade negativa")
        except ValueError:
             messagebox.showerror("Erro de Validação", "Quantidade deve ser um número válido não negativo.")
             self.cadastro_entries["QUANTIDADE"].focus_set()
             return
        try:
            valor_un = float(val_un_str)
            if valor_un < 0: raise ValueError("Valor negativo")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Valor Unitário deve ser um número válido não negativo.")
            self.cadastro_entries["VALOR UN"].focus_set()
            return

        codigo = self._obter_proximo_codigo()
        valor_total = quantidade * valor_un
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")

        confirm_msg = (f"Cadastrar o seguinte produto?\n\n"
                       f"Código: {codigo}\n"
                       f"Descrição: {desc}\n"
                       f"Quantidade: {quantidade}\n"
                       f"Valor Unitário: {valor_un:.2f}\n"
                       f"Localização: {loc if loc else '-'}\n")

        if messagebox.askyesno("Confirmar Cadastro", confirm_msg):
            novo_produto = {
                 "CODIGO": codigo, "DESCRICAO": desc, "VALOR UN": valor_un,
                 "VALOR TOTAL": valor_total, "QUANTIDADE": quantidade,
                 "DATA": data_hora, "LOCALIZACAO": loc
            }

            try:
                 # Usa modo 'a' para anexar com header=False se o arquivo existir
                 # Verifica se o arquivo existe e se está vazio para decidir sobre o cabeçalho
                 arquivo_estoque = ARQUIVOS["estoque"]
                 header = not os.path.exists(arquivo_estoque) or os.path.getsize(arquivo_estoque) == 0
                 pd.DataFrame([novo_produto]).to_csv(arquivo_estoque, mode='a', header=header, index=False, encoding='utf-8')

                 messagebox.showinfo("Sucesso", f"Produto '{desc}' (Cód: {codigo}) cadastrado com sucesso!")
                 self._update_status(f"Produto {codigo} - {desc} cadastrado.")

                 # Limpa campos
                 for entry in self.cadastro_entries.values():
                     entry.delete(0, tk.END)
                 self.cadastro_entries["DESCRICAO"].focus_set() # Foca o primeiro campo

                 # Atualiza a visualização se estiver mostrando estoque
                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual()

            except Exception as e:
                 self._update_status(f"Erro ao salvar novo produto: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o produto:\n{e}")


    def _registrar_entrada(self):
        """Registra a entrada de um produto no estoque."""
        codigo = self.entrada_codigo_entry.get().strip()
        qtd_str = self.entrada_qtd_entry.get().strip().replace(",",".")

        # Validação
        if not codigo:
             messagebox.showerror("Erro de Validação", "Código do produto não pode ser vazio.")
             self.entrada_codigo_entry.focus_set()
             return
        try:
            quantidade_adicionada = float(qtd_str)
            if quantidade_adicionada <= 0:
                 raise ValueError("Quantidade deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Quantidade de entrada deve ser um número positivo.")
            self.entrada_qtd_entry.focus_set()
            return

        produto_atual = self._buscar_produto(codigo)
        if produto_atual is None:
            messagebox.showerror("Erro", f"Produto com código '{codigo}' não encontrado no estoque.")
            self.entrada_codigo_entry.focus_set()
            return

        desc = produto_atual.get("DESCRICAO", "N/A")
        val_un = pd.to_numeric(produto_atual.get("VALOR UN", 0), errors='coerce')
        qtd_atual = pd.to_numeric(produto_atual.get("QUANTIDADE", 0), errors='coerce')

        # Verifica se a conversão para numérico falhou (NaN)
        if pd.isna(val_un): val_un = 0
        if pd.isna(qtd_atual): qtd_atual = 0


        nova_quantidade_estoque = qtd_atual + quantidade_adicionada
        valor_total_entrada = val_un * quantidade_adicionada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")

        confirm_msg = (f"Registrar Entrada?\n\n"
                       f"Código: {codigo}\n"
                       f"Descrição: {desc}\n"
                       f"Qtd. a adicionar: {quantidade_adicionada}\n"
                       f"Nova Qtd. Estoque: {nova_quantidade_estoque}\n"
                       f"Valor Total Entrada: {valor_total_entrada:.2f}")

        if messagebox.askyesno("Confirmar Entrada", confirm_msg):
            # 1. Registrar na planilha de Entrada
            entrada_data = {
                 "CODIGO": codigo, "DESCRICAO": desc, "QUANTIDADE": quantidade_adicionada,
                 "VALOR UN": val_un, "VALOR TOTAL": valor_total_entrada,
                 "DATA": data_hora, "ID": self.operador_logado_id
            }
            try:
                 arquivo_entrada = ARQUIVOS["entrada"]
                 header = not os.path.exists(arquivo_entrada) or os.path.getsize(arquivo_entrada) == 0
                 pd.DataFrame([entrada_data]).to_csv(arquivo_entrada, mode='a', header=header, index=False, encoding='utf-8')

                 # 2. Atualizar Estoque
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Entrada registrada para '{desc}' (Cód: {codigo}).\nEstoque atualizado: {nova_quantidade_estoque}")
                     self._update_status(f"Entrada registrada para {codigo}. Novo estoque: {nova_quantidade_estoque}")
                     # Limpa campos
                     self.entrada_codigo_entry.delete(0, tk.END)
                     self.entrada_qtd_entry.delete(0, tk.END)
                     self.entrada_codigo_entry.focus_set()
                     # Atualiza a visualização se estiver mostrando entrada
                     if self.active_table_name == "entrada":
                         self._atualizar_tabela_atual()
                 else:
                      # Atualização falhou, log/mensagem tratada em _atualizar_estoque_produto
                      # Talvez reverter o registro de entrada? Complexo sem transações.
                     messagebox.showwarning("Atenção", "Entrada registrada, mas houve erro ao atualizar o estoque. Verifique os dados.")

            except Exception as e:
                 self._update_status(f"Erro ao registrar entrada: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Entrada", f"Não foi possível salvar a entrada:\n{e}")


    def _registrar_saida(self):
        """Registra a saída de um produto do estoque."""
        codigo = self.saida_codigo_entry.get().strip()
        solicitante = self.saida_solicitante_entry.get().strip().upper()
        qtd_str = self.saida_qtd_entry.get().strip().replace(",",".")

        # Validação
        if not codigo:
             messagebox.showerror("Erro de Validação", "Código do produto não pode ser vazio.")
             self.saida_codigo_entry.focus_set()
             return
        if not solicitante:
             messagebox.showerror("Erro de Validação", "Nome do Solicitante não pode ser vazio.")
             self.saida_solicitante_entry.focus_set()
             return
        try:
            quantidade_retirada = float(qtd_str)
            if quantidade_retirada <= 0:
                 raise ValueError("Quantidade deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Quantidade de saída deve ser um número positivo.")
            self.saida_qtd_entry.focus_set()
            return

        produto_atual = self._buscar_produto(codigo)
        if produto_atual is None:
            messagebox.showerror("Erro", f"Produto com código '{codigo}' não encontrado no estoque.")
            self.saida_codigo_entry.focus_set()
            return

        desc = produto_atual.get("DESCRICAO", "N/A")
        qtd_atual = pd.to_numeric(produto_atual.get("QUANTIDADE", 0), errors='coerce')
        if pd.isna(qtd_atual): qtd_atual = 0 # Trata NaN

        if quantidade_retirada > qtd_atual:
             messagebox.showerror("Erro", f"Quantidade insuficiente no estoque!\n\nDisponível para '{desc}': {qtd_atual}\nSolicitado: {quantidade_retirada}")
             self.saida_qtd_entry.focus_set()
             return

        nova_quantidade_estoque = qtd_atual - quantidade_retirada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")

        confirm_msg = (f"Registrar Saída?\n\n"
                       f"Código: {codigo}\n"
                       f"Descrição: {desc}\n"
                       f"Solicitante: {solicitante}\n"
                       f"Qtd. a retirar: {quantidade_retirada}\n"
                       f"Qtd. Restante: {nova_quantidade_estoque}")

        if messagebox.askyesno("Confirmar Saída", confirm_msg):
            # 1. Registrar na planilha de Saída
            saida_data = {
                 "CODIGO": codigo, "DESCRICAO": desc, "QUANTIDADE": quantidade_retirada,
                 "SOLICITANTE": solicitante, "DATA": data_hora, "ID": self.operador_logado_id
            }
            try:
                 arquivo_saida = ARQUIVOS["saida"]
                 header = not os.path.exists(arquivo_saida) or os.path.getsize(arquivo_saida) == 0
                 pd.DataFrame([saida_data]).to_csv(arquivo_saida, mode='a', header=header, index=False, encoding='utf-8')

                 # 2. Atualizar Estoque
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Saída registrada para '{desc}' (Solicitante: {solicitante}).\nEstoque restante: {nova_quantidade_estoque}")
                     self._update_status(f"Saída registrada para {codigo}. Estoque restante: {nova_quantidade_estoque}")
                     # Limpa campos
                     self.saida_codigo_entry.delete(0, tk.END)
                     self.saida_solicitante_entry.delete(0, tk.END)
                     self.saida_qtd_entry.delete(0, tk.END)
                     self.saida_codigo_entry.focus_set()
                     # Atualiza a visualização se estiver mostrando saida
                     if self.active_table_name == "saida":
                         self._atualizar_tabela_atual()
                 else:
                     messagebox.showwarning("Atenção", "Saída registrada, mas houve erro ao atualizar o estoque. Verifique os dados.")

            except Exception as e:
                 self._update_status(f"Erro ao registrar saída: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Saída", f"Não foi possível salvar a saída:\n{e}")


    def _registrar_epi(self):
        """Registra um novo EPI ou adiciona quantidade a um existente."""
        ca = self.epi_ca_entry.get().strip().upper()
        descricao = self.epi_desc_entry.get().strip().upper()
        qtd_str = self.epi_qtd_entry.get().strip().replace(",",".")

        if not (ca or descricao):
            messagebox.showerror("Erro", "Você deve preencher pelo menos o CA ou a Descrição.", parent=self.epis_tab)
            self.epi_ca_entry.focus_set() # Foca no CA se ambos vazios
            return
        if not qtd_str:
             messagebox.showerror("Erro", "A Quantidade deve ser preenchida.", parent=self.epis_tab)
             self.epi_qtd_entry.focus_set()
             return
        try:
            quantidade_add = float(qtd_str)
            if quantidade_add <= 0:
                raise ValueError("Quantidade deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro", "Quantidade deve ser um número positivo.", parent=self.epis_tab)
            self.epi_qtd_entry.focus_set()
            return

        df_epis = self._safe_read_csv(ARQUIVOS["epis"])
        # Garante tipos e limpa strings antes de buscar
        if "CA" not in df_epis.columns: df_epis["CA"] = ""
        if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""
        if "QUANTIDADE" not in df_epis.columns: df_epis["QUANTIDADE"] = 0

        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

        found_epi_index = None
        epi_existente_data = None

        # Prioriza correspondência por CA se fornecido e não vazio
        if ca:
             match = df_epis[df_epis["CA"] == ca]
             if not match.empty:
                 found_epi_index = match.index[0]
                 epi_existente_data = match.iloc[0]
        # Se não houver correspondência de CA ou CA não foi fornecido (ou estava vazio), tenta por Descrição (se fornecida e não vazia)
        if found_epi_index is None and descricao:
            match = df_epis[df_epis["DESCRICAO"] == descricao]
            if not match.empty:
                # Verifica se esta descrição está ligada a um CA *diferente* e *não vazio* do fornecido (se CA foi fornecido e não vazio)
                existing_ca = match.iloc[0]["CA"]
                if ca and existing_ca and existing_ca != ca:
                    messagebox.showwarning("Conflito de Dados", f"A descrição '{descricao}' já existe mas está associada ao CA '{existing_ca}'.\nNão é possível adicionar com o CA '{ca}'. Verifique os dados.", parent=self.epis_tab)
                    return
                # Se o CA existente for vazio, ou se o CA fornecido for vazio, ou se os CAs coincidirem, permite a correspondência
                found_epi_index = match.index[0]
                epi_existente_data = match.iloc[0]

        # --- Lida com EPI Encontrado vs. Novo ---
        if epi_existente_data is not None and found_epi_index is not None:
             # EPI Existe - Confirma adição de quantidade
             qtd_atual = epi_existente_data["QUANTIDADE"]
             epi_display_ca = epi_existente_data['CA'] if epi_existente_data['CA'] else '-'
             epi_display_desc = epi_existente_data['DESCRICAO']
             confirm_msg = (f"EPI já existe:\n"
                            f" CA: {epi_display_ca}\n"
                            f" Descrição: {epi_display_desc}\n"
                            f" Qtd. Atual: {qtd_atual}\n\n"
                            f"Deseja adicionar {quantidade_add} à quantidade existente?")
             if messagebox.askyesno("Confirmar Adição", confirm_msg, parent=self.epis_tab):
                 nova_quantidade = qtd_atual + quantidade_add
                 df_epis.loc[found_epi_index, "QUANTIDADE"] = nova_quantidade
                 # Atualiza CA/Descrição se foram fornecidos e diferentes (e a verificação de conflito passou)
                 if ca and df_epis.loc[found_epi_index, "CA"] != ca:
                      df_epis.loc[found_epi_index, "CA"] = ca
                 if descricao and df_epis.loc[found_epi_index, "DESCRICAO"] != descricao:
                      df_epis.loc[found_epi_index, "DESCRICAO"] = descricao

                 # Salva
                 if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                      # Usa a descrição atualizada para a mensagem
                      desc_atualizada = df_epis.loc[found_epi_index, "DESCRICAO"]
                      self._update_status(f"Quantidade EPI {desc_atualizada} atualizada para {nova_quantidade}.")
                      messagebox.showinfo("Sucesso", f"Quantidade atualizada!\nNova Quantidade: {nova_quantidade}", parent=self.epis_tab)
                      self._atualizar_tabela_epis()
                      # Limpa Campos
                      self.epi_ca_entry.delete(0, tk.END)
                      self.epi_desc_entry.delete(0, tk.END)
                      self.epi_qtd_entry.delete(0, tk.END)
                      self.epi_ca_entry.focus_set()
                 # else: erro tratado em _safe_write_csv
             else:
                  messagebox.showinfo("Operação Cancelada", "A quantidade não foi alterada.", parent=self.epis_tab)

        else:
             # Novo EPI - Confirma registro
             confirm_msg = (f"Registrar novo EPI?\n\n"
                           f" CA: {ca if ca else '-'}\n"
                           f" Descrição: {descricao}\n"
                           f" Quantidade: {quantidade_add}")
             if messagebox.askyesno("Confirmar Registro", confirm_msg, parent=self.epis_tab):
                novo_epi = {"CA": ca, "DESCRICAO": descricao, "QUANTIDADE": quantidade_add}
                try:
                    # Anexa nova linha
                    arquivo_epis = ARQUIVOS["epis"]
                    header = not os.path.exists(arquivo_epis) or os.path.getsize(arquivo_epis) == 0
                    pd.DataFrame([novo_epi]).to_csv(arquivo_epis, mode='a', header=header, index=False, encoding='utf-8')
                    self._update_status(f"Novo EPI {descricao} registrado com {quantidade_add} unidades.")
                    messagebox.showinfo("Sucesso", f"EPI '{descricao}' registrado com sucesso!", parent=self.epis_tab)
                    self._atualizar_tabela_epis()
                    # Limpa Campos
                    self.epi_ca_entry.delete(0, tk.END)
                    self.epi_desc_entry.delete(0, tk.END)
                    self.epi_qtd_entry.delete(0, tk.END)
                    self.epi_ca_entry.focus_set()

                except Exception as e:
                     self._update_status(f"Erro ao registrar novo EPI: {e}", error=True)
                     messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o novo EPI:\n{e}", parent=self.epis_tab)


    def _registrar_retirada(self):
        """Registra a retirada de EPI por um colaborador."""
        identificador = self.retirar_epi_id_entry.get().strip().upper()
        qtd_ret_str = self.retirar_epi_qtd_entry.get().strip().replace(",",".")
        colaborador = self.retirar_epi_colab_entry.get().strip().upper()

        if not identificador:
             messagebox.showerror("Erro", "CA ou Descrição do EPI deve ser informado.", parent=self.epis_tab)
             self.retirar_epi_id_entry.focus_set()
             return
        if not colaborador:
             messagebox.showerror("Erro", "Nome do Colaborador deve ser informado.", parent=self.epis_tab)
             self.retirar_epi_colab_entry.focus_set()
             return
        try:
             quantidade_retirada = float(qtd_ret_str)
             if quantidade_retirada <= 0: raise ValueError("Qtd deve ser positiva.")
        except ValueError:
             messagebox.showerror("Erro", "Quantidade a retirar deve ser um número positivo.", parent=self.epis_tab)
             self.retirar_epi_qtd_entry.focus_set()
             return

        df_epis = self._safe_read_csv(ARQUIVOS["epis"])
        # Garante tipos e limpa strings antes de buscar
        if "CA" not in df_epis.columns: df_epis["CA"] = ""
        if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""
        if "QUANTIDADE" not in df_epis.columns: df_epis["QUANTIDADE"] = 0

        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

        # Encontra EPI por CA (se não vazio) ou Descrição
        epi_match = pd.DataFrame() # Inicializa vazio
        if identificador: # Só busca se o identificador não for vazio
            # Prioriza busca por CA se o identificador parece ser um CA (ex: só números) ou se a busca por descrição falhar
            # Uma heurística simples: se for só número, tenta CA primeiro. Senão, tenta Descrição primeiro.
            is_likely_ca = identificador.isdigit()

            if is_likely_ca:
                match_ca = df_epis[df_epis["CA"] == identificador]
                if not match_ca.empty:
                    epi_match = match_ca
            
            if epi_match.empty: # Se não achou por CA (ou não tentou), tenta por descrição
                match_desc = df_epis[df_epis["DESCRICAO"] == identificador]
                if not match_desc.empty:
                    epi_match = match_desc
            
            # Se ainda vazio e não tentou CA, tenta CA agora
            if epi_match.empty and not is_likely_ca:
                 match_ca = df_epis[df_epis["CA"] == identificador]
                 if not match_ca.empty:
                     epi_match = match_ca


        if epi_match.empty:
             messagebox.showerror("Erro", f"EPI com CA/Descrição '{identificador}' não encontrado.", parent=self.epis_tab)
             self.retirar_epi_id_entry.focus_set()
             return
        
        # Se houver múltiplas correspondências (ex: mesma descrição com e sem CA), pega a primeira.
        # Idealmente, a interface de busca ajudaria a desambiguar.
        epi_data = epi_match.iloc[0]
        epi_index = epi_match.index[0]
        ca_epi = epi_data["CA"] # Usa o CA real do registro
        desc_epi = epi_data["DESCRICAO"]
        qtd_disponivel = epi_data["QUANTIDADE"]

        if quantidade_retirada > qtd_disponivel:
             epi_display_ca = f"(CA: {ca_epi})" if ca_epi else ""
             messagebox.showerror("Erro", f"Quantidade insuficiente para '{desc_epi}' {epi_display_ca}.\nDisponível: {qtd_disponivel}", parent=self.epis_tab)
             self.retirar_epi_qtd_entry.focus_set()
             return

        # Verifica/Cria Pasta do Colaborador
        # Remove caracteres inválidos para nome de pasta
        safe_colaborador_name = "".join(c for c in colaborador if c.isalnum() or c in (' ', '_')).rstrip()
        if not safe_colaborador_name:
             messagebox.showerror("Erro", "Nome do Colaborador inválido para criar pasta.", parent=self.epis_tab)
             return
        pasta_colaborador = os.path.join(COLABORADORES_DIR, safe_colaborador_name)
        
        try:
            os.makedirs(pasta_colaborador, exist_ok=True) # Cria se não existir
        except OSError as e:
             messagebox.showerror("Erro", f"Não foi possível criar a pasta para o colaborador '{safe_colaborador_name}':\n{e}", parent=self.epis_tab)
             return


        nova_qtd_epi = qtd_disponivel - quantidade_retirada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y") # Formato consistente

        epi_display_ca = f"(CA: {ca_epi})" if ca_epi else ""
        confirm_msg = (f"Confirmar Retirada?\n\n"
                       f"Colaborador: {colaborador}\n"
                       f"EPI: {desc_epi} {epi_display_ca}\n"
                       f"Qtd. Retirar: {quantidade_retirada}\n"
                       f"Qtd. Restante: {nova_qtd_epi}")

        if messagebox.askyesno("Confirmar Retirada", confirm_msg, parent=self.epis_tab):
             # 1. Atualiza quantidade de EPI
             df_epis.loc[epi_index, "QUANTIDADE"] = nova_qtd_epi
             if not self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                  messagebox.showerror("Erro Crítico", "Falha ao atualizar a quantidade de EPIs. A retirada NÃO foi registrada.", parent=self.epis_tab)
                  # Considerar reverter a leitura do df_epis aqui? Ou forçar recarga?
                  self._carregar_epis() # Recarrega para refletir o estado real
                  return # Para se a atualização do EPI falhar

             # 2. Registra retirada no arquivo do colaborador
             nome_arquivo_colab = f"{safe_colaborador_name}_{datetime.now().strftime('%Y_%m')}.csv"
             caminho_arquivo_colab = os.path.join(pasta_colaborador, nome_arquivo_colab)
             colab_file_data = {
                 "CA": ca_epi if ca_epi else '', # Garante que é string
                 "DESCRICAO": desc_epi,
                 "QTD RETIRADA": quantidade_retirada, "DATA": data_hora
             }
             try:
                 # Define as colunas esperadas para o arquivo do colaborador
                 colab_cols = ["CA", "DESCRICAO", "QTD RETIRADA", "DATA"]
                 header_colab = not os.path.exists(caminho_arquivo_colab) or os.path.getsize(caminho_arquivo_colab) == 0
                 
                 # Cria DataFrame com colunas na ordem correta
                 df_colab_append = pd.DataFrame([colab_file_data], columns=colab_cols) 
                 
                 df_colab_append.to_csv(caminho_arquivo_colab, mode='a', header=header_colab, index=False, encoding='utf-8')

                 # Sucesso
                 messagebox.showinfo("Sucesso", f"Retirada de {quantidade_retirada} '{desc_epi}' registrada para {colaborador}.", parent=self.epis_tab)
                 self._update_status(f"Retirada EPI {desc_epi} para {colaborador}.")
                 self._atualizar_tabela_epis() # Atualiza a tabela principal de EPIs
                 # Limpa campos
                 self.retirar_epi_id_entry.delete(0, tk.END)
                 self.retirar_epi_qtd_entry.delete(0, tk.END)
                 self.retirar_epi_colab_entry.delete(0, tk.END)
                 self.retirar_epi_id_entry.focus_set()

             except Exception as e:
                  self._update_status(f"Erro ao salvar retirada no arquivo do colaborador {colaborador}: {e}", error=True)
                  # Tentar reverter a quantidade de EPI? Arriscado sem transações.
                  # Lê novamente o arquivo de EPIs para obter o estado antes da falha no log do colaborador
                  df_epis_revert = self._safe_read_csv(ARQUIVOS["epis"])
                  # Encontra o índice novamente (pode ter mudado se houve erro na leitura?)
                  # É mais seguro apenas informar o usuário sobre a inconsistência.
                  messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a retirada no arquivo de {colaborador}, mas a quantidade de EPIs FOI alterada.\nVerifique manualmente.\n\nDetalhe: {e}", parent=self.epis_tab)
                  # Recarrega a tabela de EPIs para mostrar o estado atual (com a quantidade já debitada)
                  self._atualizar_tabela_epis()


    # --- Lançadores de Diálogo de Busca ---
    # ... (O código dos métodos _show_product_lookup, _show_epi_lookup permanece o mesmo) ...
    def _show_product_lookup(self, target_field_prefix):
         """Mostra diálogo de busca para produtos e preenche a entrada."""
         df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
         if df_estoque.empty:
             messagebox.showinfo("Estoque Vazio", "Não há produtos cadastrados no estoque para buscar.", parent=self.movimentacao_tab)
             return

         # Passa o estilo para os botões do diálogo, se necessário (requer modificação no LookupDialog)
         dialog = LookupDialog(self.root, "Buscar Produto no Estoque", df_estoque, ["CODIGO", "DESCRICAO"], "CODIGO")
         result_code = dialog.result # Isso bloqueia até o diálogo ser fechado

         if result_code:
              if target_field_prefix == "entrada":
                   self.entrada_codigo_entry.delete(0, tk.END)
                   self.entrada_codigo_entry.insert(0, result_code)
                   self.entrada_qtd_entry.focus_set() # Foca o próximo campo
              elif target_field_prefix == "saida":
                   self.saida_codigo_entry.delete(0, tk.END)
                   self.saida_codigo_entry.insert(0, result_code)
                   self.saida_solicitante_entry.focus_set() # Foca o próximo campo

    def _show_epi_lookup(self):
         """Mostra diálogo de busca para EPIs e preenche a entrada."""
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         if df_epis.empty:
             messagebox.showinfo("EPIs Vazios", "Não há EPIs cadastrados para buscar.", parent=self.epis_tab)
             return
         
         # Garante colunas para busca
         if "CA" not in df_epis.columns: df_epis["CA"] = ""
         if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""
         
         # Colunas para buscar e coluna para retornar (pode ser CA ou Descrição)
         search_cols = ["CA", "DESCRICAO"]
         return_col = "CA" # Ou "DESCRICAO", dependendo da preferência

         dialog = LookupDialog(self.root, "Buscar EPI", df_epis, search_cols, return_col)
         result_id = dialog.result # ID retornado (CA ou Descrição, conforme return_col)

         if result_id:
              # Tenta encontrar o EPI exato com base no ID retornado (seja CA ou Descrição)
              # para obter o identificador preferencial (CA se existir, senão Descrição) para exibir no campo.
              
              # Limpa os dados do DataFrame para correspondência
              df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
              df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
              result_id_upper = str(result_id).upper() # Garante comparação correta

              found_epi = df_epis[(df_epis[return_col].astype(str).str.upper() == result_id_upper)]

              display_text = result_id # Padrão é o que foi retornado

              if not found_epi.empty:
                  # Pega o primeiro encontrado se houver múltiplos
                  epi_data = found_epi.iloc[0]
                  # Decide o que exibir: CA se não estiver vazio, senão Descrição
                  display_text = epi_data['CA'] if epi_data['CA'] else epi_data['DESCRICAO']


              self.retirar_epi_id_entry.delete(0, tk.END)
              self.retirar_epi_id_entry.insert(0, display_text) # Usa o texto decidido
              self.retirar_epi_qtd_entry.focus_set()


    # --- Funcionalidade Editar / Excluir ---
    # ... (O código dos métodos _get_selected_data, _edit_selected_item, _delete_selected_item, _edit_selected_epi, _delete_selected_epi permanece o mesmo) ...
    def _get_selected_data(self, table):
        """
        Obtém dados para a única linha selecionada na pandastable especificada.
        Usa o número da linha visual para obter o rótulo de índice correspondente do DataFrame do modelo atual.
        Retorna uma pandas Series ou None se a seleção for inválida ou os dados não puderem ser recuperados.
        """
        # Verifica se a tabela e o modelo existem
        if not table or not hasattr(table, 'model'):
            print("Debug: Tabela ou modelo não encontrado em _get_selected_data")
            return None
            
        model = table.model
        # Garante que o modelo e seu atributo dataframe existam
        if not hasattr(model, 'df') or model.df is None:
            # Este caso pode ocorrer se a tabela ainda não foi populada
            # ou se algo deu errado durante a atualização do modelo.
            # print("Debug: Modelo ou model.df não encontrado em _get_selected_data")
            return None

        # getSelectedRow() retorna o número da linha VISUAL baseado em 0 na tabela
        # Verifica se o método existe antes de chamar
        if not hasattr(table, 'getSelectedRow'):
             print("Debug: Método getSelectedRow não encontrado na tabela.")
             return None
             
        row_num = table.getSelectedRow()

        if row_num < 0:
            # Nenhuma linha está visualmente selecionada, ou ocorreu um erro ao relatar a seleção
            # print(f"Debug: Nenhuma linha selecionada (row_num={row_num})")
            return None

        # Obtém o DataFrame atualmente sendo exibido pelo modelo da tabela
        current_df = model.df

        # *** VERIFICAÇÃO CRÍTICA ***
        # Verifica se o número da linha visualmente selecionada é um índice posicional válido
        # para o DataFrame *atual* sendo exibido.
        if row_num >= len(current_df):
            # Esta condição significa que a seleção visual está apontando para fora dos limites
            # dos dados reais atualmente no modelo da tabela. Isso pode acontecer
            # durante atualizações rápidas, conflitos de filtragem ou inconsistências de estado.
            print(f"Debug: Erro - linha visual selecionada {row_num} está fora dos limites para o comprimento atual do DataFrame {len(current_df)}")
            # Opcionalmente, mostre uma mensagem ao usuário:
            # messagebox.showwarning("Seleção Inválida", f"A linha selecionada ({row_num}) parece inválida. A tabela pode ter sido atualizada. Tente selecionar novamente.", parent=self.root)
            # table.clearSelection() # Opcionalmente, limpa a seleção inválida
            return None # Indica que dados válidos não puderam ser recuperados

        try:
            # Se o número da linha visual for válido para o comprimento do DataFrame atual,
            # acessa o índice do DataFrame usando esse número posicional.
            # Isso nos dá o RÓTULO de índice real (pode ser inteiro, string, etc.)
            index_label = current_df.index[row_num]

            # Agora, recupera os dados da linha usando o RÓTULO de índice confiável via .loc
            selected_series = current_df.loc[index_label]

            # print(f"Debug: Dados recuperados com sucesso para a linha visual {row_num}, rótulo de índice {index_label}")
            return selected_series # Retorna a pandas Series

        except IndexError:
            # Este IndexError pode ocorrer se, apesar da verificação de comprimento,
            # o próprio índice estiver de alguma forma inválido ou inconsistente no momento do acesso.
            # A verificação de comprimento (`row_num >= len(current_df)`) deve prevenir a maioria destes.
            print(f"Debug: IndexError ao tentar acessar o rótulo de índice na posição {row_num} no DataFrame atual com comprimento {len(current_df)}.")
            # messagebox.showerror("Erro de Seleção", f"Não foi possível encontrar o índice para a linha selecionada ({row_num}). A tabela pode ter sido atualizada.", parent=self.root)
            return None
        except Exception as e:
            # Captura quaisquer outros erros inesperados durante a recuperação de dados
            print(f"Debug: Erro inesperado em _get_selected_data para linha visual {row_num}: {e}")
            messagebox.showerror("Erro Interno", f"Ocorreu um erro ao obter dados da linha selecionada.\n\nDetalhes: {e}", parent=self.root)
            return None


    def _edit_selected_item(self):
         """Lida com a edição do item selecionado na visualização da tabela principal."""
         if not self.pandas_table or self.active_table_name != "estoque":
              messagebox.showwarning("Ação Inválida", "A edição só está disponível para a tabela de Estoque.", parent=self.estoque_tab)
              return

         selected_item_series = self._get_selected_data(self.pandas_table)

         if selected_item_series is None:
              messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único item do estoque para editar.", parent=self.estoque_tab)
              return

         item_dict = selected_item_series.to_dict()
         # Garante que Codigo esteja presente
         if 'CODIGO' not in item_dict:
             messagebox.showerror("Erro Interno", "Não foi possível obter o código do item selecionado.", parent=self.estoque_tab)
             return
         codigo_to_edit = str(item_dict['CODIGO']) # Obtém o código antes de abrir o diálogo

         # Mostra Diálogo de Edição
         dialog = EditProductDialog(self.root, f"Editar Produto: {codigo_to_edit}", item_dict)
         updated_data = dialog.updated_data # Isso bloqueia

         if updated_data: # Usuário clicou em Salvar e a validação passou
              # Salva alterações de volta no DataFrame principal (primeiro na memória)
              df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
              if 'CODIGO' not in df_estoque.columns:
                   messagebox.showerror("Erro", "Coluna 'CODIGO' não encontrada no arquivo de estoque para salvar edição.", parent=self.estoque_tab)
                   return
              df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)

              idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_edit].tolist()
              if not idx:
                    messagebox.showerror("Erro", f"Item {codigo_to_edit} não encontrado no arquivo para salvar após edição.", parent=self.estoque_tab)
                    return
              idx = idx[0]

              try:
                    # Atualiza campos do resultado do diálogo
                    for key, value in updated_data.items():
                         if key in df_estoque.columns:
                              # Lida com possíveis problemas de conversão de tipo antes da atribuição
                              current_dtype = df_estoque[key].dtype
                              try:
                                   if pd.api.types.is_numeric_dtype(current_dtype) and not pd.isna(value):
                                        # Tenta converter para numérico, tratando strings vazias ou inválidas
                                        value_str = str(value).strip()
                                        if value_str: # Só converte se não for vazio
                                            value = pd.to_numeric(value_str.replace(',', '.'))
                                        else: # Se for string vazia, talvez definir como 0 ou NaN? Vamos usar 0.
                                            value = 0
                                   elif pd.api.types.is_string_dtype(current_dtype) or pd.api.types.is_object_dtype(current_dtype):
                                        # Garante que é string e remove espaços extras
                                        value = str(value).strip()
                                   # Adiciona outras verificações de tipo se necessário (ex: datas)
                              except (ValueError, TypeError):
                                    messagebox.showwarning("Aviso de Tipo", f"Não foi possível converter '{value}' para o tipo da coluna '{key}'. Mantendo valor original.", parent=self.estoque_tab)
                                    continue # Pula a atualização deste campo se a conversão falhar
                              df_estoque.loc[idx, key] = value

                    # Atualiza timestamp se a coluna DATA existir
                    if "DATA" in df_estoque.columns:
                        df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y")
                    else:
                         print("Aviso: Coluna 'DATA' não encontrada para atualizar timestamp durante edição.")


                    # Salva DataFrame atualizado para CSV
                    if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                        messagebox.showinfo("Sucesso", f"Produto {codigo_to_edit} atualizado.", parent=self.estoque_tab)
                        self._atualizar_tabela_atual() # Atualiza visualização
                    # else: Mensagem de erro mostrada por _safe_write_csv

              except Exception as e:
                  self._update_status(f"Erro ao aplicar edições para {codigo_to_edit}: {e}", error=True)
                  messagebox.showerror("Erro ao Salvar Edição", f"Não foi possível salvar as alterações para o produto {codigo_to_edit}.\n\nDetalhes: {e}", parent=self.estoque_tab)


    def _delete_selected_item(self):
         """Lida com a exclusão do item selecionado da tabela Estoque."""
         if not self.pandas_table or self.active_table_name != "estoque":
              messagebox.showwarning("Ação Inválida", "A exclusão só está disponível para a tabela de Estoque.", parent=self.estoque_tab)
              return

         selected_item_series = self._get_selected_data(self.pandas_table)

         if selected_item_series is None:
              messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único item do estoque para excluir.", parent=self.estoque_tab)
              return

         codigo_to_delete = str(selected_item_series.get('CODIGO', 'N/A'))
         desc_to_delete = selected_item_series.get('DESCRICAO', 'N/A')

         confirm_msg = (f"Tem certeza que deseja excluir permanentemente o item abaixo?\n\n"
                       f" Código: {codigo_to_delete}\n"
                       f" Descrição: {desc_to_delete}\n\n"
                       f"Esta ação não pode ser desfeita.")

         if messagebox.askyesno("Confirmar Exclusão", confirm_msg, icon='warning'):
              # Carrega os dados, encontra o índice, remove a linha, salva
              df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
              if 'CODIGO' not in df_estoque.columns:
                   messagebox.showerror("Erro", "Coluna 'CODIGO' não encontrada no arquivo de estoque para excluir.", parent=self.estoque_tab)
                   return
              df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
              idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_delete].tolist()

              if not idx:
                  messagebox.showerror("Erro", f"Item {codigo_to_delete} não encontrado no arquivo para excluir.", parent=self.estoque_tab)
                  return

              df_estoque.drop(idx, inplace=True) # Remove linhas pela lista de índices

              # Salva DataFrame atualizado para CSV
              if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                   messagebox.showinfo("Sucesso", f"Produto {codigo_to_delete} - {desc_to_delete} excluído.", parent=self.estoque_tab)
                   self._atualizar_tabela_atual() # Atualiza visualização
              # else: Mensagem de erro mostrada por _safe_write_csv


    def _edit_selected_epi(self):
        """Lida com a edição do EPI selecionado."""
        selected_epi_series = self._get_selected_data(self.epis_table)

        if selected_epi_series is None:
             messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único EPI para editar.", parent=self.epis_tab)
             return

        epi_dict = selected_epi_series.to_dict()
        # Precisamos de um identificador único para encontrá-lo mais tarde. Usa CA se disponível, senão Descrição.
        ca_to_edit = epi_dict.get('CA',"").strip()
        desc_to_edit = epi_dict.get('DESCRICAO',"").strip()
        if not (ca_to_edit or desc_to_edit):
             messagebox.showerror("Erro Interno", "EPI selecionado não tem CA ou Descrição para identificar.", parent=self.epis_tab)
             return

        dialog = EditEPIDialog(self.root, f"Editar EPI: {desc_to_edit if desc_to_edit else ca_to_edit}", epi_dict)
        updated_data = dialog.updated_data

        if updated_data:
             df_epis = self._safe_read_csv(ARQUIVOS["epis"])
             # Garante colunas e tipos antes de corresponder/atualizar
             if "CA" not in df_epis.columns: df_epis["CA"] = ""
             if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""
             if "QUANTIDADE" not in df_epis.columns: df_epis["QUANTIDADE"] = 0

             df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
             df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
             df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)


             # Encontra a linha do EPI original novamente com base no CA primariamente (se não vazio), depois descrição
             match = pd.DataFrame()
             if ca_to_edit:
                 match = df_epis[df_epis["CA"] == ca_to_edit]
             
             # Se não achou por CA ou CA original era vazio, tenta por descrição (se não vazia)
             if match.empty and desc_to_edit:
                 # Para evitar conflito se CA original era vazio, busca onde CA é vazio E descrição bate
                 if not ca_to_edit:
                      match = df_epis[(df_epis["DESCRICAO"] == desc_to_edit) & (df_epis["CA"] == "")]
                 else: # Se CA original existia mas não bateu, busca só pela descrição (menos seguro)
                      match = df_epis[df_epis["DESCRICAO"] == desc_to_edit]


             if match.empty:
                  messagebox.showerror("Erro", f"EPI original (CA:'{ca_to_edit}'/Desc:'{desc_to_edit}') não encontrado no arquivo para salvar.", parent=self.epis_tab)
                  return

             idx = match.index[0] # Pega o primeiro índice se houver múltiplos

             try:
                    # Verifica possíveis conflitos de CA/Descrição antes de atualizar
                    new_ca = updated_data['CA'].strip().upper()
                    new_desc = updated_data['DESCRICAO'].strip().upper()

                    # Verifica se o *novo* CA (se não vazio) existe em outra linha
                    if new_ca and any((df_epis['CA'] == new_ca) & (df_epis.index != idx)):
                        messagebox.showerror("Erro de Duplicidade", f"O CA '{new_ca}' já existe para outro EPI.", parent=self.epis_tab)
                        return
                    # Verifica se a *nova* Descrição (se não vazia) existe em outra linha E o CA é diferente (ou novo CA é vazio)
                    if new_desc:
                        desc_conflict = df_epis[
                            (df_epis['DESCRICAO'] == new_desc) &
                            (df_epis.index != idx) &
                            (df_epis['CA'] != new_ca if new_ca else df_epis['CA'] != "") # Conflito se CAs forem diferentes, ou se novo CA for vazio e o existente não
                        ]
                        if not desc_conflict.empty:
                            messagebox.showerror("Erro de Duplicidade", f"A Descrição '{new_desc}' já existe para outro EPI com CA diferente ou vazio.", parent=self.epis_tab)
                            return

                    # Atualiza campos do resultado do diálogo
                    df_epis.loc[idx, "CA"] = new_ca
                    df_epis.loc[idx, "DESCRICAO"] = new_desc
                    try:
                        df_epis.loc[idx, "QUANTIDADE"] = float(str(updated_data['QUANTIDADE']).replace(',','.'))
                    except ValueError:
                         messagebox.showwarning("Aviso", f"Quantidade inválida '{updated_data['QUANTIDADE']}' ao editar EPI. Mantendo valor anterior.", parent=self.epis_tab)
                         # Não atualiza a quantidade se for inválida


                    # Salva DataFrame atualizado para CSV
                    if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                        messagebox.showinfo("Sucesso", f"EPI '{new_desc if new_desc else new_ca}' atualizado.", parent=self.epis_tab)
                        self._atualizar_tabela_epis() # Atualiza visualização
                    # else: Erro tratado

             except Exception as e:
                 self._update_status(f"Erro ao aplicar edições EPI: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Edição EPI", f"Não foi possível salvar as alterações.\n\nDetalhes: {e}", parent=self.epis_tab)


    def _delete_selected_epi(self):
        """Lida com a exclusão do EPI selecionado."""
        selected_epi_series = self._get_selected_data(self.epis_table)

        if selected_epi_series is None:
             messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único EPI para excluir.", parent=self.epis_tab)
             return

        ca_to_delete = selected_epi_series.get('CA', '').strip()
        desc_to_delete = selected_epi_series.get('DESCRICAO', 'N/A')

        confirm_msg = (f"Tem certeza que deseja excluir permanentemente o EPI abaixo?\n\n"
                       f" CA: {ca_to_delete if ca_to_delete else '-'}\n"
                       f" Descrição: {desc_to_delete}\n\n"
                       f"Esta ação não pode ser desfeita.")

        if messagebox.askyesno("Confirmar Exclusão EPI", confirm_msg, icon='warning'):
            df_epis = self._safe_read_csv(ARQUIVOS["epis"])
            # Garante colunas e tipos
            if "CA" not in df_epis.columns: df_epis["CA"] = ""
            if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""

            df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
            df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()

            # Encontra com base no CA primariamente (se não vazio), depois descrição
            match = pd.DataFrame()
            if ca_to_delete:
                 match = df_epis[df_epis["CA"] == ca_to_delete]
            
            if match.empty and desc_to_delete:
                 # Se CA era vazio, busca onde CA é vazio E descrição bate
                 if not ca_to_delete:
                      match = df_epis[(df_epis["DESCRICAO"] == desc_to_delete) & (df_epis["CA"] == "")]
                 else: # Se CA existia mas não bateu, busca só pela descrição
                      match = df_epis[df_epis["DESCRICAO"] == desc_to_delete]


            if match.empty:
                 messagebox.showerror("Erro", f"EPI original (CA:'{ca_to_delete}'/Desc:'{desc_to_delete}') não encontrado no arquivo para excluir.", parent=self.epis_tab)
                 return

            idx = match.index.tolist() # Obtém todos os índices correspondentes
            df_epis.drop(idx, inplace=True)

            if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                 messagebox.showinfo("Sucesso", f"EPI '{desc_to_delete}' excluído.", parent=self.epis_tab)
                 self._atualizar_tabela_epis() # Atualiza visualização
            # else: Erro tratado


    # --- Backup e Exportação ---
    # ... (O código dos métodos _criar_backup_periodico, _schedule_backup, _exportar_conteudo permanece o mesmo) ...
    def _criar_backup_periodico(self):
        """Cria backups com timestamp dos arquivos de dados."""
        arquivo_ultimo_backup = os.path.join(BACKUP_DIR, "ultimo_backup_timestamp.txt")
        backup_interval_seconds = 3 * 60 * 60 # 3 horas

        # --- Limpeza de Backups Antigos (Mais robustamente) ---
        try:
            if not os.path.exists(BACKUP_DIR): # Verifica se a pasta de backup existe
                 print(f"Pasta de backup '{BACKUP_DIR}' não encontrada. Pulando limpeza.")
            else:
                now = time.time()
                cutoff_time = now - (3 * 24 * 60 * 60) # 3 dias atrás
                for filename in os.listdir(BACKUP_DIR):
                    # Verifica se é um backup CSV gerado por esta lógica
                    if filename.endswith(".csv") and ("_backup_" in filename or "_auto.csv" in filename):
                        file_path = os.path.join(BACKUP_DIR, filename)
                        try:
                            file_mod_time = os.path.getmtime(file_path)
                            if file_mod_time < cutoff_time:
                                os.remove(file_path)
                                print(f"Backup antigo removido: {filename}")
                        except OSError as e:
                             print(f"Erro ao processar/remover backup antigo {filename}: {e}") # Registra erro mas continua
        except Exception as e:
            self._update_status(f"Erro ao limpar backups antigos: {e}", error=True) # Registra erro geral de limpeza


        # --- Cria Novo Backup se o Intervalo Passou ---
        perform_backup = True
        if os.path.exists(arquivo_ultimo_backup):
            try:
                with open(arquivo_ultimo_backup, "r", encoding="utf-8") as f:
                    last_backup_content = f.read().strip()
                    if last_backup_content: # Verifica se o arquivo não está vazio
                         ultimo_backup_time = float(last_backup_content)
                         if (time.time() - ultimo_backup_time) < backup_interval_seconds:
                             perform_backup = False
                    else:
                         print("Arquivo de timestamp do último backup está vazio. Criando novo backup.")
                         perform_backup = True
            except (ValueError, FileNotFoundError, TypeError) as e: # Adicionado TypeError
                print(f"Erro ao ler timestamp do último backup: {e}. Criando novo backup.")
                perform_backup = True # Força backup se o arquivo de timestamp estiver ruim ou ilegível

        if perform_backup:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True) # Garante que a pasta existe antes de copiar
                backup_count = 0
                for nome, arquivo_origem in ARQUIVOS.items():
                    if os.path.exists(arquivo_origem):
                        try:
                            nome_backup = f"{nome}_{timestamp}_auto.csv"
                            caminho_backup = os.path.join(BACKUP_DIR, nome_backup)
                            shutil.copy2(arquivo_origem, caminho_backup) # copy2 preserva metadados
                            backup_count += 1
                        except Exception as copy_e:
                             self._update_status(f"Erro ao copiar {arquivo_origem} para backup: {copy_e}", error=True)
                             print(f"Erro ao copiar {arquivo_origem} para backup: {copy_e}")


                # Só atualiza timestamp se o backup foi bem-sucedido (pelo menos um arquivo)
                if backup_count > 0:
                    try:
                        with open(arquivo_ultimo_backup, "w", encoding="utf-8") as f:
                            f.write(str(time.time()))
                        self._update_status(f"Backup automático criado ({backup_count} arquivos).")
                        print(f"Backup automático criado em {timestamp}")
                    except Exception as ts_write_e:
                         self._update_status(f"Erro ao escrever timestamp do backup: {ts_write_e}", error=True)
                         print(f"Erro ao escrever timestamp do backup: {ts_write_e}")

                else:
                     print("Nenhum arquivo de dados encontrado ou copiado para backup.")

            except Exception as e:
                self._update_status(f"Erro durante backup automático: {e}", error=True)
                # Evita messagebox se a root não existir mais (caso raro)
                if self.root and self.root.winfo_exists():
                    messagebox.showerror("Erro de Backup", f"Falha ao criar backup automático:\n{e}", parent=self.root)

    def _schedule_backup(self):
        """Chama a função de backup periodicamente."""
        self._criar_backup_periodico()
        # Reagenda apenas se a janela ainda existir
        if self.root and self.root.winfo_exists():
            self.root.after(10800000, self._schedule_backup) # Reagenda (3 horas)


    def _exportar_conteudo(self):
        """Exporta dados atuais para Excel e gera relatório txt para estoque baixo."""
        pasta_saida = "Relatorios"
        try:
            os.makedirs(pasta_saida, exist_ok=True)
        except OSError as e:
             messagebox.showerror("Erro de Exportação", f"Não foi possível criar a pasta de relatórios '{pasta_saida}':\n{e}", parent=self.root)
             return


        data_atual = datetime.now().strftime("%d-%m-%Y_%H%M")
        caminho_excel = os.path.join(pasta_saida, f"Relatorio_Almoxarifado_{data_atual}.xlsx")
        caminho_txt = os.path.join(pasta_saida, f"Produtos_Esgotados_{data_atual}.txt")

        try:
            with pd.ExcelWriter(caminho_excel) as writer:
                # Inclui EPIs na exportação
                all_files = {**ARQUIVOS} # Cria cópia do dicionário
                self._update_status("Iniciando exportação para Excel...")
                sheet_count = 0
                for nome, arquivo in all_files.items():
                     try:
                         df_export = self._safe_read_csv(arquivo) # Lê dados frescos
                         if not df_export.empty:
                             # Usa um nome de planilha mais seguro (remove caracteres inválidos)
                             safe_sheet_name = "".join(c for c in nome.capitalize() if c.isalnum() or c in (' ', '_'))[:31] # Limita a 31 caracteres
                             df_export.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                             sheet_count +=1
                         else:
                             print(f"Planilha '{nome}' vazia ou não encontrada, não será incluída no Excel.")
                     except Exception as e:
                          print(f"Erro ao processar {arquivo} para Excel: {e}")
                          # Evita messagebox se a root não existir
                          if self.root and self.root.winfo_exists():
                              messagebox.showwarning("Aviso de Exportação", f"Erro ao incluir '{nome}' no Excel:\n{e}", parent=self.root)

            self._update_status(f"Exportação Excel concluída ({sheet_count} planilhas). Gerando relatório de esgotados...")

            # Relatório de Produtos Esgotados
            df_estoque_report = self._safe_read_csv(ARQUIVOS["estoque"])
            # Verifica se as colunas essenciais existem
            required_cols = ["QUANTIDADE", "CODIGO", "DESCRICAO"]
            if not df_estoque_report.empty and all(col in df_estoque_report.columns for col in required_cols):
                 # Garante que quantidade seja numérica antes de filtrar
                 df_estoque_report["QUANTIDADE"] = pd.to_numeric(df_estoque_report["QUANTIDADE"], errors='coerce').fillna(0)
                 produtos_esgotados = df_estoque_report[df_estoque_report["QUANTIDADE"] <= 0] # Mantido <= 0

                 try:
                     with open(caminho_txt, "w", encoding="utf-8") as f:
                         f.write(f"Relatório de Produtos Esgotados/Zerados - {data_atual}\n")
                         f.write("=" * 50 + "\n")
                         if not produtos_esgotados.empty:
                             for _, row in produtos_esgotados.iterrows():
                                  # Garante que CODIGO e DESCRICAO são strings para o f-string
                                  codigo_str = str(row['CODIGO'])
                                  desc_str = str(row['DESCRICAO'])
                                  qtd_str = str(row['QUANTIDADE'])
                                  f.write(f"Código: {codigo_str} | Descrição: {desc_str} | Qtd: {qtd_str}\n")
                         else:
                             f.write("Nenhum produto com quantidade zero ou negativa encontrado.\n")
                         f.write("=" * 50 + "\n")

                     # Evita messagebox se a root não existir
                     if self.root and self.root.winfo_exists():
                         messagebox.showinfo("Sucesso", f"Relatórios exportados com sucesso!\n\n"
                                                        f"Excel: {caminho_excel}\n"
                                                        f"Esgotados: {caminho_txt}", parent=self.root)
                     self._update_status("Relatórios exportados com sucesso.")

                 except Exception as txt_e:
                      self._update_status(f"Erro ao gerar relatório de esgotados: {txt_e}", error=True)
                      if self.root and self.root.winfo_exists():
                           messagebox.showerror("Erro de Exportação", f"Falha ao gerar relatório de produtos esgotados:\n{txt_e}", parent=self.root)


            else:
                  missing_cols = [col for col in required_cols if col not in df_estoque_report.columns]
                  reason = "dados do estoque não encontrados" if df_estoque_report.empty else f"colunas ausentes: {', '.join(missing_cols)}"
                  warning_msg = f"Não foi possível gerar relatório de produtos esgotados ({reason})."
                  if self.root and self.root.winfo_exists():
                      messagebox.showwarning("Aviso", warning_msg, parent=self.root)
                  self._update_status(f"Relatório de esgotados não gerado ({reason}).")

        except Exception as e:
            self._update_status(f"Erro geral ao exportar relatórios: {e}", error=True)
            if self.root and self.root.winfo_exists():
                messagebox.showerror("Erro de Exportação", f"Falha ao exportar relatórios:\n{e}", parent=self.root)


    # --- Fechamento da Aplicação ---
    # ... (O código do método _on_close permanece o mesmo) ...
    def _on_close(self):
        """Lida com o evento de fechamento da janela."""
        # Verifica se a janela ainda existe antes de mostrar messagebox
        if self.root and self.root.winfo_exists():
            if messagebox.askyesno("Confirmar Saída", "Deseja realmente fechar o aplicativo?", icon='warning', parent=self.root):
                print("Fechando aplicativo...")
                # Cancela qualquer agendamento pendente do 'after' para evitar erros
                # (Embora destruir a root geralmente cuide disso)
                # self.root.after_cancel(self._backup_schedule_id) # Precisaria armazenar o ID retornado por after
                self.root.destroy()
        else:
             print("Fechando aplicativo (janela já não existia).")


# --- Classe da Janela de Login ---

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Almoxarifado - Login")
        self.root.geometry("300x250")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_login) # Lida com o fechamento da janela de login

        self.logged_in_user_id = None # Armazena o ID no login bem-sucedido

        # Estilo
        style = ttk.Style()
        style.theme_use('clam')

        # --- Estilos para Login ---
        # Botão de Login Principal
        style.configure("Login.TButton",
                        foreground="white",
                        background="#107C10", # Verde
                        font=('-size', 11, '-weight', 'bold'))
        style.map("Login.TButton",
                  background=[('active', '#0A530A')]) # Verde mais escuro

        # Frame
        login_frame = ttk.Frame(root, padding="20")
        login_frame.pack(expand=True, fill="both")

        ttk.Label(login_frame, text="Login", font="-size 16 -weight bold").pack(pady=(0, 15))

        ttk.Label(login_frame, text="Usuário:", font="-size 12").pack(pady=(5, 0))
        self.usuario_entry = ttk.Entry(login_frame, width=30)
        self.usuario_entry.config(font="-size 12")
        self.usuario_entry.pack(pady=(0, 10))
        self.usuario_entry.bind("<Return>", lambda e: self.senha_entry.focus_set()) # Foca senha no Enter

        ttk.Label(login_frame, text="Senha:", font="-size 12").pack(pady=(5, 0))
        self.senha_entry = ttk.Entry(login_frame, show="*", width=30)
        self.senha_entry.config(font="-size 12")
        self.senha_entry.pack(pady=(0, 15))
        self.senha_entry.bind("<Return>", lambda e: self._validate_login()) # Valida no Enter

        # Aplicando estilo ao botão de login
        login_button = ttk.Button(login_frame, text="Entrar", command=self._validate_login, style="Login.TButton")
        # login_button.config(width=25, padding=5) # Padding já está no estilo
        login_button.pack(pady=5, fill=tk.X, ipady=4) # Faz o botão preencher horizontalmente e adiciona padding interno vertical

        # Centraliza a janela
        self.root.update_idletasks() # Garante que a geometria seja atualizada
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f'+{x}+{y}')

        self.usuario_entry.focus_set()


    def _validate_login(self):
        """Valida as credenciais do usuário."""
        usuario = self.usuario_entry.get().strip()
        senha = self.senha_entry.get() # Obtém a senha exatamente como digitada

        if not usuario or not senha:
            messagebox.showwarning("Login Inválido", "Por favor, preencha usuário e senha.", parent=self.root)
            return

        # AVISO: Comparação de texto plano! Inseguro! Use hashing (ex: bcrypt).
        # Considerar usar case-insensitive para nome de usuário?
        # usuario_lower = usuario.lower()
        # if usuario_lower in usuarios and usuarios[usuario_lower]["senha"] == senha:

        if usuario in usuarios and usuarios[usuario]["senha"] == senha:
            self.logged_in_user_id = usuarios[usuario]["id"]
            messagebox.showinfo("Sucesso", f"Login bem-sucedido!\nOperador ID: {self.logged_in_user_id}", parent=self.root)
            self.root.destroy() # Fecha a janela de login
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos!", parent=self.root)
            self.senha_entry.delete(0, tk.END) # Limpa o campo de senha na falha
            self.senha_entry.focus_set() # Foca novamente na senha


    def _on_close_login(self):
        """Lida com o fechamento direto da janela de login."""
        # Verifica se a janela ainda existe
        if self.root and self.root.winfo_exists():
            if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do aplicativo?", icon='question', parent=self.root):
                 self.logged_in_user_id = None # Garante que nenhum ID de usuário seja definido
                 self.root.destroy() # Fecha a janela
                 # sys.exit() ou os._exit(0) podem ser usados aqui se necessário, mas destruir a root geralmente é suficiente
        else:
            print("Saindo (janela de login já não existia).")


    def get_user_id(self):
         return self.logged_in_user_id


# --- Execução Principal ---

if __name__ == "__main__":
    # 1. Mostra Janela de Login
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()

    # 2. Prossegue apenas se o login foi bem-sucedido
    user_id = login_app.get_user_id()
    if user_id:
        # 3. Lança Janela Principal da Aplicação
        main_root = tk.Tk()
        app = AlmoxarifadoApp(main_root, user_id)
        main_root.mainloop()
    else:
        print("Login cancelado ou falhou. Saindo.")

