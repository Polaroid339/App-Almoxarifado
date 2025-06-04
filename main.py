import os
import csv
import time
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from usuarios import usuarios
from pandastable import Table, TableModel
import ast # <-- ADDED for parsing list strings from CSV
import re # <-- ADDED for date/time validation

# --- Configuração ---
PLANILHAS_DIR = "Planilhas"
BACKUP_DIR = "Backups"
COLABORADORES_DIR = "Colaboradores"

# Próximo ao topo do arquivo (onde já está definido)
EXPECTED_COLUMNS = {
    "estoque": ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"],
    "entrada": ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"],
    # MODIFICADO: Adiciona VALOR, SETOR, NF/PEDIDO a Saida
    "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "VALOR", "SOLICITANTE", "SETOR", "NF/PEDIDO", "DATA", "ID"],
    "epis": ["CA", "DESCRICAO", "QUANTIDADE"]
}

ARQUIVOS = {
    "estoque": os.path.join(PLANILHAS_DIR, "Estoque.csv"),
    "entrada": os.path.join(PLANILHAS_DIR, "Entrada.csv"),
    "saida": os.path.join(PLANILHAS_DIR, "Saida.csv"),
    "epis": os.path.join(PLANILHAS_DIR, "Epis.csv")
}

# --- Classes Auxiliares para Diálogos ---
# LookupDialog, EditDialogBase, EditProductDialog, EditEPIDialog remain the same for now
# (Could update EditProductDialog later to show new fields if needed)

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
        ttk.Button(button_frame, text="Selecionar", command=self._select_item, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.wait_window(self)

    def _populate_listbox(self, df):
        self.listbox.delete(0, tk.END)
        for index, row in df.iterrows():
            # Prioritize return_col for display, fallback gracefully
            id_val = row.get(self.return_col, '')
            desc_val = row.get('DESCRICAO', '')
            display_text = f"{id_val} - {desc_val}" if id_val and desc_val else str(id_val) if id_val else str(desc_val)
            self.listbox.insert(tk.END, display_text)


    def _filter_list(self, *args):
        query = self.search_var.get().strip().lower()
        df_to_search = self.df_full
        if not query:
            df_filtered = df_to_search
        else:
            try:
                mask = df_to_search.apply(
                    lambda row: any(query in str(row[col]).lower()
                                    for col in self.search_cols if pd.notna(row[col])),
                    axis=1
                )
                df_filtered = df_to_search[mask]
            except Exception as e:
                 print(f"Erro durante o filtro de busca: {e}")
                 df_filtered = df_to_search

        self._populate_listbox(df_filtered)

    def _select_item(self, event=None):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            list_item_text = self.listbox.get(index)
            # Extract identifier, handle cases where ' - ' might be missing
            parts = list_item_text.split(' - ', 1)
            self.result = parts[0].strip() if parts else "" # Return the first part or empty string
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
        self.geometry("400x350") # Adjusted size slightly

        self.item_data = item_data
        self.updated_data = None

        self.entries = {}
        self.comboboxes = {} # For dropdowns like CLASSIFICACAO
        self._create_widgets()

        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(button_frame, text="Salvar", command=self._save, style="Success.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _create_widgets(self):
        pass # Implementado por subclasses

    def _add_entry(self, frame, field_key, label_text, row, col, **kwargs):
        """Auxiliar para criar rótulo e entrada."""
        entry_state = kwargs.pop('state', tk.NORMAL)
        ttk.Label(frame, text=label_text + ":").grid(row=row, column=col*2, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(frame, **kwargs)
        entry.grid(row=row, column=col*2 + 1, sticky=tk.EW, padx=5, pady=2)

        if field_key in self.item_data:
            value = self.item_data[field_key] # Pega o valor uma vez

            # Verifica se é uma lista (o caso problemático)
            if isinstance(value, list):
                # Uma lista (mesmo vazia) não é 'na'. Apenas converte para string.
                value_to_insert = str(value)
            else:
                # Para todos os outros tipos (números, strings, None, NaN), usa pd.notna
                value_to_insert = str(value) if pd.notna(value) else ""

            entry.insert(0, value_to_insert) # Insere o valor formatado
        if entry_state == 'readonly':
            entry.config(state='readonly')

        self.entries[field_key] = entry
        frame.columnconfigure(col*2 + 1, weight=1)

    # --- NEW: Helper for Combobox ---
    def _add_combobox(self, frame, field_key, label_text, options, row, col, **kwargs):
        """Auxiliar para criar rótulo e combobox."""
        combo_state = kwargs.pop('state', 'readonly') # Default readonly
        ttk.Label(frame, text=label_text + ":").grid(row=row, column=col*2, sticky=tk.W, padx=5, pady=2)
        combo = ttk.Combobox(frame, values=options, state=combo_state, **kwargs)
        combo.grid(row=row, column=col*2 + 1, sticky=tk.EW, padx=5, pady=2)

        if field_key in self.item_data and pd.notna(self.item_data[field_key]):
            # Try to set the current value, handle if it's not in options
            try:
                combo.set(str(self.item_data[field_key]))
            except tk.TclError:
                # Value not in list, set to first option or leave blank
                if options:
                    combo.current(0)
                print(f"Aviso: Valor '{self.item_data[field_key]}' para '{field_key}' não encontrado nas opções do combobox.")

        self.comboboxes[field_key] = combo
        frame.columnconfigure(col*2 + 1, weight=1)

    def _validate_and_collect(self):
        collected_data = {}
        for key, entry in self.entries.items():
             collected_data[key] = entry.get().strip()
        # --- NEW: Collect from comboboxes ---
        for key, combo in self.comboboxes.items():
             collected_data[key] = combo.get().strip()
        return collected_data

    def _save(self):
        try:
            self.updated_data = self._validate_and_collect()
            if self.updated_data:
                self.destroy()
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e), parent=self)


class EditProductDialog(EditDialogBase):
    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        classificacao_options = ["ATIVO", "CONSUMIVEL", "PREVENTIVO"]

        self._add_entry(main_frame, "CODIGO", "Código", 0, 0, state='readonly')
        self._add_entry(main_frame, "DESCRICAO", "Descrição", 1, 0)
        # --- MODIFIED: Added CLASSIFICACAO ---
        self._add_combobox(main_frame, "CLASSIFICACAO", "Classificação", classificacao_options, 2, 0)
        self._add_entry(main_frame, "VALOR UN", "Valor Unitário", 3, 0)
        self._add_entry(main_frame, "QUANTIDADE", "Quantidade", 4, 0)
        self._add_entry(main_frame, "LOCALIZACAO", "Localização", 5, 0)
        # --- MODIFIED: Added NF/PEDIDO (read-only for now) ---
        # Editing the list here is complex, just display for now
        self._add_entry(main_frame, "NF/PEDIDO", "NF/Pedidos (Lista)", 6, 0)


    def _validate_and_collect(self):
        data = super()._validate_and_collect()

        # --- MODIFIED: Validation ---
        if not data["DESCRICAO"]:
            raise ValueError("Descrição não pode ser vazia.")
        if not data["CLASSIFICACAO"]: # Check if classification was selected
            raise ValueError("Classificação deve ser selecionada.")

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

        data["VALOR TOTAL"] = data["VALOR UN"] * data["QUANTIDADE"]
        # NF/PEDIDO is read-only here, no validation needed for it in this dialog

        return data
    
class EditEntradaDialog(EditDialogBase):
    """Diálogo para adicionar NF/Pedido a um item existente, usando um registro de entrada como referência."""

    def __init__(self, parent, title, item_data):
        # Aumentar um pouco a altura para caber os campos e botões
        # Mantido o mesmo tamanho da base por enquanto
        super().__init__(parent, title, item_data)
        # Removido geometry daqui, a base já define um tamanho padrão.
        # Se precisar ajustar especificamente, faça aqui ou na base.

    def _create_widgets(self):
        # Cria um frame principal DENTRO do Toplevel (self)
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True) # Empacota este frame

        ttk.Label(main_frame, text="Detalhes da Entrada Original (Referência)", font="-weight bold").grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky="w")

        fields_to_display = [
            ("CODIGO", "Código"), ("DESCRICAO", "Descrição"),
            ("CLASSIFICACAO", "Classificação"), ("QUANTIDADE", "Qtd. Entrada"),
            ("VALOR UN", "Valor Unit."), ("DATA", "Data Registro"),
            ("NF/PEDIDO", "NF/Ped. Original"), ("DATA EMISSAO", "Data Emissão Orig.")
        ]
        row_num = 1
        for key, label in fields_to_display:
             if key in self.item_data:
                # Passa main_frame como o parent para _add_entry
                self._add_entry(main_frame, key, label, row_num, 0, state='readonly', width=40)
                row_num += 1

        ttk.Separator(main_frame, orient=tk.HORIZONTAL).grid(row=row_num, column=0, columnspan=2, sticky="ew", pady=10)
        row_num += 1

        ttk.Label(main_frame, text="Adicionar Nova NF/Pedido ao Estoque", font="-weight bold").grid(row=row_num, column=0, columnspan=2, pady=(5, 5), sticky="w")
        row_num += 1
        # Passa main_frame como o parent para _add_entry
        self._add_entry(main_frame, "NF_PEDIDO_ADICIONAL", "Nova NF/Pedido", row_num, 0, width=40)

        if "NF_PEDIDO_ADICIONAL" in self.entries:
             self.entries["NF_PEDIDO_ADICIONAL"].focus_set()
             # Adiciona o bind de <Return> aqui também para conveniência
             self.entries["NF_PEDIDO_ADICIONAL"].bind("<Return>", lambda e: self._save())


    def _validate_and_collect(self):
        """Coleta apenas a nova NF/Pedido."""
        collected_data = super()._validate_and_collect()
        nf_adicional = collected_data.get("NF_PEDIDO_ADICIONAL", "").strip().upper()
        if not nf_adicional:
            raise ValueError("Digite o número da Nova NF/Pedido a ser adicionada ao histórico do item.")
        return {"NF_PEDIDO_ADICIONAL": nf_adicional}


    def _save(self):
        # Usa a validação personalizada para coletar apenas a NF adicional
        try:
            self.updated_data = self._validate_and_collect()
            if self.updated_data: # Validação passou se retornou a NF
                self.destroy()
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e), parent=self)


class EditEPIDialog(EditDialogBase):
     def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        # No changes needed here based on requirements
        self._add_entry(main_frame, "CA", "CA", 0, 0)
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
        data["CA"] = data["CA"].upper()
        data["DESCRICAO"] = data["DESCRICAO"].upper()
        return data

# --- Classe Principal da Aplicação ---

class AlmoxarifadoApp:
    def __init__(self, root, user_id):
        self.root = root
        self.operador_logado_id = user_id
        self.active_table_name = "estoque"
        self.current_table_df = None

        self.root.title(f"Almoxarifado - Operador: {self.operador_logado_id}")
        self.root.geometry("1150x650")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._criar_pastas_e_planilhas()
        self._setup_ui()
        self._criar_backup_periodico()
        self._load_and_display_table(self.active_table_name)
        self.root.after(10800000, self._schedule_backup) # 3 hours

    def _setup_ui(self):
            style = ttk.Style()
            style.theme_use('clam')
            # --- Definição de Estilos Coloridos ---
            style.configure("Accent.TButton", foreground="white", background="#0078D7")
            style.map("Accent.TButton", background=[('active', '#005A9E')])
            style.configure("Success.TButton", foreground="white", background="#107C10")
            style.map("Success.TButton", background=[('active', '#0A530A')])
            style.configure("Edit.TButton", foreground="white", background="#FFB900")
            style.map("Edit.TButton", background=[('active', '#D89D00')])
            style.configure("Delete.TButton", foreground="white", background="#D83B01")
            style.map("Delete.TButton", background=[('active', '#A42E00')])
            style.configure("Secondary.TButton", foreground="black", background="#CCCCCC", padding=5)
            style.map("Secondary.TButton", background=[('active', '#B3B3B3')])
            # --- Fim da Definição de Estilos ---

            main_frame = ttk.Frame(self.root, padding="5")
            main_frame.pack(expand=True, fill="both")

            self.status_var = tk.StringVar()
            self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
            # self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # Pack later

            self.notebook = ttk.Notebook(main_frame)
            self.notebook.pack(expand=True, fill="both", pady=(0, 5))

            self._create_estoque_tab()
            self._create_cadastro_tab() # Will be modified
            self._create_movimentacao_tab() # Will be modified
            self._create_epis_tab()

            self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
            self._update_status("Pronto.")

    def _create_estoque_tab(self):
        self.estoque_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.estoque_tab, text="Estoque & Tabelas")

        controls_frame = ttk.Frame(self.estoque_tab)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        search_frame = ttk.LabelFrame(controls_frame, text="Pesquisar na Tabela Atual", padding="5")
        search_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.pesquisar_entry = ttk.Entry(search_frame, width=40)
        self.pesquisar_entry.bind("<KeyRelease>", self._pesquisar_tabela_event)
        self.pesquisar_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(search_frame, text="Buscar", command=self._pesquisar_tabela, style="Success.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(search_frame, text="Limpar", command=self._limpar_pesquisa).pack(side=tk.LEFT)

        switch_frame = ttk.LabelFrame(controls_frame, text="Visualizar Tabela", padding="5")
        switch_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_view_estoque = ttk.Button(switch_frame, text="Estoque", command=lambda: self._trocar_tabela_view("estoque"))
        self.btn_view_estoque.pack(side=tk.LEFT, padx=2)
        self.btn_view_entrada = ttk.Button(switch_frame, text="Entradas", command=lambda: self._trocar_tabela_view("entrada"))
        self.btn_view_entrada.pack(side=tk.LEFT, padx=2)
        self.btn_view_saida = ttk.Button(switch_frame, text="Saídas", command=lambda: self._trocar_tabela_view("saida"))
        self.btn_view_saida.pack(side=tk.LEFT, padx=2)

        action_frame = ttk.Frame(controls_frame)
        action_frame.pack(side=tk.RIGHT)

        edit_delete_frame = ttk.LabelFrame(action_frame, text="Item Selecionado", padding="5")
        edit_delete_frame.pack(side=tk.LEFT, padx=(0,10))
        self.edit_button = ttk.Button(edit_delete_frame, text="Editar", command=self._edit_selected_item, state=tk.DISABLED, style="Edit.TButton")
        self.edit_button.pack(side=tk.LEFT, padx=2)
        self.delete_button = ttk.Button(edit_delete_frame, text="Excluir", command=self._delete_selected_item, state=tk.DISABLED, style="Delete.TButton")
        self.delete_button.pack(side=tk.LEFT, padx=2)

        general_action_frame = ttk.LabelFrame(action_frame, text="Ações", padding="5")
        general_action_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(general_action_frame, text="Atualizar", command=self._atualizar_tabela_atual, style="Success.TButton").pack(side=tk.LEFT, padx=2)

        export_frame = ttk.LabelFrame(action_frame, text="Relatórios", padding="5")
        export_frame.pack(side=tk.LEFT)
        ttk.Button(export_frame, text="Exportar", command=self._exportar_conteudo, style="Accent.TButton").pack(side=tk.LEFT, padx=2)

        self.pandas_table_frame = ttk.Frame(self.estoque_tab)
        self.pandas_table_frame.pack(expand=True, fill="both")
        self.pandas_table = Table(parent=self.pandas_table_frame, editable=False)
        self.pandas_table.show()
        self.pandas_table.bind("<<TableSelectChanged>>", self._on_table_select)

        self._atualizar_cores_botoes_view()

    # --- MODIFIED: Cadastro Tab UI ---
    def _create_cadastro_tab(self):
        self.cadastro_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.cadastro_tab, text="Cadastrar Produto")

        container = ttk.Frame(self.cadastro_tab)
        container.pack(anchor=tk.CENTER)

        ttk.Label(container, text="Cadastrar Novo Produto no Estoque", font="-weight bold -size 14").grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Updated fields and widgets
        self.cadastro_entries = {}
        self.cadastro_widgets_ordered = [] # To manage focus order

        # Descrição
        ttk.Label(container, text="Descrição:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        entry_desc = ttk.Entry(container, width=40, font="-size 12")
        entry_desc.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=5)
        self.cadastro_entries["DESCRICAO"] = entry_desc
        self.cadastro_widgets_ordered.append(entry_desc)

        # Classificação (Combobox)
        ttk.Label(container, text="Classificação:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        classificacao_options = ["", "ATIVO", "CONSUMIVEL", "PREVENTIVO"] # Add empty default
        self.cadastro_classificacao_combo = ttk.Combobox(container, values=classificacao_options, state="readonly", font="-size 12")
        self.cadastro_classificacao_combo.current(0) # Select empty by default
        self.cadastro_classificacao_combo.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5)
        # No direct Enter bind needed for Combobox typically
        self.cadastro_widgets_ordered.append(self.cadastro_classificacao_combo)

        # Quantidade
        ttk.Label(container, text="Quantidade:", font="-size 12").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        entry_qtd = ttk.Entry(container, width=40, font="-size 12")
        entry_qtd.grid(row=3, column=1, sticky=tk.EW, padx=5, pady=5)
        self.cadastro_entries["QUANTIDADE"] = entry_qtd
        self.cadastro_widgets_ordered.append(entry_qtd)

        # Valor Unitário
        ttk.Label(container, text="Valor Unitário:", font="-size 12").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        entry_val = ttk.Entry(container, width=40, font="-size 12")
        entry_val.grid(row=4, column=1, sticky=tk.EW, padx=5, pady=5)
        self.cadastro_entries["VALOR UN"] = entry_val
        self.cadastro_widgets_ordered.append(entry_val)

        # Localização
        ttk.Label(container, text="Localização:", font="-size 12").grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
        entry_loc = ttk.Entry(container, width=40, font="-size 12")
        entry_loc.grid(row=5, column=1, sticky=tk.EW, padx=5, pady=5)
        self.cadastro_entries["LOCALIZACAO"] = entry_loc
        self.cadastro_widgets_ordered.append(entry_loc)

        # NF/Pedido Inicial
        ttk.Label(container, text="NF/Pedido Inicial:", font="-size 12").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        entry_nf = ttk.Entry(container, width=40, font="-size 12")
        entry_nf.grid(row=6, column=1, sticky=tk.EW, padx=5, pady=5)
        self.cadastro_entries["NF/PEDIDO"] = entry_nf
        self.cadastro_widgets_ordered.append(entry_nf) # Add to focus order

        # Bind Enter key for focus traversal
        for i, widget in enumerate(self.cadastro_widgets_ordered):
            if i < len(self.cadastro_widgets_ordered) - 1:
                # Bind Enter to focus the *next* widget in the list
                widget.bind("<Return>", lambda e, next_widget=self.cadastro_widgets_ordered[i+1]: next_widget.focus_set())
            else:
                # Bind Enter on the last widget to trigger the cadastro action
                widget.bind("<Return>", lambda e: self._cadastrar_estoque())

        cadastrar_button = ttk.Button(container, text="Cadastrar Produto", command=self._cadastrar_estoque, style="Success.TButton")
        cadastrar_button.grid(row=7, column=0, columnspan=2, pady=(20, 0), sticky=tk.EW)

        container.columnconfigure(1, weight=1)

    # --- MODIFIED: Movimentacao Tab UI (Entrada Frame) ---
    def _create_movimentacao_tab(self):
        self.movimentacao_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.movimentacao_tab, text="Movimentação Estoque")

        self.movimentacao_tab.columnconfigure(0, weight=1)
        self.movimentacao_tab.columnconfigure(1, minsize=20) # Separator space
        self.movimentacao_tab.columnconfigure(2, weight=1)
        self.movimentacao_tab.rowconfigure(0, weight=1)

        # --- Frame de Entrada (Modified) ---
        entrada_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Entrada", padding="15")
        entrada_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        entrada_frame.columnconfigure(1, weight=1)

        self.entrada_widgets_ordered = [] # For focus

        # Código
        ttk.Label(entrada_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_codigo_entry = ttk.Entry(entrada_frame, font="-size 12")
        self.entrada_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        self.entrada_widgets_ordered.append(self.entrada_codigo_entry)
        ttk.Button(entrada_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("entrada"), style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5))

        # NF/Pedido (New)
        ttk.Label(entrada_frame, text="NF/Pedido:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_nf_pedido_entry = ttk.Entry(entrada_frame, font="-size 12")
        self.entrada_nf_pedido_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.entrada_widgets_ordered.append(self.entrada_nf_pedido_entry)

        # Data Emissão (New) - Using simple Entry + Validation for HH:MM DD/MM/YY
        ttk.Label(entrada_frame, text="Data Emissão:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_data_emissao_entry = ttk.Entry(entrada_frame, font="-size 12")
        self.entrada_data_emissao_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5)
        # Add placeholder/tooltip if possible, or just rely on label
        self.entrada_data_emissao_entry.insert(0, "HH:MM DD/MM/YY") # Placeholder text
        self.entrada_data_emissao_entry.bind("<FocusIn>", self._clear_placeholder)
        self.entrada_data_emissao_entry.bind("<FocusOut>", self._add_placeholder)
        self.entrada_widgets_ordered.append(self.entrada_data_emissao_entry)

        # Qtd. Entrada
        ttk.Label(entrada_frame, text="Qtd. Entrada:", font="-size 12").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_qtd_entry = ttk.Entry(entrada_frame, font="-size 12")
        self.entrada_qtd_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.entrada_widgets_ordered.append(self.entrada_qtd_entry)

        # Bind Enter for focus in Entrada tab
        for i, widget in enumerate(self.entrada_widgets_ordered):
            if i < len(self.entrada_widgets_ordered) - 1:
                widget.bind("<Return>", lambda e, next_widget=self.entrada_widgets_ordered[i+1]: self._focus_widget(next_widget)) # Use helper
            else:
                widget.bind("<Return>", lambda e: self._registrar_entrada())

        entrada_button = ttk.Button(entrada_frame, text="Registrar Entrada", command=self._registrar_entrada, style="Success.TButton")
        entrada_button.grid(row=4, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)


        # --- Separador ---
        sep = ttk.Separator(self.movimentacao_tab, orient=tk.VERTICAL)
        sep.grid(row=0, column=1, sticky="ns", padx=5, pady=5)

        # --- Frame de Saída (Unchanged for now) ---
        saida_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Saída", padding="15")
        saida_frame.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        saida_frame.columnconfigure(1, weight=1) # Faz as entradas expandirem

        self.saida_widgets_ordered = [] # Ordem para tecla Enter/Tab

        # Código
        ttk.Label(saida_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_codigo_entry = ttk.Entry(saida_frame, font="-size 12")
        self.saida_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        # ** Bind para popular dropdown quando o código mudar (FocusOut ou Return) **
        self.saida_codigo_entry.bind("<FocusOut>", self._update_saida_nf_dropdown_event)
        self.saida_codigo_entry.bind("<Return>", self._update_saida_nf_dropdown_event) # Também no Enter
        self.saida_widgets_ordered.append(self.saida_codigo_entry)
        ttk.Button(saida_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("saida"), style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5))

        # Solicitante
        ttk.Label(saida_frame, text="Solicitante:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_solicitante_entry = ttk.Entry(saida_frame, font="-size 12")
        self.saida_solicitante_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_widgets_ordered.append(self.saida_solicitante_entry)

        # Setor (Novo)
        ttk.Label(saida_frame, text="Setor:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_setor_entry = ttk.Entry(saida_frame, font="-size 12")
        self.saida_setor_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_widgets_ordered.append(self.saida_setor_entry)

        # Qtd. Saída
        ttk.Label(saida_frame, text="Qtd. Saída:", font="-size 12").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_qtd_entry = ttk.Entry(saida_frame, font="-size 12")
        self.saida_qtd_entry.grid(row=3, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_widgets_ordered.append(self.saida_qtd_entry)

        # NF/Pedido (Novo - Dropdown)
        ttk.Label(saida_frame, text="NF/Ped. (Origem):", font="-size 12").grid(row=4, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_nf_pedido_combo = ttk.Combobox(saida_frame, state="disabled", font="-size 12") # Começa desabilitado
        self.saida_nf_pedido_combo['values'] = ["(Informe o Código)"]
        self.saida_nf_pedido_combo.current(0)
        self.saida_nf_pedido_combo.grid(row=4, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_widgets_ordered.append(self.saida_nf_pedido_combo) # Adicionado à ordem de foco

        # Bind Enter for focus in Saida tab (Ajustado)
        for i, widget in enumerate(self.saida_widgets_ordered):
            # Se não for o último widget, foca o próximo
            if i < len(self.saida_widgets_ordered) - 1:
                 next_widget = self.saida_widgets_ordered[i+1]
                 # Tratamento especial para o campo código - não avançar foco no Return se chamar dropdown
                 if widget == self.saida_codigo_entry:
                     # O binding de <Return> no código já chama a atualização do dropdown,
                     # não queremos avançar o foco automaticamente *nele*, deixamos
                     # o usuário Tabular ou clicar para o próximo campo após ver as NFs.
                     # Ou podemos chamar `_focus_widget(next_widget)` dentro de _update_saida_nf_dropdown_event
                     # Vamos tentar não focar automaticamente no Return do código por enquanto.
                     pass
                 else:
                     widget.bind("<Return>", lambda e, nw=next_widget: self._focus_widget(nw))

            # Se for o último widget (o Combobox), liga o Enter à função de registrar
            else:
                widget.bind("<Return>", lambda e: self._registrar_saida())
                # Permitir também acionar com Enter no penúltimo (Qtd Saída) se NF for opcional
                self.saida_qtd_entry.bind("<Return>", lambda e: self._registrar_saida()) # Duplica no QTD caso o user pule NF

        # Botão Registrar Saída (Row ajustado)
        saida_button = ttk.Button(saida_frame, text="Registrar Saída", command=self._registrar_saida, style="Success.TButton")
        saida_button.grid(row=5, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)

    # --- Handler de Evento para chamar o preenchimento do dropdown ---
    def _update_saida_nf_dropdown_event(self, event=None):
        """Chamado pelo FocusOut ou Return do campo Código da Saída."""
        codigo = self.saida_codigo_entry.get().strip()
        self._populate_saida_nf_dropdown(codigo)
        # Decidir se avança o foco aqui
        # self._focus_widget(self.saida_solicitante_entry) # Avança para solicitante?
        return "break" # Impede processamento adicional do Enter se chamado por <Return>

    # --- Nova Função para Popular o Dropdown de NF/Pedido da Saída ---
    def _populate_saida_nf_dropdown(self, codigo):
        """Busca NFs/Pedidos do produto e atualiza o combobox de Saída."""
        if not codigo:
            self.saida_nf_pedido_combo.config(state="disabled")
            self.saida_nf_pedido_combo['values'] = ["(Informe o Código)"]
            self.saida_nf_pedido_combo.current(0)
            return

        produto_atual = self._buscar_produto(codigo)

        if produto_atual is None or "NF/PEDIDO" not in produto_atual:
            self.saida_nf_pedido_combo.config(state="disabled")
            self.saida_nf_pedido_combo['values'] = ["(Produto s/ NFs)"]
            self.saida_nf_pedido_combo.current(0)
            return

        nf_pedido_string = produto_atual["NF/PEDIDO"] # Já vem parseado se _buscar_produto for usado corretamente

        # Mas _buscar_produto retorna a série original, precisamos parsear aqui ou modificar _buscar_produto
        # Vamos parsear aqui por segurança, usando nosso helper
        nf_list = self._parse_nf_pedido_list(str(nf_pedido_string)) # Converte pra string antes de parsear

        if nf_list:
             # Adiciona opção vazia no início para permitir saída sem atrelar NF
            display_values = [""] + nf_list
            self.saida_nf_pedido_combo['values'] = display_values
            self.saida_nf_pedido_combo.config(state="readonly") # Permite selecionar
            self.saida_nf_pedido_combo.current(0) # Seleciona a opção vazia por padrão
        else:
            self.saida_nf_pedido_combo['values'] = ["SEM NF"]
            self.saida_nf_pedido_combo.config(state="disabled")
            self.saida_nf_pedido_combo.current(0)

    def _create_epis_tab(self):
        # This tab remains unchanged based on current requirements
        self.epis_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.epis_tab, text="EPIs")
        self.epis_tab.columnconfigure(0, weight=2)
        self.epis_tab.columnconfigure(1, weight=1)
        self.epis_tab.rowconfigure(0, weight=1)

        epis_table_frame = ttk.Frame(self.epis_tab)
        epis_table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        epi_table_controls = ttk.Frame(epis_table_frame)
        epi_table_controls.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(epi_table_controls, text="Atualizar Lista", command=self._atualizar_tabela_epis, style="Success.TButton").pack(side=tk.LEFT, padx=(0, 10))
        self.edit_epi_button = ttk.Button(epi_table_controls, text="Editar EPI Sel.", command=self._edit_selected_epi, state=tk.DISABLED, style="Accent.TButton")
        self.edit_epi_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_epi_button = ttk.Button(epi_table_controls, text="Excluir EPI Sel.", command=self._delete_selected_epi, state=tk.DISABLED, style="Accent.TButton")
        self.delete_epi_button.pack(side=tk.LEFT, padx=(0, 5))

        self.epis_pandas_frame = ttk.Frame(epis_table_frame)
        self.epis_pandas_frame.pack(expand=True, fill="both")
        self.epis_table = Table(parent=self.epis_pandas_frame, editable=False)
        self.epis_table.show()
        self._carregar_epis()
        self.epis_table.bind("<<TableSelectChanged>>", self._on_epi_table_select)

        forms_frame = ttk.Frame(self.epis_tab)
        forms_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        forms_frame.rowconfigure(1, weight=1)

        registrar_frame = ttk.LabelFrame(forms_frame, text="Registrar / Adicionar EPI", padding="10")
        registrar_frame.pack(fill=tk.X, pady=(0, 15))
        registrar_frame.columnconfigure(1, weight=1)
        ttk.Label(registrar_frame, text="CA:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_ca_entry = ttk.Entry(registrar_frame, font="-size 12")
        self.epi_ca_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_ca_entry.bind("<Return>", self._focar_proximo_simples)
        ttk.Label(registrar_frame, text="Descrição:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_desc_entry = ttk.Entry(registrar_frame, font="-size 12")
        self.epi_desc_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_desc_entry.bind("<Return>", self._focar_proximo_simples)
        ttk.Label(registrar_frame, text="Quantidade:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_qtd_entry = ttk.Entry(registrar_frame, font="-size 12")
        self.epi_qtd_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_qtd_entry.bind("<Return>", lambda e: self._registrar_epi())
        registrar_epi_button = ttk.Button(registrar_frame, text="Registrar / Adicionar", command=self._registrar_epi, style="Success.TButton")
        registrar_epi_button.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky=tk.EW)

        retirar_frame = ttk.LabelFrame(forms_frame, text="Registrar Retirada de EPI", padding="10")
        retirar_frame.pack(fill=tk.BOTH, expand=True)
        retirar_frame.columnconfigure(1, weight=1)
        ttk.Label(retirar_frame, text="CA/Descrição:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_id_entry = ttk.Entry(retirar_frame, font="-size 12")
        self.retirar_epi_id_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2), pady=5)
        self.retirar_epi_id_entry.bind("<Return>", self._focar_proximo_simples)
        ttk.Button(retirar_frame, text="Buscar", width=8, command=self._show_epi_lookup, style="Secondary.TButton").grid(row=0, column=2, sticky=tk.W, padx=(2,5), pady=5)
        ttk.Label(retirar_frame, text="Qtd. Retirada:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_qtd_entry = ttk.Entry(retirar_frame, font="-size 12")
        self.retirar_epi_qtd_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_qtd_entry.bind("<Return>", self._focar_proximo_simples)
        ttk.Label(retirar_frame, text="Colaborador:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_colab_entry = ttk.Entry(retirar_frame, font="-size 12")
        self.retirar_epi_colab_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_colab_entry.bind("<Return>", lambda e: self._registrar_retirada())
        retirar_epi_button = ttk.Button(retirar_frame, text="Registrar Retirada", command=self._registrar_retirada, style="Success.TButton")
        retirar_epi_button.grid(row=3, column=0, columnspan=3, pady=(10, 5), sticky=tk.EW)


    # --- Manipuladores de Eventos da UI & Auxiliares ---
    def _update_status(self, message, error=False):
        self.status_var.set(message)
        self.status_bar.config(foreground="red" if error else "black")
        # print(message) # Debug

    # --- NEW: Placeholder handlers for Data Emissao ---
    def _clear_placeholder(self, event):
        """Clears placeholder text on focus."""
        if event.widget.get() == "HH:MM DD/MM/YY":
            event.widget.delete(0, tk.END)
            event.widget.config(foreground='black') # Or your default text color

    def _add_placeholder(self, event):
        """Adds placeholder text if entry is empty on focus out."""
        if not event.widget.get():
            event.widget.insert(0, "HH:MM DD/MM/YY")
            event.widget.config(foreground='grey')

    # --- NEW: Simplified focus helper ---
    def _focus_widget(self, widget_to_focus):
        """Safely focuses the next widget."""
        try:
            widget_to_focus.focus_set()
            # If it's the data emissao entry, clear placeholder
            if widget_to_focus == self.entrada_data_emissao_entry and widget_to_focus.get() == "HH:MM DD/MM/YY":
                self._clear_placeholder(tk.Event()) # Simulate event
        except Exception as e:
            print(f"Error focusing widget: {e}")
        return "break"

    def _focar_proximo_simples(self, event):
        """Generic focus next (replace old _focar_proximo)."""
        try:
            event.widget.tk_focusNext().focus()
        except Exception:
            pass
        return "break"

    # --- OBSOLETE: Remove or repurpose _focar_proximo_cadastro ---
    # def _focar_proximo_cadastro(self, event):
        # ... (This logic is now handled by the loop in _create_cadastro_tab)
        # return "break"

    def _on_table_select(self, event=None):
        """Habilita/desabilita botões de editar/excluir com base na seleção da tabela."""
        selected = self.pandas_table.getSelectedRowData()
        # Assume single selection
        row_is_selected = selected is not None and len(selected) == 1
        is_estoque = self.active_table_name == "estoque"
        # --- MODIFICADO: Habilita Editar também para Entrada ---
        is_entrada = self.active_table_name == "entrada"

        # Habilita Editar para Estoque ou Entrada
        if row_is_selected and (is_estoque or is_entrada):
             self.edit_button.config(state=tk.NORMAL)
        else:
            self.edit_button.config(state=tk.DISABLED)

        # Habilita Excluir APENAS para Estoque
        if row_is_selected and is_estoque:
             self.delete_button.config(state=tk.NORMAL)
        else:
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
        self._pesquisar_tabela()

    def _pesquisar_tabela(self):
        if self.current_table_df is None or self.pandas_table is None:
            return
        query = self.pesquisar_entry.get().strip().lower()

        if query:
            try:
                # Create mask for string columns
                string_cols = self.current_table_df.select_dtypes(include='object').columns
                mask_str = self.current_table_df[string_cols].apply(lambda col: col.astype(str).str.lower().str.contains(query, na=False)).any(axis=1)

                # Create mask for non-string columns (convert to string first)
                other_cols = self.current_table_df.select_dtypes(exclude='object').columns
                mask_others = self.current_table_df[other_cols].astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)

                df_filtered = self.current_table_df[mask_str | mask_others]

            except Exception as e:
                self._update_status(f"Erro na pesquisa: {e}", error=True)
                df_filtered = self.current_table_df
        else:
            df_filtered = self.current_table_df

        try:
            # Check if model exists before updating
            if hasattr(self.pandas_table, 'model'):
                self.pandas_table.updateModel(TableModel(df_filtered))
                self.pandas_table.redraw()
                status_msg = f"Exibindo {len(df_filtered)} resultados para '{query}'" if query else f"Exibindo todos os {len(df_filtered)} registros."
                self._update_status(status_msg)
                self._on_table_select() # Update button states after filter/clear
            else:
                self._update_status("Erro: Modelo da tabela não encontrado para atualizar pesquisa.", error=True)
        except Exception as e:
            self._update_status(f"Erro ao atualizar tabela pós-pesquisa: {e}", error=True)


    def _limpar_pesquisa(self):
        """Limpa a entrada de pesquisa e redefine a visualização da tabela."""
        if self.pandas_table is None or self.current_table_df is None:
            return
        self.pesquisar_entry.delete(0, tk.END)
        try:
             # Check if model exists
            if hasattr(self.pandas_table, 'model'):
                self.pandas_table.updateModel(TableModel(self.current_table_df))
                self.pandas_table.redraw()
                self._update_status(f"Exibindo todos os {len(self.current_table_df)} registros.")
                self._on_table_select() # Redefine os estados dos botões
            else:
                 self._update_status("Erro: Modelo da tabela não encontrado para limpar pesquisa.", error=True)

        except Exception as e:
             self._update_status(f"Erro ao limpar pesquisa: {e}", error=True)

    def _atualizar_cores_botoes_view(self):
        buttons = {
            "estoque": self.btn_view_estoque,
            "entrada": self.btn_view_entrada,
            "saida": self.btn_view_saida
        }
        for name, button in buttons.items():
             if hasattr(self, f"btn_view_{name}"):
                 target_button = getattr(self, f"btn_view_{name}")
                 style = "Accent.TButton" if name == self.active_table_name else "TButton"
                 try:
                     target_button.config(style=style)
                 except tk.TclError: # Handle case where style might not exist yet
                     print(f"Aviso: Estilo '{style}' não encontrado ao atualizar cores dos botões.")
                     target_button.config(style="TButton") # Fallback


    # --- Carregamento de Dados e Operações de Arquivo ---

    # --- MODIFIED: Define new columns ---
    def _criar_pastas_e_planilhas(self):
        """Cria diretórios e arquivos CSV necessários se não existirem."""
        os.makedirs(PLANILHAS_DIR, exist_ok=True)
        os.makedirs(BACKUP_DIR, exist_ok=True)
        os.makedirs(COLABORADORES_DIR, exist_ok=True)

        colunas = {
            # ADDED: CLASSIFICACAO, NF/PEDIDO
            "estoque": ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"],
            # ADDED: NF/PEDIDO, DATA EMISSAO, CLASSIFICACAO
            "entrada": ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"],
            "saida": ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA", "ID"],
            "epis": ["CA", "DESCRICAO", "QUANTIDADE"]
        }

        for nome, arquivo in ARQUIVOS.items():
            if not os.path.exists(arquivo):
                try:
                    df = pd.DataFrame(columns=colunas.get(nome, []))
                    # Ensure NF/PEDIDO in estoque is object type if created empty
                    if nome == "estoque" and "NF/PEDIDO" in df.columns:
                         df["NF/PEDIDO"] = df["NF/PEDIDO"].astype(object)
                    df.to_csv(arquivo, index=False, encoding="utf-8")
                    print(f"Arquivo criado: {arquivo}")
                except Exception as e:
                    messagebox.showerror("Erro ao Criar Arquivo", f"Não foi possível criar {arquivo}: {e}")

    def _safe_read_csv(self, file_path):
        """Lê um arquivo CSV com segurança, retornando um DataFrame vazio em caso de erro."""
        try:
            # Define dtypes explicitly, esp for IDs and potentially list-like strings
            dtype_map = {'CODIGO': str, 'CA': str, 'ID': str, 'SOLICITANTE': str}
            # For NF/PEDIDO in estoque, read as string initially, parse later
            if "Estoque.csv" in file_path:
                dtype_map['NF/PEDIDO'] = str # Read as string
                dtype_map['CLASSIFICACAO'] = str
            elif "Entrada.csv" in file_path:
                dtype_map['NF/PEDIDO'] = str
                dtype_map['DATA EMISSAO'] = str
                dtype_map['CLASSIFICACAO'] = str

            # Attempt to read with specified dtypes
            df = pd.read_csv(file_path, encoding="utf-8", dtype=dtype_map)

            # Post-read checks for specific columns if needed
            # Example: Ensure NF/PEDIDO in estoque is filled with empty string if NaN
            if "Estoque.csv" in file_path and 'NF/PEDIDO' in df.columns:
                df['NF/PEDIDO'] = df['NF/PEDIDO'].fillna('')

            return df

        except FileNotFoundError:
            self._update_status(f"Aviso: Arquivo não encontrado: {file_path}. Verifique as pastas.", error=True)
            # Create DataFrame with expected columns if file not found
            if "Estoque.csv" in file_path:
                cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"]
                return pd.DataFrame(columns=cols).astype({'NF/PEDIDO': object, 'CODIGO': str, 'CLASSIFICACAO':str})
            elif "Entrada.csv" in file_path:
                 cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"]
                 return pd.DataFrame(columns=cols).astype({'CODIGO':str, 'ID': str, 'NF/PEDIDO': str, 'DATA EMISSAO': str, 'CLASSIFICACAO': str})
            # Add similar handlers for other files if necessary
            else:
                return pd.DataFrame()
        except pd.errors.EmptyDataError:
             self._update_status(f"Aviso: Arquivo vazio: {file_path}.", error=False)
             # Return empty DF with correct columns like above
             # This duplicates the FileNotFoundError logic, could be refactored
             if "Estoque.csv" in file_path:
                 cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"]
                 return pd.DataFrame(columns=cols).astype({'NF/PEDIDO': object, 'CODIGO': str, 'CLASSIFICACAO':str})
             elif "Entrada.csv" in file_path:
                 cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"]
                 return pd.DataFrame(columns=cols).astype({'CODIGO':str, 'ID': str, 'NF/PEDIDO': str, 'DATA EMISSAO': str, 'CLASSIFICACAO': str})
             else:
                 return pd.DataFrame()

        except Exception as e:
            self._update_status(f"Erro ao ler {file_path}: {e}", error=True)
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo {os.path.basename(file_path)}.\nVerifique se ele não está corrompido ou aberto em outro programa.\n\nDetalhes: {e}")
            return pd.DataFrame() # Return empty df on other errors

    def _safe_write_csv(self, df, file_path, create_backup=True):
        """Escreve um DataFrame em CSV com segurança, opcionalmente criando um backup."""
        backup_path = None
        try:
            if create_backup and os.path.exists(file_path):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                backup_file_name = os.path.basename(file_path).replace(".csv", f"_backup_{timestamp}.csv")
                backup_path = os.path.join(BACKUP_DIR, backup_file_name)
                shutil.copy2(file_path, backup_path)

            # Ensure consistent column order based on initial definition
            target_columns = []
            if "Estoque.csv" in file_path:
                 target_columns = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"]
            elif "Entrada.csv" in file_path:
                 target_columns = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"]
            elif "Saida.csv" in file_path:
                 target_columns = ["CODIGO", "DESCRICAO", "QUANTIDADE", "SOLICITANTE", "DATA", "ID"]
            elif "Epis.csv" in file_path:
                 target_columns = ["CA", "DESCRICAO", "QUANTIDADE"]
            # Add collaborator file columns if needed

            # Reorder and select only target columns if defined
            df_to_write = df
            if target_columns:
                 # Ensure all target columns exist, add missing ones with default value (e.g., NaN or '')
                 for col in target_columns:
                      if col not in df_to_write.columns:
                           df_to_write[col] = '' # Or pd.NA
                 df_to_write = df_to_write[target_columns] # Select and reorder


            df_to_write.to_csv(file_path, index=False, encoding="utf-8", quoting=csv.QUOTE_MINIMAL) # Use minimal quoting
            return True
        except Exception as e:
            self._update_status(f"Erro Crítico ao salvar {os.path.basename(file_path)}: {e}", error=True)
            backup_msg = f"Um backup pode ter sido criado: {os.path.basename(backup_path)}" if backup_path else "Não foi possível criar backup."
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar as alterações em {os.path.basename(file_path)}.\n\nDetalhes: {e}\n\n{backup_msg}")
            return False


    def _load_and_display_table(self, table_name):
        """Carrega dados do CSV e atualiza a pandastable principal."""
        file_path = ARQUIVOS.get(table_name)
        if not file_path:
            messagebox.showerror("Erro Interno", f"Nome de tabela inválido: {table_name}")
            return

        self._update_status(f"Carregando tabela '{table_name.capitalize()}'...")
        df = self._safe_read_csv(file_path) # Use safe read

        # Data type conversions and sanity checks
        # Ensure numeric columns are numeric, fill NaN with 0
        numeric_cols = ["VALOR UN", "VALOR TOTAL", "QUANTIDADE"]
        for col in numeric_cols:
             if col in df.columns:
                 df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # Ensure key string/identifier columns are strings
        string_like_cols = ["DESCRICAO", "LOCALIZACAO", "SOLICITANTE", "ID", "CODIGO", "CA", "CLASSIFICACAO", "NF/PEDIDO", "DATA EMISSAO"]
        for col in string_like_cols:
             if col in df.columns:
                 # Convert all to string, fill potential NaN/None before conversion
                 df[col] = df[col].fillna('').astype(str)


        # Specific handling for Estoque NF/PEDIDO (keep as string for display)
        if table_name == "estoque" and "NF/PEDIDO" in df.columns:
             # Already handled by string conversion above, just ensuring it's present
             pass

        self.current_table_df = df
        self.active_table_name = table_name

        try:
            if hasattr(self, 'pandas_table') and self.pandas_table:
                # Check if model exists before updating
                if hasattr(self.pandas_table, 'model'):
                    self.pandas_table.updateModel(TableModel(self.current_table_df))
                    self.pandas_table.redraw()
                    self._update_status(f"Tabela '{table_name.capitalize()}' carregada ({len(self.current_table_df)} registros).")
                else:
                     # If model doesn't exist (e.g., first load), show table
                     self.pandas_table.model = TableModel(self.current_table_df)
                     self.pandas_table.show()
                     self.pandas_table.redraw()
                     self._update_status(f"Tabela '{table_name.capitalize()}' carregada ({len(self.current_table_df)} registros).")

                self._atualizar_cores_botoes_view()
                self._on_table_select()
            else:
                 self._update_status(f"Erro: Tabela principal (pandas_table) não inicializada.", error=True)

        except Exception as e:
            self._update_status(f"Erro ao exibir tabela '{table_name.capitalize()}': {e}", error=True)
            print(f"Detailed error displaying table: {e}") # More detailed log


    def _atualizar_tabela_atual(self):
        """Recarrega os dados para a visualização da tabela ativa atual."""
        if self.active_table_name:
            query = self.pesquisar_entry.get().strip()
            self._load_and_display_table(self.active_table_name)
            if query:
                 self.pesquisar_entry.delete(0, tk.END)
                 self.pesquisar_entry.insert(0, query)
                 self._pesquisar_tabela()

    def _trocar_tabela_view(self, nome_tabela):
        if nome_tabela in ARQUIVOS:
             self.pesquisar_entry.delete(0, tk.END)
             self._load_and_display_table(nome_tabela)
        else:
            messagebox.showerror("Erro", f"Configuração de tabela '{nome_tabela}' não encontrada.")


    def _carregar_epis(self):
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         if not df_epis.empty:
             df_epis["CA"] = df_epis["CA"].astype(str).fillna("")
             df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("")
             df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

         try:
             if hasattr(self, 'epis_table') and self.epis_table:
                  # Handle first load vs update
                  if hasattr(self.epis_table, 'model'):
                     self.epis_table.updateModel(TableModel(df_epis))
                  else:
                     self.epis_table.model = TableModel(df_epis)
                     self.epis_table.show()

                  self.epis_table.redraw()
                  self._update_status(f"Lista de EPIs atualizada ({len(df_epis)} itens).")
                  self._on_epi_table_select()
             else:
                 self._update_status(f"Erro: Tabela de EPIs (epis_table) não inicializada.", error=True)

         except Exception as e:
              self._update_status(f"Erro ao exibir EPIs: {e}", error=True)

    def _atualizar_tabela_epis(self):
         self._carregar_epis()

    def _buscar_produto(self, codigo):
        """Busca um produto no estoque pelo código. Retorna uma Series ou None."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        if df_estoque.empty or 'CODIGO' not in df_estoque.columns:
            return None

        # Ensure correct types before comparison
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
        produto = df_estoque[df_estoque['CODIGO'] == str(codigo)]
        if not produto.empty:
            return produto.iloc[0].copy() # Return a copy to avoid SettingWithCopyWarning
        return None


    # --- NEW HELPER: Add NF/Pedido to Estoque Item ---
    def _adicionar_nf_pedido_estoque(self, codigo, nf_pedido_adicionar):
        """Adds a new NF/Pedido number to the list for a specific product in Estoque.csv."""
        if not nf_pedido_adicionar: # Don't add empty strings
            return True # Return True as technically no update failed

        arquivo_estoque = ARQUIVOS["estoque"]
        df_estoque = self._safe_read_csv(arquivo_estoque)
        if df_estoque.empty or 'CODIGO' not in df_estoque.columns or 'NF/PEDIDO' not in df_estoque.columns:
            self._update_status("Erro: Estoque vazio ou colunas ausentes para adicionar NF/Pedido.", error=True)
            return False

        codigo_str = str(codigo)
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
        df_estoque['NF/PEDIDO'] = df_estoque['NF/PEDIDO'].fillna('').astype(str) # Ensure string

        idx = df_estoque.index[df_estoque['CODIGO'] == codigo_str].tolist()

        if not idx:
             self._update_status(f"Erro interno: Produto {codigo_str} não encontrado no estoque para adicionar NF/Pedido.", error=True)
             return False

        idx = idx[0] # Use the first index

        try:
            nf_pedido_atual_str = df_estoque.loc[idx, "NF/PEDIDO"]
            nf_pedido_list = []

            # Try to parse the existing string as a list
            if nf_pedido_atual_str and nf_pedido_atual_str.startswith('[') and nf_pedido_atual_str.endswith(']'):
                try:
                    nf_pedido_list = ast.literal_eval(nf_pedido_atual_str)
                    if not isinstance(nf_pedido_list, list):
                        nf_pedido_list = [nf_pedido_atual_str] # Treat as single item if parse fails weirdly
                except (ValueError, SyntaxError):
                    # If parsing fails, assume it's a single non-list entry or corrupted
                    print(f"Aviso: Falha ao parsear NF/PEDIDO '{nf_pedido_atual_str}' para o código {codigo_str}. Tratando como item único.")
                    nf_pedido_list = [nf_pedido_atual_str] # Keep the existing string as the first item
            elif nf_pedido_atual_str: # Handle case where it's just a single string, not list format
                 nf_pedido_list = [nf_pedido_atual_str]


            # Add the new number if it's not already present
            if nf_pedido_adicionar not in nf_pedido_list:
                nf_pedido_list.append(nf_pedido_adicionar)

            # Convert back to string representation for saving
            df_estoque.loc[idx, "NF/PEDIDO"] = str(nf_pedido_list)

            # Save the updated DataFrame (without backup within this specific update)
            if self._safe_write_csv(df_estoque, arquivo_estoque, create_backup=False): # Avoid redundant backups
                 return True
            else:
                 # Error saving is handled by _safe_write_csv
                 return False

        except Exception as e:
             self._update_status(f"Erro ao processar/adicionar NF/Pedido para {codigo_str}: {e}", error=True)
             messagebox.showerror("Erro NF/Pedido", f"Não foi possível adicionar a NF/Pedido '{nf_pedido_adicionar}' ao produto {codigo_str}.\n\nDetalhes: {e}")
             return False


    def _atualizar_estoque_produto(self, codigo, nova_quantidade):
        """Atualiza a quantidade e valor total de um produto no estoque (keeps NF/PEDIDO logic separate)."""
        arquivo_estoque = ARQUIVOS["estoque"]
        df_estoque = self._safe_read_csv(arquivo_estoque)
        if 'CODIGO' not in df_estoque.columns:
             self._update_status(f"Erro interno: Coluna 'CODIGO' não encontrada no estoque.", error=True)
             return False

        codigo_str = str(codigo)
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)

        idx = df_estoque.index[df_estoque['CODIGO'] == codigo_str].tolist()
        if not idx:
             self._update_status(f"Erro interno: Produto {codigo_str} não encontrado para atualizar.", error=True)
             return False

        idx = idx[0]

        try:
            # Check required columns exist
            required_cols = ["VALOR UN", "QUANTIDADE", "VALOR TOTAL", "DATA"]
            for col in required_cols:
                 if col not in df_estoque.columns:
                      self._update_status(f"Erro: Coluna '{col}' não encontrada para atualizar produto {codigo_str}.", error=True)
                      return False

            valor_un_raw_estoque = df_estoque.loc[idx, "VALOR UN"] # Get raw value from DataFrame
            current_valor_un = pd.to_numeric(valor_un_raw_estoque, errors='coerce') # Convert to numeric
            if pd.isna(current_valor_un):
                current_valor_un = 0.0
            nova_quantidade_num = pd.to_numeric(nova_quantidade, errors='coerce')

            if pd.isna(nova_quantidade_num) or nova_quantidade_num < 0:
                 raise ValueError("Nova quantidade inválida.")

            df_estoque.loc[idx, "QUANTIDADE"] = nova_quantidade_num
            df_estoque.loc[idx, "VALOR TOTAL"] = current_valor_un * nova_quantidade_num
            df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y") # Update timestamp

            # Save changes (main save, potentially creates backup)
            if self._safe_write_csv(df_estoque, arquivo_estoque):
                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual() # Refresh view if needed
                 return True
            else:
                 return False # Save failed

        except Exception as e: # Catch potential KeyErrors, ValueErrors, etc.
             self._update_status(f"Erro ao atualizar QTD/Valor estoque para {codigo_str}: {e}", error=True)
             messagebox.showerror("Erro de Atualização", f"Não foi possível atualizar quantidade/valor do produto {codigo_str}.\nVerifique os dados.\n\nDetalhes: {e}")
             return False

    def _obter_proximo_codigo(self):
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        if df_estoque.empty or 'CODIGO' not in df_estoque.columns:
            return "1"
        codes_numeric = pd.to_numeric(df_estoque['CODIGO'], errors='coerce')
        max_code = codes_numeric.max()
        if pd.isna(max_code):
            # Find max string code if all numeric fail? Risky. Start from 1.
            numeric_codes_exist = codes_numeric.notna().any()
            if not numeric_codes_exist:
                 # Try to find max string numerically if possible
                 max_str_code = 0
                 for code in df_estoque['CODIGO'].dropna():
                     try:
                         num_code = int(code)
                         if num_code > max_str_code:
                             max_str_code = num_code
                     except ValueError:
                         continue # Skip non-integer strings
                 return str(max_str_code + 1)
            else:
                 # Some numeric codes exist, but max is NaN (mixed types likely)
                 # Fallback to finding the max int among them
                 return str(int(codes_numeric.dropna().max()) + 1)

        else:
            try:
                return str(int(max_code) + 1)
            except ValueError:
                return str(int(round(max_code)) + 1)


    # --- Métodos de Lógica Principal (Cadastro, Movimentação, EPIs) ---

    # --- MODIFIED: Cadastro Logic ---
    def _cadastrar_estoque(self):
        """Cadastra um novo produto no estoque."""
        desc = self.cadastro_entries["DESCRICAO"].get().strip().upper()
        classificacao = self.cadastro_classificacao_combo.get().strip().upper() # Get from combobox
        qtd_str = self.cadastro_entries["QUANTIDADE"].get().strip().replace(",",".")
        val_un_str = self.cadastro_entries["VALOR UN"].get().strip().replace(",",".")
        loc = self.cadastro_entries["LOCALIZACAO"].get().strip().upper()
        nf_pedido_inicial = self.cadastro_entries["NF/PEDIDO"].get().strip().upper() # Get initial NF

        # Validation
        if not desc:
             messagebox.showerror("Erro de Validação", "Descrição não pode ser vazia.", parent=self.cadastro_tab)
             self.cadastro_entries["DESCRICAO"].focus_set()
             return
        if not classificacao: # Check if classification was selected
             messagebox.showerror("Erro de Validação", "Classificação deve ser selecionada.", parent=self.cadastro_tab)
             self.cadastro_classificacao_combo.focus_set()
             return
        # Removed check: NF/Pedido can be initially empty if desired
        # if not nf_pedido_inicial:
        #      messagebox.showerror("Erro de Validação", "NF/Pedido Inicial não pode ser vazio.", parent=self.cadastro_tab)
        #      self.cadastro_entries["NF/PEDIDO"].focus_set()
        #      return

        try:
            quantidade = float(qtd_str)
            if quantidade < 0: raise ValueError("Quantidade negativa")
        except ValueError:
             messagebox.showerror("Erro de Validação", "Quantidade deve ser um número válido não negativo.", parent=self.cadastro_tab)
             self.cadastro_entries["QUANTIDADE"].focus_set()
             return
        try:
            valor_un = float(val_un_str)
            if valor_un < 0: raise ValueError("Valor negativo")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Valor Unitário deve ser um número válido não negativo.", parent=self.cadastro_tab)
            self.cadastro_entries["VALOR UN"].focus_set()
            return

        codigo = self._obter_proximo_codigo()
        valor_total = quantidade * valor_un
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")
        # Store initial NF/Pedido as a list containing one item (or empty list)
        nf_pedido_list = [nf_pedido_inicial] if nf_pedido_inicial else []

        confirm_msg = (f"Cadastrar o seguinte produto?\n\n"
                       f"Código: {codigo}\n"
                       f"Descrição: {desc}\n"
                       f"Classificação: {classificacao}\n" # Show classification
                       f"Quantidade: {quantidade}\n"
                       f"Valor Unitário: R$ {valor_un:.2f}\n"
                       f"Localização: {loc if loc else '-'}\n"
                       f"NF/Pedido Inicial: {nf_pedido_inicial if nf_pedido_inicial else '-'}\n") # Show initial NF

        if messagebox.askyesno("Confirmar Cadastro", confirm_msg, parent=self.cadastro_tab):
            novo_produto = {
                 "CODIGO": codigo, "DESCRICAO": desc, "CLASSIFICACAO": classificacao, # Added
                 "VALOR UN": valor_un, "VALOR TOTAL": valor_total, "QUANTIDADE": quantidade,
                 "DATA": data_hora, "LOCALIZACAO": loc,
                 "NF/PEDIDO": str(nf_pedido_list) # Store list as string
            }

            try:
                 arquivo_estoque = ARQUIVOS["estoque"]
                 df_append = pd.DataFrame([novo_produto])

                 # Check if file exists and is not empty to determine header
                 header = not os.path.exists(arquivo_estoque) or os.path.getsize(arquivo_estoque) == 0

                 # Ensure correct column order before appending
                 all_cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "VALOR UN", "VALOR TOTAL", "QUANTIDADE", "DATA", "LOCALIZACAO", "NF/PEDIDO"]
                 df_append = df_append.reindex(columns=all_cols)

                 df_append.to_csv(arquivo_estoque, mode='a', header=header, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)

                 messagebox.showinfo("Sucesso", f"Produto '{desc}' (Cód: {codigo}) cadastrado com sucesso!", parent=self.cadastro_tab)
                 self._update_status(f"Produto {codigo} - {desc} cadastrado.")

                 # Clear fields
                 for entry in self.cadastro_entries.values():
                     entry.delete(0, tk.END)
                 self.cadastro_classificacao_combo.current(0) # Reset combobox
                 self.cadastro_entries["DESCRICAO"].focus_set()

                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual()

            except Exception as e:
                 self._update_status(f"Erro ao salvar novo produto: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o produto:\n{e}", parent=self.cadastro_tab)


    # --- MODIFIED: Entrada Logic ---
    def _registrar_entrada(self):
        """Registra a entrada de um produto no estoque."""
        codigo = self.entrada_codigo_entry.get().strip()
        # Get the original input from the user for NF/Pedido
        nf_pedido_input = self.entrada_nf_pedido_entry.get().strip().upper()
        # Get the original input from the user for Data Emissao
        data_emissao_input = self.entrada_data_emissao_entry.get().strip()
        qtd_str = self.entrada_qtd_entry.get().strip().replace(",",".")

        # --- Basic Validation (Code and Quantity) ---
        if not codigo:
             messagebox.showerror("Erro de Validação", "Código do produto não pode ser vazio.", parent=self.movimentacao_tab)
             self.entrada_codigo_entry.focus_set()
             return
        try:
            quantidade_adicionada = float(qtd_str)
            if quantidade_adicionada <= 0:
                 raise ValueError("Quantidade deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Quantidade de entrada deve ser um número positivo.", parent=self.movimentacao_tab)
            self.entrada_qtd_entry.focus_set()
            return

        # --- Prepare NF/Pedido for Saving ---
        # If the input is empty, use "SEM NF" for the Entrada CSV
        nf_pedido_para_salvar = "SEM NF" if not nf_pedido_input else nf_pedido_input
        # We will pass the original 'nf_pedido_input' to the stock update function later

        # --- Prepare and Validate Data Emissão ---
        data_emissao_para_salvar = "" # Initialize
        # Handle placeholder and empty input
        if data_emissao_input == "HH:MM DD/MM/YY" or not data_emissao_input:
            data_emissao_para_salvar = "SEM DATA"
        else:
            # Validate format only if it's not empty or placeholder
            date_pattern = r"^\d{2}:\d{2} \d{2}/\d{2}/\d{2}$"
            if not re.match(date_pattern, data_emissao_input):
                 messagebox.showerror("Erro de Validação", "Formato da Data de Emissão inválido. Use HH:MM DD/MM/YY ou deixe vazio.", parent=self.movimentacao_tab)
                 self.entrada_data_emissao_entry.focus_set()
                 return
            # Optional: More robust date/time validation (e.g., check valid month/day/hour/minute)
            try:
                 time_part, date_part = data_emissao_input.split(' ')
                 hour, minute = map(int, time_part.split(':'))
                 day, month, year_short = map(int, date_part.split('/'))
                 # Basic range checks
                 if not (0 <= hour <= 23 and 0 <= minute <= 59): raise ValueError("Hora/Minuto inválido")
                 if not (1 <= day <= 31 and 1 <= month <= 12): raise ValueError("Dia/Mês inválido")
                 # Validation passed, use the original input for saving
                 data_emissao_para_salvar = data_emissao_input
            except ValueError as date_val_err:
                messagebox.showerror("Erro de Validação", f"Data/Hora de Emissão inválida: {date_val_err}.", parent=self.movimentacao_tab)
                self.entrada_data_emissao_entry.focus_set()
                return

        # --- Get Product Info ---
        produto_atual = self._buscar_produto(codigo)
        if produto_atual is None:
            messagebox.showerror("Erro", f"Produto com código '{codigo}' não encontrado no estoque.", parent=self.movimentacao_tab)
            self.entrada_codigo_entry.focus_set()
            return

        desc = produto_atual.get("DESCRICAO", "N/A")
        classificacao = produto_atual.get("CLASSIFICACAO", "N/A")
        # Get value and convert, handle non-numeric as NaN
        val_un_raw = produto_atual.get("VALOR UN", 0)
        val_un = pd.to_numeric(val_un_raw, errors='coerce')
        if pd.isna(val_un): val_un = 0.0
        # Get current quantity
        qtd_atual_raw = produto_atual.get("QUANTIDADE", 0)
        qtd_atual = pd.to_numeric(qtd_atual_raw, errors='coerce')
        if pd.isna(qtd_atual): qtd_atual = 0.0


        nova_quantidade_estoque = qtd_atual + quantidade_adicionada
        valor_total_entrada = val_un * quantidade_adicionada
        data_registro = datetime.now().strftime("%H:%M %d/%m/%Y") # System timestamp for registration

        # --- Confirmation Dialog ---
        # Display the value that will be *saved* in Entrada.csv
        confirm_msg = (f"Registrar Entrada?\n\n"
                       f"Código: {codigo} ({desc})\n"
                       f"Classificação: {classificacao}\n"
                       f"NF/Pedido: {nf_pedido_para_salvar}\n"
                       f"Data Emissão: {data_emissao_para_salvar}\n"
                       f"Qtd. a adicionar: {quantidade_adicionada}\n"
                       f"Nova Qtd. Estoque: {nova_quantidade_estoque}\n"
                       f"Valor Total Entrada: R$ {valor_total_entrada:.2f}")

        if messagebox.askyesno("Confirmar Entrada", confirm_msg, parent=self.movimentacao_tab):
            # 1. Registrar na planilha de Entrada
            entrada_data = {
                 "CODIGO": codigo, "DESCRICAO": desc, "CLASSIFICACAO": classificacao,
                 "QUANTIDADE": quantidade_adicionada, "VALOR UN": val_un,
                 "VALOR TOTAL": valor_total_entrada, "DATA": data_registro, # System time
                 "ID": self.operador_logado_id,
                 # Use the placeholders if original input was empty
                 "NF/PEDIDO": nf_pedido_para_salvar,
                 "DATA EMISSAO": data_emissao_para_salvar
            }
            try:
                 arquivo_entrada = ARQUIVOS["entrada"]
                 df_entrada_append = pd.DataFrame([entrada_data])
                 header = not os.path.exists(arquivo_entrada) or os.path.getsize(arquivo_entrada) == 0
                 entrada_cols = ["CODIGO", "DESCRICAO", "CLASSIFICACAO", "QUANTIDADE", "VALOR UN", "VALOR TOTAL", "DATA", "ID", "NF/PEDIDO", "DATA EMISSAO"]
                 df_entrada_append = df_entrada_append.reindex(columns=entrada_cols)
                 df_entrada_append.to_csv(arquivo_entrada, mode='a', header=header, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)

                 # 2. Adicionar NF/Pedido ao Estoque
                 #    PASS THE ORIGINAL INPUT ('nf_pedido_input') not 'nf_pedido_para_salvar'
                 #    The helper function '_adicionar_nf_pedido_estoque' already checks
                 #    if the input is empty and won't add empty strings.
                 nf_added_to_stock = self._adicionar_nf_pedido_estoque(codigo, nf_pedido_input)

                 if not nf_added_to_stock:
                     messagebox.showwarning("Atenção NF/Pedido", f"Entrada registrada, mas houve erro ao adicionar a NF/Pedido '{nf_pedido_input}' ao histórico do item {codigo} no estoque. Verifique o estoque.", parent=self.movimentacao_tab)
                 # else: NF successfully added or input was empty

                 # 3. Atualizar Quantidade/Valor no Estoque
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Entrada registrada para '{desc}' (Cód: {codigo}).\nEstoque atualizado: {nova_quantidade_estoque}", parent=self.movimentacao_tab)
                     self._update_status(f"Entrada registrada para {codigo}. Novo estoque: {nova_quantidade_estoque}")
                     # Clear fields and reset placeholder
                     self.entrada_codigo_entry.delete(0, tk.END)
                     self.entrada_nf_pedido_entry.delete(0, tk.END)
                     self.entrada_data_emissao_entry.delete(0, tk.END)
                     target_widget = self.entrada_data_emissao_entry # Reset placeholder directly
                     if not target_widget.get():
                         target_widget.insert(0, "HH:MM DD/MM/YY")
                         target_widget.config(foreground='grey')
                     self.entrada_qtd_entry.delete(0, tk.END)
                     self.entrada_codigo_entry.focus_set()
                     # Update view if showing entrada or estoque
                     if self.active_table_name == "entrada" or self.active_table_name == "estoque":
                         self._atualizar_tabela_atual()
                 else:
                     messagebox.showwarning("Atenção Estoque", "Entrada registrada (com NF/Data como preenchido), mas houve erro ao atualizar a quantidade/valor do estoque. Verifique os dados.", parent=self.movimentacao_tab)

            except Exception as e:
                 self._update_status(f"Erro ao registrar entrada: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Entrada", f"Não foi possível salvar a entrada:\n{e}", parent=self.movimentacao_tab)


    def _parse_nf_pedido_list(self, nf_pedido_string):
        """Tenta parsear uma string para uma lista de NF/Pedidos. Retorna lista."""
        nf_pedido_list = []
        if isinstance(nf_pedido_string, str) and nf_pedido_string.strip().startswith('[') and nf_pedido_string.strip().endswith(']'):
            try:
                parsed = ast.literal_eval(nf_pedido_string)
                if isinstance(parsed, list):
                    # Garante que todos os itens da lista sejam strings
                    nf_pedido_list = [str(item).strip() for item in parsed if str(item).strip()]
                else: # Se parseou algo que não é lista
                    nf_pedido_list = [str(parsed).strip()] if str(parsed).strip() else []
            except (ValueError, SyntaxError):
                    # Se parse falhou, mas parece lista, tenta pegar conteúdo interno
                    # Isso é um fallback mais complexo, pode simplificar se não necessário
                    content = nf_pedido_string.strip()[1:-1].strip()
                    if content:
                        # Tenta split por vírgula e limpar aspas/espaços
                        nf_pedido_list = [item.strip().strip("'").strip('"').strip() for item in content.split(',') if item.strip()]
                    else:
                        nf_pedido_list = [] # Lista vazia se falhar totalmente

        elif isinstance(nf_pedido_string, str) and nf_pedido_string.strip():
                # Se for só uma string não-lista, trata como lista de um item
            nf_pedido_list = [nf_pedido_string.strip()]
        # Ignora outros tipos ou vazios/NaN (retorna lista vazia)

        # Remove duplicatas e ordena (opcional, mas pode ser útil)
        # return sorted(list(set(nf_pedido_list)))
        # Ou apenas remove duplicatas sem ordenar
        return list(dict.fromkeys(nf_pedido_list)) # Preserva ordem enquanto remove duplicatas

    def _registrar_saida(self):
        """Registra a saída de um produto do estoque."""
        codigo = self.saida_codigo_entry.get().strip()
        solicitante = self.saida_solicitante_entry.get().strip().upper()
        # ** Pega valor do Setor **
        setor = self.saida_setor_entry.get().strip().upper()
        qtd_str = self.saida_qtd_entry.get().strip().replace(",",".")
        # ** Pega valor do Combobox NF/Pedido **
        nf_pedido_selecionado = self.saida_nf_pedido_combo.get().strip()

        # --- Validações ---
        if not codigo:
             messagebox.showerror("Erro de Validação", "Código do produto não pode ser vazio.", parent=self.movimentacao_tab)
             self.saida_codigo_entry.focus_set()
             return
        if not solicitante:
             messagebox.showerror("Erro de Validação", "Nome do Solicitante não pode ser vazio.", parent=self.movimentacao_tab)
             self.saida_solicitante_entry.focus_set()
             return
        # ** Validação do Setor (Opcional, exemplo: não pode ser vazio) **
        if not setor:
             messagebox.showerror("Erro de Validação", "Setor não pode ser vazio.", parent=self.movimentacao_tab)
             self.saida_setor_entry.focus_set()
             return
        # ** Validação do NF/Pedido (Opcional) **
        # Se for mandatório selecionar uma NF (e ela existir):
        # if self.saida_nf_pedido_combo['state'] == 'readonly' and not nf_pedido_selecionado:
        #      messagebox.showerror("Erro de Validação", "Selecione uma NF/Pedido de origem.", parent=self.movimentacao_tab)
        #      self.saida_nf_pedido_combo.focus_set()
        #      return
        # Se for opcional, não precisa validar aqui, mas tratar o valor vazio abaixo.

        try:
            quantidade_retirada = float(qtd_str)
            if quantidade_retirada <= 0:
                 raise ValueError("Quantidade deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro de Validação", "Quantidade de saída deve ser um número positivo.", parent=self.movimentacao_tab)
            self.saida_qtd_entry.focus_set()
            return

        # --- Busca Produto e Calcula Valor ---
        produto_atual = self._buscar_produto(codigo)
        if produto_atual is None:
            messagebox.showerror("Erro", f"Produto com código '{codigo}' não encontrado no estoque.", parent=self.movimentacao_tab)
            self.saida_codigo_entry.focus_set()
            return

        desc = produto_atual.get("DESCRICAO", "N/A")

        # --- CORREÇÃO APLICADA AQUI (val_un) ---
        val_un_raw = produto_atual.get("VALOR UN", 0)
        val_un = pd.to_numeric(val_un_raw, errors='coerce')
        if pd.isna(val_un):
             val_un = 0.0

        # --- CORREÇÃO APLICADA AQUI (qtd_atual) ---
        qtd_atual_raw = produto_atual.get("QUANTIDADE", 0)
        qtd_atual = pd.to_numeric(qtd_atual_raw, errors='coerce')
        if pd.isna(qtd_atual):
            qtd_atual = 0.0
        # --- FIM DA SEÇÃO CORRIGIDA ---

        if quantidade_retirada > qtd_atual:
             messagebox.showerror("Erro", f"Quantidade insuficiente no estoque!\n\nDisponível para '{desc}': {qtd_atual}\nSolicitado: {quantidade_retirada}", parent=self.movimentacao_tab)
             self.saida_qtd_entry.focus_set()
             return

        # ** Calcula o Valor da Saída **
        valor_saida = round(quantidade_retirada * val_un, 2)

        nova_quantidade_estoque = qtd_atual - quantidade_retirada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")

        # Trata NF/Pedido selecionado vazio para o save (salva string vazia)
        nf_pedido_para_salvar = nf_pedido_selecionado if nf_pedido_selecionado else ""

        # --- Mensagem de Confirmação Atualizada ---
        confirm_msg = (f"Registrar Saída?\n\n"
                       f"Código: {codigo} ({desc})\n"
                       f"Solicitante: {solicitante}\n"
                       f"Setor: {setor}\n"                   # Adicionado
                       f"Qtd. a retirar: {quantidade_retirada}\n"
                       f"Valor da Saída: R$ {valor_saida:.2f}\n" # Adicionado
                       f"NF/Ped. Origem: {nf_pedido_para_salvar if nf_pedido_para_salvar else '-'}\n" # Adicionado/Ajustado
                       f"Qtd. Restante: {nova_quantidade_estoque}")

        if messagebox.askyesno("Confirmar Saída", confirm_msg, parent=self.movimentacao_tab):
            # 1. Registrar na planilha de Saída (com novos campos)
            saida_data = {
                 "CODIGO": codigo, "DESCRICAO": desc, "QUANTIDADE": quantidade_retirada,
                 "VALOR": valor_saida,            # Adicionado
                 "SOLICITANTE": solicitante,
                 "SETOR": setor,                  # Adicionado
                 "NF/PEDIDO": nf_pedido_para_salvar, # Adicionado
                 "DATA": data_hora, "ID": self.operador_logado_id
            }
            try:
                 arquivo_saida = ARQUIVOS["saida"]
                 df_saida_append = pd.DataFrame([saida_data])
                 header = not os.path.exists(arquivo_saida) or os.path.getsize(arquivo_saida) == 0

                 # Usa a ordem definida em EXPECTED_COLUMNS
                 saida_cols = EXPECTED_COLUMNS["saida"]
                 df_saida_append = df_saida_append.reindex(columns=saida_cols)

                 df_saida_append.to_csv(arquivo_saida, mode='a', header=header, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)

                 # 2. Atualizar Estoque (Quantidade)
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Saída registrada para '{desc}' (Solic.: {solicitante} / Setor: {setor}).\nEstoque restante: {nova_quantidade_estoque}", parent=self.movimentacao_tab)
                     self._update_status(f"Saída registrada para {codigo}. Estoque restante: {nova_quantidade_estoque}")
                     # Limpa campos (incluindo os novos)
                     self.saida_codigo_entry.delete(0, tk.END)
                     self.saida_solicitante_entry.delete(0, tk.END)
                     self.saida_setor_entry.delete(0, tk.END) # Limpa setor
                     self.saida_qtd_entry.delete(0, tk.END)
                     # Reseta o combobox NF/Pedido
                     self.saida_nf_pedido_combo['values'] = ["(Informe o Código)"]
                     self.saida_nf_pedido_combo.current(0)
                     self.saida_nf_pedido_combo.config(state="disabled")

                     self.saida_codigo_entry.focus_set() # Foca no código novamente
                     # Atualiza a visualização se estiver mostrando saida ou estoque
                     if self.active_table_name == "saida" or self.active_table_name == "estoque":
                         self._atualizar_tabela_atual()
                 else:
                     # A atualização do estoque falhou, mas a saída FOI registrada. Alerta crucial.
                     messagebox.showwarning("Atenção Crítica", "A Saída foi registrada no histórico, mas OCORREU UM ERRO ao atualizar a quantidade no estoque. Verifique os dados IMEDIATAMENTE para evitar inconsistência.", parent=self.movimentacao_tab)
                     # Poderia tentar reverter a escrita em Saida.csv aqui, mas é complexo. A melhor opção é alertar fortemente.

            except Exception as e:
                 self._update_status(f"Erro ao registrar saída: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Saída", f"Não foi possível salvar a saída:\n{e}", parent=self.movimentacao_tab)

    # --- EPI Methods (Unchanged for now) ---
    def _registrar_epi(self):
        # ... (Keep existing logic) ...
        ca = self.epi_ca_entry.get().strip().upper()
        descricao = self.epi_desc_entry.get().strip().upper()
        qtd_str = self.epi_qtd_entry.get().strip().replace(",",".")

        if not (ca or descricao):
            messagebox.showerror("Erro", "Você deve preencher pelo menos o CA ou a Descrição.", parent=self.epis_tab)
            self.epi_ca_entry.focus_set()
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
        if df_epis.empty: # Initialize if empty
             df_epis = pd.DataFrame(columns=["CA", "DESCRICAO", "QUANTIDADE"])

        # Ensure types before processing
        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)


        found_epi_index = None
        epi_existente_data = None

        if ca:
             match = df_epis[df_epis["CA"] == ca]
             if not match.empty:
                 found_epi_index = match.index[0]
                 epi_existente_data = match.iloc[0]

        if found_epi_index is None and descricao:
            match = df_epis[df_epis["DESCRICAO"] == descricao]
            if not match.empty:
                existing_ca = match.iloc[0]["CA"]
                if ca and existing_ca and existing_ca != ca:
                    messagebox.showwarning("Conflito de Dados", f"A descrição '{descricao}' já existe mas está associada ao CA '{existing_ca}'.\nNão é possível adicionar com o CA '{ca}'. Verifique os dados.", parent=self.epis_tab)
                    return
                found_epi_index = match.index[0]
                epi_existente_data = match.iloc[0]

        if epi_existente_data is not None and found_epi_index is not None:
             qtd_atual = epi_existente_data["QUANTIDADE"]
             epi_display_ca = epi_existente_data['CA'] if epi_existente_data['CA'] else '-'
             epi_display_desc = epi_existente_data['DESCRICAO']
             confirm_msg = (f"EPI já existe:\n"
                            f" CA: {epi_display_ca}\n Descrição: {epi_display_desc}\n"
                            f" Qtd. Atual: {qtd_atual}\n\n"
                            f"Deseja adicionar {quantidade_add} à quantidade existente?")
             if messagebox.askyesno("Confirmar Adição", confirm_msg, parent=self.epis_tab):
                 nova_quantidade = qtd_atual + quantidade_add
                 df_epis.loc[found_epi_index, "QUANTIDADE"] = nova_quantidade
                 if ca and df_epis.loc[found_epi_index, "CA"] != ca:
                      df_epis.loc[found_epi_index, "CA"] = ca
                 if descricao and df_epis.loc[found_epi_index, "DESCRICAO"] != descricao:
                      df_epis.loc[found_epi_index, "DESCRICAO"] = descricao

                 if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                      desc_atualizada = df_epis.loc[found_epi_index, "DESCRICAO"]
                      ca_atualizado = df_epis.loc[found_epi_index, "CA"]
                      display_name = desc_atualizada if desc_atualizada else ca_atualizado
                      self._update_status(f"Quantidade EPI {display_name} atualizada para {nova_quantidade}.")
                      messagebox.showinfo("Sucesso", f"Quantidade atualizada!\nNova Quantidade: {nova_quantidade}", parent=self.epis_tab)
                      self._atualizar_tabela_epis()
                      self.epi_ca_entry.delete(0, tk.END)
                      self.epi_desc_entry.delete(0, tk.END)
                      self.epi_qtd_entry.delete(0, tk.END)
                      self.epi_ca_entry.focus_set()
             else:
                  messagebox.showinfo("Operação Cancelada", "A quantidade não foi alterada.", parent=self.epis_tab)

        else:
             confirm_msg = (f"Registrar novo EPI?\n\n"
                           f" CA: {ca if ca else '-'}\n"
                           f" Descrição: {descricao}\n"
                           f" Quantidade: {quantidade_add}")
             if messagebox.askyesno("Confirmar Registro", confirm_msg, parent=self.epis_tab):
                novo_epi = {"CA": ca, "DESCRICAO": descricao, "QUANTIDADE": quantidade_add}
                try:
                    arquivo_epis = ARQUIVOS["epis"]
                    df_epi_append = pd.DataFrame([novo_epi])
                    header = not os.path.exists(arquivo_epis) or os.path.getsize(arquivo_epis) == 0
                    # Ensure order
                    epi_cols = ["CA", "DESCRICAO", "QUANTIDADE"]
                    df_epi_append = df_epi_append.reindex(columns=epi_cols)

                    df_epi_append.to_csv(arquivo_epis, mode='a', header=header, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)

                    display_name = descricao if descricao else ca
                    self._update_status(f"Novo EPI {display_name} registrado com {quantidade_add} unidades.")
                    messagebox.showinfo("Sucesso", f"EPI '{display_name}' registrado com sucesso!", parent=self.epis_tab)
                    self._atualizar_tabela_epis()
                    self.epi_ca_entry.delete(0, tk.END)
                    self.epi_desc_entry.delete(0, tk.END)
                    self.epi_qtd_entry.delete(0, tk.END)
                    self.epi_ca_entry.focus_set()

                except Exception as e:
                     self._update_status(f"Erro ao registrar novo EPI: {e}", error=True)
                     messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o novo EPI:\n{e}", parent=self.epis_tab)


    def _registrar_retirada(self):
        # ... (Keep existing logic) ...
        identificador = self.retirar_epi_id_entry.get().strip().upper()
        qtd_ret_str = self.retirar_epi_qtd_entry.get().strip().replace(",", ".")
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
            if quantidade_retirada <= 0:
                raise ValueError("Qtd deve ser positiva.")
        except ValueError:
            messagebox.showerror("Erro", "Quantidade a retirar deve ser um número positivo.", parent=self.epis_tab)
            self.retirar_epi_qtd_entry.focus_set()
            return

        df_epis = self._safe_read_csv(ARQUIVOS["epis"])
        if df_epis.empty:
             messagebox.showerror("Erro", "Não há EPIs cadastrados para retirar.", parent=self.epis_tab)
             return

        # Ensure types
        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)


        epi_match = pd.DataFrame()
        if identificador:
            # Try CA first if it looks numeric or is non-empty
            if identificador.isdigit() or identificador:
                match_ca = df_epis[df_epis["CA"] == identificador]
                if not match_ca.empty:
                    epi_match = match_ca

            # If no CA match (or identifier wasn't CA-like), try Description
            if epi_match.empty:
                match_desc = df_epis[df_epis["DESCRICAO"] == identificador]
                if not match_desc.empty:
                    epi_match = match_desc

            # Final check: if it wasn't numeric but still no match, try CA again
            if epi_match.empty and not identificador.isdigit():
                 match_ca_final = df_epis[df_epis["CA"] == identificador]
                 if not match_ca_final.empty:
                     epi_match = match_ca_final


        if epi_match.empty:
            messagebox.showerror("Erro", f"EPI com CA/Descrição '{identificador}' não encontrado.", parent=self.epis_tab)
            self.retirar_epi_id_entry.focus_set()
            return

        epi_data = epi_match.iloc[0]
        epi_index = epi_match.index[0]
        ca_epi = epi_data["CA"]
        desc_epi = epi_data["DESCRICAO"]
        qtd_disponivel = epi_data["QUANTIDADE"]

        if quantidade_retirada > qtd_disponivel:
            epi_display_ca = f"(CA: {ca_epi})" if ca_epi else ""
            messagebox.showerror("Erro", f"Quantidade insuficiente para '{desc_epi}' {epi_display_ca}.\nDisponível: {qtd_disponivel}", parent=self.epis_tab)
            self.retirar_epi_qtd_entry.focus_set()
            return

        safe_colaborador_name = "".join(c for c in colaborador if c.isalnum() or c in (' ', '_')).rstrip()
        if not safe_colaborador_name:
            messagebox.showerror("Erro", "Nome do Colaborador inválido para criar pasta.", parent=self.epis_tab)
            return
        pasta_colaborador = os.path.join(COLABORADORES_DIR, safe_colaborador_name)

        if not os.path.exists(pasta_colaborador):
            if not messagebox.askyesno("Criar Pasta", f"A pasta para o colaborador '{safe_colaborador_name}' não existe.\nDeseja criá-la?", parent=self.epis_tab):
                return
            try:
                os.makedirs(pasta_colaborador, exist_ok=True)
            except OSError as e:
                messagebox.showerror("Erro", f"Não foi possível criar a pasta para o colaborador '{safe_colaborador_name}':\n{e}", parent=self.epis_tab)
                return

        nova_qtd_epi = qtd_disponivel - quantidade_retirada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y")
        display_name = desc_epi if desc_epi else ca_epi
        epi_display_ca_confirm = f"(CA: {ca_epi})" if ca_epi else ""

        confirm_msg = (f"Confirmar Retirada?\n\n"
                    f"Colaborador: {colaborador}\n"
                    f"EPI: {display_name} {epi_display_ca_confirm}\n"
                    f"Qtd. Retirar: {quantidade_retirada}\n"
                    f"Qtd. Restante: {nova_qtd_epi}")

        if messagebox.askyesno("Confirmar Retirada", confirm_msg, parent=self.epis_tab):
            df_epis.loc[epi_index, "QUANTIDADE"] = nova_qtd_epi
            if not self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                messagebox.showerror("Erro Crítico", "Falha ao atualizar a quantidade de EPIs. A retirada NÃO foi registrada.", parent=self.epis_tab)
                self._carregar_epis() # Reload to show original state
                return

            nome_arquivo_colab = f"{safe_colaborador_name}_{datetime.now().strftime('%Y_%m')}.csv"
            caminho_arquivo_colab = os.path.join(pasta_colaborador, nome_arquivo_colab)
            colab_file_data = {
                "CA": ca_epi if ca_epi else '', "DESCRICAO": desc_epi,
                "QTD RETIRADA": quantidade_retirada, "DATA": data_hora
            }
            try:
                colab_cols = ["CA", "DESCRICAO", "QTD RETIRADA", "DATA"]
                header_colab = not os.path.exists(caminho_arquivo_colab) or os.path.getsize(caminho_arquivo_colab) == 0
                df_colab_append = pd.DataFrame([colab_file_data], columns=colab_cols)
                df_colab_append.to_csv(caminho_arquivo_colab, mode='a', header=header_colab, index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)

                messagebox.showinfo("Sucesso", f"Retirada de {quantidade_retirada} '{display_name}' registrada para {colaborador}.", parent=self.epis_tab)
                self._update_status(f"Retirada EPI {display_name} para {colaborador}.")
                self._atualizar_tabela_epis()
                self.retirar_epi_id_entry.delete(0, tk.END)
                self.retirar_epi_qtd_entry.delete(0, tk.END)
                self.retirar_epi_colab_entry.delete(0, tk.END)
                self.retirar_epi_id_entry.focus_set()

            except Exception as e:
                self._update_status(f"Erro ao salvar retirada no arquivo do colaborador {colaborador}: {e}", error=True)
                messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a retirada no arquivo de {colaborador}, mas a quantidade de EPIs FOI alterada.\nVerifique manualmente.\n\nDetalhe: {e}", parent=self.epis_tab)
                # Reload EPI table as quantity WAS changed
                self._atualizar_tabela_epis()


    # --- Lançadores de Diálogo de Busca ---
    def _show_product_lookup(self, target_field_prefix):
         df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
         if df_estoque.empty:
             # Para a saída, o messagebox deve pertencer à aba movimentacao
             parent_tab = self.movimentacao_tab if hasattr(self, 'movimentacao_tab') else self.root
             messagebox.showinfo("Estoque Vazio", "Não há produtos cadastrados no estoque para buscar.", parent=parent_tab)
             return

         dialog = LookupDialog(self.root, "Buscar Produto no Estoque", df_estoque, ["CODIGO", "DESCRICAO"], "CODIGO")
         result_code = dialog.result

         if result_code:
              entry_to_fill = None
              next_focus_widget = None
              if target_field_prefix == "entrada":
                   entry_to_fill = self.entrada_codigo_entry
                   next_focus_widget = self.entrada_nf_pedido_entry # Focus NF/Pedido Entrada next
              elif target_field_prefix == "saida":
                   entry_to_fill = self.saida_codigo_entry
                   # ** Chamar atualização do dropdown APÓS preencher o código **
                   # next_focus_widget = self.saida_solicitante_entry # Opcional, talvez não focar auto

              if entry_to_fill:
                   entry_to_fill.delete(0, tk.END)
                   entry_to_fill.insert(0, result_code)

              # ** Atualiza o dropdown para Saida após o lookup **
              if target_field_prefix == "saida":
                   self._populate_saida_nf_dropdown(result_code)
                   # Opcional: Mover o foco para o próximo campo desejado
                   if hasattr(self, 'saida_solicitante_entry'):
                       self._focus_widget(self.saida_solicitante_entry)


    def _show_epi_lookup(self):
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         if df_epis.empty:
             messagebox.showinfo("EPIs Vazios", "Não há EPIs cadastrados para buscar.", parent=self.epis_tab)
             return

         # Ensure columns for search
         if "CA" not in df_epis.columns: df_epis["CA"] = ""
         if "DESCRICAO" not in df_epis.columns: df_epis["DESCRICAO"] = ""
         df_epis["CA"] = df_epis["CA"].astype(str).fillna("")
         df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("")

         search_cols = ["CA", "DESCRICAO"]
         # Return CA if available, otherwise Descrição
         # Modify LookupDialog might be better, but for now, check result
         dialog = LookupDialog(self.root, "Buscar EPI", df_epis, search_cols, "CA") # Prefer returning CA
         result_id = dialog.result

         if result_id:
              # Find the EPI based on the returned ID (which should be CA if possible)
              found_epi = df_epis[df_epis["CA"] == result_id]
              display_text = result_id # Default to returned ID

              if found_epi.empty:
                 # If CA didn't match, maybe it returned Description? Try matching that.
                 found_epi_desc = df_epis[df_epis["DESCRICAO"] == result_id]
                 if not found_epi_desc.empty:
                     epi_data = found_epi_desc.iloc[0]
                     # Prefer CA for display if it exists, else use the description
                     display_text = epi_data['CA'] if epi_data['CA'] else epi_data['DESCRICAO']
              elif not found_epi.empty:
                  # Found by CA, ensure display text is CA if present
                  epi_data = found_epi.iloc[0]
                  display_text = epi_data['CA'] if epi_data['CA'] else epi_data['DESCRICAO']


              self.retirar_epi_id_entry.delete(0, tk.END)
              self.retirar_epi_id_entry.insert(0, display_text)
              self._focus_widget(self.retirar_epi_qtd_entry) # Focus Qtd


    # --- Funcionalidade Editar / Excluir ---
    def _get_selected_data(self, table):
        if not table or not hasattr(table, 'model') or not hasattr(table.model, 'df') or table.model.df is None:
            # print("Debug: Tabela ou modelo inválido em _get_selected_data")
            return None

        model = table.model
        current_df = model.df

        if not hasattr(table, 'getSelectedRow'):
             print("Debug: Método getSelectedRow não encontrado na tabela.")
             return None

        row_num = table.getSelectedRow()
        if row_num < 0:
            # print(f"Debug: Nenhuma linha selecionada (row_num={row_num})")
            return None

        if row_num >= len(current_df):
            print(f"Debug: Erro - linha visual selecionada {row_num} fora dos limites do DF atual {len(current_df)}")
            # messagebox.showwarning("Seleção Inválida", "A linha selecionada parece inválida após atualização. Tente selecionar novamente.", parent=self.root)
            # table.clearSelected() # Optional
            return None

        try:
            index_label = current_df.index[row_num]
            selected_series = current_df.loc[index_label].copy() # Return a copy

            # --- PARSE NF/PEDIDO LIST HERE USING HELPER ---
            if table == self.pandas_table and self.active_table_name == "estoque" and "NF/PEDIDO" in selected_series:
                nf_str = selected_series["NF/PEDIDO"]
                # Chama a função auxiliar para parsear
                selected_series["NF/PEDIDO"] = self._parse_nf_pedido_list(nf_str)

            return selected_series

        except IndexError:
            print(f"Debug: IndexError ao acessar índice na posição {row_num} no DF de tamanho {len(current_df)}.")
            return None
        except Exception as e:
            print(f"Debug: Erro inesperado em _get_selected_data para linha {row_num}: {e}")
            messagebox.showerror("Erro Interno", f"Ocorreu um erro ao obter dados da linha selecionada.\n\nDetalhes: {e}", parent=self.root)
            return None


    def _edit_selected_item(self):
         """Lida com a edição do item selecionado na visualização da tabela principal."""

         if not self.pandas_table or self.active_table_name not in ["estoque", "entrada"]:
              # Ajuste para permitir 'entrada'
              messagebox.showwarning("Ação Inválida", "A edição está disponível apenas para as tabelas de Estoque ou Entrada.", parent=self.estoque_tab)
              return

         selected_item_series = self._get_selected_data(self.pandas_table)

         if selected_item_series is None:
              messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único item para editar.", parent=self.estoque_tab)
              return

         item_dict = selected_item_series.to_dict()

         # --- Lógica para Estoque (sem mudanças) ---
         if self.active_table_name == "estoque":
             if 'CODIGO' not in item_dict:
                 messagebox.showerror("Erro Interno", "Não foi possível obter o código do item selecionado.", parent=self.estoque_tab)
                 return
             codigo_to_edit = str(item_dict['CODIGO'])
             dialog = EditProductDialog(self.root, f"Editar Produto: {codigo_to_edit}", item_dict)
             updated_data = dialog.updated_data # Bloqueia até fechar

             if updated_data: # Usuário salvou
                # --- (O código para salvar as edições do estoque permanece o mesmo) ---
                df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
                # ... (restante do código de atualização do estoque) ...
                if 'CODIGO' not in df_estoque.columns:
                   messagebox.showerror("Erro", "Coluna 'CODIGO' não encontrada no estoque para salvar edição.", parent=self.estoque_tab)
                   return
                df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
                idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_edit].tolist()
                if not idx:
                      messagebox.showerror("Erro", f"Item {codigo_to_edit} não encontrado para salvar após edição.", parent=self.estoque_tab)
                      return
                idx = idx[0]
                try:
                      original_nf_pedido = df_estoque.loc[idx, "NF/PEDIDO"] # Get original list string
                      for key, value in updated_data.items():
                           if key == "NF/PEDIDO": continue # Skip NF/PEDIDO list editing here
                           if key in df_estoque.columns:
                                try:
                                     current_dtype = df_estoque[key].dtype
                                     if pd.api.types.is_numeric_dtype(current_dtype) and not pd.isna(value):
                                          value_str = str(value).strip().replace(',', '.')
                                          value = pd.to_numeric(value_str) if value_str else 0
                                     elif pd.api.types.is_string_dtype(current_dtype) or pd.api.types.is_object_dtype(current_dtype):
                                          value = str(value).strip()
                                     df_estoque.loc[idx, key] = value
                                except (ValueError, TypeError) as conv_err:
                                      print(f"Warning: Conversion failed for {key}='{value}': {conv_err}")
                                      continue
                      df_estoque.loc[idx, "NF/PEDIDO"] = original_nf_pedido # Ensure NF/PEDIDO remains unchanged
                      if "DATA" in df_estoque.columns: df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y")

                      if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                          messagebox.showinfo("Sucesso", f"Produto {codigo_to_edit} atualizado.", parent=self.estoque_tab)
                          self._atualizar_tabela_atual()

                except Exception as e:
                    self._update_status(f"Erro ao aplicar edições estoque para {codigo_to_edit}: {e}", error=True)
                    messagebox.showerror("Erro ao Salvar Edição", f"Não foi possível salvar as alterações para {codigo_to_edit}.\n\nDetalhes: {e}", parent=self.estoque_tab)
                # --- Fim do Bloco de Salvar Estoque ---

         # --- NOVA Lógica para Entrada ---
         elif self.active_table_name == "entrada":
             if 'CODIGO' not in item_dict:
                  # A linha selecionada na tabela de entrada pode não ter todos os dados se o arquivo estiver corrompido
                  messagebox.showerror("Erro Interno", "Não foi possível obter o Código do produto da linha de entrada selecionada.", parent=self.estoque_tab)
                  return
             codigo_produto = str(item_dict['CODIGO']) # Código do produto relacionado a esta entrada
             desc_produto = item_dict.get('DESCRICAO', 'N/A')

             # Abre o diálogo de "edição" de entrada
             dialog = EditEntradaDialog(self.root, f"Adicionar NF/Pedido (Ref: {codigo_produto} - {desc_produto})", item_dict)
             result_data = dialog.updated_data # Bloqueia até fechar

             if result_data and "NF_PEDIDO_ADICIONAL" in result_data:
                  # Usuário clicou em Salvar e digitou uma nova NF
                  nf_adicional = result_data["NF_PEDIDO_ADICIONAL"]

                  # Chama a função auxiliar para adicionar a NF ao arquivo de Estoque
                  if self._adicionar_nf_pedido_estoque(codigo_produto, nf_adicional):
                       messagebox.showinfo("Sucesso", f"NF/Pedido '{nf_adicional}' adicionada com sucesso ao histórico do item {codigo_produto}.", parent=self.estoque_tab)
                       self._update_status(f"NF/Pedido {nf_adicional} adicionado ao item {codigo_produto}.")
                       # Atualiza a visualização do ESTOQUE se ela estiver ativa
                       if self.active_table_name == "estoque":
                            self._atualizar_tabela_atual()
                       # Não precisa atualizar a visualização de ENTRADA pois o registro histórico não mudou
                  else:
                       # Mensagem de erro já foi mostrada por _adicionar_nf_pedido_estoque
                       self._update_status(f"Falha ao adicionar NF/Pedido {nf_adicional} ao item {codigo_produto}.", error=True)

             # else: O usuário cancelou ou não digitou uma NF/Pedido adicional.


    def _delete_selected_item(self):
         # No changes needed here based on requirements
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

         if messagebox.askyesno("Confirmar Exclusão", confirm_msg, icon='warning', parent=self.estoque_tab):
              df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
              if 'CODIGO' not in df_estoque.columns:
                   messagebox.showerror("Erro", "Coluna 'CODIGO' não encontrada no estoque para excluir.", parent=self.estoque_tab)
                   return
              df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
              idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_delete].tolist()

              if not idx:
                  messagebox.showerror("Erro", f"Item {codigo_to_delete} não encontrado no arquivo para excluir.", parent=self.estoque_tab)
                  return

              df_estoque.drop(idx, inplace=True)

              if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                   messagebox.showinfo("Sucesso", f"Produto {codigo_to_delete} - {desc_to_delete} excluído.", parent=self.estoque_tab)
                   self._atualizar_tabela_atual()
              # else: Error shown by safe_write


    def _edit_selected_epi(self):
        # No changes needed here based on requirements
        selected_epi_series = self._get_selected_data(self.epis_table)
        if selected_epi_series is None:
             messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único EPI para editar.", parent=self.epis_tab)
             return

        epi_dict = selected_epi_series.to_dict()
        ca_to_edit = epi_dict.get('CA',"").strip()
        desc_to_edit = epi_dict.get('DESCRICAO',"").strip()
        if not (ca_to_edit or desc_to_edit):
             messagebox.showerror("Erro Interno", "EPI selecionado não tem CA ou Descrição.", parent=self.epis_tab)
             return

        dialog = EditEPIDialog(self.root, f"Editar EPI: {desc_to_edit if desc_to_edit else ca_to_edit}", epi_dict)
        updated_data = dialog.updated_data

        if updated_data:
             df_epis = self._safe_read_csv(ARQUIVOS["epis"])
             if df_epis.empty:
                  messagebox.showerror("Erro", "Arquivo de EPIs vazio ou não encontrado para editar.", parent=self.epis_tab)
                  return

             # Ensure types
             df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
             df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
             df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

             # Find original row
             match = pd.DataFrame()
             if ca_to_edit:
                 match = df_epis[df_epis["CA"] == ca_to_edit]
             if match.empty and desc_to_edit:
                 if not ca_to_edit:
                      match = df_epis[(df_epis["DESCRICAO"] == desc_to_edit) & (df_epis["CA"] == "")]
                 else:
                      match = df_epis[df_epis["DESCRICAO"] == desc_to_edit]

             if match.empty:
                  messagebox.showerror("Erro", f"EPI original (CA:'{ca_to_edit}'/Desc:'{desc_to_edit}') não encontrado para salvar.", parent=self.epis_tab)
                  return
             idx = match.index[0]

             try:
                    new_ca = updated_data['CA'].strip().upper()
                    new_desc = updated_data['DESCRICAO'].strip().upper()

                    # Check for conflicts before updating
                    if new_ca and any((df_epis['CA'] == new_ca) & (df_epis.index != idx)):
                        messagebox.showerror("Erro de Duplicidade", f"O CA '{new_ca}' já existe para outro EPI.", parent=self.epis_tab)
                        return
                    if new_desc:
                        desc_conflict = df_epis[
                            (df_epis['DESCRICAO'] == new_desc) & (df_epis.index != idx) &
                            ( (df_epis['CA'] != new_ca) if new_ca else (df_epis['CA'] != "") )
                        ]
                        if not desc_conflict.empty:
                            messagebox.showerror("Erro de Duplicidade", f"A Descrição '{new_desc}' já existe para outro EPI com CA diferente/vazio.", parent=self.epis_tab)
                            return

                    df_epis.loc[idx, "CA"] = new_ca
                    df_epis.loc[idx, "DESCRICAO"] = new_desc
                    try:
                        new_qty = float(str(updated_data['QUANTIDADE']).replace(',','.'))
                        if new_qty >= 0:
                            df_epis.loc[idx, "QUANTIDADE"] = new_qty
                        else:
                             messagebox.showwarning("Aviso", "Quantidade negativa inválida ao editar EPI. Mantendo valor anterior.", parent=self.epis_tab)
                    except ValueError:
                         messagebox.showwarning("Aviso", f"Quantidade inválida '{updated_data['QUANTIDADE']}' ao editar EPI. Mantendo valor anterior.", parent=self.epis_tab)

                    if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                        display_name = new_desc if new_desc else new_ca
                        messagebox.showinfo("Sucesso", f"EPI '{display_name}' atualizado.", parent=self.epis_tab)
                        self._atualizar_tabela_epis()
                    # else: Error handled

             except Exception as e:
                 self._update_status(f"Erro ao aplicar edições EPI: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Edição EPI", f"Não foi possível salvar as alterações.\n\nDetalhes: {e}", parent=self.epis_tab)


    def _delete_selected_epi(self):
        # No changes needed here
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

        if messagebox.askyesno("Confirmar Exclusão EPI", confirm_msg, icon='warning', parent=self.epis_tab):
            df_epis = self._safe_read_csv(ARQUIVOS["epis"])
            if df_epis.empty: return # Nothing to delete

            # Ensure types
            df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
            df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()

            # Find original row(s)
            match = pd.DataFrame()
            if ca_to_delete:
                 match = df_epis[df_epis["CA"] == ca_to_delete]
            if match.empty and desc_to_delete:
                 if not ca_to_delete:
                      match = df_epis[(df_epis["DESCRICAO"] == desc_to_delete) & (df_epis["CA"] == "")]
                 else:
                      match = df_epis[df_epis["DESCRICAO"] == desc_to_delete]

            if match.empty:
                 messagebox.showerror("Erro", f"EPI original (CA:'{ca_to_delete}'/Desc:'{desc_to_delete}') não encontrado para excluir.", parent=self.epis_tab)
                 return

            idx = match.index.tolist()
            df_epis.drop(idx, inplace=True)

            if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                 display_name = desc_to_delete if desc_to_delete else ca_to_delete
                 messagebox.showinfo("Sucesso", f"EPI '{display_name}' excluído.", parent=self.epis_tab)
                 self._atualizar_tabela_epis()
            # else: Error handled


    # --- Backup e Exportação ---
    def _criar_backup_periodico(self):
        arquivo_ultimo_backup = os.path.join(BACKUP_DIR, "ultimo_backup_timestamp.txt")
        backup_interval_seconds = 3 * 60 * 60 # 3 hours

        # --- Limpeza de Backups Antigos ---
        try:
            if os.path.exists(BACKUP_DIR):
                now = time.time()
                cutoff_time = now - (3 * 24 * 60 * 60) # 3 dias
                for filename in os.listdir(BACKUP_DIR):
                    if filename.endswith(".csv") and ("_backup_" in filename or "_auto.csv" in filename):
                        file_path = os.path.join(BACKUP_DIR, filename)
                        try:
                            if os.path.getmtime(file_path) < cutoff_time:
                                os.remove(file_path)
                                # print(f"Backup antigo removido: {filename}")
                        except OSError as e:
                             print(f"Erro ao processar/remover backup antigo {filename}: {e}")
        except Exception as e:
            self._update_status(f"Erro ao limpar backups antigos: {e}", error=True)

        # --- Cria Novo Backup ---
        perform_backup = True
        if os.path.exists(arquivo_ultimo_backup):
            try:
                with open(arquivo_ultimo_backup, "r", encoding="utf-8") as f:
                    last_backup_content = f.read().strip()
                    if last_backup_content:
                         ultimo_backup_time = float(last_backup_content)
                         if (time.time() - ultimo_backup_time) < backup_interval_seconds:
                             perform_backup = False
                    # else: empty timestamp file, proceed with backup
            except (ValueError, FileNotFoundError, TypeError) as e:
                print(f"Erro ao ler timestamp do último backup: {e}. Criando novo backup.")
                perform_backup = True

        if perform_backup:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            try:
                os.makedirs(BACKUP_DIR, exist_ok=True)
                backup_count = 0
                for nome, arquivo_origem in ARQUIVOS.items():
                    if os.path.exists(arquivo_origem):
                        try:
                            nome_backup = f"{nome}_{timestamp}_auto.csv"
                            caminho_backup = os.path.join(BACKUP_DIR, nome_backup)
                            shutil.copy2(arquivo_origem, caminho_backup)
                            backup_count += 1
                        except Exception as copy_e:
                             err_msg = f"Erro ao copiar {arquivo_origem} para backup: {copy_e}"
                             print(err_msg)
                             self._update_status(err_msg, error=True)


                if backup_count > 0:
                    try:
                        with open(arquivo_ultimo_backup, "w", encoding="utf-8") as f:
                            f.write(str(time.time()))
                        status_msg = f"Backup automático criado ({backup_count} arquivos)."
                        self._update_status(status_msg)
                        print(status_msg)
                    except Exception as ts_write_e:
                         err_msg = f"Erro ao escrever timestamp do backup: {ts_write_e}"
                         print(err_msg)
                         self._update_status(err_msg, error=True)
                # else: No files backed up

            except Exception as e:
                err_msg = f"Erro durante backup automático: {e}"
                print(err_msg)
                self._update_status(err_msg, error=True)
                if self.root and self.root.winfo_exists():
                    messagebox.showerror("Erro de Backup", f"Falha ao criar backup automático:\n{e}", parent=self.root)

    def _schedule_backup(self):
        self._criar_backup_periodico()
        if self.root and self.root.winfo_exists():
            self.root.after(10800000, self._schedule_backup) # Reagenda


    def _exportar_conteudo(self):
        """Exporta dados atuais para Excel e gera relatório txt para estoque baixo."""
        # No changes needed here based on requirements
        pasta_saida = "Relatorios"
        try:
            os.makedirs(pasta_saida, exist_ok=True)
        except OSError as e:
             messagebox.showerror("Erro de Exportação", f"Não foi possível criar a pasta '{pasta_saida}':\n{e}", parent=self.root)
             return

        data_atual = datetime.now().strftime("%d-%m-%Y_%H%M")
        caminho_excel = os.path.join(pasta_saida, f"Relatorio_Almoxarifado_{data_atual}.xlsx")
        caminho_txt = os.path.join(pasta_saida, f"Produtos_Esgotados_{data_atual}.txt")

        try:
            with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer: # Specify engine if needed
                all_files = {**ARQUIVOS}
                self._update_status("Iniciando exportação para Excel...")
                sheet_count = 0
                for nome, arquivo in all_files.items():
                     try:
                         df_export = self._safe_read_csv(arquivo)
                         if not df_export.empty:
                             safe_sheet_name = "".join(c for c in nome.capitalize() if c.isalnum() or c in (' ', '_'))[:31]
                             df_export.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                             sheet_count +=1
                         # else: print(f"Planilha '{nome}' vazia, não incluída.")
                     except Exception as e:
                          print(f"Erro ao processar {arquivo} para Excel: {e}")
                          if self.root and self.root.winfo_exists():
                              messagebox.showwarning("Aviso de Exportação", f"Erro ao incluir '{nome}' no Excel:\n{e}", parent=self.root)

            self._update_status(f"Exportação Excel concluída ({sheet_count} planilhas). Gerando relatório...")

            df_estoque_report = self._safe_read_csv(ARQUIVOS["estoque"])
            required_cols = ["QUANTIDADE", "CODIGO", "DESCRICAO"]
            if not df_estoque_report.empty and all(col in df_estoque_report.columns for col in required_cols):
                 df_estoque_report["QUANTIDADE"] = pd.to_numeric(df_estoque_report["QUANTIDADE"], errors='coerce').fillna(0)
                 produtos_esgotados = df_estoque_report[df_estoque_report["QUANTIDADE"] <= 0]

                 try:
                     with open(caminho_txt, "w", encoding="utf-8") as f:
                         f.write(f"Relatório de Produtos Esgotados/Zerados - {data_atual}\n")
                         f.write("=" * 50 + "\n")
                         if not produtos_esgotados.empty:
                             for _, row in produtos_esgotados.iterrows():
                                  codigo_str = str(row.get('CODIGO', 'N/A'))
                                  desc_str = str(row.get('DESCRICAO', 'N/A'))
                                  f.write(f"Código: {codigo_str} | Descrição: {desc_str}\n")
                         else:
                             f.write("Nenhum produto com quantidade zero ou negativa encontrado.\n")
                         f.write("=" * 50 + "\n")

                     if self.root and self.root.winfo_exists():
                         messagebox.showinfo("Sucesso", f"Relatórios exportados!\n\nExcel: {caminho_excel}\nEsgotados: {caminho_txt}", parent=self.root)
                     self._update_status("Relatórios exportados com sucesso.")

                 except Exception as txt_e:
                      err_msg = f"Erro ao gerar relatório de esgotados: {txt_e}"
                      self._update_status(err_msg, error=True)
                      if self.root and self.root.winfo_exists():
                           messagebox.showerror("Erro de Exportação", f"Falha ao gerar relatório de esgotados:\n{txt_e}", parent=self.root)

            else:
                  reason = "dados do estoque vazios" if df_estoque_report.empty else f"colunas ausentes ({', '.join(c for c in required_cols if c not in df_estoque_report.columns)})"
                  warning_msg = f"Não foi possível gerar relatório de esgotados ({reason})."
                  if self.root and self.root.winfo_exists():
                      messagebox.showwarning("Aviso", warning_msg, parent=self.root)
                  self._update_status(f"Relatório de esgotados não gerado ({reason}).")

        except Exception as e:
            err_msg = f"Erro geral ao exportar relatórios: {e}"
            self._update_status(err_msg, error=True)
            if self.root and self.root.winfo_exists():
                messagebox.showerror("Erro de Exportação", f"Falha ao exportar relatórios:\n{e}", parent=self.root)

    # --- Fechamento da Aplicação ---
    def _on_close(self):
        if self.root and self.root.winfo_exists():
            if messagebox.askyesno("Confirmar Saída", "Deseja realmente fechar o aplicativo?", icon='warning', parent=self.root):
                print("Fechando aplicativo...")
                # Cancel scheduled tasks if ID was stored
                # try: self.root.after_cancel(self.backup_schedule_id)
                # except: pass
                self.root.destroy()
        else:
             print("Fechando aplicativo (janela já não existia).")


# --- Classe da Janela de Login ---
class LoginWindow:
    # No changes needed here based on requirements
    def __init__(self, root):
        self.root = root
        self.root.title("Almoxarifado - Login")
        self.root.geometry("300x250")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_login)

        self.logged_in_user_id = None

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Login.TButton", foreground="white", background="#107C10", font=('-size', 11, '-weight', 'bold'))
        style.map("Login.TButton", background=[('active', '#0A530A')])

        login_frame = ttk.Frame(root, padding="20")
        login_frame.pack(expand=True, fill="both")
        ttk.Label(login_frame, text="Login", font="-size 16 -weight bold").pack(pady=(0, 15))
        ttk.Label(login_frame, text="Usuário:", font="-size 12").pack(pady=(5, 0))
        self.usuario_entry = ttk.Entry(login_frame, width=30, font="-size 12")
        self.usuario_entry.pack(pady=(0, 10))
        self.usuario_entry.bind("<Return>", lambda e: self.senha_entry.focus_set())
        ttk.Label(login_frame, text="Senha:", font="-size 12").pack(pady=(5, 0))
        self.senha_entry = ttk.Entry(login_frame, show="*", width=30, font="-size 12")
        self.senha_entry.pack(pady=(0, 15))
        self.senha_entry.bind("<Return>", lambda e: self._validate_login())

        login_button = ttk.Button(login_frame, text="Entrar", command=self._validate_login, style="Login.TButton")
        login_button.pack(pady=5, fill=tk.X, ipady=4)

        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f'+{x}+{y}')
        self.usuario_entry.focus_set()

    def _validate_login(self):
        usuario = self.usuario_entry.get().strip()
        senha = self.senha_entry.get()

        if not usuario or not senha:
            messagebox.showwarning("Login Inválido", "Por favor, preencha usuário e senha.", parent=self.root)
            return

        # !!! INSECURE: Plain text password comparison !!!
        # Replace with hashed password check (e.g., bcrypt) in a real application
        if usuario in usuarios and usuarios[usuario]["senha"] == senha:
            self.logged_in_user_id = usuarios[usuario]["id"]
            messagebox.showinfo("Sucesso", f"Login bem-sucedido!\nOperador ID: {self.logged_in_user_id}", parent=self.root)
            self.root.destroy()
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos!", parent=self.root)
            self.senha_entry.delete(0, tk.END)
            self.senha_entry.focus_set()

    def _on_close_login(self):
        if self.root and self.root.winfo_exists():
            if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair?", icon='question', parent=self.root):
                 self.logged_in_user_id = None
                 self.root.destroy()
        else:
            print("Saindo (janela de login já não existia).")

    def get_user_id(self):
         return self.logged_in_user_id


# --- Execução Principal ---
if __name__ == "__main__":
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()

    user_id = login_app.get_user_id()
    if user_id:
        main_root = tk.Tk()
        app = AlmoxarifadoApp(main_root, user_id)
        main_root.mainloop()
    else:
        print("Login cancelado ou falhou. Saindo.")
