import os
import csv
import time
import shutil
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from usuarios import usuarios  # Assuming usuarios.py is in the same directory
from pandastable import Table, TableModel

# --- Configuration ---
PLANILHAS_DIR = "Planilhas"
BACKUP_DIR = "Backups"
COLABORADORES_DIR = "Colaboradores"

ARQUIVOS = {
    "estoque": os.path.join(PLANILHAS_DIR, "Estoque.csv"),
    "entrada": os.path.join(PLANILHAS_DIR, "Entrada.csv"),
    "saida": os.path.join(PLANILHAS_DIR, "Saida.csv"),
    "epis": os.path.join(PLANILHAS_DIR, "Epis.csv")
}

# --- Helper Classes for Dialogs ---

class LookupDialog(tk.Toplevel):
    """A simple searchable lookup dialog."""
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

        # Search Frame
        search_frame = ttk.Frame(self, padding="5")
        search_frame.pack(fill=tk.X)
        ttk.Label(search_frame, text="Buscar:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._filter_list)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        search_entry.focus_set()

        # Listbox Frame
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

        # Button Frame
        button_frame = ttk.Frame(self, padding="5")
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Selecionar", command=self._select_item).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.wait_window(self) # Wait until the window is closed

    def _populate_listbox(self, df):
        self.listbox.delete(0, tk.END)
        # Adjust display format as needed
        for index, row in df.iterrows():
            # Example display: Code - Description (or CA - Description for EPIs)
            display_text = f"{row[self.return_col]} - {row.get('DESCRICAO', '')}"
            self.listbox.insert(tk.END, display_text)

    def _filter_list(self, *args):
        query = self.search_var.get().strip().lower()
        df_to_search = self.df_full # Use the original df for filtering
        if not query:
            df_filtered = df_to_search
        else:
            try:
                # --- CORRECTED SEARCH LOGIC ---
                mask = df_to_search.apply(
                    lambda row: any(query in str(row[col]).lower() # Use 'in' for substring check
                                    for col in self.search_cols if pd.notna(row[col])),
                    axis=1
                )
                # --- END CORRECTION ---
                df_filtered = df_to_search[mask]
            except Exception as e:
                 print(f"Error during lookup filter: {e}") # Log error
                 df_filtered = df_to_search # Show all on error

        self._populate_listbox(df_filtered)

    def _select_item(self, event=None):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            # Retrieve the stored data
            list_item_text = self.listbox.get(index)
            # Extract the identifier (first part before ' - ')
            self.result = list_item_text.split(' - ')[0]
            self.destroy()
        else:
            messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um item da lista.", parent=self)


class EditDialogBase(tk.Toplevel):
    """Base class for edit dialogs."""
    def __init__(self, parent, title, item_data):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.geometry("400x300") # Adjust as needed

        self.item_data = item_data # Dictionary with original data
        self.updated_data = None # Will hold the result if saved

        self.entries = {}
        self._create_widgets()

        # Button Frame
        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Button(button_frame, text="Salvar", command=self._save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    def _create_widgets(self):
        # To be implemented by subclasses
        pass

    def _add_entry(self, frame, field_key, label_text, row, col, **kwargs):
        """Helper to create label and entry."""
        ttk.Label(frame, text=label_text + ":").grid(row=row, column=col*2, sticky=tk.W, padx=5, pady=2)
        entry = ttk.Entry(frame, **kwargs)
        entry.grid(row=row, column=col*2 + 1, sticky=tk.EW, padx=5, pady=2)
        if field_key in self.item_data:
             entry.insert(0, str(self.item_data[field_key]))
        self.entries[field_key] = entry
        frame.columnconfigure(col*2 + 1, weight=1) # Make entry expand

    def _validate_and_collect(self):
        # To be implemented by subclasses for specific validation
        collected_data = {}
        for key, entry in self.entries.items():
             collected_data[key] = entry.get().strip()
        return collected_data # Return collected data

    def _save(self):
        try:
            self.updated_data = self._validate_and_collect()
            if self.updated_data: # validation passed if it returned data
                self.destroy()
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e), parent=self)


class EditProductDialog(EditDialogBase):
    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._add_entry(main_frame, "CODIGO", "Código", 0, 0, state='readonly') # Read-only
        self._add_entry(main_frame, "DESCRICAO", "Descrição", 1, 0)
        self._add_entry(main_frame, "VALOR UN", "Valor Unitário", 2, 0)
        self._add_entry(main_frame, "QUANTIDADE", "Quantidade", 3, 0)
        self._add_entry(main_frame, "LOCALIZACAO", "Localização", 4, 0)
        # DATA and VALOR TOTAL are usually calculated, not directly edited here

    def _validate_and_collect(self):
        data = super()._validate_and_collect()

        # Validation
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

        # Calculate Valor Total based on updated values
        data["VALOR TOTAL"] = data["VALOR UN"] * data["QUANTIDADE"]

        return data

class EditEPIDialog(EditDialogBase):
     def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._add_entry(main_frame, "CA", "CA", 0, 0) # Allow editing CA maybe? Or make readonly?
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

        # Make sure CA/Desc are upper
        data["CA"] = data["CA"].upper()
        data["DESCRICAO"] = data["DESCRICAO"].upper()
        return data


# --- Main Application Class ---

class AlmoxarifadoApp:
    def __init__(self, root, user_id):
        self.root = root
        self.operador_logado_id = user_id
        self.active_table_name = "estoque"
        self.current_table_df = None # Holds the dataframe for the active pandastable

        self.root.title(f"Almoxarifado - Operador: {self.operador_logado_id}")
        self.root.geometry("1150x650") # Slightly larger for status bar
        self.root.config(bg="#E0E0E0") # Lighter background
        # self.root.resizable(False, False) # Consider allowing resizing

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # --- CORRECTED ORDER ---
        self._criar_pastas_e_planilhas()
        self._setup_ui()                # <--- Call UI setup FIRST
        self._criar_backup_periodico()  # <--- THEN call backup check
        self._load_and_display_table(self.active_table_name) # Load initial table

        # Schedule periodic backup check (e.g., every 3 hours)
        self.root.after(10800000, self._schedule_backup) # 10800000 ms = 3 hours

    def _setup_ui(self):
            # Style
            style = ttk.Style()
            style.theme_use('clam') # Or 'alt', 'default', 'classic'

            # Main frame
            main_frame = ttk.Frame(self.root, padding="5")
            main_frame.pack(expand=True, fill="both")

            self.status_var = tk.StringVar()
            # Temporarily pack status bar at top to reserve space, then move it
            self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
            # self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # Pack later


            # Notebook (Tabs)
            self.notebook = ttk.Notebook(main_frame)
            self.notebook.pack(expand=True, fill="both", pady=(0, 5))


            # Create Tabs (Now they can safely call _update_status)
            self._create_estoque_tab()
            self._create_cadastro_tab()
            self._create_movimentacao_tab()
            self._create_epis_tab()


            # --- CORRECTED ORDER: Pack Status Bar at the BOTTOM now ---
            self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # Pack it in the final desired position
            self._update_status("Pronto.")

    def _create_estoque_tab(self):
        self.estoque_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.estoque_tab, text="Estoque & Tabelas")

        # --- Top Controls ---
        controls_frame = ttk.Frame(self.estoque_tab)
        controls_frame.pack(fill=tk.X, pady=(0, 10))

        # Search
        search_frame = ttk.LabelFrame(controls_frame, text="Pesquisar na Tabela Atual", padding="5")
        search_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.pesquisar_entry = ttk.Entry(search_frame, width=40)
        self.pesquisar_entry.bind("<KeyRelease>", self._pesquisar_tabela_event)
        self.pesquisar_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(search_frame, text="Buscar", command=self._pesquisar_tabela).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(search_frame, text="Limpar", command=self._limpar_pesquisa).pack(side=tk.LEFT)

        # Table Switching
        switch_frame = ttk.LabelFrame(controls_frame, text="Visualizar Tabela", padding="5")
        switch_frame.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_view_estoque = ttk.Button(switch_frame, text="Estoque", command=lambda: self._trocar_tabela_view("estoque"))
        self.btn_view_estoque.pack(side=tk.LEFT, padx=2)
        self.btn_view_entrada = ttk.Button(switch_frame, text="Entradas", command=lambda: self._trocar_tabela_view("entrada"))
        self.btn_view_entrada.pack(side=tk.LEFT, padx=2)
        self.btn_view_saida = ttk.Button(switch_frame, text="Saídas", command=lambda: self._trocar_tabela_view("saida"))
        self.btn_view_saida.pack(side=tk.LEFT, padx=2)

        # Actions Frame (Right Aligned)
        action_frame = ttk.Frame(controls_frame)
        action_frame.pack(side=tk.RIGHT)

        # Edit/Delete Frame (Grouped)
        edit_delete_frame = ttk.LabelFrame(action_frame, text="Item Selecionado", padding="5")
        edit_delete_frame.pack(side=tk.LEFT, padx=(0,10))
        self.edit_button = ttk.Button(edit_delete_frame, text="Editar", command=self._edit_selected_item, state=tk.DISABLED) # Initially disabled
        self.edit_button.pack(side=tk.LEFT, padx=2)
        self.delete_button = ttk.Button(edit_delete_frame, text="Excluir", command=self._delete_selected_item, state=tk.DISABLED) # Initially disabled
        self.delete_button.pack(side=tk.LEFT, padx=2)


        # General Actions Frame (Grouped)
        general_action_frame = ttk.LabelFrame(action_frame, text="Ações", padding="5")
        general_action_frame.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(general_action_frame, text="Atualizar", command=self._atualizar_tabela_atual).pack(side=tk.LEFT, padx=2)
        # Save button is removed - edits are saved via Edit/Delete dialogs or movements


        # Export Frame (Grouped)
        export_frame = ttk.LabelFrame(action_frame, text="Relatórios", padding="5")
        export_frame.pack(side=tk.LEFT)
        ttk.Button(export_frame, text="Exportar", command=self._exportar_conteudo).pack(side=tk.LEFT, padx=2)


        # --- Table Frame ---
        self.pandas_table_frame = ttk.Frame(self.estoque_tab)
        self.pandas_table_frame.pack(expand=True, fill="both")

        # Create dummy table initially, will be replaced
        self.pandas_table = Table(parent=self.pandas_table_frame)
        self.pandas_table.show()
        # Disable direct editing
        self.pandas_table.editable = False
        # Add binding to enable/disable buttons on selection change
        self.pandas_table.bind("<<TableSelectChanged>>", self._on_table_select)

        self._atualizar_cores_botoes_view() # Set initial button state


    def _create_cadastro_tab(self):
        self.cadastro_tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(self.cadastro_tab, text="Cadastrar Produto")

        container = ttk.Frame(self.cadastro_tab)
        container.pack(anchor=tk.CENTER) # Center the form elements

        ttk.Label(container, text="Cadastrar Novo Produto no Estoque", font="-weight bold -size 14").grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Use a consistent structure for labels and entries
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
            entry.config(font="-size 12") # Set font size for entries
            entry.grid(row=i+1, column=1, sticky=tk.EW, padx=5, pady=5)
            entry.bind("<Return>", self._focar_proximo_cadastro) # Bind Enter key
            self.cadastro_entries[key] = entry

        # Bind last entry to the cadastrar function on Return
        self.cadastro_entries["LOCALIZACAO"].bind("<Return>", lambda e: self._cadastrar_estoque())

        # Button
        cadastrar_button = ttk.Button(container, text="Cadastrar Produto", command=self._cadastrar_estoque, style="Accent.TButton") # Use accent style if available
        cadastrar_button.grid(row=len(fields)+1, column=0, columnspan=2, pady=(20, 0), sticky=tk.EW)

        # Make the entry column expandable
        container.columnconfigure(1, weight=1)

    def _create_movimentacao_tab(self):
        self.movimentacao_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.movimentacao_tab, text="Movimentação Estoque")

        # Main Grid Layout
        self.movimentacao_tab.columnconfigure(0, weight=1)
        self.movimentacao_tab.columnconfigure(1, minsize=20) # Separator space
        self.movimentacao_tab.columnconfigure(2, weight=1)
        self.movimentacao_tab.rowconfigure(0, weight=1)

        # --- Entrada Frame ---
        entrada_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Entrada", padding="15")
        entrada_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        entrada_frame.columnconfigure(1, weight=1) # Make entries expand

        # Entrada Widgets
        ttk.Label(entrada_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_codigo_entry = ttk.Entry(entrada_frame)
        self.entrada_codigo_entry.config(font="-size 12") # Set font size for entries
        self.entrada_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        self.entrada_codigo_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        ttk.Button(entrada_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("entrada")).grid(row=0, column=2, sticky=tk.W, padx=(2,5))


        ttk.Label(entrada_frame, text="Qtd. Entrada:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.entrada_qtd_entry = ttk.Entry(entrada_frame)
        self.entrada_qtd_entry.config(font="-size 12") # Set font size for entries
        self.entrada_qtd_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.entrada_qtd_entry.bind("<Return>", lambda e: self._registrar_entrada())

        entrada_button = ttk.Button(entrada_frame, text="Registrar Entrada", command=self._registrar_entrada, style="Accent.TButton")
        entrada_button.grid(row=2, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)

        # --- Separator ---
        sep = ttk.Separator(self.movimentacao_tab, orient=tk.VERTICAL)
        sep.grid(row=0, column=1, sticky="ns", padx=5, pady=5)

        # --- Saida Frame ---
        saida_frame = ttk.LabelFrame(self.movimentacao_tab, text="Registrar Saída", padding="15")
        saida_frame.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        saida_frame.columnconfigure(1, weight=1) # Make entries expand

        # Saida Widgets
        ttk.Label(saida_frame, text="Código:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_codigo_entry = ttk.Entry(saida_frame)
        self.saida_codigo_entry.config(font="-size 12") # Set font size for entries
        self.saida_codigo_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2))
        self.saida_codigo_entry.bind("<Return>", lambda e: self._focar_proximo(e))
        ttk.Button(saida_frame, text="Buscar", width=8, command=lambda: self._show_product_lookup("saida")).grid(row=0, column=2, sticky=tk.W, padx=(2,5))


        ttk.Label(saida_frame, text="Solicitante:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_solicitante_entry = ttk.Entry(saida_frame)
        self.saida_solicitante_entry.config(font="-size 12") # Set font size for entries
        self.saida_solicitante_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_solicitante_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(saida_frame, text="Qtd. Saída:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.saida_qtd_entry = ttk.Entry(saida_frame)
        self.saida_qtd_entry.config(font="-size 12") # Set font size for entries
        self.saida_qtd_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5)
        self.saida_qtd_entry.bind("<Return>", lambda e: self._registrar_saida())

        saida_button = ttk.Button(saida_frame, text="Registrar Saída", command=self._registrar_saida, style="Accent.TButton")
        saida_button.grid(row=3, column=0, columnspan=3, pady=(15, 5), sticky=tk.EW)


    def _create_epis_tab(self):
        self.epis_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.epis_tab, text="EPIs")

        # Layout: Table on left, Forms on right
        self.epis_tab.columnconfigure(0, weight=2) # Table area
        self.epis_tab.columnconfigure(1, weight=1) # Forms area
        self.epis_tab.rowconfigure(0, weight=1)

        # --- EPI Table Frame ---
        epis_table_frame = ttk.Frame(self.epis_tab)
        epis_table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        # Controls for EPI Table
        epi_table_controls = ttk.Frame(epis_table_frame)
        epi_table_controls.pack(fill=tk.X, pady=(0, 5))

        ttk.Button(epi_table_controls, text="Atualizar Lista", command=self._atualizar_tabela_epis).pack(side=tk.LEFT, padx=(0, 10))

        self.edit_epi_button = ttk.Button(epi_table_controls, text="Editar EPI Sel.", command=self._edit_selected_epi, state=tk.DISABLED)
        self.edit_epi_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_epi_button = ttk.Button(epi_table_controls, text="Excluir EPI Sel.", command=self._delete_selected_epi, state=tk.DISABLED)
        self.delete_epi_button.pack(side=tk.LEFT, padx=(0, 5))


        # EPI Pandastable
        self.epis_pandas_frame = ttk.Frame(epis_table_frame)
        self.epis_pandas_frame.pack(expand=True, fill="both")
        self.epis_table = Table(parent=self.epis_pandas_frame, editable=False) # Disable direct editing
        self.epis_table.show()
        self._carregar_epis() # Initial load
        self.epis_table.bind("<<TableSelectChanged>>", self._on_epi_table_select)

        # --- Forms Frame (Right) ---
        forms_frame = ttk.Frame(self.epis_tab)
        forms_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        forms_frame.rowconfigure(1, weight=1) # Allow retirar frame to expand if needed


        # --- Registrar EPI Frame ---
        registrar_frame = ttk.LabelFrame(forms_frame, text="Registrar / Adicionar EPI", padding="10")
        registrar_frame.pack(fill=tk.X, pady=(0, 15))
        registrar_frame.columnconfigure(1, weight=1)

        ttk.Label(registrar_frame, text="CA:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_ca_entry = ttk.Entry(registrar_frame)
        self.epi_ca_entry.config(font="-size 12") # Set font size for entries
        self.epi_ca_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_ca_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        ttk.Label(registrar_frame, text="Descrição:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_desc_entry = ttk.Entry(registrar_frame)
        self.epi_desc_entry.config(font="-size 12") # Set font size for entries
        self.epi_desc_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_desc_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(registrar_frame, text="Quantidade:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.epi_qtd_entry = ttk.Entry(registrar_frame)
        self.epi_qtd_entry.config(font="-size 12") # Set font size for entries
        self.epi_qtd_entry.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=2)
        self.epi_qtd_entry.bind("<Return>", lambda e: self._registrar_epi())

        registrar_epi_button = ttk.Button(registrar_frame, text="Registrar / Adicionar", command=self._registrar_epi, style="Accent.TButton")
        registrar_epi_button.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky=tk.EW)


        # --- Retirar EPI Frame ---
        retirar_frame = ttk.LabelFrame(forms_frame, text="Registrar Retirada de EPI", padding="10")
        retirar_frame.pack(fill=tk.BOTH, expand=True) # Fill remaining space
        retirar_frame.columnconfigure(1, weight=1)


        ttk.Label(retirar_frame, text="CA/Descrição:", font="-size 12").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_id_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_id_entry.config(font="-size 12") # Set font size for entries
        self.retirar_epi_id_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0,2), pady=5)
        self.retirar_epi_id_entry.bind("<Return>", lambda e: self._focar_proximo(e))

        ttk.Button(retirar_frame, text="Buscar", width=8, command=self._show_epi_lookup).grid(row=0, column=2, sticky=tk.W, padx=(2,5), pady=5)


        ttk.Label(retirar_frame, text="Qtd. Retirada:", font="-size 12").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_qtd_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_qtd_entry.config(font="-size 12") # Set font size for entries
        self.retirar_epi_qtd_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_qtd_entry.bind("<Return>", lambda e: self._focar_proximo(e))


        ttk.Label(retirar_frame, text="Colaborador:", font="-size 12").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.retirar_epi_colab_entry = ttk.Entry(retirar_frame)
        self.retirar_epi_colab_entry.config(font="-size 12") # Set font size for entries
        self.retirar_epi_colab_entry.grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        self.retirar_epi_colab_entry.bind("<Return>", lambda e: self._registrar_retirada())


        retirar_epi_button = ttk.Button(retirar_frame, text="Registrar Retirada", command=self._registrar_retirada, style="Accent.TButton")
        retirar_epi_button.grid(row=3, column=0, columnspan=3, pady=(10, 5), sticky=tk.EW)


    # --- UI Event Handlers & Helpers ---

    def _update_status(self, message, error=False):
        """Updates the status bar."""
        self.status_var.set(message)
        if error:
            self.status_bar.config(foreground="red")
        else:
            self.status_bar.config(foreground="black")
        # print(message) # Also print to console for debugging

    def _focar_proximo(self, event):
        """Move focus to the next widget on Enter key."""
        try:
            event.widget.tk_focusNext().focus()
        except Exception: # Ignore if no next widget
            pass
        return "break" # Prevents default Enter behavior (like adding newline)

    def _focar_proximo_cadastro(self, event):
        """Focus next specifically for cadastro fields, ordered."""
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
                 # If last entry, maybe focus the button? Or call submit? Let's focus button
                 # Assuming the button is accessible, otherwise call _cadastrar_estoque()
                 # Find the button widget if needed, or just call the function
                 # For simplicity here, we call the function if Return is pressed on last field
                 if widget == self.cadastro_entries["LOCALIZACAO"]:
                      self._cadastrar_estoque()
                 else: #Should not happen based on index check
                      pass

        except ValueError:
             # Widget not in the expected list, try default focus next
            try:
                 event.widget.tk_focusNext().focus()
            except: pass # ignore further errors
        return "break"


    def _on_table_select(self, event=None):
        """Enable/disable edit/delete buttons based on table selection."""
        selected = self.pandas_table.getSelectedRowData()
        # Assuming single selection for simplicity. Adjust if multi-select needed.
        if selected is not None and len(selected) == 1:
            self.edit_button.config(state=tk.NORMAL)
            # Only allow deleting from estoque view? Or base on table type?
            if self.active_table_name == "estoque":
                 self.delete_button.config(state=tk.NORMAL)
            else:
                 self.delete_button.config(state=tk.DISABLED)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)


    def _on_epi_table_select(self, event=None):
        """Enable/disable EPI edit/delete buttons."""
        selected = self.epis_table.getSelectedRowData()
        if selected is not None and len(selected) == 1:
            self.edit_epi_button.config(state=tk.NORMAL)
            self.delete_epi_button.config(state=tk.NORMAL)
        else:
            self.edit_epi_button.config(state=tk.DISABLED)
            self.delete_epi_button.config(state=tk.DISABLED)


    def _pesquisar_tabela_event(self, event=None):
        """Handler for key release in search entry."""
        # Optional: Add a small delay here if needed to avoid searching on every keystroke
        self._pesquisar_tabela()

    def _pesquisar_tabela(self):
        """Filters the currently displayed pandas table."""
        if self.current_table_df is None or self.pandas_table is None:
            return
        query = self.pesquisar_entry.get().strip().lower()

        if query:
            try:
                # Filter based on string columns containing the query
                string_cols = self.current_table_df.select_dtypes(include='object').columns
                mask = self.current_table_df[string_cols].apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)

                # Also try converting other columns to string and checking
                other_cols = self.current_table_df.select_dtypes(exclude='object').columns
                mask_others = self.current_table_df[other_cols].astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)

                df_filtered = self.current_table_df[mask | mask_others]

            except Exception as e:
                self._update_status(f"Erro na pesquisa: {e}", error=True)
                df_filtered = self.current_table_df # Show all on error
        else:
            df_filtered = self.current_table_df # Show all if query is empty

        # Update the table model safely
        try:
            self.pandas_table.updateModel(TableModel(df_filtered))
            self.pandas_table.redraw()
            self._update_status(f"Exibindo {len(df_filtered)} resultados para '{query}'" if query else f"Exibindo todos os {len(df_filtered)} registros.")
        except Exception as e:
            self._update_status(f"Erro ao atualizar tabela: {e}", error=True)

    def _limpar_pesquisa(self):
        """Clears the search entry and resets the table view."""
        if self.pandas_table is None or self.current_table_df is None:
            return
        self.pesquisar_entry.delete(0, tk.END)
        try:
            self.pandas_table.updateModel(TableModel(self.current_table_df))
            self.pandas_table.redraw()
            self._update_status(f"Exibindo todos os {len(self.current_table_df)} registros.")
            self._on_table_select() # Reset button states
        except Exception as e:
             self._update_status(f"Erro ao limpar pesquisa: {e}", error=True)


    def _atualizar_cores_botoes_view(self):
        """Highlights the button for the currently viewed table."""
        buttons = {
            "estoque": self.btn_view_estoque,
            "entrada": self.btn_view_entrada,
            "saida": self.btn_view_saida
        }
        for name, button in buttons.items():
             # Check if button exists before configuring
            if hasattr(self, button.winfo_name().replace("!button", "btn_view_" + name)):
                if name == self.active_table_name:
                     button.config(style="Accent.TButton") # Use accent style for active
                else:
                     button.config(style="TButton") # Default style


    # --- Data Loading and File Operations ---

    def _criar_pastas_e_planilhas(self):
        """Creates necessary directories and CSV files if they don't exist."""
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
                    # Consider exiting if essential files can't be created

    def _safe_read_csv(self, file_path):
        """Safely reads a CSV file, returning an empty DataFrame on error."""
        try:
            # Explicitly define dtypes to avoid confusion, esp for codes/CAs
            dtype_map = {'CODIGO': str, 'CA': str} # Add others if needed
            return pd.read_csv(file_path, encoding="utf-8", dtype=dtype_map)
        except FileNotFoundError:
            self._update_status(f"Aviso: Arquivo não encontrado: {file_path}. Verifique as pastas.", error=True)
            return pd.DataFrame() # Return empty df
        except Exception as e:
            self._update_status(f"Erro ao ler {file_path}: {e}", error=True)
            messagebox.showerror("Erro de Leitura", f"Não foi possível ler o arquivo {os.path.basename(file_path)}.\nVerifique se ele não está corrompido ou aberto em outro programa.\n\nDetalhes: {e}")
            return pd.DataFrame()

    def _safe_write_csv(self, df, file_path, create_backup=True):
        """Safely writes a DataFrame to CSV, optionally creating a backup."""
        backup_path = None
        try:
            # Backup before writing
            if create_backup and os.path.exists(file_path):
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                backup_path = file_path.replace(".csv", f"_backup_{timestamp}.csv")
                shutil.copy2(file_path, backup_path) # copy2 preserves metadata

            df.to_csv(file_path, index=False, encoding="utf-8")
            return True # Indicate success
        except Exception as e:
            self._update_status(f"Erro Crítico ao salvar {os.path.basename(file_path)}: {e}", error=True)
            messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar as alterações em {os.path.basename(file_path)}.\n\nDetalhes: {e}\n\n{'Um backup pode ter sido criado: ' + os.path.basename(backup_path) if backup_path else 'Não foi possível criar backup.'}")
            # Consider recovery options or warning the user data might be inconsistent
            return False # Indicate failure

    def _load_and_display_table(self, table_name):
        """Loads data from CSV and updates the main pandastable."""
        file_path = ARQUIVOS.get(table_name)
        if not file_path:
            messagebox.showerror("Erro Interno", f"Nome de tabela inválido: {table_name}")
            return

        self._update_status(f"Carregando tabela '{table_name.capitalize()}'...")
        df = self._safe_read_csv(file_path)

        # Perform sanity checks/corrections if needed *carefully*
        # Example: Ensure numeric columns are numeric, fill NaNs reasonably
        if table_name == "estoque":
             numeric_cols = ["VALOR UN", "VALOR TOTAL", "QUANTIDADE"]
             df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
             # Recalculate VALOR TOTAL for consistency? Risky if data is bad.
             # df["VALOR TOTAL"] = df["VALOR UN"] * df["QUANTIDADE"]
             # Only save back immediately if calculation is non-destructive
             # self._safe_write_csv(df, file_path, create_backup=False) # Avoid backup loops

        # Ensure string columns are strings
        string_like_cols = ["DESCRICAO", "LOCALIZACAO", "SOLICITANTE", "ID", "CODIGO", "CA"]
        for col in string_like_cols:
             if col in df.columns:
                 df[col] = df[col].astype(str).fillna('') # Fill NaN with empty string


        self.current_table_df = df # Store the loaded dataframe
        self.active_table_name = table_name

        # Update pandastable
        try:
            self.pandas_table.updateModel(TableModel(self.current_table_df))
            # Adjust column widths if needed - might need more specific logic
            # self.pandas_table.autoResizeColumns() # Use with caution on large tables
            self.pandas_table.redraw()
            self._update_status(f"Tabela '{table_name.capitalize()}' carregada ({len(self.current_table_df)} registros).")
            self._atualizar_cores_botoes_view()
            self._on_table_select() # Update button states
        except Exception as e:
            self._update_status(f"Erro ao exibir tabela '{table_name.capitalize()}': {e}", error=True)

    def _atualizar_tabela_atual(self):
        """Reloads data for the currently active table view."""
        if self.active_table_name:
             # Re-apply search filter if active
            query = self.pesquisar_entry.get().strip()
            self._load_and_display_table(self.active_table_name)
            if query:
                 self.pesquisar_entry.delete(0, tk.END)
                 self.pesquisar_entry.insert(0, query)
                 self._pesquisar_tabela() # Reapply search

    def _trocar_tabela_view(self, nome_tabela):
        """Changes the table being viewed in the 'Estoque & Tabelas' tab."""
        if nome_tabela in ARQUIVOS:
             # Clear search before switching
             self.pesquisar_entry.delete(0, tk.END)
             self._load_and_display_table(nome_tabela)
        else:
            messagebox.showerror("Erro", f"Configuração de tabela '{nome_tabela}' não encontrada.")


    def _carregar_epis(self):
         """Loads EPI data into the EPIs tab table."""
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         # Basic type handling for EPIs
         df_epis["CA"] = df_epis["CA"].astype(str).fillna("")
         df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("")
         df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

         try:
             self.epis_table.updateModel(TableModel(df_epis))
             self.epis_table.redraw()
             self._update_status(f"Lista de EPIs atualizada ({len(df_epis)} itens).")
             self._on_epi_table_select() # Update button state
         except Exception as e:
              self._update_status(f"Erro ao exibir EPIs: {e}", error=True)


    def _atualizar_tabela_epis(self):
         """Convenience method to reload EPI data."""
         self._carregar_epis()

    def _buscar_produto(self, codigo):
        """Busca um produto no estoque pelo código. Returns a Series or None."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        # Ensure Codigo column is string for comparison
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
        produto = df_estoque[df_estoque['CODIGO'] == str(codigo)]
        if not produto.empty:
            return produto.iloc[0] # Return the first match as a Series
        return None

    def _atualizar_estoque_produto(self, codigo, nova_quantidade):
        """Atualiza a quantidade e valor total de um produto no estoque."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        codigo_str = str(codigo)
        df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str) # Ensure consistency

        # Find the index
        idx = df_estoque.index[df_estoque['CODIGO'] == codigo_str].tolist()

        if not idx:
             self._update_status(f"Erro interno: Produto {codigo_str} não encontrado para atualizar.", error=True)
             return False # Product not found

        try:
            # Update quantity (handle potential type issues)
            idx = idx[0] # Use first index if multiple somehow exist
            current_valor_un = pd.to_numeric(df_estoque.loc[idx, "VALOR UN"], errors='coerce')
            if pd.isna(current_valor_un):
                current_valor_un = 0 # Default if conversion failed
            nova_quantidade = pd.to_numeric(nova_quantidade, errors='coerce')
            if pd.isna(nova_quantidade) or nova_quantidade < 0:
                 raise ValueError("Nova quantidade inválida.")


            df_estoque.loc[idx, "QUANTIDADE"] = nova_quantidade
            df_estoque.loc[idx, "VALOR TOTAL"] = current_valor_un * nova_quantidade
            df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y") # Update timestamp

            # Save the changes
            if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                 # If the estoque table is currently displayed, refresh it
                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual()
                 return True
            else:
                 return False # Save failed

        except (ValueError, KeyError, IndexError) as e:
             self._update_status(f"Erro ao atualizar estoque para {codigo_str}: {e}", error=True)
             messagebox.showerror("Erro de Atualização", f"Não foi possível atualizar o produto {codigo_str}.\nVerifique os dados na planilha.\n\nDetalhes: {e}")
             return False

    def _obter_proximo_codigo(self):
        """Obtém o próximo código disponível para um novo produto."""
        df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
        if df_estoque.empty or 'CODIGO' not in df_estoque.columns:
            return "1"  # Start from 1 if empty or no codigo column

        # Convert to numeric safely, find max, handle non-numeric codes
        codes_numeric = pd.to_numeric(df_estoque['CODIGO'], errors='coerce')
        max_code = codes_numeric.max()

        if pd.isna(max_code):
            # If all codes were non-numeric or conversion failed, start fresh
            return "1"
        else:
            return str(int(max_code) + 1)

    # --- Core Logic Methods (Cadastro, Movimentação, EPIs) ---

    def _cadastrar_estoque(self):
        """Cadastra um novo produto no estoque."""
        desc = self.cadastro_entries["DESCRICAO"].get().strip().upper()
        qtd_str = self.cadastro_entries["QUANTIDADE"].get().strip().replace(",",".")
        val_un_str = self.cadastro_entries["VALOR UN"].get().strip().replace(",",".")
        loc = self.cadastro_entries["LOCALIZACAO"].get().strip().upper()

        # Validation
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
                 # Use 'a' mode for appending with header=False if file exists
                 header = not os.path.exists(ARQUIVOS["estoque"]) or os.path.getsize(ARQUIVOS["estoque"]) == 0
                 pd.DataFrame([novo_produto]).to_csv(ARQUIVOS["estoque"], mode='a', header=header, index=False, encoding='utf-8')

                 messagebox.showinfo("Sucesso", f"Produto '{desc}' (Cód: {codigo}) cadastrado com sucesso!")
                 self._update_status(f"Produto {codigo} - {desc} cadastrado.")

                 # Clear fields
                 for entry in self.cadastro_entries.values():
                     entry.delete(0, tk.END)
                 self.cadastro_entries["DESCRICAO"].focus_set() # Focus first field

                 # Refresh view if showing estoque
                 if self.active_table_name == "estoque":
                      self._atualizar_tabela_atual()

            except Exception as e:
                 self._update_status(f"Erro ao salvar novo produto: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar o produto:\n{e}")


    def _registrar_entrada(self):
        """Registra a entrada de um produto no estoque."""
        codigo = self.entrada_codigo_entry.get().strip()
        qtd_str = self.entrada_qtd_entry.get().strip().replace(",",".")

        # Validation
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
                 header = not os.path.exists(ARQUIVOS["entrada"]) or os.path.getsize(ARQUIVOS["entrada"]) == 0
                 pd.DataFrame([entrada_data]).to_csv(ARQUIVOS["entrada"], mode='a', header=header, index=False, encoding='utf-8')

                 # 2. Atualizar Estoque
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Entrada registrada para '{desc}' (Cód: {codigo}).\nEstoque atualizado: {nova_quantidade_estoque}")
                     self._update_status(f"Entrada registrada para {codigo}. Novo estoque: {nova_quantidade_estoque}")
                     # Clear fields
                     self.entrada_codigo_entry.delete(0, tk.END)
                     self.entrada_qtd_entry.delete(0, tk.END)
                     self.entrada_codigo_entry.focus_set()
                     # Refresh view if showing entrada
                     if self.active_table_name == "entrada":
                         self._atualizar_tabela_atual()
                 else:
                      # Update failed, log/message handled in _atualizar_estoque_produto
                      # Maybe rollback the entry record? Complex without transactions.
                     messagebox.showwarning("Atenção", "Entrada registrada, mas houve erro ao atualizar o estoque. Verifique os dados.")

            except Exception as e:
                 self._update_status(f"Erro ao registrar entrada: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Entrada", f"Não foi possível salvar a entrada:\n{e}")


    def _registrar_saida(self):
        """Registra a saída de um produto do estoque."""
        codigo = self.saida_codigo_entry.get().strip()
        solicitante = self.saida_solicitante_entry.get().strip().upper()
        qtd_str = self.saida_qtd_entry.get().strip().replace(",",".")

        # Validation
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
                 header = not os.path.exists(ARQUIVOS["saida"]) or os.path.getsize(ARQUIVOS["saida"]) == 0
                 pd.DataFrame([saida_data]).to_csv(ARQUIVOS["saida"], mode='a', header=header, index=False, encoding='utf-8')

                 # 2. Atualizar Estoque
                 if self._atualizar_estoque_produto(codigo, nova_quantidade_estoque):
                     messagebox.showinfo("Sucesso", f"Saída registrada para '{desc}' (Solicitante: {solicitante}).\nEstoque restante: {nova_quantidade_estoque}")
                     self._update_status(f"Saída registrada para {codigo}. Estoque restante: {nova_quantidade_estoque}")
                     # Clear fields
                     self.saida_codigo_entry.delete(0, tk.END)
                     self.saida_solicitante_entry.delete(0, tk.END)
                     self.saida_qtd_entry.delete(0, tk.END)
                     self.saida_codigo_entry.focus_set()
                     # Refresh view if showing saida
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
            return
        if not qtd_str:
             messagebox.showerror("Erro", "A Quantidade deve ser preenchida.", parent=self.epis_tab)
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
        # Ensure types and clean strings before searching
        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

        found_epi_index = None
        epi_existente_data = None

        # Prioritize matching by CA if provided
        if ca:
             match = df_epis[df_epis["CA"] == ca]
             if not match.empty:
                 found_epi_index = match.index[0]
                 epi_existente_data = match.iloc[0]
        # If no CA match or CA wasn't provided, try matching by Description
        if found_epi_index is None and descricao:
            match = df_epis[df_epis["DESCRICAO"] == descricao]
            if not match.empty:
                # Check if this description is tied to a *different* CA than provided (if CA was provided)
                if ca and match.iloc[0]["CA"] and match.iloc[0]["CA"] != ca:
                    messagebox.showwarning("Conflito de Dados", f"A descrição '{descricao}' já existe mas está associada ao CA '{match.iloc[0]['CA']}'.\nNão é possível adicionar com o CA '{ca}'. Verifique os dados.", parent=self.epis_tab)
                    return
                found_epi_index = match.index[0]
                epi_existente_data = match.iloc[0]

        # --- Handle Found vs. New EPI ---
        if epi_existente_data is not None:
             # EPI Exists - Confirm adding quantity
             qtd_atual = epi_existente_data["QUANTIDADE"]
             confirm_msg = (f"EPI já existe:\n"
                            f" CA: {epi_existente_data['CA'] if epi_existente_data['CA'] else '-'}\n"
                            f" Descrição: {epi_existente_data['DESCRICAO']}\n"
                            f" Qtd. Atual: {qtd_atual}\n\n"
                            f"Deseja adicionar {quantidade_add} à quantidade existente?")
             if messagebox.askyesno("Confirmar Adição", confirm_msg, parent=self.epis_tab):
                 nova_quantidade = qtd_atual + quantidade_add
                 df_epis.loc[found_epi_index, "QUANTIDADE"] = nova_quantidade
                 # Also update description/ca if they were provided and different but deemed okay to update
                 if ca and df_epis.loc[found_epi_index, "CA"] != ca:
                      df_epis.loc[found_epi_index, "CA"] = ca
                 if descricao and df_epis.loc[found_epi_index, "DESCRICAO"] != descricao:
                      df_epis.loc[found_epi_index, "DESCRICAO"] = descricao

                 # Save
                 if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                      self._update_status(f"Quantidade EPI {epi_existente_data['DESCRICAO']} atualizada para {nova_quantidade}.")
                      messagebox.showinfo("Sucesso", f"Quantidade atualizada!\nNova Quantidade: {nova_quantidade}", parent=self.epis_tab)
                      self._atualizar_tabela_epis()
                      # Clear Fields
                      self.epi_ca_entry.delete(0, tk.END)
                      self.epi_desc_entry.delete(0, tk.END)
                      self.epi_qtd_entry.delete(0, tk.END)
                      self.epi_ca_entry.focus_set()
                 # else: error handled in _safe_write_csv
             else:
                  messagebox.showinfo("Operação Cancelada", "A quantidade não foi alterada.", parent=self.epis_tab)

        else:
             # New EPI - Confirm registration
             confirm_msg = (f"Registrar novo EPI?\n\n"
                           f" CA: {ca if ca else '-'}\n"
                           f" Descrição: {descricao}\n"
                           f" Quantidade: {quantidade_add}")
             if messagebox.askyesno("Confirmar Registro", confirm_msg, parent=self.epis_tab):
                novo_epi = {"CA": ca, "DESCRICAO": descricao, "QUANTIDADE": quantidade_add}
                try:
                    # Append new row
                    header = not os.path.exists(ARQUIVOS["epis"]) or os.path.getsize(ARQUIVOS["epis"]) == 0
                    pd.DataFrame([novo_epi]).to_csv(ARQUIVOS["epis"], mode='a', header=header, index=False, encoding='utf-8')
                    self._update_status(f"Novo EPI {descricao} registrado com {quantidade_add} unidades.")
                    messagebox.showinfo("Sucesso", f"EPI '{descricao}' registrado com sucesso!", parent=self.epis_tab)
                    self._atualizar_tabela_epis()
                    # Clear Fields
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
             return
        if not colaborador:
             messagebox.showerror("Erro", "Nome do Colaborador deve ser informado.", parent=self.epis_tab)
             return
        try:
             quantidade_retirada = float(qtd_ret_str)
             if quantidade_retirada <= 0: raise ValueError("Qtd deve ser positiva.")
        except ValueError:
             messagebox.showerror("Erro", "Quantidade a retirar deve ser um número positivo.", parent=self.epis_tab)
             return

        df_epis = self._safe_read_csv(ARQUIVOS["epis"])
        # Clean data for matching
        df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
        df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
        df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)

        # Find EPI by CA or Description
        epi_match = df_epis[(df_epis["CA"] == identificador) | (df_epis["DESCRICAO"] == identificador)]

        if epi_match.empty:
             messagebox.showerror("Erro", f"EPI com CA/Descrição '{identificador}' não encontrado.", parent=self.epis_tab)
             return

        epi_data = epi_match.iloc[0]
        epi_index = epi_match.index[0]
        ca_epi = epi_data["CA"] # Use the actual CA from the record
        desc_epi = epi_data["DESCRICAO"]
        qtd_disponivel = epi_data["QUANTIDADE"]

        if quantidade_retirada > qtd_disponivel:
             messagebox.showerror("Erro", f"Quantidade insuficiente para '{desc_epi}' (CA: {ca_epi}).\nDisponível: {qtd_disponivel}", parent=self.epis_tab)
             return

        # Check/Create Colaborador Folder
        pasta_colaborador = os.path.join(COLABORADORES_DIR, colaborador)
        os.makedirs(pasta_colaborador, exist_ok=True) # Create if not exists

        nova_qtd_epi = qtd_disponivel - quantidade_retirada
        data_hora = datetime.now().strftime("%H:%M %d/%m/%Y") # Consistent format

        confirm_msg = (f"Confirmar Retirada?\n\n"
                       f"Colaborador: {colaborador}\n"
                       f"EPI: {desc_epi} (CA: {ca_epi if ca_epi else '-'}) \n"
                       f"Qtd. Retirar: {quantidade_retirada}\n"
                       f"Qtd. Restante: {nova_qtd_epi}")

        if messagebox.askyesno("Confirmar Retirada", confirm_msg, parent=self.epis_tab):
             # 1. Update EPI quantity
             df_epis.loc[epi_index, "QUANTIDADE"] = nova_qtd_epi
             if not self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                  messagebox.showerror("Erro Crítico", "Falha ao atualizar a quantidade de EPIs. A retirada NÃO foi registrada.", parent=self.epis_tab)
                  return # Stop if EPI update fails

             # 2. Record withdrawal in collaborator's file
             nome_arquivo_colab = f"{colaborador}_{datetime.now().strftime('%Y_%m')}.csv"
             caminho_arquivo_colab = os.path.join(pasta_colaborador, nome_arquivo_colab)
             colab_file_data = {
                 "CA": ca_epi, "DESCRICAO": desc_epi,
                 "QTD RETIRADA": quantidade_retirada, "DATA": data_hora
             }
             try:
                 header_colab = not os.path.exists(caminho_arquivo_colab) or os.path.getsize(caminho_arquivo_colab) == 0
                 pd.DataFrame([colab_file_data]).to_csv(caminho_arquivo_colab, mode='a', header=header_colab, index=False, encoding='utf-8')

                 # Success
                 messagebox.showinfo("Sucesso", f"Retirada de {quantidade_retirada} '{desc_epi}' registrada para {colaborador}.", parent=self.epis_tab)
                 self._update_status(f"Retirada EPI {desc_epi} para {colaborador}.")
                 self._atualizar_tabela_epis()
                 # Clear fields
                 self.retirar_epi_id_entry.delete(0, tk.END)
                 self.retirar_epi_qtd_entry.delete(0, tk.END)
                 self.retirar_epi_colab_entry.delete(0, tk.END)
                 self.retirar_epi_id_entry.focus_set()

             except Exception as e:
                  self._update_status(f"Erro ao salvar retirada no arquivo do colaborador {colaborador}: {e}", error=True)
                  # Attempt to rollback EPI quantity? Risky without transactions.
                  # df_epis.loc[epi_index, "QUANTIDADE"] = qtd_disponivel # Rollback
                  # self._safe_write_csv(df_epis, ARQUIVOS["epis"]) # Try saving back
                  messagebox.showerror("Erro ao Salvar", f"Não foi possível salvar a retirada no arquivo de {colaborador}, mas a quantidade de EPIs FOI alterada.\nVerifique manualmente.\n\nDetalhe: {e}", parent=self.epis_tab)


    # --- Lookup Dialog Launchers ---

    def _show_product_lookup(self, target_field_prefix):
         """Shows lookup dialog for products and fills entry."""
         df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
         if df_estoque.empty:
             messagebox.showinfo("Estoque Vazio", "Não há produtos cadastrados no estoque para buscar.", parent=self.movimentacao_tab)
             return

         dialog = LookupDialog(self.root, "Buscar Produto no Estoque", df_estoque, ["CODIGO", "DESCRICAO"], "CODIGO")
         result_code = dialog.result # This blocks until dialog is closed

         if result_code:
              if target_field_prefix == "entrada":
                   self.entrada_codigo_entry.delete(0, tk.END)
                   self.entrada_codigo_entry.insert(0, result_code)
                   self.entrada_qtd_entry.focus_set() # Focus next field
              elif target_field_prefix == "saida":
                   self.saida_codigo_entry.delete(0, tk.END)
                   self.saida_codigo_entry.insert(0, result_code)
                   self.saida_solicitante_entry.focus_set() # Focus next field

    def _show_epi_lookup(self):
         """Shows lookup dialog for EPIs and fills entry."""
         df_epis = self._safe_read_csv(ARQUIVOS["epis"])
         if df_epis.empty:
             messagebox.showinfo("EPIs Vazios", "Não há EPIs cadastrados para buscar.", parent=self.epis_tab)
             return
         # Allow searching by CA or Description, return the identifier used
         dialog = LookupDialog(self.root, "Buscar EPI", df_epis, ["CA", "DESCRICAO"], "CA") # Preference return CA? Or Description? Or a combined key? Let's return CA if exists else description? Or maybe always description? For now, return CA as primary key. Modify display logic in LookupDialog if needed.
         result_id = dialog.result

         if result_id:
              # User selected an item. Find it again to ensure we use its standard representation
              found_epi = df_epis[df_epis["CA"] == result_id]
              display_id = result_id # Default to CA if found by CA

              if found_epi.empty: # Maybe it was selected by Description
                  found_epi = df_epis[df_epis["DESCRICAO"] == result_id] # Should maybe return desc from dialog if chosen by desc?
                  if not found_epi.empty:
                      # Decide what to display: CA or Desc? Let's favor Desc for clarity if CA is blank.
                      display_id = found_epi.iloc[0]['CA'] if found_epi.iloc[0]['CA'] else found_epi.iloc[0]['DESCRICAO']


              self.retirar_epi_id_entry.delete(0, tk.END)
              self.retirar_epi_id_entry.insert(0, display_id) # Use consistent ID
              self.retirar_epi_qtd_entry.focus_set()


    # --- Edit / Delete Functionality ---

    def _get_selected_data(self, table):
        """
        Gets data for the single selected row in the specified pandastable.
        Uses the visual row number to get the corresponding index label from the current model's DataFrame.
        Returns a pandas Series or None if selection is invalid or data cannot be retrieved.
        """
        model = table.model
        # Ensure the model and its dataframe attribute exist
        if not hasattr(model, 'df') or model.df is None:
            # This case might occur if the table hasn't been populated yet
            # or if something went wrong during model update.
            # print("Debug: Model or model.df not found in _get_selected_data")
            return None

        # getSelectedRow() returns the 0-based VISUAL row number in the table
        row_num = table.getSelectedRow()

        if row_num < 0:
            # No row is visually selected, or an error occurred during selection reporting
            # print(f"Debug: No row selected (row_num={row_num})")
            return None

        # Get the DataFrame currently being displayed by the table model
        current_df = model.df

        # *** CRITICAL CHECK ***
        # Verify that the visually selected row number is a valid positional index
        # for the *current* DataFrame being displayed.
        if row_num >= len(current_df):
            # This condition means the visual selection is pointing outside the bounds
            # of the actual data currently in the table model. This can happen
            # during rapid updates, filtering conflicts, or state inconsistencies.
            print(f"Debug: Error - selected visual row {row_num} is out of bounds for current DataFrame length {len(current_df)}")
            # Optionally show a user message:
            # messagebox.showwarning("Seleção Inválida", f"A linha selecionada ({row_num}) parece inválida. A tabela pode ter sido atualizada. Tente selecionar novamente.", parent=self.root)
            # table.clearSelection() # Optionally clear the invalid selection
            return None # Indicate that valid data couldn't be retrieved

        try:
            # If the visual row number is valid for the current DataFrame's length,
            # access the DataFrame's index using that positional number.
            # This gives us the actual index LABEL (could be integer, string, etc.)
            index_label = current_df.index[row_num]

            # Now, retrieve the row data using the reliable index LABEL via .loc
            selected_series = current_df.loc[index_label]

            # print(f"Debug: Successfully retrieved data for visual row {row_num}, index label {index_label}")
            return selected_series # Return the pandas Series

        except IndexError:
            # This IndexError might occur if, despite the length check,
            # the index itself is somehow invalid or inconsistent at the time of access.
            # The length check (`row_num >= len(current_df)`) should prevent most of these.
            print(f"Debug: IndexError trying to access index label at position {row_num} in current DataFrame with length {len(current_df)}.")
            # messagebox.showerror("Erro de Seleção", f"Não foi possível encontrar o índice para a linha selecionada ({row_num}). A tabela pode ter sido atualizada.", parent=self.root)
            return None
        except Exception as e:
            # Catch any other unexpected errors during data retrieval
            print(f"Debug: Unexpected error in _get_selected_data for visual row {row_num}: {e}")
            messagebox.showerror("Erro Interno", f"Ocorreu um erro ao obter dados da linha selecionada.\n\nDetalhes: {e}", parent=self.root)
            return None


    def _edit_selected_item(self):
         """Handles editing the selected item from the main table view."""
         if not self.pandas_table or self.active_table_name != "estoque":
              messagebox.showwarning("Ação Inválida", "A edição só está disponível para a tabela de Estoque.", parent=self.estoque_tab)
              return

         selected_item_series = self._get_selected_data(self.pandas_table)

         if selected_item_series is None:
              messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único item do estoque para editar.", parent=self.estoque_tab)
              return

         item_dict = selected_item_series.to_dict()
         # Ensure Codigo is present
         if 'CODIGO' not in item_dict:
             messagebox.showerror("Erro Interno", "Não foi possível obter o código do item selecionado.", parent=self.estoque_tab)
             return
         codigo_to_edit = str(item_dict['CODIGO']) # Get code before opening dialog

         # Show Edit Dialog
         dialog = EditProductDialog(self.root, f"Editar Produto: {codigo_to_edit}", item_dict)
         updated_data = dialog.updated_data # This blocks

         if updated_data: # User clicked Save and validation passed
              # Save changes back to the main DataFrame (in memory first)
              df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
              df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)

              idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_edit].tolist()
              if not idx:
                    messagebox.showerror("Erro", f"Item {codigo_to_edit} não encontrado no arquivo para salvar após edição.", parent=self.estoque_tab)
                    return
              idx = idx[0]

              try:
                    # Update fields from dialog result
                    for key, value in updated_data.items():
                         if key in df_estoque.columns:
                              # Handle potential type conversion issues before assignment
                              current_dtype = df_estoque[key].dtype
                              try:
                                   if pd.api.types.is_numeric_dtype(current_dtype):
                                        value = pd.to_numeric(value)
                                   # Add other type checks if necessary (e.g., dates)
                              except (ValueError, TypeError):
                                    messagebox.showwarning("Aviso de Tipo", f"Não foi possível converter '{value}' para o tipo da coluna '{key}'. Mantendo valor original.", parent=self.estoque_tab)
                                    continue # Skip updating this field if conversion fails
                              df_estoque.loc[idx, key] = value

                    # Update timestamp
                    df_estoque.loc[idx, "DATA"] = datetime.now().strftime("%H:%M %d/%m/%Y")

                    # Save updated DataFrame to CSV
                    if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                        messagebox.showinfo("Sucesso", f"Produto {codigo_to_edit} atualizado.", parent=self.estoque_tab)
                        self._atualizar_tabela_atual() # Refresh view
                    # else: Error message shown by _safe_write_csv

              except Exception as e:
                  self._update_status(f"Erro ao aplicar edições para {codigo_to_edit}: {e}", error=True)
                  messagebox.showerror("Erro ao Salvar Edição", f"Não foi possível salvar as alterações para o produto {codigo_to_edit}.\n\nDetalhes: {e}", parent=self.estoque_tab)


    def _delete_selected_item(self):
         """Handles deleting the selected item from the Estoque table."""
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
              # Load the data, find index, drop row, save
              df_estoque = self._safe_read_csv(ARQUIVOS["estoque"])
              df_estoque['CODIGO'] = df_estoque['CODIGO'].astype(str)
              idx = df_estoque.index[df_estoque['CODIGO'] == codigo_to_delete].tolist()

              if not idx:
                  messagebox.showerror("Erro", f"Item {codigo_to_delete} não encontrado no arquivo para excluir.", parent=self.estoque_tab)
                  return

              df_estoque.drop(idx, inplace=True) # Drop rows by index list

              # Save updated DataFrame to CSV
              if self._safe_write_csv(df_estoque, ARQUIVOS["estoque"]):
                   messagebox.showinfo("Sucesso", f"Produto {codigo_to_delete} - {desc_to_delete} excluído.", parent=self.estoque_tab)
                   self._atualizar_tabela_atual() # Refresh view
              # else: Error message shown by _safe_write_csv


    def _edit_selected_epi(self):
        """Handles editing the selected EPI."""
        selected_epi_series = self._get_selected_data(self.epis_table)

        if selected_epi_series is None:
             messagebox.showwarning("Nenhuma Seleção", "Por favor, selecione um único EPI para editar.", parent=self.epis_tab)
             return

        epi_dict = selected_epi_series.to_dict()
        # We need a unique identifier to find it later. Use CA if available, otherwise Description.
        # This assumes CA or Description is unique enough. Ideally, have a hidden unique ID.
        ca_to_edit = epi_dict.get('CA',"").strip()
        desc_to_edit = epi_dict.get('DESCRICAO',"").strip()
        if not (ca_to_edit or desc_to_edit):
             messagebox.showerror("Erro Interno", "EPI selecionado não tem CA ou Descrição para identificar.", parent=self.epis_tab)
             return

        dialog = EditEPIDialog(self.root, f"Editar EPI: {desc_to_edit if desc_to_edit else ca_to_edit}", epi_dict)
        updated_data = dialog.updated_data

        if updated_data:
             df_epis = self._safe_read_csv(ARQUIVOS["epis"])
             # Clean types before matching/updating
             df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
             df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()
             df_epis["QUANTIDADE"] = pd.to_numeric(df_epis["QUANTIDADE"], errors='coerce').fillna(0)


             # Find the original EPI row again based on CA primarily, then description
             match = df_epis[df_epis["CA"] == ca_to_edit] if ca_to_edit else pd.DataFrame()
             if match.empty and desc_to_edit:
                 match = df_epis[(df_epis["DESCRICAO"] == desc_to_edit) & (df_epis["CA"] == ca_to_edit)] # Be precise

             if match.empty:
                  messagebox.showerror("Erro", f"EPI original (CA:{ca_to_edit}/Desc:{desc_to_edit}) não encontrado no arquivo para salvar.", parent=self.epis_tab)
                  return

             idx = match.index[0]

             try:
                    # Check for potential CA/Description conflicts before updating
                    new_ca = updated_data['CA'].strip().upper()
                    new_desc = updated_data['DESCRICAO'].strip().upper()

                    # Check if the *new* CA exists elsewhere (excluding current row)
                    if new_ca and any((df_epis['CA'] == new_ca) & (df_epis.index != idx)):
                        messagebox.showerror("Erro de Duplicidade", f"O CA '{new_ca}' já existe para outro EPI.", parent=self.epis_tab)
                        return
                    # Check if the *new* Description exists elsewhere (excluding current row) AND the CA is different or blank
                    desc_conflict = df_epis[(df_epis['DESCRICAO'] == new_desc) & (df_epis.index != idx) & (df_epis['CA'] != new_ca if new_ca else True)]
                    if not desc_conflict.empty:
                        messagebox.showerror("Erro de Duplicidade", f"A Descrição '{new_desc}' já existe para outro EPI com CA diferente ou vazio.", parent=self.epis_tab)
                        return

                    # Update fields from dialog result
                    df_epis.loc[idx, "CA"] = new_ca
                    df_epis.loc[idx, "DESCRICAO"] = new_desc
                    df_epis.loc[idx, "QUANTIDADE"] = pd.to_numeric(updated_data['QUANTIDADE'])

                    # Save updated DataFrame to CSV
                    if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                        messagebox.showinfo("Sucesso", f"EPI '{new_desc}' atualizado.", parent=self.epis_tab)
                        self._atualizar_tabela_epis() # Refresh view
                    # else: Error handled

             except Exception as e:
                 self._update_status(f"Erro ao aplicar edições EPI: {e}", error=True)
                 messagebox.showerror("Erro ao Salvar Edição EPI", f"Não foi possível salvar as alterações.\n\nDetalhes: {e}", parent=self.epis_tab)


    def _delete_selected_epi(self):
        """Handles deleting the selected EPI."""
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
            df_epis["CA"] = df_epis["CA"].astype(str).fillna("").str.strip().str.upper()
            df_epis["DESCRICAO"] = df_epis["DESCRICAO"].astype(str).fillna("").str.strip().str.upper()

            # Find based on CA primarily, then description as fallback ID
            match = df_epis[df_epis["CA"] == ca_to_delete] if ca_to_delete else pd.DataFrame()
            if match.empty and desc_to_delete:
                  # Match description only if CA was originally blank too
                 match = df_epis[(df_epis["DESCRICAO"] == desc_to_delete) & (df_epis["CA"] == ca_to_delete)] # Precise match

            if match.empty:
                 messagebox.showerror("Erro", f"EPI original (CA:{ca_to_delete}/Desc:{desc_to_delete}) não encontrado no arquivo para excluir.", parent=self.epis_tab)
                 return

            idx = match.index.tolist() # Get all matching indices
            df_epis.drop(idx, inplace=True)

            if self._safe_write_csv(df_epis, ARQUIVOS["epis"]):
                 messagebox.showinfo("Sucesso", f"EPI '{desc_to_delete}' excluído.", parent=self.epis_tab)
                 self._atualizar_tabela_epis() # Refresh view
            # else: Error handled


    # --- Backup and Export ---

    def _criar_backup_periodico(self):
        """Creates timestamped backups of data files."""
        arquivo_ultimo_backup = os.path.join(BACKUP_DIR, "ultimo_backup_timestamp.txt")
        backup_interval_seconds = 3 * 60 * 60 # 3 hours

        # --- Cleanup Old Backups (More robustly) ---
        try:
            now = time.time()
            cutoff_time = now - (3 * 24 * 60 * 60) # 3 days ago
            for filename in os.listdir(BACKUP_DIR):
                if filename.endswith(".csv") and "_backup_" in filename: # Target specific backups
                    file_path = os.path.join(BACKUP_DIR, filename)
                    try:
                        file_mod_time = os.path.getmtime(file_path)
                        if file_mod_time < cutoff_time:
                            os.remove(file_path)
                            print(f"Backup antigo removido: {filename}")
                    except OSError as e:
                         print(f"Erro ao processar/remover backup antigo {filename}: {e}") # Log error but continue
        except Exception as e:
            self._update_status(f"Erro ao limpar backups antigos: {e}", error=True) # Log general cleanup error


        # --- Create New Backup if Interval Elapsed ---
        perform_backup = True
        if os.path.exists(arquivo_ultimo_backup):
            try:
                with open(arquivo_ultimo_backup, "r", encoding="utf-8") as f:
                    ultimo_backup_time = float(f.read().strip())
                if (time.time() - ultimo_backup_time) < backup_interval_seconds:
                    perform_backup = False
            except (ValueError, FileNotFoundError) as e:
                print(f"Erro ao ler timestamp do último backup: {e}. Criando novo backup.")
                perform_backup = True # Force backup if timestamp file is bad

        if perform_backup:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            try:
                backup_count = 0
                for nome, arquivo_origem in ARQUIVOS.items():
                    if os.path.exists(arquivo_origem):
                        nome_backup = f"{nome}_{timestamp}_auto.csv"
                        caminho_backup = os.path.join(BACKUP_DIR, nome_backup)
                        shutil.copy2(arquivo_origem, caminho_backup) # copy2 preserves metadata
                        backup_count += 1

                # Only update timestamp if backup was successful
                if backup_count > 0:
                    with open(arquivo_ultimo_backup, "w", encoding="utf-8") as f:
                        f.write(str(time.time()))
                    self._update_status(f"Backup automático criado ({backup_count} arquivos).")
                    print(f"Backup automático criado em {timestamp}")
                else:
                     print("Nenhum arquivo de dados encontrado para backup.")

            except Exception as e:
                self._update_status(f"Erro durante backup automático: {e}", error=True)
                messagebox.showerror("Erro de Backup", f"Falha ao criar backup automático:\n{e}", parent=self.root) # Show error if GUI is up

    def _schedule_backup(self):
        """Calls backup function periodically."""
        self._criar_backup_periodico()
        self.root.after(10800000, self._schedule_backup) # Reschedule


    def _exportar_conteudo(self):
        """Exports current data to Excel and generates txt report for low stock."""
        pasta_saida = "Relatorios"
        os.makedirs(pasta_saida, exist_ok=True)

        data_atual = datetime.now().strftime("%d-%m-%Y_%H%M")
        caminho_excel = os.path.join(pasta_saida, f"Relatorio_Almoxarifado_{data_atual}.xlsx")
        caminho_txt = os.path.join(pasta_saida, f"Produtos_Esgotados_{data_atual}.txt")

        try:
            with pd.ExcelWriter(caminho_excel) as writer:
                # Include EPIs in export
                all_files = {**ARQUIVOS}
                self._update_status("Iniciando exportação para Excel...")
                sheet_count = 0
                for nome, arquivo in all_files.items():
                     try:
                         df_export = self._safe_read_csv(arquivo) # Read fresh data
                         if not df_export.empty:
                             df_export.to_excel(writer, sheet_name=nome.capitalize(), index=False)
                             sheet_count +=1
                         else:
                             print(f"Planilha '{nome}' vazia ou não encontrada, não será incluída no Excel.")
                     except Exception as e:
                          print(f"Erro ao processar {arquivo} para Excel: {e}")
                          messagebox.showwarning("Aviso de Exportação", f"Erro ao incluir '{nome}' no Excel:\n{e}", parent=self.root)

            self._update_status(f"Exportação Excel concluída ({sheet_count} planilhas). Gerando relatório de esgotados...")

            # Produtos Esgotados Report
            df_estoque_report = self._safe_read_csv(ARQUIVOS["estoque"])
            if not df_estoque_report.empty and "QUANTIDADE" in df_estoque_report.columns and "CODIGO" in df_estoque_report.columns and "DESCRICAO" in df_estoque_report.columns:
                 # Ensure quantidade is numeric before filtering
                 df_estoque_report["QUANTIDADE"] = pd.to_numeric(df_estoque_report["QUANTIDADE"], errors='coerce').fillna(0)
                 produtos_esgotados = df_estoque_report[df_estoque_report["QUANTIDADE"] <= 0] # Changed to <= 0

                 with open(caminho_txt, "w", encoding="utf-8") as f:
                     f.write(f"Relatório de Produtos Esgotados/Zerados - {data_atual}\n")
                     f.write("=" * 50 + "\n")
                     if not produtos_esgotados.empty:
                         for _, row in produtos_esgotados.iterrows():
                              f.write(f"Código: {row['CODIGO']} | Descrição: {row['DESCRICAO']} | Qtd: {row['QUANTIDADE']}\n")
                     else:
                         f.write("Nenhum produto com quantidade zero ou negativa encontrado.\n")
                     f.write("=" * 50 + "\n")

                 messagebox.showinfo("Sucesso", f"Relatórios exportados com sucesso!\n\n"
                                                f"Excel: {caminho_excel}\n"
                                                f"Esgotados: {caminho_txt}", parent=self.root)
                 self._update_status("Relatórios exportados com sucesso.")

            else:
                  messagebox.showwarning("Aviso", "Não foi possível gerar relatório de produtos esgotados (dados do estoque incompletos ou não encontrados).", parent=self.root)
                  self._update_status("Relatório de esgotados não gerado (erro nos dados de estoque).")

        except Exception as e:
            self._update_status(f"Erro geral ao exportar relatórios: {e}", error=True)
            messagebox.showerror("Erro de Exportação", f"Falha ao exportar relatórios:\n{e}", parent=self.root)


    # --- Application Closing ---

    def _on_close(self):
        """Handles window close event."""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente fechar o aplicativo?", icon='warning'):
            print("Fechando aplicativo...")
            self.root.destroy()


# --- Login Window Class ---

class LoginWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Almoxarifado - Login")
        self.root.geometry("300x250")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_login) # Handle closing login window

        self.logged_in_user_id = None # Store the ID on successful login

        # Style
        style = ttk.Style()
        style.theme_use('clam')

        # Frame
        login_frame = ttk.Frame(root, padding="20")
        login_frame.pack(expand=True, fill="both")

        ttk.Label(login_frame, text="Login", font="-size 16 -weight bold").pack(pady=(0, 15))

        ttk.Label(login_frame, text="Usuário:", font="-size 12").pack(pady=(5, 0))
        self.usuario_entry = ttk.Entry(login_frame, width=30)
        self.usuario_entry.config(font="-size 12")
        self.usuario_entry.pack(pady=(0, 10))
        self.usuario_entry.bind("<Return>", lambda e: self.senha_entry.focus_set()) # Focus password on Enter

        ttk.Label(login_frame, text="Senha:", font="-size 12").pack(pady=(5, 0))
        self.senha_entry = ttk.Entry(login_frame, show="*", width=30)
        self.senha_entry.config(font="-size 12")   
        self.senha_entry.pack(pady=(0, 15))
        self.senha_entry.bind("<Return>", lambda e: self._validate_login()) # Validate on Enter

        # Use Accent style for login button if available
        style.configure("Accent.TButton", foreground="white", background="#0078D7") # Example accent color
        login_button = ttk.Button(login_frame, text="Entrar", command=self._validate_login, style="Accent.TButton")
        login_button.config(width=25, padding=5)
        login_button.pack(pady=5)

        # Center the window
        self.root.update_idletasks() # Ensure geometry is updated
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f'+{x}+{y}')

        self.usuario_entry.focus_set()


    def _validate_login(self):
        """Validates user credentials."""
        usuario = self.usuario_entry.get().strip()
        senha = self.senha_entry.get() # Get password exactly as entered

        if not usuario or not senha:
            messagebox.showwarning("Login Inválido", "Por favor, preencha usuário e senha.", parent=self.root)
            return

        # WARNING: Plain text comparison! Insecure! Use hashing (e.g., bcrypt).
        if usuario in usuarios and usuarios[usuario]["senha"] == senha:
            self.logged_in_user_id = usuarios[usuario]["id"]
            messagebox.showinfo("Sucesso", f"Login bem-sucedido!\nOperador ID: {self.logged_in_user_id}", parent=self.root)
            self.root.destroy() # Close login window
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos!", parent=self.root)
            self.senha_entry.delete(0, tk.END) # Clear password field on failure


    def _on_close_login(self):
        """Handles closing the login window directly."""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do aplicativo?", icon='question', parent=self.root):
             self.logged_in_user_id = None # Ensure no user ID is set
             self.root.destroy() # Close window
             # sys.exit() or os._exit(0) can be used here if needed, but destroying root usually suffices

    def get_user_id(self):
         return self.logged_in_user_id


# --- Main Execution ---

if __name__ == "__main__":
    # 1. Show Login Window
    login_root = tk.Tk()
    login_app = LoginWindow(login_root)
    login_root.mainloop()

    # 2. Proceed only if login was successful
    user_id = login_app.get_user_id()
    if user_id:
        # 3. Launch Main Application Window
        main_root = tk.Tk()
        app = AlmoxarifadoApp(main_root, user_id)
        main_root.mainloop()
    else:
        print("Login cancelado ou falhou. Saindo.")
