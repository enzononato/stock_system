import os
import csv
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

# Importa de nossos próprios arquivos
from inventory_manager_db import InventoryDBManager
from utils import format_cpf, format_date
from config import (
    REVENDAS_OPTIONS, CENTER_COST_OPTIONS, ADMIN_PASS, ADMIN_USER
)

# --- Paleta de Cores e Fontes (você pode alterar aqui para mudar o app inteiro) ---
PRIMARY_COLOR = "#0078D4"
PRIMARY_COLOR_HOVER = "#005a9e"
SECONDARY_COLOR = "#E1E1E1"
SECONDARY_COLOR_HOVER = "#CCCCCC"
DANGER_COLOR = "#D32F2F"
DANGER_COLOR_HOVER = "#C62828"
SUCCESS_COLOR = "#388E3C"
INFO_COLOR = "#1976D2"

BG_COLOR = "#F5F5F5"
FRAME_BG_COLOR = "#FFFFFF"
TEXT_COLOR = "#333333"
LABEL_COLOR = "#555555"

FONT_FAMILY = "Segoe UI"
FONT_NORMAL = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_TITLE = (FONT_FAMILY, 12, "bold")
FONT_TREEVIEW_HEADER = (FONT_FAMILY, 10, "bold")
FONT_TREEVIEW_ROW = (FONT_FAMILY, 10)

class App(tk.Tk):
    def __init__(self, user, role):
        super().__init__()
        self.title("Gestão de Estoque de Celulares")
        self.geometry("1280x800")
        self.configure(background=BG_COLOR)
        
        self.inv = InventoryDBManager()
        self.logged_user = user
        self.role = role

        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        """Configura todos os estilos customizados para a aplicação."""
        style = ttk.Style(self)
        style.theme_use('clam')

        # Estilo Base
        style.configure(".", font=FONT_NORMAL, background=BG_COLOR, foreground=TEXT_COLOR)
        style.configure("TFrame", background=FRAME_BG_COLOR)
        
        # Estilo Notebook (Abas) - BORDA ADICIONADA
        style.configure("TNotebook", background=BG_COLOR, borderwidth=0)
        style.configure("TNotebook.Tab",
            font=FONT_BOLD,
            padding=[12, 6],
            background=SECONDARY_COLOR,
            foreground=LABEL_COLOR,
            borderwidth=1,
            relief="raised") # Alterado para "raised" para borda mais visível
        style.map("TNotebook.Tab",
            background=[("selected", FRAME_BG_COLOR), ("active", PRIMARY_COLOR_HOVER)],
            foreground=[("selected", PRIMARY_COLOR), ("active", "white")])

        # Estilo Treeview (Tabelas)
        style.configure("Treeview",
            rowheight=28,
            font=FONT_TREEVIEW_ROW,
            background=FRAME_BG_COLOR,
            fieldbackground=FRAME_BG_COLOR)
        style.configure("Treeview.Heading",
            font=FONT_TREEVIEW_HEADER,
            padding=(8, 8),
            background=PRIMARY_COLOR,
            foreground="white",
            borderwidth=0)
        style.map("Treeview.Heading", background=[('active', PRIMARY_COLOR_HOVER)])
        
        # Estilos de Botões - BORDA AJUSTADA
        button_padding = (10, 5)
        # O relevo "raised" cria um efeito 3D que destaca a borda.
        # Outras opções: "solid", "sunken", "groove", "ridge"
        style.configure("TButton", font=FONT_BOLD, padding=button_padding, borderwidth=1, relief="raised")
        style.configure("Primary.TButton", background=PRIMARY_COLOR, foreground="white")
        style.map("Primary.TButton", background=[('active', PRIMARY_COLOR_HOVER)], relief=[('pressed', 'sunken')])
        style.configure("Secondary.TButton", background=SECONDARY_COLOR, foreground=TEXT_COLOR)
        style.map("Secondary.TButton", background=[('active', SECONDARY_COLOR_HOVER)], relief=[('pressed', 'sunken')])
        style.configure("Danger.TButton", background=DANGER_COLOR, foreground="white")
        style.map("Danger.TButton", background=[('active', DANGER_COLOR_HOVER)], relief=[('pressed', 'sunken')])

        # Estilos de Labels, Entries, Comboboxes - BORDAS ADICIONADAS
        style.configure("TLabel", font=FONT_NORMAL, background=FRAME_BG_COLOR, foreground=LABEL_COLOR)
        style.configure("Title.TLabel", font=FONT_TITLE, foreground=TEXT_COLOR)
        style.configure("TEntry", font=FONT_NORMAL, padding=5, borderwidth=1, relief='solid')
        style.configure("TCombobox", font=FONT_NORMAL, padding=5, borderwidth=1, relief='solid')
        style.configure("Success.TLabel", foreground=SUCCESS_COLOR, background=FRAME_BG_COLOR, font=FONT_BOLD)
        style.configure("Danger.TLabel", foreground=DANGER_COLOR, background=FRAME_BG_COLOR, font=FONT_BOLD)

        # Estilos de Entry e Combobox para validação
        style.configure("Error.TEntry", fieldbackground="#FFA2AC", foreground=DANGER_COLOR, borderwidth=1, relief='solid')
        style.configure("Error.TCombobox", fieldbackground="#FFCDD2", borderwidth=1, relief='solid')

        # Para o TCombobox, é prciso configurar o estilo da lista também
        style.map('Error.TCombobox', fieldbackground=[('readonly', '#FFA2AC')])


    # ==========================================================
    # Construção das abas
    # ==========================================================
    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_stock   = ttk.Frame(notebook, padding=15)
        self.tab_add     = ttk.Frame(notebook, padding=15)
        self.tab_edit    = ttk.Frame(notebook, padding=15)
        self.tab_issue   = ttk.Frame(notebook, padding=15)
        self.tab_return  = ttk.Frame(notebook, padding=15)
        self.tab_remove  = ttk.Frame(notebook, padding=15)
        self.tab_history = ttk.Frame(notebook, padding=15)
        self.tab_report  = ttk.Frame(notebook, padding=15)
        self.tab_terms   = ttk.Frame(notebook, padding=15)
        self.tab_graphs  = ttk.Frame(notebook, padding=15)

        notebook.add(self.tab_stock,  text="Estoque")
        notebook.add(self.tab_add,    text="Cadastrar")
        notebook.add(self.tab_edit,   text="Editar")
        notebook.add(self.tab_issue,  text="Emprestar")
        notebook.add(self.tab_return, text="Devolver")
        notebook.add(self.tab_remove, text="Remover")
        notebook.add(self.tab_history,text="Histórico")
        notebook.add(self.tab_report, text="Relatório")
        notebook.add(self.tab_terms,  text="Termos")
        notebook.add(self.tab_graphs, text="Gráficos")

        self.build_stock_tab()
        self.build_add_tab()
        self.build_edit_tab()
        self.build_issue_tab()
        self.build_return_tab()
        self.build_remove_tab()
        self.build_history_tab()
        self.build_report_tab()
        self.build_terms_tab()
        self.build_graph_tab()
        
        # Restrições por role
        if self.role == "Jovem Aprendiz":
            notebook.hide(self.tab_remove)
            notebook.hide(self.tab_history)
        elif self.role == "Técnico":
            notebook.hide(self.tab_remove)

        self.update_all_views()

    # ==========================================================
    # Helper: ordenar colunas da treeview
    # ==========================================================
    def treeview_sort_column(self, tv, col, reverse):
        items = [(tv.set(k, col), k) for k in tv.get_children("")]
        
        try:
            # Tenta ordenar como número, removendo pontos e traços
            items.sort(key=lambda t: int(str(t[0]).replace('.', '').replace('-', '').replace('/', '')), reverse=reverse)
        except (ValueError, IndexError):
            # Se falhar (ex: texto), ordena como string
            items.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)


        for index, (val, k) in enumerate(items):
            tv.move(k, "", index)
            
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))
        
    # ==========================================================
    # Aba de Estoque
    # ==========================================================
    def build_stock_tab(self):
        tab = self.tab_stock

        frm_filters = ttk.Frame(tab, padding=(0, 0, 0, 10))
        frm_filters.pack(fill="x")

        ttk.Label(frm_filters, text="Status:").grid(row=0, column=0, padx=(0, 5), pady=5, sticky="e")
        self.cb_status_filter = ttk.Combobox(frm_filters, values=["", "Disponível", "Indisponível"], state="readonly", width=15)
        self.cb_status_filter.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frm_filters, text="Tipo:").grid(row=0, column=2, padx=(15, 5), pady=5, sticky="e")
        self.cb_tipo_filter = ttk.Combobox(frm_filters, values=["", "Celular", "Notebook", "Desktop", "Impressora", "Tablet"], state="readonly", width=15)
        self.cb_tipo_filter.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frm_filters, text="Revenda:").grid(row=0, column=4, padx=(15, 5), pady=5, sticky="e")
        self.cb_revenda_filter = ttk.Combobox(frm_filters, values=[""] + REVENDAS_OPTIONS, state="readonly", width=20)
        self.cb_revenda_filter.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frm_filters, text="Buscar:").grid(row=0, column=6, padx=(15, 5), pady=5, sticky="e")
        self.e_search_stock = ttk.Entry(frm_filters, width=25)
        self.e_search_stock.grid(row=0, column=7, padx=5, pady=5, sticky="ew")

        ttk.Button(frm_filters, text="Aplicar", command=self.update_stock, style="Primary.TButton").grid(row=0, column=8, padx=15, pady=5)
        ttk.Button(frm_filters, text="Limpar", command=self.clear_filters, style="Secondary.TButton").grid(row=0, column=9, padx=5, pady=5)
        
        frm_filters.grid_columnconfigure(10, weight=1)
        ttk.Button(frm_filters, text="Exportar CSV", command=lambda: self.exportar_csv(self.tree_stock, "Exportar Estoque", "estoque"), style="Secondary.TButton").grid(row=0, column=10, padx=5, pady=5, sticky="e")

        cols = [
            "ID", "Revenda", "Tipo", "Marca", "Modelo", "Status", "Usuário", "CPF",
            "Identificador", "Domínio", "Host", "Endereço Físico", "CPU",
            "RAM", "Storage", "Sistema", "Licença", "AnyDesk",
            "Setor", "IP", "MAC", "Data Cadastro"
        ]

        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill="both", expand=True)

        self.tree_stock = ttk.Treeview(tree_frame, columns=cols, show="headings")
        self.tree_stock.pack(side="left", fill="both", expand=True)

        col_widths = {
            "ID": 40, "Revenda": 130, "Tipo": 100, "Marca": 100, "Modelo": 120, "Status": 100,
            "Usuário": 140, "CPF": 110, "Identificador": 140
        }

        for col in cols:
            self.tree_stock.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree_stock, c, False))
            self.tree_stock.column(col, width=col_widths.get(col, 120), anchor="w", stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_stock.yview)
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tab, orient="horizontal", command=self.tree_stock.xview)
        hsb.pack(side="bottom", fill="x", pady=(5,0))
        self.tree_stock.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)


    def clear_filters(self):
        self.cb_status_filter.set("")
        self.cb_tipo_filter.set("")
        self.cb_revenda_filter.set("")
        self.e_search_stock.delete(0, "end")
        self.update_stock()


    def build_add_tab(self):
        # Frame container principal para centralizar o conteúdo
        container = ttk.Frame(self.tab_add)
        container.pack(anchor="n", pady=20)

        frm = ttk.Frame(container)
        frm.pack()

        # Configura colunas para alinhamento
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="Tipo de Equipamento:", font=FONT_BOLD).grid(row=0, column=0, sticky="e", padx=5, pady=(0, 10))
        self.cb_tipo = ttk.Combobox(frm, values=["Celular", "Notebook", "Desktop", "Impressora", "Tablet"], state="readonly")
        self.cb_tipo.grid(row=0, column=1, sticky="ew", padx=5, pady=(0, 10))
        self.cb_tipo.bind("<<ComboboxSelected>>", self.on_tipo_selected)

        ttk.Separator(frm, orient="horizontal").grid(row=1, column=0, columnspan=2, sticky="ew", pady=15)

        self.frm_dynamic = ttk.Frame(frm)
        self.frm_dynamic.grid(row=2, column=0, columnspan=2, sticky="ew")
        # Configura a coluna do frame dinâmico também
        self.frm_dynamic.grid_columnconfigure(1, weight=1)

        self.lbl_add = ttk.Label(frm, text="", anchor="center")
        self.lbl_add.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Button(frm, text="Cadastrar", command=self.cmd_add, style="Primary.TButton").grid(row=4, column=0, columnspan=2, pady=10)


    def on_tipo_selected(self, event=None):
        for widget in self.frm_dynamic.winfo_children():
            widget.destroy()

        tipo = self.cb_tipo.get()
        if tipo:
            self.add_widgets = self._create_dynamic_form(self.frm_dynamic, tipo)

    def build_edit_tab(self):
        self.current_edit_id = None
        self.edit_widgets = {}

        container = ttk.Frame(self.tab_edit)
        container.pack(anchor="n", pady=20, fill="x")
        
        frm = ttk.Frame(container)
        frm.pack(fill="x")

        frm_select = ttk.Frame(frm)
        frm_select.pack(fill="x", pady=5)
        frm_select.grid_columnconfigure(1, weight=1)
        
        ttk.Label(frm_select, text="Selecione o item para editar:", font=FONT_BOLD).grid(row=0, column=0, padx=(0, 10))
        self.cb_edit = ttk.Combobox(frm_select, state="readonly")
        self.cb_edit.grid(row=0, column=1, sticky="ew")
        
        ttk.Button(frm_select, text="Carregar Dados", command=self.cmd_load_item_for_edit, style="Primary.TButton").grid(row=0, column=2, padx=10)

        ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=15)

        self.frm_edit_dynamic = ttk.Frame(frm)
        self.frm_edit_dynamic.pack(pady=10)
        self.frm_edit_dynamic.grid_columnconfigure(1, weight=1)

        self.lbl_edit = ttk.Label(frm, text="", anchor="center")
        self.lbl_edit.pack(pady=10, fill="x")
        
        ttk.Button(frm, text="Salvar Alterações", command=self.cmd_save_edit, style="Primary.TButton").pack(pady=10)
        
        self.update_edit_cb()

    def build_issue_tab(self):
        container = ttk.Frame(self.tab_issue)
        container.pack(anchor="n", pady=20)

        frm = ttk.Frame(container)
        frm.pack()

        # Faz a segunda coluna (dos inputs) se expandir
        frm.grid_columnconfigure(1, weight=1)
        
        fields = [
            ("Selecione aparelho:", 'cb_issue', ttk.Combobox, {'state': 'readonly'}),
            ("Funcionário:", 'e_issue_user', ttk.Entry, {}),
            ("CPF:", 'e_issue_cpf', ttk.Entry, {}),
            ("Centro de Custo:", 'cb_center', ttk.Combobox, {'state': 'readonly', 'values': CENTER_COST_OPTIONS}),
            ("Cargo:", 'e_cargo', ttk.Entry, {}),
            ("Revenda:", 'cb_revenda', ttk.Combobox, {'state': 'readonly', 'values': REVENDAS_OPTIONS}),
            ("Data Empréstimo (dd/mm/aaaa):", 'e_date_issue', ttk.Entry, {})
        ]

        for i, (label_text, attr, widget_class, opts) in enumerate(fields):
            ttk.Label(frm, text=label_text).grid(row=i, column=0, sticky='e', padx=5, pady=5)
            widget = widget_class(frm, **opts)
            widget.grid(row=i, column=1, pady=5, padx=5, sticky="ew")
            setattr(self, attr, widget)

        # Associa os eventos para limpar o erro ao interagir
        self.cb_issue.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        self.e_issue_user.bind("<KeyRelease>", self.on_widget_interaction)
        self.cb_center.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        self.e_cargo.bind("<KeyRelease>", self.on_widget_interaction)
        self.cb_revenda.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        
        self.e_issue_cpf.bind("<KeyRelease>", self.on_cpf_entry)
        self.e_date_issue.bind("<KeyRelease>", lambda e: self.on_date_entry(e, self.e_date_issue))

        self.lbl_issue = ttk.Label(frm, text="", anchor="center")
        self.lbl_issue.grid(row=len(fields), column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Button(frm, text="Emprestar", command=self.cmd_issue, style="Primary.TButton").grid(row=len(fields) + 1, column=0, columnspan=2, pady=10)


    def build_return_tab(self):
        container = ttk.Frame(self.tab_return)
        container.pack(anchor="n", pady=20)

        frm = ttk.Frame(container)
        frm.pack()
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.cb_return = ttk.Combobox(frm, state='readonly')
        self.cb_return.grid(row=0, column=1, pady=5, padx=5, sticky="ew")

        ttk.Label(frm, text="Data Devolução (dd/mm/aaaa):").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.e_date_return = ttk.Entry(frm)
        self.e_date_return.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
        self.e_date_return.bind("<KeyRelease>", lambda e: self.on_date_entry(e, self.e_date_return))

        self.lbl_ret = ttk.Label(frm, text="", anchor="center")
        self.lbl_ret.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Button(frm, text="Devolver", command=self.cmd_return, style="Primary.TButton").grid(row=3, column=0, columnspan=2, pady=10)

    def build_remove_tab(self):
        container = ttk.Frame(self.tab_remove)
        container.pack(anchor="n", pady=20)
        
        frm = ttk.Frame(container)
        frm.pack()
        frm.grid_columnconfigure(1, weight=1)
        
        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.cb_remove = ttk.Combobox(frm, state='readonly')
        self.cb_remove.grid(row=0, column=1, pady=5, padx=5, sticky="ew")

        self.lbl_rem = ttk.Label(frm, text="", anchor="center")
        self.lbl_rem.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Button(frm, text="Remover", command=self.cmd_remove, style="Danger.TButton").grid(row=2, column=0, columnspan=2, pady=10)

    def build_history_tab(self):
        tab = self.tab_history

        top_frame = ttk.Frame(tab, padding=(0,0,0,10))
        top_frame.pack(fill="x")

        ttk.Label(top_frame, text="Buscar:").pack(side="left", padx=(0, 5), pady=5)
        self.e_search_history = ttk.Entry(top_frame, width=30)
        self.e_search_history.pack(side="left", pady=5)

        ttk.Button(top_frame, text="Buscar", command=self.update_history_table, style="Primary.TButton").pack(side="left", padx=10, pady=5)
        ttk.Button(top_frame, text="Limpar", command=self.clear_history_filter, style="Secondary.TButton").pack(side="left", padx=5, pady=5)

        ttk.Button(top_frame, text="Exportar CSV",
                command=lambda: self.exportar_csv(self.tree_history, "Exportar Histórico", "histórico"),
                style="Secondary.TButton").pack(side="right", padx=5, pady=5)

        cols = (
            "ID Item", "Operador", "Operação", "Data", "Tipo", "Marca", "Modelo", "Identificador",
            "Usuário", "CPF", "Cargo", "Centro de Custo", "Revenda"
        )
        
        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill="both", expand=True)
        
        self.tree_history = ttk.Treeview(tree_frame, columns=cols, show="headings")

        col_widths = { "ID Item": 60, "Operador": 100, "Operação": 100, "Data": 100, "CPF": 120 }

        for c in cols:
            self.tree_history.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_history, col, False))
            self.tree_history.column(c, width=col_widths.get(c, 120), anchor="w", stretch=False)

        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_history.yview)
        xsb = ttk.Scrollbar(tab, orient="horizontal", command=self.tree_history.xview)
        self.tree_history.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.tree_history.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        xsb.pack(side="bottom", fill="x", pady=(5,0))

    def clear_history_filter(self):
        self.e_search_history.delete(0, "end")
        self.update_history_table()

    def build_report_tab(self):
        frm = ttk.Frame(self.tab_report)
        frm.pack(fill="both", expand=True)

        top_frm = ttk.Frame(frm, padding=(0,0,0,10))
        top_frm.pack(fill='x')

        ttk.Label(top_frm, text="Ano:").grid(row=0, column=0, sticky='e', padx=(0,5), pady=5)
        self.e_report_year  = ttk.Entry(top_frm, width=6)
        self.e_report_year.insert(0, datetime.now().year)
        self.e_report_year.grid(row=0, column=1, pady=5)

        ttk.Label(top_frm, text="Mês:").grid(row=0, column=2, sticky='e', padx=(10,5), pady=5)
        self.cb_report_month = ttk.Combobox(top_frm, values=list(range(1, 13)), width=4, state="readonly")
        self.cb_report_month.set(datetime.now().month)
        self.cb_report_month.grid(row=0, column=3, pady=5)

        ttk.Button(top_frm, text="Gerar Relatório", command=self.cmd_generate_report, style="Primary.TButton").grid(row=0, column=4, padx=10, pady=5)

        ttk.Label(top_frm, text="Buscar no Relatório:").grid(row=0, column=5, padx=(10, 5), pady=5)
        self.e_search_report = ttk.Entry(top_frm, width=30)
        self.e_search_report.grid(row=0, column=6, padx=5, pady=5)
        ttk.Button(top_frm, text="Limpar Busca", command=self.clear_report_filter, style="Secondary.TButton").grid(row=0, column=7, padx=5, pady=5)

        action_frm = ttk.Frame(top_frm)
        action_frm.grid(row=0, column=8, sticky="e", padx=(20,0))
        top_frm.grid_columnconfigure(8, weight=1)
        
        ttk.Button(action_frm, text="Exportar CSV", command=lambda: self.exportar_csv(self.tree_report, "Exportar Relatório", "relatório"), style="Secondary.TButton").pack(side="left")
        ttk.Button(action_frm, text="Estornar Lançamento", command=self.cmd_delete_report_entry, style="Danger.TButton").pack(side="left", padx=10)
        
        # ALTERADO: Adicionada a coluna "ID Histórico" que ficará oculta
        cols = ("ID Histórico", "ID Item", "Operador", "Tipo", "Marca", "Modelo", "Identificador", "Usuário", "CPF", "Operação", "Data Empréstimo", "Data Devolução", "Centro de Custo", "Cargo", "Revenda")
        
        tree_frame = ttk.Frame(frm)
        tree_frame.pack(fill="both", expand=True)

        self.tree_report = ttk.Treeview(tree_frame, columns=cols, show='headings')

        col_widths = {"ID Item": 60, "Operador": 100, "Operação": 110, "CPF": 120}

        for c in cols:
            self.tree_report.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_report, col, False))
            self.tree_report.column(c, width=col_widths.get(c, 120), anchor='w', stretch=False)
        
        # OCULTA a coluna ID Histórico
        self.tree_report.column("ID Histórico", width=0, stretch=False)

        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_report.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=self.tree_report.xview)
        self.tree_report.configure(yscrollcommand=ysb.set, xscrollcommand=hsb.set)
        
        self.tree_report.pack(side="left", fill='both', expand=True)
        ysb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x", pady=(5,0))

    def clear_report_filter(self):
        self.e_search_report.delete(0, "end")
        self.cmd_generate_report()

    def build_terms_tab(self):
        frm = ttk.Frame(self.tab_terms)
        frm.pack(fill="both", expand=True)

        search_frame = ttk.Frame(frm, padding=(0,0,0,10))
        search_frame.pack(fill="x")
        ttk.Label(search_frame, text="Buscar:").grid(row=0, column=0, sticky="e", padx=(0,5), pady=5)
        self.e_search_terms = ttk.Entry(search_frame, width=30)
        self.e_search_terms.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="Buscar", command=self.update_terms_table, style="Primary.TButton").grid(row=0, column=2, padx=10, pady=5)

        cols = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        
        tree_frame = ttk.Frame(frm)
        tree_frame.pack(fill="both", expand=True)

        self.tree_terms = ttk.Treeview(tree_frame, columns=cols, show="headings")
        
        col_widths = { "ID": 50, "CPF": 120, "Data Empréstimo": 120 }
        for c in cols:
            self.tree_terms.heading(c, text=c, anchor="w", command=lambda col=c: self.treeview_sort_column(self.tree_terms, col, False))
            self.tree_terms.column(c, width=col_widths.get(c, 150), anchor="w")

        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_terms.yview)
        self.tree_terms.configure(yscrollcommand=ysb.set)
        self.tree_terms.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        
        self.lbl_terms = ttk.Label(frm, text="")
        self.lbl_terms.pack(pady=10)
        ttk.Button(frm, text="Gerar Termo Selecionado", command=self.cmd_generate_term, style="Primary.TButton").pack(pady=5)

    def build_graph_tab(self):
        frm = ttk.Frame(self.tab_graphs)
        frm.pack(fill='both', expand=True, anchor="n")

        ctrl = ttk.Frame(frm)
        ctrl.pack(pady=5, anchor="center")
        ttk.Label(ctrl, text="Ano:").pack(side='left')
        self.e_graph_year = ttk.Entry(ctrl, width=6)
        self.e_graph_year.insert(0, datetime.now().year)
        self.e_graph_year.pack(side='left', padx=(5, 15))

        ttk.Label(ctrl, text="Mês:").pack(side='left')
        self.cb_graph_month = ttk.Combobox(ctrl, values=list(range(1, 13)), width=4, state='readonly')
        self.cb_graph_month.set(datetime.now().month)
        self.cb_graph_month.pack(side='left', padx=5)

        lbl = ttk.Label(frm, text="Visualizações Gráficas", style="Title.TLabel")
        lbl.pack(pady=(20, 10), anchor="center")
        
        btns = ttk.Frame(frm)
        btns.pack(anchor="center")
        ttk.Button(btns, text="Empréstimos x Devoluções (Mensal)", command=self.graph_issue_return, style="Secondary.TButton").pack(pady=5)
        ttk.Button(btns, text="Cadastros de Novos Itens (Mensal)", command=self.graph_registration, style="Secondary.TButton").pack(pady=5)

    def exportar_csv(self, tree, titulo="Exportar", nome="dados"):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                
                # Pega apenas os cabeçalhos das colunas visíveis
                visible_columns = [c for c in tree["columns"] if tree.column(c, "width") > 0]
                headers = [tree.heading(c)["text"] for c in visible_columns]
                writer.writerow(headers)

                # Pega os valores apenas das colunas visíveis
                col_indices = {col: i for i, col in enumerate(tree["columns"])}
                visible_indices = [col_indices[c] for c in visible_columns]

                for row_id in tree.get_children():
                    all_values = tree.item(row_id)["values"]
                    visible_values = [all_values[i] for i in visible_indices]
                    writer.writerow(visible_values)

            messagebox.showinfo(titulo, f"{nome.capitalize()} exportado para:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro ao Exportar", f"Ocorreu um erro: {e}")

    # --- Comandos e Ações ---

    def _create_dynamic_form(self, parent_frame, tipo, item_data=None):
        if item_data is None: item_data = {}
        
        widgets = {}
        
        # Campos comuns
        ttk.Label(parent_frame, text="Marca:").grid(row=1, column=0, sticky="e", pady=5, padx=5)
        e_brand = ttk.Entry(parent_frame)
        e_brand.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
        e_brand.insert(0, item_data.get('brand') or '')
        e_brand.bind("<KeyRelease>", self.on_widget_interaction)
        widgets['brand'] = e_brand

        ttk.Label(parent_frame, text="Revenda:").grid(row=2, column=0, sticky="e", pady=5, padx=5)
        cb_revenda = ttk.Combobox(parent_frame, values=REVENDAS_OPTIONS, state="readonly")
        cb_revenda.grid(row=2, column=1, pady=5, padx=5, sticky="ew")
        cb_revenda.set(item_data.get('revenda') or '')
        cb_revenda.bind("<KeyRelease>", self.on_widget_interaction)
        cb_revenda.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        widgets['revenda'] = cb_revenda

        # Campos específicos
        row_start = 3
        if tipo == "Celular":
            ttk.Label(parent_frame, text="Modelo:").grid(row=row_start, column=0, sticky="e", pady=5, padx=5)
            e_model = ttk.Entry(parent_frame); e_model.grid(row=row_start, column=1, pady=5, padx=5, sticky="ew")
            e_model.insert(0, item_data.get('model') or ''); widgets['model'] = e_model
            e_model.bind("<KeyRelease>", self.on_widget_interaction)
            
            ttk.Label(parent_frame, text="IMEI:").grid(row=row_start + 1, column=0, sticky="e", pady=5, padx=5)
            e_identificador = ttk.Entry(parent_frame); e_identificador.grid(row=row_start + 1, column=1, pady=5, padx=5, sticky="ew")
            e_identificador.insert(0, item_data.get('identificador') or ''); widgets['identificador'] = e_identificador
            e_identificador.bind("<KeyRelease>", self.on_widget_interaction)
            
        elif tipo in ["Notebook", "Desktop"]:
            fields = {
                'dominio': ("Domínio:", ttk.Combobox, {"values": ["Sim", "Não"], "state": "readonly"}),
                'host': ("Host:", ttk.Entry, {}), 'endereco_fisico': ("Endereço Físico:", ttk.Entry, {}),
                'storage': ("Armazenamento (GB):", ttk.Entry, {}), 'sistema': ("Sistema Operacional:", ttk.Entry, {}),
                'cpu': ("Processador:", ttk.Entry, {}), 'ram': ("Memória RAM:", ttk.Entry, {}),
                'licenca': ("Licença Windows:", ttk.Entry, {}), 'anydesk': ("AnyDesk:", ttk.Entry, {})
            }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts); widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                if isinstance(widget, ttk.Combobox): widget.set(item_data.get(key) or '')
                else: widget.insert(0, item_data.get(key) or '')
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                if isinstance(widget, ttk.Combobox):
                    widget.bind("<<ComboboxSelected>>", self.on_widget_interaction)
                widgets[key] = widget
        
        elif tipo == "Impressora":
            fields = { 'setor': ("Setor:", ttk.Entry, {}), 'ip': ("IP:", ttk.Entry, {}), 'mac': ("MAC:", ttk.Entry, {}) }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts); widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                widget.insert(0, item_data.get(key) or '')
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                widgets[key] = widget

        elif tipo == "Tablet":
            fields = { 'model': ("Modelo:", ttk.Entry,{}), 'identificador': ("Nº de Série:", ttk.Entry, {}), 'storage': ("Armazenamento (GB):", ttk.Entry, {}) }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts); widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                widget.insert(0, item_data.get(key) or '')
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                widgets[key] = widget
                
        return widgets


    def cmd_add(self):
        self.lbl_add.config(text="")
        tipo = self.cb_tipo.get().strip()
        if not tipo:
            self.lbl_add.config(text="Selecione o tipo de equipamento.", style="Danger.TLabel")
            return
        dados = {}
        for widget in self.add_widgets.values():
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)
        for key, widget in self.add_widgets.items():
            if widget.winfo_exists():
                dados[key] = widget.get().strip()
        erros = []
        if not dados.get("brand"): erros.append((self.add_widgets['brand'], "Informe a marca."))
        
        if not dados.get("revenda"): erros.append((self.add_widgets['revenda'], "Informe o campo Revenda."))
        
        if tipo in ["Celular", "Tablet"]:
            if not dados.get("model"): erros.append((self.add_widgets['model'], "Informe o modelo."))
            if not dados.get("identificador"): erros.append((self.add_widgets['identificador'], "Informe o Identificador (IMEI/Nº de Série)."))
        
        if tipo == "Tablet" and not dados.get("storage"): erros.append((self.add_widgets['storage'], "Informe o armazenamento."))
        
        if tipo == "Impressora":
            for campo, msg in {"ip": "Informe o IP.", "setor": "Informe o setor.", "mac": "Informe o MAC."}.items():
                if not dados.get(campo): erros.append((self.add_widgets[campo], msg))
        
        if tipo in ["Desktop", "Notebook"]:
            for campo, msg in {"dominio": "Informe o domínio.", "host": "Informe o host.",
                               "endereco_fisico": "Informe o endereço físico.",
                               "storage": "Informe o armazenamento.",
                               "sistema": "Informe o sistema operacional.", "cpu": "Informe o processador.",
                               "ram": "Informe a RAM.", "licenca": "Informe a licença.",
                               "anydesk": "Informe o AnyDesk."}.items():
                if not dados.get(campo): erros.append((self.add_widgets[campo], msg))
        
        if erros:
            for widget, _ in erros:
                widget_class = str(widget.winfo_class())
                if 'TCombobox' in widget_class: widget.configure(style="Error.TCombobox")
                elif 'TEntry' in widget_class: widget.configure(style="Error.TEntry")
            
            if len(erros) == 1:
                self.lbl_add.config(text=erros[0][1], style="Danger.TLabel")
            else:
                self.lbl_add.config(text="Por favor, corrija os campos destacados.", style="Danger.TLabel")
            return

        dados["tipo"] = tipo
        dados["date_registered"] = datetime.now().strftime("%Y-%m-%d")
        item_id = self.inv.add_item(dados, self.logged_user)
        
        if not item_id:
            self.lbl_add.config(text="Erro ao cadastrar item.", style="Danger.TLabel")
            return
        
        for widget in self.add_widgets.values():
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)
        self.lbl_add.config(text=f"{tipo} ID {item_id} cadastrado com sucesso!", style="Success.TLabel")
        self.cb_tipo.set("")
        
        for widget in self.frm_dynamic.winfo_children():
            widget.destroy()
        self.update_all_views()


    def cmd_load_item_for_edit(self):
        selection = self.cb_edit.get()
        if not selection:
            messagebox.showwarning("Aviso", "Por favor, selecione um item para editar.")
            return

        for widget in self.frm_edit_dynamic.winfo_children():
            widget.destroy()

        item_id = int(selection.split(' - ')[0])
        item = self.inv.find(item_id)
        if not item:
            messagebox.showerror("Erro", "Item não encontrado.")
            return

        self.current_edit_id = item_id
        self.edit_widgets = self._create_dynamic_form(self.frm_edit_dynamic, item.get('tipo'), item)

    def cmd_save_edit(self):
        if self.current_edit_id is None:
            messagebox.showerror("Erro", "Nenhum item foi carregado para edição.")
            return

        new_data = {key: widget.get().strip() for key, widget in self.edit_widgets.items() if widget.winfo_exists()}
        
        ok, msg = self.inv.update_item(self.current_edit_id, new_data, self.logged_user)
        self.lbl_edit.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        
        if ok:
            for widget in self.frm_edit_dynamic.winfo_children():
                widget.destroy()
            self.current_edit_id = None
            self.cb_edit.set('')
            self.update_all_views()

    def cmd_issue(self):
        self.lbl_issue.config(text="")
        widgets_to_reset = [self.cb_issue, self.e_issue_user, self.e_issue_cpf, self.cb_center, self.e_cargo, self.cb_revenda, self.e_date_issue]
        for widget in widgets_to_reset:
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)

        sel = self.cb_issue.get()
        user = self.e_issue_user.get().strip()
        cpf = self.e_issue_cpf.get().strip()
        cc = self.cb_center.get().strip()
        cargo = self.e_cargo.get().strip()
        revenda = self.cb_revenda.get().strip()
        date_issue = self.e_date_issue.get().strip()
        erros = []
        if not sel: erros.append((self.cb_issue, "Selecione um aparelho."))
        if not user: erros.append((self.e_issue_user, "Informe o nome do funcionário."))
        if not cc: erros.append((self.cb_center, "Selecione o centro de custo."))
        if not cargo: erros.append((self.e_cargo, "Informe o cargo."))
        if not revenda: erros.append((self.cb_revenda, "Selecione a revenda."))
        if not date_issue: erros.append((self.e_date_issue, "Informe a data do empréstimo."))
        if not cpf: erros.append((self.e_issue_cpf, "Informe o CPF."))
        elif len("".join(filter(str.isdigit, cpf))) < 11:
            erros.append((self.e_issue_cpf, "CPF inválido. Informe os 11 dígitos."))

        if erros:
            for widget, _ in erros:
                widget_class = str(widget.winfo_class())
                if 'TCombobox' in widget_class: widget.configure(style="Error.TCombobox")
                elif 'TEntry' in widget_class: widget.configure(style="Error.TEntry")
            
            if len(erros) == 1:
                # Se houver apenas um erro, exibe a mensagem específica dele.
                self.lbl_issue.config(text=erros[0][1], style="Danger.TLabel")
            else:
                # Se houver múltiplos erros, exibe a mensagem genérica.
                self.lbl_issue.config(text="Por favor, corrija os campos destacados.", style="Danger.TLabel")
            return

        pid = int(sel.split(' - ', 1)[0])
        ok, msg = self.inv.issue(pid, user, cpf, cc, cargo, revenda, date_issue, self.logged_user)
        self.lbl_issue.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        if ok:
            self.cb_issue.set('')
            self.e_issue_user.delete(0, 'end')
            self.e_issue_cpf.delete(0, 'end')
            self.cb_center.set('')
            self.e_cargo.delete(0, 'end')
            self.cb_revenda.set('')
            self.e_date_issue.delete(0, 'end')
            self.update_all_views()

    def cmd_return(self):
        sel = self.cb_return.get()
        date_return = self.e_date_return.get().strip()
        if not (sel and date_return):
            self.lbl_ret.config(text="Selecione aparelho e informe data de devolução.", style="Danger.TLabel")
            return
            
        pid = int(sel.split(' - ', 1)[0])
        ok, msg = self.inv.ret(pid, date_return, self.logged_user)
        self.lbl_ret.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        
        if ok:
            self.e_date_return.delete(0, 'end')
            self.update_all_views()

    def cmd_remove(self):
        sel = self.cb_remove.get()
        if not sel:
            self.lbl_rem.config(text="Selecione um aparelho para remover.", style="Danger.TLabel"); return

        pid = int(sel.split(' - ', 1)[0])
        item = self.inv.find(pid)
        if not item:
            self.lbl_rem.config(text="Item não encontrado.", style='Danger.TLabel'); return

        if item['status'] != 'Disponível':
            self.lbl_rem.config(text="Não é possível remover um produto emprestado.", style='Danger.TLabel')
            return

        pwd = simpledialog.askstring("Autorização", "Digite a senha de administrador:", show='*')
        if pwd != ADMIN_PASS:
            self.lbl_rem.config(text="Senha incorreta. Ação não autorizada.", style='Danger.TLabel'); return
            
        ok, msg = self.inv.remove(pid, self.logged_user)
        self.lbl_rem.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        
        if ok:
            self.update_all_views()



    def cmd_generate_report(self):
        # ALTERADO: A lógica para determinar se um empréstimo virou devolução foi movida para a query SQL.
        # Agora o frontend apenas exibe os dados que vêm do banco.
        self.tree_report.heading("Data Empréstimo", text="Data Inicial")

        for i in self.tree_report.get_children():
            self.tree_report.delete(i)
            
        try:
            year = int(self.e_report_year.get())
            month = int(self.cb_report_month.get())
        except ValueError:
            messagebox.showerror("Erro", "Ano/Mês inválidos.")
            return

        search_text = self.e_search_report.get().lower().strip()
        report_data = self.inv.generate_monthly_report(year, month)
        
        for log in report_data:
            op_type = log.get('operation_type')
            data_devolucao = log.get('data_devolucao')
            
            operation_display = ''
            data_inicial_display = format_date(log.get('data_operacao'))
            data_devolucao_display = ''

            if op_type == 'Empréstimo':
                # Agora a query já nos diz se foi devolvido ou não
                operation_display = "Devolvido" if data_devolucao else "Emprestado"
                data_devolucao_display = format_date(data_devolucao) if data_devolucao else ''
            elif op_type == 'Cadastro':
                operation_display = "Cadastro"
            
            row_values = (
                log.get('history_id'), log.get('item_id'), log.get('operador'), log.get('tipo'), log.get('brand'),
                log.get('model'), log.get('identificador'), log.get('usuario'),
                format_cpf(log.get('cpf')), operation_display, data_inicial_display,
                data_devolucao_display, log.get('center_cost'), log.get('cargo'), log.get('revenda')
            )
            
            row_values_cleaned = tuple(v or '' for v in row_values)

            if search_text and search_text not in " ".join(str(v).lower() for v in row_values_cleaned):
                continue
                
            self.tree_report.insert('', 'end', values=row_values_cleaned)

    # ALTERADO: Lógica de estorno implementada.
    def cmd_delete_report_entry(self):
        selected_items = self.tree_report.selection()
        if not selected_items:
            messagebox.showwarning("Aviso", "Selecione um lançamento do relatório para estornar.")
            return

        selected_item = selected_items[0]
        values = self.tree_report.item(selected_item, "values")
        
        try:
            history_id = int(values[0])
            item_id = values[1]
            op_display = values[9] 
        except (ValueError, IndexError):
            messagebox.showerror("Erro", "Não foi possível identificar o lançamento. Tente gerar o relatório novamente.")
            return

        # Mapeia o texto da tela para a operação real no banco
        op_real = "Cadastro" if op_display == "Cadastro" else "Empréstimo"
        if op_display == "Devolvido":
            # Para estornar uma devolução, a lógica é mais complexa, pois teríamos que
            # achar a ID da devolução, não do empréstimo. A query atual do relatório
            # não facilita isso. A implementação atual estorna o empréstimo.
            # Vamos manter a simplicidade por enquanto.
            pass

        confirm_msg = f"Você está prestes a registrar um estorno para a operação de '{op_real}' (Item ID {item_id}).\n\nA situação do item será revertida ao estado anterior a esta operação.\n\nConfirma a ação?"
        if op_real == "Cadastro":
            confirm_msg += "\n\nATENÇÃO: Estornar um cadastro tornará o item INATIVO no estoque."

        if not messagebox.askyesno("Confirmar Estorno", confirm_msg):
            return

        pwd = simpledialog.askstring("Autorização", "Digite a senha de administrador:", show='*')
        if pwd != ADMIN_PASS:
            messagebox.showerror("Erro", "Senha incorreta. Ação não autorizada.")
            return
        
        # Chama a nova função, passando o usuário logado para auditoria
        ok, msg = self.inv.reverse_history_entry(history_id, self.logged_user)

        if ok:
            messagebox.showinfo("Sucesso", msg)
            self.update_all_views() # Atualiza tudo para refletir a mudança
        else:
            messagebox.showerror("Erro no Estorno", msg)


    def cmd_generate_term(self):
        selected = self.tree_terms.selection()
        if not selected:
            self.lbl_terms.config(text="Selecione um empréstimo na tabela.", style='Danger.TLabel')
            return

        values = self.tree_terms.item(selected[0])["values"]
        pid, user = int(values[0]), values[3]
        
        ok, result = self.inv.generate_term(pid, user)
        if ok:
            self.lbl_terms.config(text=f"Termo gerado com sucesso: {os.path.basename(result)}", style='Success.TLabel')
            if messagebox.askyesno("Sucesso", f"Termo gerado em:\n{result}\n\nDeseja abri-lo agora?"):
                try:
                    os.startfile(result)
                except Exception as e:
                    messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo automaticamente:\n{e}")
        else:
            self.lbl_terms.config(text=f"Erro: {result}", style='Danger.TLabel')

    # --- Funções de Máscara ---
    def on_cpf_entry(self, event):
        self.on_widget_interaction(event)
        
        text = "".join(filter(str.isdigit, self.e_issue_cpf.get()))[:11]
        formatted = format_cpf(text)
        self.e_issue_cpf.delete(0, tk.END)
        self.e_issue_cpf.insert(0, formatted)
        self.e_issue_cpf.icursor(tk.END)

    def on_date_entry(self, event, entry_widget):
        self.on_widget_interaction(event)
        
        text = "".join(filter(str.isdigit, entry_widget.get()))[:8]
        formatted = text
        if len(text) > 2: formatted = f"{text[:2]}/{text[2:]}"
        if len(text) > 4: formatted = f"{text[:2]}/{text[2:4]}/{text[4:]}"
        
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, formatted)
        entry_widget.icursor(tk.END)
        
    def on_widget_interaction(self, event):
        """Limpa o estilo de erro de um widget quando o usuário interage com ele."""
        widget = event.widget
        widget_class = str(widget.winfo_class())
        
        if 'TEntry' in widget_class:
            widget.configure(style="TEntry")
        elif 'TCombobox' in widget_class:
            widget.configure(style="TCombobox")

    # --- Funções de Atualização de UI ---
    def update_stock(self):
        self.tree_stock.delete(*self.tree_stock.get_children())
        
        status_filter = self.cb_status_filter.get()
        tipo_filter = self.cb_tipo_filter.get()
        revenda_filter = self.cb_revenda_filter.get()
        search_text = self.e_search_stock.get().lower().strip()

        for p in self.inv.list_items():
            if status_filter and p.get('status') != status_filter: continue
            if tipo_filter and p.get('tipo') != tipo_filter: continue
            if revenda_filter and p.get('revenda') != revenda_filter: continue
            
            row_str = " ".join(str(v or '').lower() for v in p.values())
            if search_text and search_text not in row_str: continue

            row = (
                p.get('id'), p.get('revenda'), p.get('tipo'), p.get('brand'), 
                p.get('model'), p.get('status'), p.get('assigned_to'), 
                format_cpf(p.get('cpf')), p.get('identificador'), p.get('dominio'), 
                p.get('host'), p.get('endereco_fisico'), p.get('cpu'), p.get('ram'), 
                p.get('storage'), p.get('sistema'), p.get('licenca'), p.get('anydesk'), 
                p.get('setor'), p.get('ip'), p.get('mac'), format_date(p.get('date_registered'))
            )
            
            cleaned_row = tuple(v or '' for v in row)
            tag = "disp" if p.get('status') == "Disponível" else "indisp"
            self.tree_stock.insert('', 'end', values=cleaned_row, tags=(tag,))

        self.tree_stock.tag_configure("disp", background="#D4EDDA") # Verde mais suave
        self.tree_stock.tag_configure("indisp", background="#F8D7DA") # Vermelho mais suave

    def update_issue_cb(self):
        items = self.inv.list_items()
        disponiveis = [f"{p['id']} - {p.get('brand') or ''} {p.get('model') or ''}" for p in items if p['status'] == 'Disponível']
        self.cb_issue['values'] = disponiveis
        self.cb_issue.set('')

    def update_return_cb(self):
        items = self.inv.list_items()
        indisponiveis = [f"{p['id']} - {p.get('brand') or ''} {p.get('model') or ''} ({p.get('assigned_to') or ''})" for p in items if p['status'] == 'Indisponível']
        self.cb_return['values'] = indisponiveis
        self.cb_return.set('')

    def update_remove_cb(self):
        items = self.inv.list_items()
        todos = [f"{p['id']} - {p.get('brand') or ''} {p.get('model') or ''}" for p in items]
        self.cb_remove['values'] = todos
        self.cb_remove.set('')

    def update_edit_cb(self):
        items = self.inv.list_items()
        lst = [f"{p['id']} - {p.get('tipo') or ''} {p.get('brand') or ''} {p.get('model') or ''}" for p in items]
        self.cb_edit['values'] = lst
        self.cb_edit.set('')

    def update_history_table(self):
        self.tree_history.delete(*self.tree_history.get_children())
        search_text = self.e_search_history.get().lower().strip()

        for h in self.inv.list_history():
            row_values = (
                h.get("item_id"), h.get("operador"), h.get("operation"), 
                format_date(h.get("data_operacao")), h.get("tipo"), h.get("marca"), 
                h.get("modelo"), h.get("identificador"), h.get("usuario"), 
                format_cpf(h.get("cpf")), h.get("cargo"), h.get("center_cost"), 
                h.get("revenda")
            )
            cleaned_row = tuple(v or '' for v in row_values)
            
            if search_text and search_text not in " ".join(str(v).lower() for v in cleaned_row):
                continue
                
            self.tree_history.insert("", "end", values=cleaned_row)

    def update_terms_table(self):
        self.tree_terms.delete(*self.tree_terms.get_children())
        search_text = self.e_search_terms.get().lower().strip()

        for p in self.inv.list_items():
            if p['status'] == 'Indisponível':
                row_str = " ".join(str(v or '').lower() for v in p.values())
                if search_text and search_text not in row_str:
                    continue
                
                row = (
                    p.get('id'), p.get('tipo'), p.get('brand'), p.get('assigned_to'),
                    format_cpf(p.get('cpf')), format_date(p.get('date_issued')), p.get('revenda')
                )
                self.tree_terms.insert('', 'end', values=tuple(v or '' for v in row))

    def update_all_views(self):
        """Chama todas as funções de atualização da UI."""
        self.update_stock()
        self.update_issue_cb()
        self.update_return_cb()
        self.update_remove_cb()
        self.update_edit_cb()
        self.update_terms_table()
        self.update_history_table()
        self.cmd_generate_report()

    # --- Funções de Gráficos ---
    def graph_issue_return(self):
        plt.style.use('seaborn-v0_8-whitegrid')
        try:
            year = int(self.e_graph_year.get())
            month = int(self.cb_graph_month.get())
        except ValueError:
            messagebox.showerror("Erro", "Ano/Mês inválidos.")
            return

        days, issues, returns = self.inv.get_issue_return_counts(year, month)

        if not any(issues) and not any(returns):
            messagebox.showinfo("Gráfico", f"Nenhum empréstimo ou devolução encontrado para {month}/{year}.")
            return

        x = np.arange(len(days))
        width = 0.35

        fig, ax = plt.subplots(figsize=(12, 6))
        rects1 = ax.bar(x - width/2, issues, width, label='Empréstimos', color=PRIMARY_COLOR)
        rects2 = ax.bar(x + width/2, returns, width, label='Devoluções', color='#FFA726')

        ax.set_ylabel('Quantidade')
        ax.set_title(f'Empréstimos e Devoluções Diários - {month}/{year}')
        ax.set_xticks(x)
        ax.set_xticklabels(days)
        ax.legend()

        ax.bar_label(rects1, padding=3)
        ax.bar_label(rects2, padding=3)

        fig.tight_layout()
        plt.show()


    def graph_registration(self):
        plt.style.use('seaborn-v0_8-whitegrid')
        try:
            year = int(self.e_graph_year.get())
            month = int(self.cb_graph_month.get())
        except ValueError:
            messagebox.showerror("Erro", "Ano/Mês inválidos.")
            return

        days, registrations = self.inv.get_registration_counts(year, month)
        
        if not any(registrations):
            messagebox.showinfo("Gráfico", f"Nenhum cadastro encontrado para {month}/{year}.")
            return
            
        x = np.arange(len(days))
        width = 0.5
        
        fig, ax = plt.subplots(figsize=(12, 6))
        rects = ax.bar(x, registrations, width, label='Cadastros', color=SUCCESS_COLOR)
        
        ax.set_ylabel('Quantidade')
        ax.set_title(f'Cadastros Diários - {month}/{year}')
        ax.set_xticks(x)
        ax.set_xticklabels(days)
        ax.legend()
        
        ax.bar_label(rects, padding=3)
        
        fig.tight_layout()
        plt.show()