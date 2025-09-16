import os
import csv
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from tkinter.font import Font
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

# Importa de nossos próprios arquivos
from inventory_manager_db import InventoryDBManager
from utils import format_cpf, format_date
from config import (
    REVENDAS_OPTIONS, CENTER_COST_OPTIONS, ADMIN_PASS, ADMIN_USER
)

class App(tk.Tk):
    def __init__(self, user, role):
        super().__init__()
        self.title("Gestão de Estoque")
        self.geometry("1200x800")
        self.inv = InventoryDBManager()
        self.logged_user = user
        self.role = role
        self.create_widgets()

    # ==========================================================
    # Construção das abas
    # ==========================================================
    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True)

        self.tab_stock   = ttk.Frame(notebook)
        self.tab_add     = ttk.Frame(notebook)
        self.tab_edit    = ttk.Frame(notebook)
        self.tab_issue   = ttk.Frame(notebook)
        self.tab_return  = ttk.Frame(notebook)
        self.tab_remove  = ttk.Frame(notebook)
        self.tab_history = ttk.Frame(notebook)
        self.tab_report  = ttk.Frame(notebook)
        self.tab_terms   = ttk.Frame(notebook)
        self.tab_graphs  = ttk.Frame(notebook)

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
        
        # Tenta ordenar como número, se falhar, ordena como texto
        try:
            items.sort(key=lambda t: int(str(t[0]).replace('.', '').replace('-', '')), reverse=reverse)
        except ValueError:
            items.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)

        for index, (val, k) in enumerate(items):
            tv.move(k, "", index)
            
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))
        
    # ==========================================================
    # Aba de Estoque
    # ==========================================================
    def build_stock_tab(self):
        tab = self.tab_stock
        style = ttk.Style()
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))

        frm_filters = ttk.Frame(tab)
        frm_filters.pack(fill="x", padx=10, pady=5)

        ttk.Label(frm_filters, text="Status:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
        self.cb_status_filter = ttk.Combobox(frm_filters, values=["", "Disponível", "Indisponível"], state="readonly", width=15)
        self.cb_status_filter.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frm_filters, text="Tipo:").grid(row=0, column=2, padx=5, pady=2, sticky="e")
        self.cb_tipo_filter = ttk.Combobox(frm_filters, values=["", "Celular", "Notebook", "Desktop", "Impressora", "Tablet"], state="readonly", width=15)
        self.cb_tipo_filter.grid(row=0, column=3, padx=5, pady=2)

        ttk.Label(frm_filters, text="Revenda:").grid(row=0, column=4, padx=5, pady=2, sticky="e")
        self.cb_revenda_filter = ttk.Combobox(frm_filters, values=[""] + REVENDAS_OPTIONS, state="readonly", width=20)
        self.cb_revenda_filter.grid(row=0, column=5, padx=5, pady=2)

        ttk.Label(frm_filters, text="Buscar:").grid(row=0, column=6, padx=5, pady=2, sticky="e")
        self.e_search_stock = ttk.Entry(frm_filters, width=20)
        self.e_search_stock.grid(row=0, column=7, padx=5, pady=2)

        ttk.Button(frm_filters, text="Aplicar", command=self.update_stock).grid(row=0, column=8, padx=10, pady=2)
        ttk.Button(frm_filters, text="Limpar", command=self.clear_filters).grid(row=0, column=9, padx=5, pady=2)
        ttk.Button(frm_filters, text="Exportar CSV", command=lambda: self.exportar_csv(self.tree_stock, "Exportar Estoque", "estoque")).grid(row=0, column=10, padx=5, pady=2)

        cols = [
            "ID", "Revenda", "Tipo", "Marca", "Modelo", "Status", "Usuário", "CPF",
            "Identificador", "Domínio", "Host", "Endereço Físico", "CPU",
            "RAM", "Storage", "Sistema", "Licença", "AnyDesk",
            "Setor", "IP", "MAC", "Data Cadastro"
        ]

        self.tree_stock = ttk.Treeview(tab, columns=cols, show="headings", height=18)
        self.tree_stock.pack(fill="both", expand=True, padx=10, pady=5)

        col_widths = {
            "ID": 40, "Revenda": 130, "Tipo": 100, "Marca": 100, "Modelo": 120, "Status": 100,
            "Usuário": 140, "CPF": 110, "Identificador": 140
        }

        for col in cols:
            self.tree_stock.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree_stock, c, False))
            self.tree_stock.column(col, width=col_widths.get(col, 120), anchor="w", stretch=False)

        vsb = ttk.Scrollbar(tab, orient="vertical", command=self.tree_stock.yview)
        hsb = ttk.Scrollbar(tab, orient="horizontal", command=self.tree_stock.xview)
        self.tree_stock.configure(yscroll=vsb.set, xscroll=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def clear_filters(self):
        self.cb_status_filter.set("")
        self.cb_tipo_filter.set("")
        self.cb_revenda_filter.set("")
        self.e_search_stock.delete(0, "end")
        self.update_stock()

    def build_add_tab(self):
        frm = ttk.Frame(self.tab_add, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Tipo de Equipamento:").grid(row=0, column=0, sticky="e")
        self.cb_tipo = ttk.Combobox(frm, values=["Celular", "Notebook", "Desktop", "Impressora", "Tablet"], state="readonly")
        self.cb_tipo.grid(row=0, column=1)
        self.cb_tipo.bind("<<ComboboxSelected>>", self.on_tipo_selected)

        self.frm_dynamic = ttk.Frame(frm, padding=10)
        self.frm_dynamic.grid(row=1, column=0, columnspan=2, sticky="nsew")

        self.lbl_add = ttk.Label(frm, text="", foreground="red")
        self.lbl_add.grid(row=99, column=0, columnspan=2)

        ttk.Button(frm, text="Cadastrar", command=self.cmd_add).grid(row=100, column=0, columnspan=2, pady=5)

    def on_tipo_selected(self, event=None):
        for widget in self.frm_dynamic.winfo_children():
            widget.destroy()

        tipo = self.cb_tipo.get()
        if tipo:
            self.add_widgets = self._create_dynamic_form(self.frm_dynamic, tipo)

    def build_edit_tab(self):
        self.current_edit_id = None
        self.edit_widgets = {}

        frm = ttk.Frame(self.tab_edit, padding=10)
        frm.pack(fill="both", expand=True)

        frm_select = ttk.Frame(frm)
        frm_select.pack(fill="x", pady=5)
        
        ttk.Label(frm_select, text="Selecione o item para editar:").pack(side="left", padx=(0, 5))
        
        self.cb_edit = ttk.Combobox(frm_select, state="readonly", width=40)
        self.cb_edit.pack(side="left", expand=True, fill="x")
        
        ttk.Button(frm_select, text="Carregar Dados", command=self.cmd_load_item_for_edit).pack(side="left", padx=5)

        self.frm_edit_dynamic = ttk.Frame(frm, padding=10)
        self.frm_edit_dynamic.pack(fill="both", expand=True, pady=10)

        self.lbl_edit = ttk.Label(frm, text="", foreground="red")
        self.lbl_edit.pack()
        
        ttk.Button(frm, text="Salvar Alterações", command=self.cmd_save_edit).pack(pady=10)
        
        self.update_edit_cb()

    def build_issue_tab(self):
        frm = ttk.Frame(self.tab_issue, padding=10)
        frm.pack()
        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', pady=2)
        self.cb_issue = ttk.Combobox(frm, state='readonly', width=30)
        self.cb_issue.grid(row=0, column=1, pady=2)
        
        ttk.Label(frm, text="Funcionário:").grid(row=1, column=0, sticky='e', pady=2)
        self.e_issue_user = ttk.Entry(frm, width=32)
        self.e_issue_user.grid(row=1, column=1, pady=2)

        ttk.Label(frm, text="CPF:").grid(row=2, column=0, sticky='e', pady=2)
        self.e_issue_cpf = ttk.Entry(frm, width=32)
        self.e_issue_cpf.grid(row=2, column=1, pady=2)

        ttk.Label(frm, text="Centro de Custo:").grid(row=3, column=0, sticky='e', pady=2)
        self.cb_center = ttk.Combobox(frm, state='readonly', values=CENTER_COST_OPTIONS, width=28)
        self.cb_center.grid(row=3, column=1, pady=2)
        
        ttk.Label(frm, text="Cargo:").grid(row=4, column=0, sticky='e', pady=2)
        self.e_cargo = ttk.Entry(frm, width=32)
        self.e_cargo.grid(row=4, column=1, pady=2)
        
        ttk.Label(frm, text="Revenda:").grid(row=5, column=0, sticky='e', pady=2)
        self.cb_revenda = ttk.Combobox(frm, state='readonly', values=REVENDAS_OPTIONS, width=28)
        self.cb_revenda.grid(row=5, column=1, pady=2)
        
        ttk.Label(frm, text="Data Empréstimo (dd/mm/aaaa):").grid(row=6, column=0, sticky='e', pady=2)
        self.e_date_issue = ttk.Entry(frm, width=32)
        self.e_date_issue.grid(row=6, column=1, pady=2)

        self.lbl_issue = ttk.Label(frm, text="", foreground='red')
        self.lbl_issue.grid(row=7, column=0, columnspan=2, pady=5)
        
        ttk.Button(frm, text="Emprestar", command=self.cmd_issue).grid(row=8, column=0, columnspan=2, pady=5)
        
        self.e_issue_cpf.bind("<KeyRelease>", self.on_cpf_entry)
        self.e_date_issue.bind("<KeyRelease>", lambda e: self.on_date_entry(e, self.e_date_issue))


    def build_return_tab(self):
        frm = ttk.Frame(self.tab_return, padding=10)
        frm.pack()

        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', pady=2)
        self.cb_return = ttk.Combobox(frm, state='readonly', width=30)
        self.cb_return.grid(row=0, column=1, pady=2)

        ttk.Label(frm, text="Data Devolução (dd/mm/aaaa):").grid(row=1, column=0, sticky='e', pady=2)
        self.e_date_return = ttk.Entry(frm, width=32)
        self.e_date_return.grid(row=1, column=1, pady=2)

        self.e_date_return.bind("<KeyRelease>", lambda e: self.on_date_entry(e, self.e_date_return))

        self.lbl_ret = ttk.Label(frm, text="", foreground='red')
        self.lbl_ret.grid(row=2, column=0, columnspan=2, pady=5)

        ttk.Button(frm, text="Devolver", command=self.cmd_return).grid(row=3, column=0, columnspan=2, pady=5)

    def build_remove_tab(self):
        frm = ttk.Frame(self.tab_remove, padding=10)
        frm.pack()
        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', pady=2)
        self.cb_remove = ttk.Combobox(frm, state='readonly', width=30)
        self.cb_remove.grid(row=0, column=1, pady=2)

        self.lbl_rem = ttk.Label(frm, text="", foreground='red')
        self.lbl_rem.grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Button(frm, text="Remover", command=self.cmd_remove).grid(row=2, column=0, columnspan=2, pady=5)

    def build_history_tab(self):
        tab = self.tab_history

        top_frame = ttk.Frame(tab)
        top_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(top_frame, text="Buscar:").pack(side="left", padx=(0, 5))
        self.e_search_history = ttk.Entry(top_frame, width=30)
        self.e_search_history.pack(side="left")

        ttk.Button(top_frame, text="Buscar", command=self.update_history_table).pack(side="left", padx=5)
        ttk.Button(top_frame, text="Limpar", command=self.clear_history_filter).pack(side="left", padx=5)

        ttk.Button(top_frame, text="Exportar CSV",
                command=lambda: self.exportar_csv(self.tree_history, "Exportar Histórico", "histórico")
                ).pack(side="right")

        cols = (
            "ID Item", "Operador", "Operação", "Data", "Tipo", "Marca", "Modelo", "Identificador",
            "Usuário", "CPF", "Cargo", "Centro de Custo", "Revenda"
        )
        self.tree_history = ttk.Treeview(tab, columns=cols, show="headings", height=18)

        col_widths = { "ID Item": 60, "Operador": 100, "Operação": 100, "Data": 100, "CPF": 120 }

        for c in cols:
            self.tree_history.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_history, col, False))
            self.tree_history.column(c, width=col_widths.get(c, 120), anchor="w", stretch=False)

        ysb = ttk.Scrollbar(tab, orient="vertical", command=self.tree_history.yview)
        xsb = ttk.Scrollbar(tab, orient="horizontal", command=self.tree_history.xview)
        self.tree_history.configure(yscroll=ysb.set, xscroll=xsb.set)

        self.tree_history.pack(fill="both", expand=True, padx=10, pady=(0,5))
        ysb.pack(side="right", fill="y")
        xsb.pack(side="bottom", fill="x")

    def clear_history_filter(self):
        self.e_search_history.delete(0, "end")
        self.update_history_table()

    def build_report_tab(self):
        frm = ttk.Frame(self.tab_report, padding=10)
        frm.pack(fill="both", expand=True)

        top_frm = ttk.Frame(frm)
        top_frm.pack(fill='x', pady=5)

        ttk.Label(top_frm, text="Ano:").grid(row=0, column=0, sticky='e', padx=(0,5))
        self.e_report_year  = ttk.Entry(top_frm, width=6)
        self.e_report_year.insert(0, datetime.now().year)
        self.e_report_year.grid(row=0, column=1)

        ttk.Label(top_frm, text="Mês:").grid(row=0, column=2, sticky='e', padx=(10,5))
        self.cb_report_month = ttk.Combobox(top_frm, values=list(range(1, 13)), width=4, state="readonly")
        self.cb_report_month.set(datetime.now().month)
        self.cb_report_month.grid(row=0, column=3)

        ttk.Button(top_frm, text="Gerar Relatório", command=self.cmd_generate_report).grid(row=0, column=4, padx=10)

        # Frame para busca na linha de baixo
        search_frm = ttk.Frame(frm)
        search_frm.pack(fill='x', padx=5, pady=(5,10))
        ttk.Label(search_frm, text="Buscar no Relatório:").pack(side="left", padx=5)
        self.e_search_report = ttk.Entry(search_frm, width=30)
        self.e_search_report.pack(side="left", padx=5)
        ttk.Button(search_frm, text="Limpar Busca", command=self.clear_report_filter).pack(side="left", padx=5)

        # Botões de ação movidos para o final
        action_frm = ttk.Frame(frm)
        action_frm.pack(fill='x', padx=5, pady=5)
        ttk.Button(action_frm, text="Estornar Lançamento", command=self.cmd_delete_report_entry).pack(side="right", padx=5)
        ttk.Button(action_frm, text="Exportar CSV", command=lambda: self.exportar_csv(self.tree_report, "Exportar Relatório", "relatório")).pack(side="right", padx=5)

        cols = ("ID Item", "Operador", "Tipo", "Marca", "Modelo", "Identificador", "Usuário", "CPF", "Operação", "Data Empréstimo", "Data Devolução", "Centro de Custo", "Cargo", "Revenda")
        self.tree_report = ttk.Treeview(frm, columns=cols, show='headings', height=16)

        col_widths = {"ID Item": 60, "Operador": 100, "Operação": 110, "CPF": 120}

        for c in cols:
            self.tree_report.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_report, col, False))
            self.tree_report.column(c, width=col_widths.get(c, 120), anchor='w', stretch=False)
        
        ysb = ttk.Scrollbar(frm, orient="vertical", command=self.tree_report.yview)
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=self.tree_report.xview)
        self.tree_report.configure(yscroll=ysb.set, xscroll=hsb.set)
        
        self.tree_report.pack(fill='both', expand=True, padx=5, pady=5)
        ysb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def clear_report_filter(self):
        self.e_search_report.delete(0, "end")
        self.cmd_generate_report()

    def build_terms_tab(self):
        frm = ttk.Frame(self.tab_terms, padding=10)
        frm.pack(fill="both", expand=True)

        search_frame = ttk.Frame(frm)
        search_frame.pack(fill="x", pady=5)
        ttk.Label(search_frame, text="Buscar:").grid(row=0, column=0, sticky="e")
        self.e_search_terms = ttk.Entry(search_frame, width=30)
        self.e_search_terms.grid(row=0, column=1, padx=5)
        ttk.Button(search_frame, text="Buscar", command=self.update_terms_table).grid(row=0, column=2, padx=5)

        cols = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        self.tree_terms = ttk.Treeview(frm, columns=cols, show="headings", height=15)
        
        col_widths = { "ID": 50, "CPF": 120, "Data Empréstimo": 120 }
        for c in cols:
            self.tree_terms.heading(c, text=c, anchor="w", command=lambda col=c: self.treeview_sort_column(self.tree_terms, col, False))
            self.tree_terms.column(c, width=col_widths.get(c, 150), anchor="w")

        self.tree_terms.pack(fill="both", expand=True, pady=10)
        
        self.lbl_terms = ttk.Label(frm, text="", foreground="red")
        self.lbl_terms.pack()
        ttk.Button(frm, text="Gerar Termo Selecionado", command=self.cmd_generate_term).pack(pady=5)

    def build_graph_tab(self):
        frm = ttk.Frame(self.tab_graphs, padding=10)
        frm.pack(fill='both', expand=True)

        ctrl = ttk.Frame(frm)
        ctrl.pack(pady=5)
        ttk.Label(ctrl, text="Ano:").pack(side='left')
        self.e_graph_year = ttk.Entry(ctrl, width=6)
        self.e_graph_year.insert(0, datetime.now().year)
        self.e_graph_year.pack(side='left', padx=(5, 15))

        ttk.Label(ctrl, text="Mês:").pack(side='left')
        self.cb_graph_month = ttk.Combobox(ctrl, values=list(range(1, 13)), width=4, state='readonly')
        self.cb_graph_month.set(datetime.now().month)
        self.cb_graph_month.pack(side='left', padx=5)

        lbl = ttk.Label(frm, text="Gráficos:")
        lbl.pack(pady=(20, 5))
        btns = ttk.Frame(frm)
        btns.pack()
        ttk.Button(btns, text="Empréstimos x Devoluções (Mensal)", command=self.graph_issue_return).pack(side='left', padx=10)
        ttk.Button(btns, text="Cadastros (Mensal)", command=self.graph_registration).pack(side='left', padx=10)

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
                headers = [tree.heading(c)["text"] for c in tree["columns"]]
                writer.writerow(headers)
                for row_id in tree.get_children():
                    values = tree.item(row_id)["values"]
                    writer.writerow(values)
            messagebox.showinfo(titulo, f"{nome.capitalize()} exportado para:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro ao Exportar", f"Ocorreu um erro: {e}")

    # --- Comandos e Ações ---

    def _create_dynamic_form(self, parent_frame, tipo, item_data=None):
        if item_data is None:
            item_data = {}
        
        widgets = {}

        # Campos comuns
        ttk.Label(parent_frame, text="Marca:").grid(row=1, column=0, sticky="e", pady=2)
        e_brand = ttk.Entry(parent_frame, width=30)
        e_brand.grid(row=1, column=1, pady=2)
        e_brand.insert(0, item_data.get('brand') or '')
        widgets['brand'] = e_brand

        ttk.Label(parent_frame, text="Revenda:").grid(row=2, column=0, sticky="e", pady=2)
        cb_revenda = ttk.Combobox(parent_frame, values=REVENDAS_OPTIONS, state="readonly", width=28)
        cb_revenda.grid(row=2, column=1, pady=2)
        cb_revenda.set(item_data.get('revenda') or '')
        widgets['revenda'] = cb_revenda

        # Campos específicos
        if tipo == "Celular":
            ttk.Label(parent_frame, text="Modelo:").grid(row=3, column=0, sticky="e", pady=2)
            e_model = ttk.Entry(parent_frame, width=30); e_model.grid(row=3, column=1, pady=2)
            e_model.insert(0, item_data.get('model') or ''); widgets['model'] = e_model
            
            ttk.Label(parent_frame, text="IMEI:").grid(row=4, column=0, sticky="e", pady=2)
            e_identificador = ttk.Entry(parent_frame, width=30); e_identificador.grid(row=4, column=1, pady=2)
            e_identificador.insert(0, item_data.get('identificador') or ''); widgets['identificador'] = e_identificador

        elif tipo in ["Notebook", "Desktop"]:
            fields = {
                'dominio': ("Domínio:", ttk.Combobox, {"values": ["Sim", "Não"], "state": "readonly"}),
                'host': ("Host:", ttk.Entry, {}), 'endereco_fisico': ("Endereço Físico:", ttk.Entry, {}),
                'storage': ("Armazenamento (GB):", ttk.Entry, {}), 'sistema': ("Sistema Operacional:", ttk.Entry, {}),
                'cpu': ("Processador:", ttk.Entry, {}), 'ram': ("Memória RAM:", ttk.Entry, {}),
                'licenca': ("Licença Windows:", ttk.Entry, {}), 'anydesk': ("AnyDesk:", ttk.Entry, {})
            }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=3 + i, column=0, sticky="e", pady=2)
                widget = w_class(parent_frame, width=30, **opts); widget.grid(row=3 + i, column=1, pady=2)
                if isinstance(widget, ttk.Combobox): widget.set(item_data.get(key) or '')
                else: widget.insert(0, item_data.get(key) or '')
                widgets[key] = widget
        
        elif tipo == "Impressora":
            fields = { 'setor': ("Setor:", ttk.Entry, {}), 'ip': ("IP:", ttk.Entry, {}), 'mac': ("MAC:", ttk.Entry, {}) }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=3 + i, column=0, sticky="e", pady=2)
                widget = w_class(parent_frame, width=30, **opts); widget.grid(row=3 + i, column=1, pady=2)
                widget.insert(0, item_data.get(key) or ''); widgets[key] = widget

        elif tipo == "Tablet":
            fields = { 'model': ("Modelo:", ttk.Entry,{}), 'identificador': ("Nº de Série:", ttk.Entry, {}), 'storage': ("Armazenamento (GB):", ttk.Entry, {}) }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=3 + i, column=0, sticky="e", pady=2)
                widget = w_class(parent_frame, width=30, **opts); widget.grid(row=3 + i, column=1, pady=2)
                widget.insert(0, item_data.get(key) or ''); widgets[key] = widget
        return widgets

    def cmd_add(self):
        tipo = self.cb_tipo.get().strip()
        if not tipo:
            self.lbl_add.config(text="Selecione o tipo de equipamento.", foreground="red")
            return
        
        dados = {}
        for key, widget in self.add_widgets.items():
            if widget.winfo_exists():
                dados[key] = widget.get().strip()

        erros = []
        if not dados.get("brand"): erros.append("Informe a marca.")
        if not dados.get("revenda"): erros.append("Informe o campo Revenda.")

        if tipo in ["Celular", "Tablet"]:
            if not dados.get("model"):
                erros.append("Informe o modelo.")
            if not dados.get("identificador"):
                erros.append("Informe o Identificador (IMEI/Nº de Série).")

        if tipo == "Tablet" and not dados.get("storage"):
            erros.append("Informe o armazenamento.")

        if tipo == "Impressora":
            obrigatorios_impressora = {
                "ip": "Informe o IP.",
                "setor": "Informe o setor.",
                "mac": "Informe o MAC."
            }
            for campo, mensagem in obrigatorios_impressora.items():
                if not dados.get(campo):
                    erros.append(mensagem)

        if tipo in ["Desktop", "Notebook"]:
            obrigatorios_pc = {
                "dominio": "Informe o domínio.",
                "host": "Informe o host.",
                "endereco_fisico": "Informe o endereço físico.",
                "storage": "Informe o tamanho do armazenamento.",
                "sistema": "Informe o sistema operacional.",
                "cpu": "Informe o processador.",
                "ram": "Informe a quantidade de memória RAM.",
                "licenca": "Informe a licença do Windows.",
                "anydesk": "Informe o código do AnyDesk."
            }
            for campo, mensagem in obrigatorios_pc.items():
                if not dados.get(campo):
                    erros.append(mensagem)

        if erros:
            self.lbl_add.config(text="\n".join(erros), foreground="red")
            return

        dados["tipo"] = tipo
        dados["date_registered"] = datetime.now().strftime("%Y-%m-%d")

        item_id = self.inv.add_item(dados, self.logged_user)
        if not item_id:
            self.lbl_add.config(text="Erro ao cadastrar item.", foreground="red")
            return

        self.lbl_add.config(text=f"{tipo} ID {item_id} cadastrado com sucesso!", foreground="green")
        
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
        self.lbl_edit.config(text=msg, foreground="green" if ok else "red")
        
        if ok:
            for widget in self.frm_edit_dynamic.winfo_children():
                widget.destroy()
            self.current_edit_id = None
            self.cb_edit.set('')
            self.update_all_views()

    def cmd_issue(self):
        sel = self.cb_issue.get()
        user = self.e_issue_user.get().strip()
        cpf = self.e_issue_cpf.get().strip()
        cc = self.cb_center.get().strip()
        cargo = self.e_cargo.get().strip()
        revenda = self.cb_revenda.get().strip()
        date_issue = self.e_date_issue.get().strip()

        if not all([sel, user, cc, cargo, revenda, date_issue]):
            self.lbl_issue.config(text="Preencha todos os campos de empréstimo.", foreground="red")
            return
            
        if len("".join(filter(str.isdigit, cpf))) < 11:
            self.lbl_issue.config(text="CPF inválido. Informe os 11 dígitos.", foreground="red")
            return

        pid = int(sel.split(' - ', 1)[0])
        ok, msg = self.inv.issue(pid, user, cpf, cc, cargo, revenda, date_issue, self.logged_user)
        self.lbl_issue.config(text=msg, foreground='green' if ok else 'red')
        
        if ok:
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
            self.lbl_ret.config(text="Selecione aparelho e informe data de devolução.")
            return
            
        pid = int(sel.split(' - ', 1)[0])
        ok, msg = self.inv.ret(pid, date_return, self.logged_user)
        self.lbl_ret.config(text=msg, foreground='green' if ok else 'red')
        
        if ok:
            self.e_date_return.delete(0, 'end')
            self.update_all_views()

    def cmd_remove(self):
        sel = self.cb_remove.get()
        if not sel:
            self.lbl_rem.config(text="Selecione um aparelho para remover."); return

        pid = int(sel.split(' - ', 1)[0])
        item = self.inv.find(pid)
        if not item:
            self.lbl_rem.config(text="Item não encontrado.", foreground='red'); return

        if item['status'] != 'Disponível':
            self.lbl_rem.config(text="Não é possível remover um produto emprestado.", foreground='red')
            return

        pwd = simpledialog.askstring("Autorização", "Digite a senha de administrador:", show='*')
        if pwd != ADMIN_PASS:
            self.lbl_rem.config(text="Senha incorreta. Ação não autorizada.", foreground='red')
            return
            
        ok, msg = self.inv.remove(pid, self.logged_user)
        self.lbl_rem.config(text=msg, foreground='green' if ok else 'red')
        
        if ok:
            self.update_all_views()



    def cmd_generate_report(self):
        # Primeiro, renomeia a coluna para ser mais genérica, já que agora inclui cadastros
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
            
            # Lógica para definir o que será exibido em cada tipo de operação
            operation_display = ''
            data_inicial_display = format_date(log.get('data_operacao'))
            data_devolucao_display = ''

            if op_type == 'Empréstimo':
                operation_display = "Devolvido" if data_devolucao else "Emprestado"
                data_devolucao_display = format_date(data_devolucao) if data_devolucao else ''
            elif op_type == 'Cadastro':
                operation_display = "Cadastro"
            
            row_values = (
                log.get('item_id'),
                log.get('operador'),
                log.get('tipo'),
                log.get('brand'),
                log.get('model'),
                log.get('identificador'),
                log.get('usuario'),
                format_cpf(log.get('cpf')),
                operation_display,
                data_inicial_display,
                data_devolucao_display,
                log.get('center_cost'),
                log.get('cargo'),
                log.get('revenda')
            )
            
            row_values_cleaned = tuple(v or '' for v in row_values)

            if search_text and search_text not in " ".join(str(v).lower() for v in row_values_cleaned):
                continue
                
            self.tree_report.insert('', 'end', values=row_values_cleaned)

    def cmd_delete_report_entry(self):
        if not self.tree_report.selection():
            messagebox.showwarning("Aviso", "Selecione um lançamento do relatório para estornar.")
            return

        pwd = simpledialog.askstring("Autorização", "Digite a senha de administrador:", show='*')
        if pwd != ADMIN_PASS:
            messagebox.showerror("Erro", "Senha incorreta. Ação não autorizada.")
            return

        # ATENÇÃO: A lógica para estornar (deletar um registro do histórico)
        # não foi implementada no `inventory_manager_db.py`.
        # Se for necessária, ela precisará ser adicionada lá.
        messagebox.showinfo("Funcionalidade Pendente", "A lógica de estorno ainda não foi implementada no backend.")

    def cmd_generate_term(self):
        selected = self.tree_terms.selection()
        if not selected:
            self.lbl_terms.config(text="Selecione um empréstimo na tabela.", foreground='red')
            return

        values = self.tree_terms.item(selected[0])["values"]
        pid, user = int(values[0]), values[3]
        
        ok, result = self.inv.generate_term(pid, user)
        if ok:
            self.lbl_terms.config(text=f"Termo gerado com sucesso: {os.path.basename(result)}", foreground='green')
            if messagebox.askyesno("Sucesso", f"Termo gerado em:\n{result}\n\nDeseja abri-lo agora?"):
                os.startfile(result)
        else:
            self.lbl_terms.config(text=f"Erro: {result}", foreground='red')

    # --- Funções de Máscara ---
    def on_cpf_entry(self, event):
        text = "".join(filter(str.isdigit, self.e_issue_cpf.get()))[:11]
        formatted = format_cpf(text)
        self.e_issue_cpf.delete(0, tk.END)
        self.e_issue_cpf.insert(0, formatted)
        self.e_issue_cpf.icursor(tk.END)

    def on_date_entry(self, event, entry_widget):
        text = "".join(filter(str.isdigit, entry_widget.get()))[:8]
        formatted = text
        if len(text) > 2: formatted = f"{text[:2]}/{text[2:]}"
        if len(text) > 4: formatted = f"{text[:2]}/{text[2:4]}/{text[4:]}"
        
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, formatted)
        entry_widget.icursor(tk.END)

    # --- Funções de Atualização de UI ---
    def update_stock(self):
        self.tree_stock.delete(*self.tree_stock.get_children())
        
        status_filter = self.cb_status_filter.get()
        tipo_filter = self.cb_tipo_filter.get()
        revenda_filter = self.cb_revenda_filter.get()
        search_text = self.e_search_stock.get().lower().strip()

        for p in self.inv.list_items():
            # Lógica de filtro
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

        self.tree_stock.tag_configure("disp", background="#77ff77")
        self.tree_stock.tag_configure("indisp", background="#ff7a7a")

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
        rects1 = ax.bar(x - width/2, issues, width, label='Empréstimos', color='skyblue')
        rects2 = ax.bar(x + width/2, returns, width, label='Devoluções', color='salmon')

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
        rects = ax.bar(x, registrations, width, label='Cadastros', color='lightgreen')
        
        ax.set_ylabel('Quantidade')
        ax.set_title(f'Cadastros Diários - {month}/{year}')
        ax.set_xticks(x)
        ax.set_xticklabels(days)
        ax.legend()
        
        ax.bar_label(rects, padding=3)
        
        fig.tight_layout()
        plt.show()