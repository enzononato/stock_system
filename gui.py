import os
import csv
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np
import os
import sys

# --- Lógica para encontrar recursos como o logo.ico para o pyinstaller ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Importa de nossos próprios arquivos
from inventory_manager_db import InventoryDBManager
from user_manager_db import UserDBManager
from utils import format_cpf, format_date, format_title_case, format_time
from config import (
    REVENDAS_OPTIONS, CENTER_COST_OPTIONS, ADMIN_PASS, ADMIN_USER, SETORES_OPTIONS
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


# --- NOVA CLASSE PARA FRAME COM ROLAGEM ---
class ScrollableFrame(ttk.Frame):
    """
    Um frame que contém uma barra de rolagem vertical.
    Use o atributo 'scrollable_frame' para colocar seus widgets.
    """
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        
        # O Canvas permite "desenhar" widgets e rolar a visualização
        # A cor de fundo é ajustada para combinar com o fundo principal da janela
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, background="#F5F5F5") # <-- ALTERAÇÃO

        # A barra de rolagem que controla o Canvas
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        
        # Este é o frame interno que conterá os widgets.
        # Ele também precisa ter a cor de fundo correta para parecer invisível.
        self.scrollable_frame = ttk.Frame(canvas, style="TFrame") # <-- ALTERAÇÃO (garante o estilo correto)

        # Atualiza a região de rolagem do canvas sempre que o frame interno muda de tamanho
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        # Coloca o frame interno dentro do canvas
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Faz o frame interno ter a mesma largura do canvas, permitindo a centralização do conteúdo
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width)) # <-- ALTERAÇÃO (essencial para centralizar)

        # Configura o canvas para ser controlado pela barra de rolagem
        canvas.configure(yscrollcommand=scrollbar.set)

        # --- Ativa a rolagem com o mouse ---
        canvas.bind("<Enter>", lambda e: self._bind_mouse_scroll(e, canvas))
        canvas.bind("<Leave>", lambda e: self._unbind_mouse_scroll(e, canvas))

        # Empacota os componentes
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _on_mouse_wheel(self, event, canvas):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _bind_mouse_scroll(self, event, canvas):
        # Vincula a rolagem quando o mouse entra no canvas
        canvas.bind_all("<MouseWheel>", lambda e: self._on_mouse_wheel(e, canvas))

    def _unbind_mouse_scroll(self, event, canvas):
        # Desvincula a rolagem quando o mouse sai do canvas
        canvas.unbind_all("<MouseWheel>")


class App(tk.Tk):
    def __init__(self, user_id, user, role):
        super().__init__()
        self.title("Gestão de Estoque de Equipamentos Eletrônicos")
        # Definir ícone personalizado
        self.iconbitmap(resource_path("logo.ico"))

        self.geometry("1280x800")
        self.configure(background=BG_COLOR)
        
        self.inv = InventoryDBManager()
        self.user_db = UserDBManager()
        
        self.logged_user_id = user_id
        self.logged_user = user
        self.role = role
        
        # Registra a função de validação para ser usada pelo Tcl/Tk
        # %P é o valor do campo APÓS a edição
        # %L é o comprimento máximo que passaremos (para nota fiscal é o '9')
        vcmd = (self.register(self._validate_numeric_input), '%P', '9')
        self.numeric_validation = vcmd

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
        style.configure("Success.TButton", background=SUCCESS_COLOR, foreground="white")
        style.map("Success.TButton", background=[('active', '#2E7D32')], relief=[('pressed', 'sunken')])

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
        

    def _on_tab_changed(self, event):
        """
        Esta função é chamada sempre que o usuário muda de aba.
        Ela reconfigura a tecla 'Enter' para a ação da aba atual.
        """
        # Primeiro, remove qualquer ação anterior da tecla Enter
        self.unbind('<Return>')

        # Pega o texto da aba que foi selecionada
        selected_tab_text = event.widget.tab(event.widget.select(), "text")

        # Mapeia o texto da aba para a função que o Enter deve chamar
        tab_actions = {
            "Cadastrar": self.cmd_add,
            "Remover": self.cmd_remove,
            "Emprestar": self.cmd_issue,
            "Editar": self.cmd_save_edit,
            "Usuários": self.cmd_add_user # Na aba usuários, o Enter vai cadastrar
        }

        # Se a aba selecionada tem uma ação definida, vincula o Enter a ela
        if selected_tab_text in tab_actions:
            action_function = tab_actions[selected_tab_text]
            self.bind('<Return>', action_function)


    # ==========================================================
    # Construção das abas
    # ==========================================================
    def create_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Vincula o evento de mudança de aba à função _on_tab_changed
        notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

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
        self.tab_users   = ttk.Frame(notebook, padding=15)
        self.tab_peripherals = ttk.Frame(notebook, padding=15)
        self.tab_linking     = ttk.Frame(notebook, padding=15)

        notebook.add(self.tab_stock,  text="Estoque")
        notebook.add(self.tab_add,    text="Cadastrar")
        notebook.add(self.tab_edit,   text="Editar")
        notebook.add(self.tab_peripherals, text="Periféricos")
        notebook.add(self.tab_linking,text="Vincular")
        notebook.add(self.tab_issue,  text="Emprestar")
        notebook.add(self.tab_return, text="Devolver")
        notebook.add(self.tab_remove, text="Remover")
        notebook.add(self.tab_history,text="Histórico")
        notebook.add(self.tab_report, text="Relatório")
        notebook.add(self.tab_terms,  text="Termos")
        notebook.add(self.tab_graphs, text="Gráficos")
        notebook.add(self.tab_users,  text="Usuários")

        self.build_stock_tab()
        self.build_add_tab()
        self.build_edit_tab()
        self.build_peripherals_tab()
        self.build_linking_tab()
        self.build_issue_tab()
        self.build_return_tab()
        self.build_remove_tab()
        self.build_history_tab()
        self.build_report_tab()
        self.build_terms_tab()
        self.build_graph_tab()
        self.build_users_tab()
        
        # Restrições por role
        if self.role != "Gestor":
            notebook.hide(self.tab_remove)
            notebook.hide(self.tab_users)
            if self.role == "Jovem Aprendiz":
                notebook.hide(self.tab_history)
                notebook.hide(self.tab_issue)
                notebook.hide(self.tab_return)
                notebook.hide(self.tab_edit)
                notebook.hide(self.tab_report)


        self.update_all_views()
        

    def _validate_numeric_input(self, P, max_len):
        """
        Função de validação para Tcl/Tk.
        P: O valor que o Entry terá APÓS a edição.
        max_len: O comprimento máximo permitido.
        """
        # Permite que o campo fique vazio (se o usuário apagar tudo)
        if P == "":
            return True
        
        # Verifica se o novo valor é numérico E se está dentro do limite de tamanho
        if P.isdigit() and len(P) <= int(max_len):
            return True
        
        # Se não atender às condições, rejeita a alteração
        return False

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
        self.cb_status_filter = ttk.Combobox(frm_filters, values=["", "Disponível", "Indisponível", "Pendente"], state="readonly", width=15)
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
            "ID", "Revenda", "Tipo", "Marca", "Modelo", "Status", "Periféricos (Qtd)", "Usuário", "CPF", "Nota Fiscal", "Fornecedor",
            "Identificador", "Domínio", "Host", "Endereço Físico", "CPU",
            "RAM", "Storage", "Sistema", "Licença", "AnyDesk",
            "Setor", "IP", "MAC",
            "POE", "Qtd. Portas",
            "Data Cadastro"
        ]

        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill="both", expand=True)

        self.tree_stock = ttk.Treeview(tree_frame, columns=cols, show="headings")
        self.tree_stock.pack(side="left", fill="both", expand=True)

        col_widths = {
            "ID": 40, "Revenda": 130, "Tipo": 100, "Marca": 100, "Modelo": 120, "Status": 100,
            "Periféricos (Qtd)": 110, "PoE": 60, "Qtd. Portas": 90,
            "Usuário": 140, "CPF": 110, "Identificador": 140, "Nota Fiscal": 100, "Fornecedor": 140,
        }

        for col in cols:
            self.tree_stock.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree_stock, c, False))
            self.tree_stock.column(col, width=col_widths.get(col, 120), anchor="center", stretch=False)

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
        # Cria o container rolável que preenche toda a aba
        scroll_container = ScrollableFrame(self.tab_add)
        scroll_container.pack(fill="both", expand=True)
        
        # Pega o frame interno (o que realmente rola) para colocar nosso conteúdo
        container = scroll_container.scrollable_frame
        
        # O resto do código permanece o mesmo, mas agora está dentro da área rolável
        frm = ttk.Frame(container)
        frm.pack(anchor="n", pady=20, padx=20)

        # Configura colunas para alinhamento
        frm.grid_columnconfigure(1, weight=1)

        ttk.Label(frm, text="Tipo de Equipamento:", font=FONT_BOLD).grid(row=0, column=0, sticky="e", padx=5, pady=(0, 10))
        self.cb_tipo = ttk.Combobox(frm, values=["Celular", "Notebook", "Desktop", "Impressora", "Tablet", "Switch", "HD"], state="readonly")
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

        scroll_container = ScrollableFrame(self.tab_edit)
        scroll_container.pack(fill="both", expand=True)
        
        container = scroll_container.scrollable_frame
        
        frm = ttk.Frame(container)
        frm.pack(anchor="n", pady=20, padx=20, fill="x")

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


    def build_peripherals_tab(self):
        tab = self.tab_peripherals
        
        # --- Frame Superior: Cadastro ---
        frm_add = ttk.LabelFrame(tab, text=" Cadastrar Novo Periférico ", padding=15)
        frm_add.pack(fill="x", pady=(0, 15))
        frm_add.grid_columnconfigure(1, weight=1)
        frm_add.grid_columnconfigure(3, weight=1)

        # Linha 1
        ttk.Label(frm_add, text="Tipo:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.cb_peri_tipo = ttk.Combobox(frm_add, state="readonly", values=["Mouse", "Teclado", "Monitor", "Mousepad", "Suporte", "Estabilizador", "Mochila", "Fone", "Webcam", "Adaptador", "Outro"])
        self.cb_peri_tipo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(frm_add, text="Marca:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.e_peri_brand = ttk.Entry(frm_add)
        self.e_peri_brand.grid(row=0, column=3, sticky="ew", padx=5, pady=5)

        # Linha 2
        ttk.Label(frm_add, text="Modelo:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.e_peri_model = ttk.Entry(frm_add)
        self.e_peri_model.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        
        ttk.Label(frm_add, text="Identificador:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        self.e_peri_id = ttk.Entry(frm_add)
        self.e_peri_id.grid(row=1, column=3, sticky="ew", padx=5, pady=5)

        # Linha 3
        ttk.Label(frm_add, text="Nota Fiscal:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.e_peri_nf = ttk.Entry(frm_add)
        self.e_peri_nf.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(frm_add, text="Fornecedor:").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.e_peri_fornecedor = ttk.Entry(frm_add)
        self.e_peri_fornecedor.grid(row=2, column=3, sticky="ew", padx=5, pady=5)

        # Ações
        frm_add_actions = ttk.Frame(frm_add)
        frm_add_actions.grid(row=3, column=0, columnspan=4, pady=(10,0))
        ttk.Button(frm_add_actions, text="Cadastrar Periférico", command=self.cmd_add_peripheral, style="Primary.TButton").pack(side="left")
        self.lbl_peri_add = ttk.Label(frm_add_actions, text="")
        self.lbl_peri_add.pack(side="left", padx=15)

        # --- Frame Inferior: Lista de Periféricos ---
        frm_list = ttk.LabelFrame(tab, text=" Lista de Periféricos ", padding=15)
        frm_list.pack(fill="both", expand=True)

        # Frame de filtros/busca
        frm_filters = ttk.Frame(frm_list)
        frm_filters.pack(fill="x", pady=(0, 10))

        ttk.Label(frm_filters, text="Buscar:").pack(side="left", padx=(0, 5))
        self.e_search_peripherals = ttk.Entry(frm_filters, width=30)
        self.e_search_peripherals.pack(side="left")

        ttk.Button(frm_filters, text="Buscar", command=self.update_peripherals_table, style="Primary.TButton").pack(side="left", padx=10)
        ttk.Button(frm_filters, text="Limpar", command=self.cmd_clear_peripheral_filter, style="Secondary.TButton").pack(side="left")

        cols = ("ID", "Status", "Tipo", "Marca", "Modelo", "Fornecedor", "Nota Fiscal", "Identificador (S/N)")
        self.tree_peripherals = ttk.Treeview(frm_list, columns=cols, show="headings")
        self.tree_peripherals.pack(fill="both", expand=True)
        
        for c in cols:
            self.tree_peripherals.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_peripherals, col, False))
            self.tree_peripherals.column(c, anchor="center")
        
        # Define larguras específicas
        self.tree_peripherals.column("ID", width=50, stretch=False)
        self.tree_peripherals.column("Status", width=100, stretch=False)

        self.tree_peripherals.tag_configure("disp", background="#7EFF9C")
        self.tree_peripherals.tag_configure("emuso", background="#FFD966")
        self.tree_peripherals.tag_configure("defeito", background="#FF7E89")


    def build_linking_tab(self):
        tab = self.tab_linking

        # --- Frame de Seleção de Equipamento ---
        frm_select = ttk.Frame(tab, padding=(0, 0, 0, 10))
        frm_select.pack(fill="x")
        frm_select.grid_columnconfigure(1, weight=1)

        ttk.Label(frm_select, text="Selecione o Equipamento Principal:", font=FONT_BOLD).grid(row=0, column=0, sticky="e", padx=(0,10))
        self.cb_link_equip = ttk.Combobox(frm_select, state="readonly", values=[])
        self.cb_link_equip.grid(row=0, column=1, sticky="ew")
        self.cb_link_equip.bind("<<ComboboxSelected>>", self.cmd_load_equipment_for_linking)

        # --- Frame Principal com as duas listas ---
        frm_main = ttk.Frame(tab)
        frm_main.pack(fill="both", expand=True)
        frm_main.grid_columnconfigure(0, weight=1)
        frm_main.grid_columnconfigure(2, weight=1) # Faz as colunas das tabelas terem o mesmo peso

        # Frame Esquerdo: Vinculados
        frm_linked = ttk.LabelFrame(frm_main, text=" Periféricos Vinculados a este Equipamento ", padding=10)
        frm_linked.grid(row=0, column=0, sticky="nsew", padx=(0,5))
        
        cols_linked = ("ID Vínculo", "ID Per.", "Tipo", "Marca/Modelo", "Identificador")
        self.tree_linked_peripherals = ttk.Treeview(frm_linked, columns=cols_linked, show="headings")
        self.tree_linked_peripherals.pack(fill="both", expand=True)
        for c in cols_linked:
            self.tree_linked_peripherals.heading(c, text=c)
            self.tree_linked_peripherals.column(c, anchor="center")
        self.tree_linked_peripherals.column("ID Vínculo", width=0, stretch=False) # Oculta a coluna ID Vínculo
        self.tree_linked_peripherals.column("ID Per.", width=60, stretch=False)

        # Frame Central: Botões de Ação
        frm_actions = ttk.Frame(frm_main)
        frm_actions.grid(row=0, column=1, padx=10, sticky="ns")
        ttk.Button(frm_actions, text="< Desvincular", command=self.cmd_unlink_peripheral, style="Danger.TButton").pack(pady=5)
        ttk.Button(frm_actions, text="Vincular >", command=self.cmd_link_peripheral, style="Success.TButton").pack(pady=5)
        ttk.Separator(frm_actions, orient="horizontal").pack(fill="x", pady=20)
        ttk.Button(frm_actions, text="Substituir...", command=self.cmd_replace_peripheral_dialog, style="Secondary.TButton").pack(pady=5)
        
        # Frame Direito: Disponíveis
        frm_available = ttk.LabelFrame(frm_main, text=" Periféricos Disponíveis ", padding=10)
        frm_available.grid(row=0, column=2, sticky="nsew", padx=(5,0))
        
        cols_avail = ("ID", "Tipo", "Marca/Modelo", "Identificador")
        self.tree_available_peripherals = ttk.Treeview(frm_available, columns=cols_avail, show="headings")
        self.tree_available_peripherals.pack(fill="both", expand=True)
        for c in cols_avail:
            self.tree_available_peripherals.heading(c, text=c)
            self.tree_available_peripherals.column(c, anchor="center")
        self.tree_available_peripherals.column("ID", width=60, stretch=False)

        # Label de Mensagens
        self.lbl_link_msg = ttk.Label(tab, text="Selecione um equipamento para começar.", anchor="center")
        self.lbl_link_msg.pack(fill="x", pady=(10,0))

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
            ("Setor:", 'cb_setor', ttk.Combobox, {'state': 'readonly', 'values': SETORES_OPTIONS}),
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
        self.cb_setor.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        self.e_cargo.bind("<KeyRelease>", self.on_widget_interaction)
        self.cb_revenda.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        
        self.e_issue_cpf.bind("<KeyRelease>", self.on_cpf_entry)
        self.e_date_issue.bind("<KeyRelease>", lambda e: self.on_date_entry(e, self.e_date_issue))
        
        self.lbl_issue = ttk.Label(frm, text="", anchor="center")
        self.lbl_issue.grid(row=len(fields), column=0, columnspan=2, pady=10, sticky="ew")
        
        ttk.Button(frm, text="Emprestar", command=self.cmd_issue, style="Primary.TButton").grid(row=len(fields) + 1, column=0, columnspan=2, pady=10)


    def build_return_tab(self):
        # Variáveis para guardar os caminhos dos anexos
        self.signed_return_term_path = None
        
        # Usamos nossa classe rolável para toda a aba
        scroll_container = ScrollableFrame(self.tab_return)
        scroll_container.pack(fill="both", expand=True)

        # Todo o conteúdo vai dentro do 'scrollable_frame'
        tab = scroll_container.scrollable_frame
        
        # --- Frame Superior: Empréstimos Ativos (para gerar termo) ---
        frm_active = ttk.LabelFrame(tab, text=" Empréstimos Ativos (Para Gerar Termo de Devolução) ", padding=10)
        frm_active.pack(fill="x", expand=False, padx=10, pady=(10, 5))

        cols_active = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        self.tree_return_active = ttk.Treeview(frm_active, columns=cols_active, show="headings", height=8)
        self.tree_return_active.pack(fill="x", expand=True)
        for c in cols_active:
            self.tree_return_active.heading(c, text=c, anchor="center")
        self.tree_return_active.column("ID", width=50, stretch=False)
        self.tree_return_active.column("Usuário", width=200)

        frm_active_actions = ttk.Frame(frm_active)
        frm_active_actions.pack(fill="x", pady=(10, 0))
        
        ttk.Button(frm_active_actions, text="Gerar Termo de Devolução", command=self.cmd_generate_and_initiate_return, style="Primary.TButton").pack(side="left")

        # --- Frame Inferior: Devoluções Pendentes ---
        frm_pending = ttk.LabelFrame(tab, text=" Devoluções Pendentes (Aguardando Anexo do Termo Assinado) ", padding=10)
        frm_pending.pack(fill="x", expand=False, padx=10, pady=(5, 10))
        
        cols_pending = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        self.tree_return_pending = ttk.Treeview(frm_pending, columns=cols_pending, show="headings", height=8)
        self.tree_return_pending.pack(fill="both", expand=True)
        for c in cols_pending:
            self.tree_return_pending.heading(c, text=c, anchor="center")
        self.tree_return_pending.column("ID", width=50, stretch=False)
        self.tree_return_pending.column("Usuário", width=200)
        
        # Ações para pendentes
        frm_pending_actions = ttk.Frame(frm_pending)
        frm_pending_actions.pack(fill="x", pady=(10, 0))
        
        ttk.Button(frm_pending_actions, text="Anexar Termo de Devolução Assinado (PDF)...", command=self.select_signed_return_term_attachment).pack(side="left")
        self.lbl_signed_return_term = ttk.Label(frm_pending_actions, text=" Nenhum arquivo selecionado.", width=40)
        self.lbl_signed_return_term.pack(side="left", padx=10)

        ttk.Button(frm_pending_actions, text="Confirmar Devolução", command=self.cmd_confirm_return, style="Success.TButton").pack(side="left", padx=(10,0))
        
        # Label para mensagens
        self.lbl_ret = ttk.Label(tab, text="", anchor="center")
        self.lbl_ret.pack(pady=5, fill="x")

    def build_remove_tab(self):
        container = ttk.Frame(self.tab_remove)
        container.pack(anchor="n", pady=20)
        
        frm = ttk.Frame(container)
        frm.pack()
        frm.grid_columnconfigure(1, weight=1)
        
        # --- Campo de Seleção do Aparelho ---
        ttk.Label(frm, text="Selecione aparelho:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.cb_remove = ttk.Combobox(frm, state='readonly')
        self.cb_remove.grid(row=0, column=1, pady=5, padx=5, sticky="ew")
        
        # --- Campo de Anexo ---
        ttk.Label(frm, text="Anexar Nota (PDF):").grid(row=1, column=0, sticky='e', padx=5, pady=8)
        
        frm_anexo = ttk.Frame(frm)
        frm_anexo.grid(row=1, column=1, sticky="ew", padx=5, pady=8)
        frm_anexo.grid_columnconfigure(1, weight=1) 
        
        self.btn_anexo = ttk.Button(frm_anexo, text="Selecionar Arquivo...", command=self.select_removal_attachment)
        # Coloca o botão na primeira coluna (coluna 0)
        self.btn_anexo.grid(row=0, column=0, sticky='w') 
        
        self.lbl_remove_attachment = ttk.Label(frm_anexo, text=" Nenhum arquivo selecionado.", anchor="w")
        # Coloca o label na segunda coluna (coluna 1), fazendo-o esticar (sticky='ew')
        self.lbl_remove_attachment.grid(row=0, column=1, sticky='ew', padx=(10, 0)) 
        
        self.remove_attachment_path = None

        self.lbl_rem = ttk.Label(frm, text="", anchor="center")
        self.lbl_rem.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

        ttk.Button(frm, text="Remover", command=self.cmd_remove, style="Danger.TButton").grid(row=3, column=0, columnspan=2, pady=10)

    def select_removal_attachment(self): 
        # =======================================
        # Função para selecionar o arquivo da nota de remoção
        # =======================================
        """Abre uma caixa de diálogo para o usuário selecionar um arquivo PDF."""
        file_path = filedialog.askopenfilename(
            title="Selecione a Nota Fiscal de Remoção",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        
        if file_path:
            self.remove_attachment_path = file_path
            # Mostra apenas o nome do arquivo, e não o caminho completo
            filename = os.path.basename(file_path)
            self.lbl_remove_attachment.config(text=f" {filename}")
        else:
            self.remove_attachment_path = None
            self.lbl_remove_attachment.config(text=" Nenhum arquivo selecionado.")

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
            "ID Item", "Id Per.", "Operador", "Operação", "Revenda", "Data", "Hora", "Tipo", "Marca", "Modelo", "Nota Fiscal", "Fornecedor",
            "Identificador", "Usuário", "CPF", "Cargo", "Centro de Custo", "Setor", "Detalhes"
        )
        
        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill="both", expand=True)
        
        self.tree_history = ttk.Treeview(tree_frame, columns=cols, show="headings")

        col_widths = { "ID Item": 30, "ID Per.": 30, "Operador": 100, "Operação": 140, "Data": 90, "Hora": 70, "CPF": 110, "Detalhes": 200, "Setor": 100}

        for c in cols:
            self.tree_history.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_history, col, False))
            self.tree_history.column(c, width=col_widths.get(c, 120), anchor="center", stretch=False)

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
        
        # adicionado a coluna "ID Histórico" que ficará oculta
        cols = ("ID Histórico", "ID Item", "Operador", "Revenda", "Tipo", "Marca", "Modelo", "Nota Fiscal", "Fornecedor", "Identificador", "Usuário", "CPF", "Operação", "Data Empréstimo", "Data Devolução", "Centro de Custo", "Setor", "Cargo")
        
        tree_frame = ttk.Frame(frm)
        tree_frame.pack(fill="both", expand=True)

        self.tree_report = ttk.Treeview(tree_frame, columns=cols, show='headings')

        col_widths = {"ID Item": 60, "Operador": 100, "Operação": 110, "CPF": 120}

        for c in cols:
            self.tree_report.heading(c, text=c, command=lambda col=c: self.treeview_sort_column(self.tree_report, col, False))
            self.tree_report.column(c, width=col_widths.get(c, 120), anchor='center', stretch=False)
        
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
        # Variável para guardar o caminho do termo assinado selecionado
        self.signed_term_path = None
        
        tab = self.tab_terms
        
        # --- Frame Superior: Empréstimos Pendentes ---
        frm_pending = ttk.LabelFrame(tab, text=" Empréstimos Pendentes (Aguardando Termo Assinado) ", padding=10)
        frm_pending.pack(fill="x", expand=False, padx=10, pady=(10, 5))

        # Tabela de pendentes
        cols_pending = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        self.tree_terms_pending = ttk.Treeview(frm_pending, columns=cols_pending, show="headings", height=8)
        self.tree_terms_pending.pack(fill="x", expand=True)
        for c in cols_pending:
            self.tree_terms_pending.heading(c, text=c, anchor="w")
        self.tree_terms_pending.column("ID", width=50, stretch=False)
        self.tree_terms_pending.column("Usuário", width=200)

        # Ações para pendentes
        frm_pending_actions = ttk.Frame(frm_pending)
        frm_pending_actions.pack(fill="x", pady=(10, 0))
        
        ttk.Button(frm_pending_actions, text="Gerar Termo de Responsabilidade", command=self.cmd_generate_term, style="Primary.TButton").pack(side="left")
        
        ttk.Separator(frm_pending_actions, orient="vertical").pack(side="left", fill="y", padx=20, pady=5)
        
        ttk.Button(frm_pending_actions, text="Anexar Termo Assinado (PDF)...", command=self.select_signed_term_attachment).pack(side="left")
        self.lbl_signed_term = ttk.Label(frm_pending_actions, text=" Nenhum arquivo selecionado.", width=40)
        self.lbl_signed_term.pack(side="left", padx=10)

        ttk.Button(frm_pending_actions, text="Confirmar Empréstimo", command=self.cmd_confirm_loan, style="Success.TButton").pack(side="left", padx=(10,0))


        # --- Frame Inferior: Empréstimos Ativos ---
        frm_active = ttk.LabelFrame(tab, text=" Empréstimos Ativos (Termo OK) ", padding=10)
        frm_active.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        
        cols_active = ("ID", "Tipo", "Marca", "Usuário", "CPF", "Data Empréstimo", "Revenda")
        tree_frame = ttk.Frame(frm_active)
        tree_frame.pack(fill="both", expand=True)

        self.tree_terms_active = ttk.Treeview(tree_frame, columns=cols_active, show="headings")
        self.tree_terms_active.pack(side="left", fill="both", expand=True)
        for c in cols_active:
            self.tree_terms_active.heading(c, text=c, anchor="w")
        self.tree_terms_active.column("ID", width=50, stretch=False)
        self.tree_terms_active.column("Usuário", width=200)
        
        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_terms_active.yview)
        ysb.pack(side="right", fill="y")
        self.tree_terms_active.configure(yscrollcommand=ysb.set)
        
        # Label para mensagens
        self.lbl_terms = ttk.Label(tab, text="", anchor="center")
        self.lbl_terms.pack(pady=5, fill="x")

    def select_signed_term_attachment(self):
        """Abre uma caixa de diálogo para selecionar o termo assinado em PDF."""
        file_path = filedialog.askopenfilename(
            title="Selecione o Termo de Responsabilidade Assinado",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if file_path:
            self.signed_term_path = file_path
            filename = os.path.basename(file_path)
            self.lbl_signed_term.config(text=f" {filename}")
        else:
            self.signed_term_path = None
            self.lbl_signed_term.config(text=" Nenhum arquivo selecionado.")

    def cmd_confirm_loan(self):
        """Confirma um empréstimo pendente após anexar o termo."""
        selected = self.tree_terms_pending.selection()
        if not selected:
            self.lbl_terms.config(text="Selecione um empréstimo pendente na tabela acima.", style='Danger.TLabel')
            return

        if not self.signed_term_path:
            self.lbl_terms.config(text="É obrigatório anexar o termo assinado (PDF) para confirmar.", style='Danger.TLabel')
            return

        item_id = int(self.tree_terms_pending.item(selected[0])["values"][0])
        
        ok, msg = self.inv.confirm_loan(item_id, self.logged_user, self.signed_term_path)
        self.lbl_terms.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        
        if ok:
            # Limpa os campos e atualiza tudo
            self.signed_term_path = None
            self.lbl_signed_term.config(text=" Nenhum arquivo selecionado.")
            self.update_all_views()

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
        
        
    def build_users_tab(self):
        tab = self.tab_users

        # --- Frame da Esquerda (Lista de Usuários) ---
        frm_left = ttk.Frame(tab, padding=10)
        frm_left.pack(side="left", fill="both", expand=True, padx=(0, 5))

        ttk.Label(frm_left, text="Usuários Cadastrados", style="Title.TLabel").pack(pady=(0, 10), anchor="center")

        cols = ("ID", "Usuário", "Função")
        self.tree_users = ttk.Treeview(frm_left, columns=cols, show="headings")
        self.tree_users.pack(fill="both", expand=True)

        for col in cols:
            self.tree_users.heading(col, text=col)
        self.tree_users.column("ID", width=50, stretch=False)
        self.tree_users.column("Usuário", width=200)
        self.tree_users.column("Função", width=150)

        # --- Frame da Direita (Ações) ---
        frm_right = ttk.Frame(tab, padding=10)
        frm_right.pack(side="right", fill="y", padx=(5, 0))

        # -- Sub-frame para Adicionar Usuário --
        frm_add = ttk.LabelFrame(frm_right, text=" Cadastrar Novo Usuário ", padding=15)
        frm_add.pack(fill="x", pady=(0, 20))

        ttk.Label(frm_add, text="Nome de Usuário:").grid(row=0, column=0, sticky="w", pady=2)
        self.e_new_username = ttk.Entry(frm_add, width=30)
        self.e_new_username.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(frm_add, text="Senha:").grid(row=2, column=0, sticky="w", pady=2)
        self.e_new_password = ttk.Entry(frm_add, width=30, show="*")
        self.e_new_password.grid(row=3, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(frm_add, text="Função:").grid(row=4, column=0, sticky="w", pady=2)
        self.cb_new_role = ttk.Combobox(frm_add, values=["Técnico", "Jovem Aprendiz", "Gestor"], state="readonly")
        self.cb_new_role.grid(row=5, column=0, sticky="ew", pady=(0, 15))

        ttk.Button(frm_add, text="Cadastrar Usuário", command=self.cmd_add_user, style="Primary.TButton").grid(row=6, column=0, sticky="ew")

        # -- Sub-frame para Ações no Usuário Selecionado --
        frm_actions = ttk.LabelFrame(frm_right, text=" Ações no Usuário Selecionado ", padding=15)
        frm_actions.pack(fill="x", expand=True)

        ttk.Button(frm_actions, text="Alterar Senha", command=self.cmd_change_password, style="Secondary.TButton").pack(fill="x", pady=5)
        ttk.Button(frm_actions, text="Remover Usuário", command=self.cmd_remove_user, style="Danger.TButton").pack(fill="x", pady=5)

        self.lbl_users = ttk.Label(frm_right, text="", wraplength=250, anchor="center")
        self.lbl_users.pack(pady=10, fill="x", expand=True, side="bottom")
        
        
    

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
        
        # --- CAMPOS COMUNS ---
        # Marca
        ttk.Label(parent_frame, text="Marca:").grid(row=0, column=0, sticky="e", pady=5, padx=5)
        e_brand = ttk.Entry(parent_frame)
        e_brand.grid(row=0, column=1, pady=5, padx=5, sticky="ew")
        e_brand.insert(0, item_data.get('brand') or '')
        e_brand.bind("<KeyRelease>", self.on_widget_interaction)
        widgets['brand'] = e_brand

        # Modelo
        ttk.Label(parent_frame, text="Modelo:").grid(row=1, column=0, sticky="e", pady=5, padx=5)
        e_model = ttk.Entry(parent_frame)
        e_model.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
        e_model.insert(0, item_data.get('model') or '')
        e_model.bind("<KeyRelease>", self.on_widget_interaction)
        widgets['model'] = e_model

        # Revenda
        ttk.Label(parent_frame, text="Revenda:").grid(row=2, column=0, sticky="e", pady=5, padx=5)
        cb_revenda = ttk.Combobox(parent_frame, values=REVENDAS_OPTIONS, state="readonly")
        cb_revenda.grid(row=2, column=1, pady=5, padx=5, sticky="ew")
        cb_revenda.set(item_data.get('revenda') or '')
        cb_revenda.bind("<KeyRelease>", self.on_widget_interaction)
        cb_revenda.bind("<<ComboboxSelected>>", self.on_widget_interaction)
        widgets['revenda'] = cb_revenda
        
        # Nota Fiscal
        ttk.Label(parent_frame, text="Nota Fiscal:").grid(row=3, column=0, sticky="e", pady=5, padx=5)
        vcmd_nf = (self.register(self._validate_numeric_input), '%P', '9')
        e_nota_fiscal = ttk.Entry(parent_frame, validate="key", validatecommand=vcmd_nf)
        e_nota_fiscal.grid(row=3, column=1, pady=5, padx=5, sticky="ew")
        e_nota_fiscal.insert(0, item_data.get('nota_fiscal') or '')
        e_nota_fiscal.bind("<KeyRelease>", self.on_widget_interaction)
        widgets['nota_fiscal'] = e_nota_fiscal

        # Fornecedor
        ttk.Label(parent_frame, text="Fornecedor:").grid(row=4, column=0, sticky="e", pady=5, padx=5)
        e_fornecedor = ttk.Entry(parent_frame)
        e_fornecedor.grid(row=4, column=1, pady=5, padx=5, sticky="ew")
        e_fornecedor.insert(0, item_data.get('fornecedor') or '')
        e_fornecedor.bind("<KeyRelease>", self.on_widget_interaction)
        widgets['fornecedor'] = e_fornecedor

        # DATA DE CADASTRO
        ttk.Label(parent_frame, text="Data de Cadastro:").grid(row=5, column=0, sticky="e", pady=5, padx=5)
        e_date_registered = ttk.Entry(parent_frame)
        e_date_registered.grid(row=5, column=1, pady=5, padx=5, sticky="ew")
        widgets['date_registered'] = e_date_registered
        
        if item_data: # Se estamos editando um item...
            # Formata a data que vem do banco e insere no campo
            e_date_registered.insert(0, format_date(item_data.get('date_registered')))
            # Bloqueia a edição da data de cadastro original
            e_date_registered.config(state="readonly")
        else: # Se estamos cadastrando um novo item...
            # Aplica a máscara de data dd/mm/aaaa
            e_date_registered.bind("<KeyRelease>", lambda e: self.on_date_entry(e, e_date_registered))
            e_date_registered.bind("<KeyRelease>", self.on_widget_interaction, add="+")

        # --- FIM DOS CAMPOS COMUNS ---

        # Campos específicos
        row_start = 6
        if tipo == "Celular":
            
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
            fields = {'identificador': ("Nº de Série:", ttk.Entry, {}), 'storage': ("Armazenamento (GB):", ttk.Entry, {}) }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts); widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                widget.insert(0, item_data.get(key) or '')
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                widgets[key] = widget
                
        elif tipo == "Switch":
            fields = {
                'poe': ("POE:", ttk.Combobox, {"values": ["Sim", "Não"], "state": "readonly"}),
                'quantidade_portas': ("Qtd. de Portas:", ttk.Entry, {})
            }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts)
                widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                
                if isinstance(widget, ttk.Combobox):
                    widget.set(item_data.get(key) or '')
                    widget.bind("<<ComboboxSelected>>", self.on_widget_interaction)
                else:
                    widget.insert(0, item_data.get(key) or '')
                
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                widgets[key] = widget

        # --- BLOCO ADICIONADO PARA O HD ---
        elif tipo == "HD":
            fields = {
                'storage': ("Armazenamento (GB):", ttk.Entry, {})
            }
            for i, (key, (label, w_class, opts)) in enumerate(fields.items()):
                ttk.Label(parent_frame, text=label).grid(row=row_start + i, column=0, sticky="e", pady=5, padx=5)
                widget = w_class(parent_frame, **opts)
                widget.grid(row=row_start + i, column=1, pady=5, padx=5, sticky="ew")
                widget.insert(0, item_data.get(key) or '')
                widget.bind("<KeyRelease>", self.on_widget_interaction)
                widgets[key] = widget

        return widgets


    def cmd_add(self, event=None):
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

        if 'brand' in dados:
            dados['brand'] = format_title_case(dados['brand'])
        if 'model' in dados:
            dados['model'] = format_title_case(dados['model'])
        if 'fornecedor' in dados:
            dados['fornecedor'] = format_title_case(dados['fornecedor'])
        if 'sistema' in dados:
            dados['sistema'] = format_title_case(dados['sistema'])

        erros = []
        if not dados.get("brand"): erros.append((self.add_widgets['brand'], "Informe a marca."))
        if not dados.get("revenda"): erros.append((self.add_widgets['revenda'], "Informe o campo Revenda."))
        if not dados.get("model"): erros.append((self.add_widgets['model'], "Informe o modelo."))
        if not dados.get("fornecedor"): erros.append((self.add_widgets['fornecedor'], "Informe o fornecedor."))
        nota_fiscal = dados.get("nota_fiscal")
        if not nota_fiscal:
            erros.append((self.add_widgets['nota_fiscal'], "Informe a Nota Fiscal."))
        elif not nota_fiscal.isdigit() or len(nota_fiscal) != 9:
            erros.append((self.add_widgets['nota_fiscal'], "A Nota Fiscal deve conter exatamente 9 números."))

        date_str = dados.get("date_registered")
        if not date_str:
            erros.append((self.add_widgets['date_registered'], "Informe a data de cadastro."))
        else:
            try:
                # Tenta converter a data para garantir que o formato está correto
                # e já a armazena no formato que o banco de dados entende
                dt_obj = datetime.strptime(date_str, "%d/%m/%Y")
                dados['date_registered'] = dt_obj
            except ValueError:
                erros.append((self.add_widgets['date_registered'], "Formato de data inválido. Use dd/mm/aaaa."))

            
        if tipo in ["Celular", "Tablet"]:
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

        if tipo == "Switch":
            for campo, msg in {"poe": "Informe se possui POE.", "quantidade_portas": "Informe a quantidade de portas."}.items():
                if not dados.get(campo):
                    erros.append((self.add_widgets[campo], msg))

        if tipo == "HD":
            for campo, msg in {"storage": "Informe o armazenamento."}.items():
                if not dados.get(campo):
                    erros.append((self.add_widgets[campo], msg))
        
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

        # A função add_item retorna apenas o ID ou None/False em caso de erro.
        item_id = self.inv.add_item(dados, self.logged_user)
        
        # Se o backend retornar um ID válido (um número), o cadastro deu certo.
        if isinstance(item_id, int):
            ok = True
            result = item_id
        else:
            # Se não, item_id contém a mensagem de erro (que é uma string)
            ok = False
            result = item_id

        ok, result = self.inv.add_item(dados, self.logged_user)
        
        if not ok:
            # Se 'ok' for False, 'result' é a mensagem de erro do backend
            self.lbl_add.config(text=result, style="Danger.TLabel")
            return
        
        # Se 'ok' for True, 'result' é o item_id
        item_id = result
        
        for widget in self.add_widgets.values():
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)

        self.lbl_add.config(text=f"{tipo} ID {item_id} cadastrado com sucesso!", style="Success.TLabel")
        self.cb_tipo.set("")
        
        for widget in self.frm_dynamic.winfo_children():
            widget.destroy()
        self.update_all_views()



    def cmd_add_peripheral(self):
    
        identificador_str = self.e_peri_id.get().strip()

        data = {
            "tipo": self.cb_peri_tipo.get(),
            "brand": self.e_peri_brand.get().strip(),
            "model": self.e_peri_model.get().strip(),
            "nota_fiscal": self.e_peri_nf.get().strip(),
            "fornecedor": self.e_peri_fornecedor.get(),
            "identificador": identificador_str if identificador_str else None
        }

        data['brand'] = format_title_case(data['brand'])
        data['model'] = format_title_case(data['model'])
        data['fornecedor'] = format_title_case(data['fornecedor'])


        if not data["tipo"]:
            self.lbl_peri_add.config(text="O campo 'Tipo' é obrigatório.", style="Danger.TLabel")
            return
        if not data["nota_fiscal"]:
            self.lbl_peri_add.config(text="O campo 'Nota Fiscal' é obrigatório.", style="Danger.TLabel")
            return
        if not data["fornecedor"]:
            self.lbl_peri_add.config(text="O campo 'Fornecedor' é obrigatório.", style="Danger.TLabel")
            return
        if not data["identificador"]:
            self.lbl_peri_add.config(text="O campo 'Identificador' é obrigatório.", style="Danger.TLabel")
            return
            
        ok, msg = self.inv.add_peripheral(data, self.logged_user)
        self.lbl_peri_add.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        if ok:
            self.cb_peri_tipo.set("")
            self.e_peri_brand.delete(0, "end")
            self.e_peri_model.delete(0, "end")
            self.e_peri_id.delete(0, "end")
            self.e_peri_nf.delete(0, "end")
            self.e_peri_fornecedor.delete(0, "end")
            self.update_all_views()

    def cmd_clear_peripheral_filter(self):
        """Limpa o campo de busca da aba de periféricos e atualiza a tabela."""
        self.e_search_peripherals.delete(0, "end")
        self.update_peripherals_table()


    def cmd_load_equipment_for_linking(self, event=None):
        selection = self.cb_link_equip.get()
        if not selection:
            return
        
        equipment_id = int(selection.split(' - ')[0])

        self.tree_linked_peripherals.delete(*self.tree_linked_peripherals.get_children())
        self.tree_available_peripherals.delete(*self.tree_available_peripherals.get_children())
        self.lbl_link_msg.config(text="")


        linked_peripherals = list(self.inv.list_peripherals_for_equipment(equipment_id))
        linked_peripherals.sort(key=lambda p: p['tipo']) 
        for p in linked_peripherals: 
            self.tree_linked_peripherals.insert("", "end", values=(
                p['link_id'], p['id'], p.get('tipo'), f"{p.get('brand','')} {p.get('model','')}", p.get('identificador')
            ))
        
        available_peripherals = list(self.inv.list_peripherals(status_filter="Disponível"))
        available_peripherals.sort(key=lambda p: p['tipo']) 
        for p in available_peripherals:
            self.tree_available_peripherals.insert("", "end", values=(
                p['id'], p.get('tipo'), f"{p.get('brand','')} {p.get('model','')}", p.get('identificador')
            ))

    def cmd_link_peripheral(self):
        equip_selection = self.cb_link_equip.get()
        peri_selection = self.tree_available_peripherals.selection()

        if not equip_selection or not peri_selection:
            self.lbl_link_msg.config(text="Selecione um equipamento E um periférico disponível para vincular.", style="Danger.TLabel")
            return

        equipment_id = int(equip_selection.split(' - ')[0])
        peripheral_id = int(self.tree_available_peripherals.item(peri_selection[0])['values'][0])

        ok, msg = self.inv.link_peripheral_to_equipment(equipment_id, peripheral_id, self.logged_user)
        self.lbl_link_msg.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        if ok:
            self.cmd_load_equipment_for_linking() # Recarrega as listas
            self.update_stock() # Atualiza o contador na aba de estoque

    def cmd_unlink_peripheral(self):
        selection = self.tree_linked_peripherals.selection()
        if not selection:
            self.lbl_link_msg.config(text="Selecione um periférico vinculado para desvincular.", style="Danger.TLabel")
            return

        link_id = int(self.tree_linked_peripherals.item(selection[0])['values'][0])
        
        ok, msg = self.inv.unlink_peripheral_from_equipment(link_id, self.logged_user)
        self.lbl_link_msg.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        if ok:
            self.cmd_load_equipment_for_linking()
            self.update_stock()

    def cmd_replace_peripheral_dialog(self):
        equip_selection = self.cb_link_equip.get()
        old_peri_selection = self.tree_linked_peripherals.selection()

        if not equip_selection or not old_peri_selection:
            self.lbl_link_msg.config(text="Selecione um equipamento E o periférico a ser substituído.", style="Danger.TLabel")
            return
        
        equipment_id = int(equip_selection.split(' - ')[0])
        values = self.tree_linked_peripherals.item(old_peri_selection[0])['values']
        old_peripheral_id = int(values[1])
        peripheral_type = values[2]

        # Abre a janela de diálogo para substituição
        dialog = ReplacePeripheralDialog(self, equipment_id, old_peripheral_id, peripheral_type)
        self.wait_window(dialog) # Pausa a janela principal até o diálogo ser fechado

        if dialog.success:
            self.lbl_link_msg.config(text="Substituição realizada com sucesso.", style="Success.TLabel")
            self.cmd_load_equipment_for_linking() # Recarrega tudo
            self.update_peripherals_table()
        else:
            self.lbl_link_msg.config(text="", style="TLabel")


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


    def cmd_save_edit(self, event=None):
        if self.current_edit_id is None:
            messagebox.showerror("Erro", "Nenhum item foi carregado para edição.")
            return

        # Limpa estilos de erro anteriores
        for widget in self.edit_widgets.values():
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)

        new_data = {
            key: widget.get().strip()
            for key, widget in self.edit_widgets.items()
            if widget.winfo_exists() and key != 'date_registered'
        }
        
        if 'brand' in new_data:
            new_data['brand'] = format_title_case(new_data['brand'])
        if 'model' in new_data:
            new_data['model'] = format_title_case(new_data['model'])
        if 'sistema' in new_data:
            new_data['sistema'] = format_title_case(new_data['sistema'])

        erros = []
        if not new_data.get("brand"): erros.append((self.edit_widgets['brand'], "Informe a marca."))
        if not new_data.get("revenda"): erros.append((self.edit_widgets['revenda'], "Informe a revenda."))
        
        nota_fiscal = new_data.get("nota_fiscal")
        if not nota_fiscal:
            erros.append((self.edit_widgets['nota_fiscal'], "Informe a Nota Fiscal."))
        elif not nota_fiscal.isdigit() or len(nota_fiscal) != 9:
            erros.append((self.edit_widgets['nota_fiscal'], "A Nota Fiscal deve conter exatamente 9 números."))

        if erros:
            for widget, _ in erros:
                widget_class = str(widget.winfo_class())
                if 'TCombobox' in widget_class: widget.configure(style="Error.TCombobox")
                elif 'TEntry' in widget_class: widget.configure(style="Error.TEntry")
            
            self.lbl_edit.config(text=erros[0][1], style="Danger.TLabel") # Mostra o primeiro erro
            return


        ok, msg = self.inv.update_item(self.current_edit_id, new_data, self.logged_user)
        self.lbl_edit.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        
        if ok:
            for widget in self.frm_edit_dynamic.winfo_children():
                widget.destroy()
            self.current_edit_id = None
            self.cb_edit.set('')
            self.update_all_views()

    def cmd_issue(self, event=None):
        self.lbl_issue.config(text="")
        widgets_to_reset = [self.cb_issue, self.e_issue_user, self.e_issue_cpf, self.cb_center, self.e_cargo, self.cb_revenda, self.e_date_issue]
        for widget in widgets_to_reset:
            style_name = str(widget.cget('style')).replace('Error.', '')
            widget.configure(style=style_name)

        sel = self.cb_issue.get()
        user = self.e_issue_user.get().strip()
        cpf = self.e_issue_cpf.get().strip()
        cc = self.cb_center.get().strip()
        setor = self.cb_setor.get().strip()
        cargo = self.e_cargo.get().strip()
        revenda = self.cb_revenda.get().strip()
        date_issue = self.e_date_issue.get().strip()

        user = format_title_case(user)
        cargo = format_title_case(cargo)

        erros = []
        if not sel: erros.append((self.cb_issue, "Selecione um aparelho."))
        if not user: erros.append((self.e_issue_user, "Informe o nome do funcionário."))
        if not cc: erros.append((self.cb_center, "Selecione o centro de custo."))
        if not setor: erros.append((self.cb_setor, "Selecione o setor."))
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
        ok, msg = self.inv.issue(pid, user, cpf, cc, setor, cargo, revenda, date_issue, self.logged_user)
        self.lbl_issue.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        if ok:
            self.cb_issue.set('')
            self.e_issue_user.delete(0, 'end')
            self.e_issue_cpf.delete(0, 'end')
            self.cb_center.set('')
            self.cb_setor.set('')
            self.e_cargo.delete(0, 'end')
            self.cb_revenda.set('')
            self.e_date_issue.delete(0, 'end')
            messagebox.showinfo("Status: Pendente", "Empréstimo iniciado!\n\nVá para a aba 'Termos' para gerar o termo de responsabilidade e confirmar o empréstimo.")
            self.update_all_views()



    # ANTIGA A FUNÇÃO cmd_return REMOVIDA (ATO DE DEVOLUÇÃO SÃO AS PRÓXIMAS 4 FUNCOES)

    def cmd_generate_and_initiate_return(self):
        selected = self.tree_return_active.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um empréstimo ativo na lista para gerar o termo.")
            return

        item_id = int(self.tree_return_active.item(selected[0])["values"][0])
        
        # Chama a nova função única do backend
        ok, result = self.inv.generate_and_initiate_return(item_id, self.logged_user)
        
        if ok:
            self.lbl_ret.config(text=f"Termo de devolução gerado. Status alterado para 'Pendente Devolução'.", style='Success.TLabel')
            self.update_all_views() # Atualiza as listas
            
            # Pergunta se o usuário quer abrir o arquivo gerado
            if messagebox.askyesno("Sucesso", f"Termo gerado em:\n{result}\n\nDeseja abri-lo agora?"):
                try:
                    os.startfile(result)
                except Exception as e:
                    messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo automaticamente:\n{e}")
        else:
            # Se der erro, mostra a mensagem e não faz mais nada
            messagebox.showerror("Erro ao Gerar Termo", result)
            self.lbl_ret.config(text=f"Erro: {result}", style='Danger.TLabel')

    def select_signed_return_term_attachment(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o Termo de Devolução Assinado",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if file_path:
            self.signed_return_term_path = file_path
            filename = os.path.basename(file_path)
            self.lbl_signed_return_term.config(text=f" {filename}")
        else:
            self.signed_return_term_path = None
            self.lbl_signed_return_term.config(text=" Nenhum arquivo selecionado.")

    def cmd_confirm_return(self):
        selected = self.tree_return_pending.selection()
        if not selected:
            self.lbl_ret.config(text="Selecione uma devolução pendente na tabela para confirmar.", style='Danger.TLabel')
            return

        if not self.signed_return_term_path:
            self.lbl_ret.config(text="É obrigatório anexar o termo de devolução assinado para confirmar.", style='Danger.TLabel')
            return

        item_id = int(self.tree_return_pending.item(selected[0])["values"][0])
        
        ok, msg = self.inv.confirm_return(item_id, self.logged_user, self.signed_return_term_path)
        self.lbl_ret.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        
        if ok:
            self.signed_return_term_path = None
            self.lbl_signed_return_term.config(text=" Nenhum arquivo selecionado.")
            self.update_all_views()
            

    def cmd_remove(self, event=None):
        sel = self.cb_remove.get()
        if not sel:
            self.lbl_rem.config(text="Selecione um aparelho para remover.", style="Danger.TLabel"); return
        
        # --- NOVA VERIFICAÇÃO ---
        if not self.remove_attachment_path:
            self.lbl_rem.config(text="É obrigatório anexar a nota fiscal (PDF) para remover.", style="Danger.TLabel")
            return
        # ------------------------

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
            
        # Passe o caminho do anexo para a função 'remove'
        ok, msg = self.inv.remove(pid, self.logged_user, self.remove_attachment_path)
        self.lbl_rem.config(text=msg, style='Success.TLabel' if ok else 'Danger.TLabel')
        
        if ok:
            # Limpa os campos após o sucesso
            self.remove_attachment_path = None
            self.lbl_remove_attachment.config(text=" Nenhum arquivo selecionado.")
            self.update_all_views()


    def cmd_generate_report(self):
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
            data_confirmacao = log.get('data_confirmacao')
            
            operation_display = ''
            data_inicial_display = format_date(log.get('data_emprestimo'))
            data_devolucao_display = ''

            if op_type == 'Empréstimo':
                if data_devolucao:
                    operation_display = "Devolvido"
                    data_devolucao_display = format_date(data_devolucao)
                elif data_confirmacao:
                    operation_display = "Confirmado"
                else:
                    operation_display = "Pendente"
            elif op_type == 'Cadastro':
                operation_display = "Cadastro"
            # -----------------------------
            
            row_values = (
            log.get('history_id'),            
            log.get('item_id'),                   
            log.get('operador'),                  
            log.get('revenda'),                   
            log.get('tipo'),                      
            log.get('brand'),                     
            log.get('model'),                     
            log.get('nota_fiscal'),
            log.get('fornecedor'),             
            log.get('identificador'),             
            log.get('usuario'),                   
            format_cpf(log.get('cpf')),           
            operation_display,                    
            data_inicial_display,                 
            data_devolucao_display,               
            log.get('center_cost'),
            log.get('setor'),             
            log.get('cargo')
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
        selected = self.tree_terms_pending.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um empréstimo pendente na tabela para gerar o termo.")
            return

        values = self.tree_terms_pending.item(selected[0])["values"]
        pid, user = int(values[0]), values[3]
        
        ok, result = self.inv.generate_term(pid, user)
        
        if ok:
            self.lbl_terms.config(text=f"Termo gerado: {os.path.basename(result)}", style='Success.TLabel')
            if messagebox.askyesno("Sucesso", f"Termo gerado em:\n{result}\n\nDeseja abri-lo agora?"):
                try:
                    os.startfile(result)
                except Exception as e:
                    messagebox.showerror("Erro ao Abrir", f"Não foi possível abrir o arquivo automaticamente:\n{e}")
        else:
            messagebox.showerror("Erro ao Gerar Termo", result)
            self.lbl_terms.config(text=f"Erro: {result}", style='Danger.TLabel')            
            
    # --- Comandos e Ações da Aba Usuários (NOVOS) ---

    def cmd_add_user(self, event=None):
        username = self.e_new_username.get().strip()
        password = self.e_new_password.get().strip()
        role = self.cb_new_role.get()

        ok, msg = self.user_db.add_user(username, password, role)
        self.lbl_users.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")

        if ok:
            self.e_new_username.delete(0, "end")
            self.e_new_password.delete(0, "end")
            self.cb_new_role.set("")
            self.update_users_table()

    def cmd_remove_user(self):
        selected = self.tree_users.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um usuário na lista para remover.")
            return

        item = self.tree_users.item(selected[0])
        user_id = item["values"][0]
        username = item["values"][1]

        if user_id == self.logged_user_id:
            messagebox.showerror("Ação Inválida", "Você não pode remover seu próprio usuário.")
            return

        if messagebox.askyesno("Confirmar Remoção", f"Tem certeza que deseja remover o usuário '{username}'?\nEsta ação não pode ser desfeita."):
            ok, msg = self.user_db.remove_user(user_id)
            self.lbl_users.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
            if ok:
                self.update_users_table()

    def cmd_change_password(self):
        selected = self.tree_users.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um usuário na lista para alterar a senha.")
            return

        item = self.tree_users.item(selected[0])
        user_id = item["values"][0]
        username = item["values"][1]

        new_password = simpledialog.askstring("Alterar Senha", f"Digite a NOVA senha para o usuário '{username}':", show='*')

        if new_password: # Se o usuário não clicou em "Cancelar"
            ok, msg = self.user_db.update_password(user_id, new_password)
            self.lbl_users.config(text=msg, style="Success.TLabel" if ok else "Danger.TLabel")
        else:
            self.lbl_users.config(text="Alteração de senha cancelada.", style="TLabel")

    def update_users_table(self):
        """Atualiza a tabela de usuários com os dados do banco."""
        self.tree_users.delete(*self.tree_users.get_children())
        users = self.user_db.get_all_users()
        for user in users:
            self.tree_users.insert("", "end", values=(user['id'], user['username'], user['role']))

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
                p.get('model'), p.get('status'), p.get('peripheral_count'), p.get('assigned_to'), 
                format_cpf(p.get('cpf')), p.get('nota_fiscal'), p.get('fornecedor'), p.get('identificador'), p.get('dominio'), 
                p.get('host'), p.get('endereco_fisico'), p.get('cpu'), p.get('ram'), 
                p.get('storage'), p.get('sistema'), p.get('licenca'), p.get('anydesk'), 
                p.get('setor'), p.get('ip'), p.get('mac'),
                p.get('poe'), p.get('quantidade_portas'),
                format_date(p.get('date_registered'))
            )
            
            cleaned_row = tuple(v or '' for v in row)
            
            status = p.get('status')
            tag = ""
            if status == "Disponível":
                tag = "disp"
            elif status == "Indisponível":
                tag = "indisp"
            elif status == "Pendente":
                tag = "pend"
            elif status == "Pendente Devolução":
                tag = "pend_ret"
                
            self.tree_stock.insert('', 'end', values=cleaned_row, tags=(tag,))

        self.tree_stock.tag_configure("disp", background="#7EFF9C") # Verde
        self.tree_stock.tag_configure("indisp", background="#FF7E89") # Vermelho
        self.tree_stock.tag_configure("pend", background="#FFD966") # Amarelo
        self.tree_stock.tag_configure("pend_ret", background="#87CEEB") # Azul Claro

    def update_issue_cb(self):
        items = self.inv.list_items()
        disponiveis = [f"{p['id']} - {p.get('brand') or ''} {p.get('model') or ''}" for p in items if p['status'] == 'Disponível']
        self.cb_issue['values'] = disponiveis
        self.cb_issue.set('')

    def update_return_views(self):
        """Atualiza as tabelas da aba de devolução."""
        self.tree_return_active.delete(*self.tree_return_active.get_children())
        self.tree_return_pending.delete(*self.tree_return_pending.get_children())

        for p in self.inv.list_items():
            row_data = (
                p.get('id'), p.get('tipo'), p.get('brand'), p.get('assigned_to'),
                format_cpf(p.get('cpf')), format_date(p.get('date_issued')), p.get('revenda')
            )
            values = tuple(v or '' for v in row_data)

            if p['status'] == 'Indisponível':
                self.tree_return_active.insert('', 'end', values=values)
            elif p['status'] == 'Pendente Devolução':
                self.tree_return_pending.insert('', 'end', values=values)

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

    def update_peripherals_table(self):
        """Atualiza a tabela de periféricos, aplicando o filtro de busca."""
        self.tree_peripherals.delete(*self.tree_peripherals.get_children())
        search_text = self.e_search_peripherals.get().lower().strip() # Pega o texto da busca

        for p in self.inv.list_peripherals():
            row_values = (
                p['id'], p.get('status'), p.get('tipo'), 
                p.get('brand'), p.get('model'),
                p.get('fornecedor'), p.get('nota_fiscal'),
                p.get('identificador')
            )
            
            # Converte a linha em texto para a busca
            row_str = " ".join(str(v or '').lower() for v in row_values)
            if search_text and search_text not in row_str:
                continue # Pula a linha se não corresponder à busca

            status = p.get('status')
            tag = ""
            if status == "Disponível": tag = "disp"
            elif status == "Em Uso": tag = "emuso"
            elif status == "Com Defeito": tag = "defeito"

            self.tree_peripherals.insert("", "end", values=row_values, tags=(tag,))

    def update_linking_combobox(self):
        """Atualiza o combobox de seleção de equipamento na aba de vínculos."""
        items = self.inv.list_items()
        # Filtra para mostrar apenas equipamentos que podem ter periféricos (ex: não celulares)
        linkable_items = [
            f"{i['id']} - {i.get('tipo')} {i.get('brand','')} {i.get('model','')}"
            for i in items if i.get('tipo') in ["Desktop", "Notebook", "Switch", "Impressora"]
        ]
        self.cb_link_equip['values'] = linkable_items

    def update_history_table(self):
        self.tree_history.delete(*self.tree_history.get_children())
        search_text = self.e_search_history.get().lower().strip()

        for h in self.inv.list_history():
            row_values = (
            h.get("item_id"),
            h.get("peripheral_id"),        
            h.get("operador"),        
            h.get("operation"),       
            h.get("revenda"),         
            format_date(h.get("data_operacao")), 
            format_time(h.get("data_operacao")), 
            h.get("tipo"),            
            h.get("marca"),         
            h.get("modelo"),          
            h.get("nota_fiscal"),
            h.get("fornecedor"),
            h.get("identificador"),   
            h.get("usuario"),         
            format_cpf(h.get("cpf")), 
            h.get("cargo"),           
            h.get("center_cost"),
            h.get("setor"),
            h.get("details")    
        )
            cleaned_row = tuple(v or '' for v in row_values)
            
            if search_text and search_text not in " ".join(str(v).lower() for v in cleaned_row):
                continue
                
            self.tree_history.insert("", "end", values=cleaned_row)


    def update_terms_table(self):
        self.tree_terms_pending.delete(*self.tree_terms_pending.get_children())
        self.tree_terms_active.delete(*self.tree_terms_active.get_children())

        for p in self.inv.list_items():
            row = (
                p.get('id'), p.get('tipo'), p.get('brand'), p.get('assigned_to'),
                format_cpf(p.get('cpf')), format_date(p.get('date_issued')), p.get('revenda')
            )
            values = tuple(v or '' for v in row)

            if p['status'] == 'Pendente':
                self.tree_terms_pending.insert('', 'end', values=values)
            elif p['status'] == 'Indisponível':
                self.tree_terms_active.insert('', 'end', values=values)

    def update_all_views(self):
        """Chama todas as funções de atualização da UI."""
        self.update_stock()
        self.update_issue_cb()
        self.update_return_views()
        self.update_remove_cb()
        self.update_edit_cb()
        self.update_peripherals_table()
        self.update_linking_combobox()
        self.update_terms_table()
        self.update_history_table()
        self.cmd_generate_report()
        self.update_users_table()

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





class ReplacePeripheralDialog(tk.Toplevel):
    def __init__(self, parent, equipment_id, old_peripheral_id, peripheral_type):
        super().__init__(parent)
        self.parent = parent
        self.inv = parent.inv
        self.equipment_id = equipment_id
        self.old_peripheral_id = old_peripheral_id
        self.peripheral_type = peripheral_type
        self.success = False

        self.title("Substituir Periférico")
        self.geometry("400x300")
        self.transient(parent)
        self.grab_set()

        frm = ttk.Frame(self, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=f"Selecione um(a) novo(a) {peripheral_type} disponível:", font=FONT_BOLD).pack(anchor="w")
        
        # Lista de periféricos disponíveis do mesmo tipo
        available = self.inv.list_peripherals(status_filter="Disponível", type_filter=peripheral_type)
        self.cb_new_peripheral = ttk.Combobox(frm, state="readonly", values=[f"{p['id']} - {p.get('brand','')} {p.get('model','')}" for p in available])
        self.cb_new_peripheral.pack(fill="x", pady=10)

        ttk.Label(frm, text="Motivo da Substituição:", font=FONT_BOLD).pack(anchor="w", pady=(10,0))
        self.e_reason = ttk.Entry(frm)
        self.e_reason.pack(fill="x", pady=5)
        
        self.lbl_msg = ttk.Label(frm, text="")
        self.lbl_msg.pack(pady=10)

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(fill="x", side="bottom")
        ttk.Button(btn_frm, text="Confirmar Substituição", command=self.confirm, style="Primary.TButton").pack(side="right")
        ttk.Button(btn_frm, text="Cancelar", command=self.destroy).pack(side="right", padx=10)

    def confirm(self):
        new_peri_selection = self.cb_new_peripheral.get()
        reason = self.e_reason.get().strip()
        if not new_peri_selection:
            self.lbl_msg.config(text="Selecione um novo periférico.", style="Danger.TLabel")
            return
        if not reason:
            self.lbl_msg.config(text="O motivo é obrigatório.", style="Danger.TLabel")
            return
        
        new_peripheral_id = int(new_peri_selection.split(' - ')[0])
        
        ok, msg = self.inv.replace_peripheral(
            self.equipment_id, self.old_peripheral_id, new_peripheral_id,
            reason, self.parent.logged_user
        )

        if ok:
            self.success = True
            self.destroy()
        else:
            self.lbl_msg.config(text=msg, style="Danger.TLabel")