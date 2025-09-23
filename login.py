import tkinter as tk
from tkinter import ttk, messagebox
from database_mysql import get_connection
from gui import App
import bcrypt


    # --- 1. ESTILIZAÇÃO DA TELA DE LOGIN ---

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login - Gestão de Estoque")
        self.iconbitmap("logo.ico")
        self.geometry("350x280")
        self.resizable(False, False)
        self.configure(background='#F0F0F0')

        self.eval('tk::PlaceWindow . center')

        style = ttk.Style(self)
        style.theme_use('clam')

        style.configure("Primary.TButton",
            background="#0078D4",
            foreground="white",
            font=("Segoe UI", 10, "bold"),
            padding=(10, 8),
            borderwidth=0)
        style.map("Primary.TButton",
            background=[('active', '#005a9e')])

        main_frame = ttk.Frame(self, padding=(20, 20), style="TFrame")
        main_frame.pack(fill="both", expand=True)

        ttk.Label(
            main_frame,
            text="Acesso ao Sistema",
            font=("Segoe UI", 16, "bold"),
            foreground="#333333"
        ).pack(pady=(0, 20))

        ttk.Label(main_frame, text="Usuário:", font=("Segoe UI", 10)).pack(pady=(5, 2), anchor="w")
        self.username_entry = ttk.Entry(main_frame, font=("Segoe UI", 10), width=35)
        self.username_entry.pack(fill="x", pady=(0, 10))
        self.username_entry.focus()

        ttk.Label(main_frame, text="Senha:", font=("Segoe UI", 10)).pack(pady=(5, 2), anchor="w")
        self.password_entry = ttk.Entry(main_frame, show="*", font=("Segoe UI", 10), width=35)
        self.password_entry.pack(fill="x", pady=(0, 20))

        ttk.Button(main_frame, text="Entrar", command=self.login, style="Primary.TButton").pack(fill="x")
        # Vincula a tecla Enter (Return) à função de login para a janela inteira
        self.bind('<Return>', self.login)

    # --- 2. LÓGICA DE LOGIN MODIFICADA IMPLEMENTANDO O HASH ---
    def login(self, event=None):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("Atenção", "Preencha usuário e senha.")
            return

        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        # Primeiro, busca o usuário APENAS pelo nome de usuário
        cur.execute("SELECT * FROM usuarios WHERE username = %s", (username,))
        user = cur.fetchone()
        conn.close()

        # Verifica se o usuário existe E se a senha digitada corresponde ao hash salvo
        if user and bcrypt.checkpw(password.encode('utf-8'), user['password'].encode('utf-8')):
            self.destroy()
            # MODIFICAÇÃO AQUI: Passe o ID do usuário para a classe App
            app = App(user["id"], user["username"], user["role"])
            app.mainloop()
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos.")