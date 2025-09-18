import tkinter as tk
from tkinter import ttk, messagebox
from database_mysql import get_connection
from gui import App

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login - Gestão de Estoque")
        self.geometry("350x280")
        self.resizable(False, False)
        self.configure(background='#F0F0F0')

        # Centralizar a janela
        self.eval('tk::PlaceWindow . center')

        # Estilo
        style = ttk.Style(self)
        style.theme_use('clam')

        # Estilo do botão primário
        style.configure("Primary.TButton",
            background="#0078D4",
            foreground="white",
            font=("Segoe UI", 10, "bold"),
            padding=(10, 8),
            borderwidth=0)
        style.map("Primary.TButton",
            background=[('active', '#005a9e')]) # Cor ao passar o mouse

        # Frame principal
        main_frame = ttk.Frame(self, padding=(20, 20), style="TFrame")
        main_frame.pack(fill="both", expand=True)

        # Título
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

    def login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("Atenção", "Preencha usuário e senha.")
            return

        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM usuarios WHERE username=%s AND password=%s", (username, password))
        user = cur.fetchone()
        conn.close()

        if user:
            self.destroy()
            app = App(user["username"], user["role"])
            app.mainloop()
        else:
            messagebox.showerror("Erro", "Usuário ou senha inválidos.")