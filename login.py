import tkinter as tk
from tkinter import messagebox
from database_mysql import get_connection
from gui import App

class LoginWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Login - Gestão de Estoque")
        self.geometry("300x200")
        self.resizable(False, False)

        tk.Label(self, text="Usuário:").pack(pady=5)
        self.username_entry = tk.Entry(self)
        self.username_entry.pack(pady=5)

        tk.Label(self, text="Senha:").pack(pady=5)
        self.password_entry = tk.Entry(self, show="*")
        self.password_entry.pack(pady=5)

        tk.Button(self, text="Entrar", command=self.login).pack(pady=10)

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
