from login import LoginWindow
from inicializar_db import setup_database

if __name__ == "__main__":
    # Execute a configuração do banco de dados antes de iniciar a interface
    setup_database() 
    
    login = LoginWindow()
    login.mainloop()