import os
from datetime import datetime
import mysql.connector
from mysql.connector import Error
from docx import Document
import calendar

from database_mysql import get_connection
from config import TERMS_DIR, TERMO_MODELOS
from utils import format_cpf, format_date

class InventoryDBManager:
    def __init__(self):
        self._create_tables()

    def _create_tables(self):
        """Cria as tabelas no banco, caso ainda não existam"""
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS items (
            id INT AUTO_INCREMENT PRIMARY KEY,
            tipo VARCHAR(50), brand VARCHAR(100), model VARCHAR(100), identificador VARCHAR(100),
            status ENUM('Disponível','Indisponível') DEFAULT 'Disponível',
            assigned_to VARCHAR(100), cpf VARCHAR(20), revenda VARCHAR(100),
            dominio VARCHAR(50), host VARCHAR(100), endereco_fisico VARCHAR(150),
            cpu VARCHAR(100), ram VARCHAR(50), storage VARCHAR(50),
            sistema VARCHAR(100), licenca VARCHAR(100), anydesk VARCHAR(50),
            setor VARCHAR(100), ip VARCHAR(50), mac VARCHAR(50),
            date_registered DATE NOT NULL, date_issued DATE,
            is_active TINYINT(1) DEFAULT 1 
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_id INT,
            operador VARCHAR(100),
            usuario VARCHAR(100),
            cpf VARCHAR(20),
            cargo VARCHAR(100),
            center_cost VARCHAR(100),
            revenda VARCHAR(100),
            data_operacao DATE,
            operation ENUM('Cadastro','Empréstimo','Devolução', 'Edição', 'Exclusão') DEFAULT 'Cadastro',
            -- Campos para guardar dados do item no momento da exclusão
            tipo VARCHAR(50),
            brand VARCHAR(100),
            model VARCHAR(100),
            identificador VARCHAR(100),
            FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE SET NULL
        )
        """)
        conn.commit()
        cur.close()
        conn.close()

    def add_item(self, item_data: dict, logged_user: str):
        """Adiciona novo item no estoque"""
        keys = ", ".join(item_data.keys())
        placeholders = ", ".join(["%s"] * len(item_data))
        values = list(item_data.values())
        sql = f"INSERT INTO items ({keys}) VALUES ({placeholders})"
        
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(sql, values)
        item_id = cur.lastrowid

        cur.execute("""
            INSERT INTO history (item_id, operador, data_operacao, operation)
            VALUES (%s, %s, %s, 'Cadastro')
        """, (item_id, logged_user, datetime.now().strftime("%Y-%m-%d")))
        
        conn.commit()
        cur.close()
        conn.close()
        return item_id

    def update_item(self, item_id: int, item_data: dict, logged_user: str):
        """Atualiza os dados de um item e LOGA A AÇÃO"""
        sets = ", ".join([f"{k}=%s" for k in item_data.keys()])
        values = list(item_data.values())
        values.append(item_id)
        sql = f"UPDATE items SET {sets} WHERE id=%s"

        conn = get_connection()
        cur = conn.cursor()
        cur.execute(sql, values)

        cur.execute("""
            INSERT INTO history (item_id, operador, data_operacao, operation)
            VALUES (%s, %s, %s, 'Edição')
        """, (item_id, logged_user, datetime.now().strftime("%Y-%m-%d")))

        conn.commit()
        cur.close()
        conn.close()
        return True, f"Item {item_id} atualizado com sucesso."


    def remove(self, item_id: int, logged_user: str):
        """"Remove" (desativa) item do estoque, salvando seus dados no histórico."""
        item = self.find(item_id)
        if not item: return False, "ID não encontrado."
        if item["status"] != "Disponível": return False, "Não é possível remover produto emprestado."

        conn = get_connection()
        cur = conn.cursor()

        # Salva os dados do item no log de histórico ANTES de desativar
        cur.execute("""
            INSERT INTO history (item_id, operador, data_operacao, operation, tipo, brand, model, identificador)
            VALUES (%s, %s, %s, 'Exclusão', %s, %s, %s, %s)
        """, (item_id, logged_user, datetime.now().strftime("%Y-%m-%d"),
              item.get('tipo'), item.get('brand'), item.get('model'), item.get('identificador')))

        # EM VEZ DE DELETAR, APENAS MARCA COMO INATIVO (Exclusão Lógica)
        cur.execute("UPDATE items SET is_active = 0 WHERE id=%s", (item_id,))
        
        conn.commit()
        cur.close()
        conn.close()
        return True, f"Aparelho {item_id} removido do estoque."


    def list_items(self):
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        # Adicionado "WHERE is_active = 1" para listar apenas itens não excluídos
        cur.execute("SELECT * FROM items WHERE is_active = 1")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows


    def find(self, item_id: int):
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        # Adicionado "AND is_active = 1" para não encontrar itens já excluídos
        cur.execute("SELECT * FROM items WHERE id=%s AND is_active = 1", (item_id,))
        row = cur.fetchone()
        cur.close()
        conn.close()
        return row



    def issue(self, pid, user, cpf, center_cost, cargo, revenda, date_issue, logged_user: str):
        """Realiza empréstimo de item"""
        item = self.find(pid)
        if not item:
            return False, "Item não encontrado."

        # ---> ADICIONADO: Validação para impedir empréstimo de item já indisponível <---
        if item['status'] != 'Disponível':
            return False, f"Este item já está emprestado para {item.get('assigned_to', 'desconhecido')}."
        
        try:
            dt_issue = datetime.strptime(date_issue, "%d/%m/%Y")
        except ValueError:
            return False, "Data de empréstimo inválida (use dd/mm/aaaa)."
        
        #Validar que a data do empréstimo não pode ser maior que a data atual
        if dt_issue.date() > datetime.now().date():
            return False, "A data de empréstimo não pode ser no futuro."

        if item.get('date_registered') and dt_issue.date() < item['date_registered']:
            data_formatada = format_date(str(item['date_registered']))
            return False, f"Data de empréstimo não pode ser anterior ao cadastro ({data_formatada})."

        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
            UPDATE items SET status='Indisponível', assigned_to=%s, cpf=%s, date_issued=%s WHERE id=%s
        """, (user, cpf, dt_issue.strftime("%Y-%m-%d"), pid))

        cur.execute("""
            INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, revenda, data_operacao, operation)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Empréstimo')
        """, (pid, logged_user, user, cpf, cargo, center_cost, revenda, dt_issue.strftime("%Y-%m-%d")))

        conn.commit()
        cur.close()
        conn.close()
        return True, f"Aparelho {pid} emprestado para {user}."


    # Em inventory_manager_db.py

    def ret(self, pid, date_return, logged_user: str):
        """Realiza devolução de item"""
        item = self.find(pid)
        if not item:
            return False, "Item não encontrado."

        try:
            dt_return = datetime.strptime(date_return, "%d/%m/%Y")
        except ValueError:
            return False, "Data de devolução inválida (use dd/mm/aaaa)."
        
        #data de devolução não pode ser maior que a data atual
        if dt_return.date() > datetime.now().date():
            return False, "A data de devolução não pode ser no futuro."

        if item.get('date_issued') and dt_return.date() < item['date_issued']:
            data_formatada = format_date(str(item['date_issued']))
            return False, f"Data de devolução não pode ser anterior ao empréstimo ({data_formatada})."

        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        cur.execute("""
            SELECT usuario, cpf, cargo, center_cost, revenda FROM history
            WHERE item_id = %s AND operation = 'Empréstimo'
            ORDER BY data_operacao DESC LIMIT 1
        """, (pid,))
        last_loan_details = cur.fetchone() or {}

        # ---> CORREÇÃO: "revenda=NULL" foi REMOVIDO desta query <---
        # A revenda é um atributo do item e não deve ser apagada na devolução.
        cur.execute("""
            UPDATE items SET status='Disponível', assigned_to=NULL, cpf=NULL, date_issued=NULL WHERE id=%s
        """, (pid,))

        cur.execute("""
            INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, revenda, data_operacao, operation)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Devolução')
        """, (
            pid,
            logged_user,
            last_loan_details.get('usuario'),
            last_loan_details.get('cpf'),
            last_loan_details.get('cargo'),
            last_loan_details.get('center_cost'),
            last_loan_details.get('revenda'),
            dt_return.strftime("%Y-%m-%d")
        ))

        conn.commit()
        cur.close()
        conn.close()
        return True, f"Aparelho {pid} devolvido."

    def list_history(self):
        """Lista histórico de TODAS as operações, tratando itens excluídos."""
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        cur.execute("""
            SELECT
                h.id, h.item_id, h.operador,
                COALESCE(i.tipo, h.tipo) as tipo,
                COALESCE(i.brand, h.brand) as marca,
                COALESCE(i.model, h.model) as modelo,
                COALESCE(i.identificador, h.identificador) as identificador,
                h.usuario, h.cpf, h.cargo, h.center_cost, h.revenda,
                h.data_operacao, h.operation
            FROM history h
            LEFT JOIN items i ON i.id = h.item_id
            ORDER BY h.data_operacao DESC, h.id DESC
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows




    def generate_monthly_report(self, ano, mes):
        """Relatório consolidado de um mês, usando Funções de Janela para maior precisão."""
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        sql = """
            WITH EventosOrdenados AS (
                -- Passo 1: Cria a 'fila' de eventos para cada item
                SELECT
                    id,
                    item_id,
                    operation,
                    data_operacao,
                    operador,
                    usuario,
                    cpf,
                    cargo,
                    center_cost,
                    revenda,
                    -- Passo 2: "Espia" a operação e a data do próximo evento na fila do mesmo item
                    LEAD(operation, 1) OVER (PARTITION BY item_id ORDER BY data_operacao, id) as proxima_operacao,
                    LEAD(data_operacao, 1) OVER (PARTITION BY item_id ORDER BY data_operacao, id) as proxima_data
                FROM
                    history
            ),
            RelatorioEmprestimos AS (
                -- Passo 3: Filtra apenas os empréstimos e verifica se o próximo evento é uma devolução
                SELECT
                    e.item_id,
                    e.operador,
                    e.usuario,
                    e.cpf,
                    e.cargo,
                    e.center_cost,
                    e.revenda,
                    e.data_operacao,
                    'Empréstimo' AS operation_type,
                    -- Se a próxima operação for 'Devolução', usa a data dela. Senão, é NULL.
                    CASE
                        WHEN e.proxima_operacao = 'Devolução' THEN e.proxima_data
                        ELSE NULL
                    END AS data_devolucao
                FROM
                    EventosOrdenados e
                WHERE
                    e.operation = 'Empréstimo'
                    AND YEAR(e.data_operacao) = %s AND MONTH(e.data_operacao) = %s
            )
            -- Consulta final que une os resultados
            SELECT
                re.item_id, re.operador, re.usuario, re.cpf, re.cargo, re.center_cost, re.revenda,
                re.data_operacao, re.operation_type,
                COALESCE(i.tipo, h.tipo) AS tipo,
                COALESCE(i.brand, h.brand) AS brand,
                COALESCE(i.model, h.model) AS model,
                COALESCE(i.identificador, h.identificador) AS identificador,
                re.data_devolucao
            FROM RelatorioEmprestimos re
            LEFT JOIN items i ON i.id = re.item_id
            LEFT JOIN history h ON h.item_id = re.item_id AND h.id = (SELECT MAX(id) FROM history WHERE item_id = re.item_id) -- Pega os dados mais recentes do histórico em caso de item deletado

            UNION ALL

            -- Parte de Cadastros (inalterada)
            SELECT
                cad.item_id, cad.operador, NULL, NULL, NULL, NULL, i.revenda,
                cad.data_operacao,
                'Cadastro' AS operation_type,
                COALESCE(i.tipo, cad.tipo) AS tipo,
                COALESCE(i.brand, cad.brand) AS brand,
                COALESCE(i.model, cad.model) AS model,
                COALESCE(i.identificador, cad.identificador) AS identificador,
                NULL AS data_devolucao
            FROM history cad
            LEFT JOIN items i ON i.id = cad.item_id
            WHERE cad.operation = 'Cadastro'
            AND YEAR(cad.data_operacao) = %s AND MONTH(cad.data_operacao) = %s
            
            ORDER BY data_operacao, item_id;
        """
        cur.execute(sql, (ano, mes, ano, mes))
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    def generate_term(self, item_id, user):
        item = self.find(item_id)
        if not item: return False, "Equipamento não encontrado."
        if item["status"] != "Indisponível" or item["assigned_to"] != user: return False, "Equipamento não está emprestado para este usuário."
        revenda = item.get("revenda")
        modelo_path = TERMO_MODELOS.get(revenda)
        if not modelo_path or not os.path.exists(modelo_path): return False, f"Modelo de termo não encontrado para {revenda}."
        safe_user_name = user.replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saida_path = os.path.join(TERMS_DIR, f"termo_{item_id}_{safe_user_name}_{revenda}_{timestamp}.docx")
        doc = Document(modelo_path)
        substituicoes = {
            "{{nome}}": user, "{{data_hoje}}": datetime.now().strftime("%d/%m/%Y"),
            "{{cpf}}": format_cpf(item.get("cpf", "")), "{{data_cadastro}}": format_date(item.get("date_registered", "")),
            "{{data_emprestimo}}": format_date(item.get("date_issued", "")), "{{marca}}": f" {item.get('brand', '')}" if item.get("brand") else "",
            "{{modelo}}": f" {item.get('model', '')}" if item.get("model") else "", "{{tipo}}": item.get("tipo", ""),
            "{{identificador}}": f" {item.get('identificador', '')}" if item.get("identificador") else ""
        }
        for key, value in item.items():
            placeholder = f"{{{{{key}}}}}"
            str_value = str(value) if value not in [None, ""] else ""
            if str_value: str_value = " " + str_value
            substituicoes[placeholder] = str_value
        for p in doc.paragraphs:
            for chave, valor in substituicoes.items():
                if chave in p.text: p.text = p.text.replace(chave, str(valor))
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for p in celula.paragraphs:
                        for chave, valor in substituicoes.items():
                            if chave in p.text: p.text = p.text.replace(chave, str(valor))
        doc.save(saida_path)
        return True, saida_path

    # --- NOVAS FUNÇÕES PARA GRÁFICOS ---
    def get_issue_return_counts(self, year, month):
        """Busca o número de empréstimos e devoluções por dia para um dado mês/ano."""
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        
        # Inicializa dicionários para empréstimos e devoluções com 0 para todos os dias
        issues = {day: 0 for day in days}
        returns = {day: 0 for day in days}

        sql = """
            SELECT
                DAY(data_operacao) as dia,
                operation,
                COUNT(id) as total
            FROM history
            WHERE
                YEAR(data_operacao) = %s AND
                MONTH(data_operacao) = %s AND
                operation IN ('Empréstimo', 'Devolução')
            GROUP BY dia, operation
        """
        cur.execute(sql, (year, month))
        
        for row in cur.fetchall():
            if row['operation'] == 'Empréstimo':
                issues[row['dia']] = row['total']
            elif row['operation'] == 'Devolução':
                returns[row['dia']] = row['total']
                
        cur.close()
        conn.close()
        
        return list(issues.keys()), list(issues.values()), list(returns.values())

    def get_registration_counts(self, year, month):
        """Busca o número de cadastros por dia para um dado mês/ano."""
        conn = get_connection()
        cur = conn.cursor(dictionary=True)
        
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        registrations = {day: 0 for day in days}

        sql = """
            SELECT
                DAY(data_operacao) as dia,
                COUNT(id) as total
            FROM history
            WHERE
                YEAR(data_operacao) = %s AND
                MONTH(data_operacao) = %s AND
                operation = 'Cadastro'
            GROUP BY dia
        """
        cur.execute(sql, (year, month))
        
        for row in cur.fetchall():
            registrations[row['dia']] = row['total']
            
        cur.close()
        conn.close()
        
        return list(registrations.keys()), list(registrations.values())