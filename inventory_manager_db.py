import os
import shutil
from datetime import datetime
import pymysql
from docx import Document
import calendar

from database_mysql import get_connection
from config import TERMS_DIR, TERMO_MODELOS, REMOVAL_NOTES_DIR, SIGNED_TERMS_DIR
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
            nota_fiscal VARCHAR(9) UNIQUE,
            status ENUM('Disponível','Indisponível', 'Pendente') DEFAULT 'Disponível',
            assigned_to VARCHAR(100), cpf VARCHAR(20), revenda VARCHAR(100),
            dominio VARCHAR(50), host VARCHAR(100), endereco_fisico VARCHAR(150),
            cpu VARCHAR(100), ram VARCHAR(50), storage VARCHAR(50),
            sistema VARCHAR(100), licenca VARCHAR(100), anydesk VARCHAR(50),
            setor VARCHAR(100), ip VARCHAR(50), mac VARCHAR(50),
            date_registered DATE NOT NULL, date_issued DATE,
            is_active TINYINT(1) DEFAULT 1 
        )
        """)
        # ALTERADO: Adicionado 'Estorno' ao ENUM e a coluna 'is_reversed'
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
            operation ENUM('Cadastro','Empréstimo','Devolução', 'Edição', 'Exclusão', 'Estorno', 'Confirmação Empréstimo') DEFAULT 'Cadastro',
            is_reversed TINYINT(1) DEFAULT 0,
            -- Campos para guardar dados do item
            tipo VARCHAR(50),
            brand VARCHAR(100),
            model VARCHAR(100),
            identificador VARCHAR(100),
            nota_fiscal VARCHAR(9),
            nota_fiscal_anexo VARCHAR(255) NULL,
            termo_assinado_anexo VARCHAR(255) NULL,
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


    def remove(self, item_id: int, logged_user: str, attachment_path: str):
        """"Remove" (desativa) item do estoque, salvando seus dados no histórico."""
        item = self.find(item_id)
        if not item: return False, "ID não encontrado."
        if item["status"] != "Disponível": return False, "Não é possível remover produto emprestado."
        
        # --- Lógica para copiar o anexo ---
        try:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            original_filename = os.path.basename(attachment_path)
            # Cria um nome de arquivo seguro e único
            new_filename = f"remocao_{item_id}_{timestamp}_{original_filename}"
            destination_path = os.path.join(REMOVAL_NOTES_DIR, new_filename)
            
            # Copia o arquivo para a pasta de notas
            shutil.copy(attachment_path, destination_path)
        except Exception as e:
            return False, f"Erro ao processar o anexo: {e}"
        # ------------------------------------

        conn = get_connection()
        cur = conn.cursor()

        cur.execute("""
            INSERT INTO history (item_id, operador, data_operacao, operation, tipo, brand, model, identificador, nota_fiscal, nota_fiscal_anexo)
            VALUES (%s, %s, %s, 'Exclusão', %s, %s, %s, %s, %s, %s)
        """, (item_id, logged_user, datetime.now().strftime("%Y-%m-%d"),
              item.get('tipo'), item.get('brand'), item.get('model'), item.get('identificador'), item.get('nota_fiscal'), destination_path))

        cur.execute("UPDATE items SET is_active = 0 WHERE id=%s", (item_id,))
        
        conn.commit()
        cur.close()
        conn.close()
        return True, f"Aparelho {item_id} removido do estoque."


    def list_items(self):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("SELECT * FROM items WHERE is_active = 1")
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows


    def find(self, item_id: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("SELECT * FROM items WHERE id=%s AND is_active = 1", (item_id,))
        row = cur.fetchone()
        cur.close()
        conn.close()
        return row



    def issue(self, pid, user, cpf, center_cost, cargo, revenda, date_issue, logged_user: str):
        item = self.find(pid)
        if not item:
            return False, "Item não encontrado."

        if item['status'] != 'Disponível':
            return False, f"Este item não está disponível para empréstimo."
        
        try:
            dt_issue = datetime.strptime(date_issue, "%d/%m/%Y")
        except ValueError:
            return False, "Data de empréstimo inválida (use dd/mm/aaaa)."
        
        if dt_issue.date() > datetime.now().date():
            return False, "A data de empréstimo não pode ser no futuro."

        if item.get('date_registered') and dt_issue.date() < item['date_registered']:
            data_formatada = format_date(str(item['date_registered']))
            return False, f"Data de empréstimo não pode ser anterior ao cadastro ({data_formatada})."

        conn = get_connection()
        cur = conn.cursor()
        # ALTERAÇÃO AQUI: O status agora é 'Pendente'
        cur.execute("""
            UPDATE items 
            SET status='Pendente', assigned_to=%s, cpf=%s, date_issued=%s, revenda=%s 
            WHERE id=%s
        """, (user, cpf, dt_issue.strftime("%Y-%m-%d"), revenda, pid))

        cur.execute("""
            INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, revenda, data_operacao, operation)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Empréstimo')
        """, (pid, logged_user, user, cpf, cargo, center_cost, revenda, dt_issue.strftime("%Y-%m-%d")))

        conn.commit()
        cur.close()
        conn.close()
        return True, f"Empréstimo do item {pid} para {user} iniciado. Status: Pendente."
    
    def confirm_loan(self, item_id: int, logged_user: str, signed_term_path: str):
        """
        finaliza o empréstimo, anexa o termo e muda o status para Indisponível.
        """
        item = self.find(item_id)
        if not item or item['status'] != 'Pendente':
            return False, "Apenas itens com status 'Pendente' podem ser confirmados."

        # --- Lógica para copiar o anexo ---
        try:
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            original_filename = os.path.basename(signed_term_path)
            new_filename = f"termo_{item_id}_{timestamp}_{original_filename}"
            destination_path = os.path.join(SIGNED_TERMS_DIR, new_filename)
            shutil.copy(signed_term_path, destination_path)
        except Exception as e:
            return False, f"Erro ao processar o termo assinado: {e}"
        # ------------------------------------

        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)

        # Encontra o último registro de empréstimo para este item
        cur.execute("""
            SELECT id, usuario, cpf, cargo, center_cost, revenda FROM history
            WHERE item_id = %s AND operation = 'Empréstimo' AND is_reversed = 0
            ORDER BY data_operacao DESC, id DESC LIMIT 1
        """, (item_id,))
        last_loan_details = cur.fetchone()

        if not last_loan_details:
            conn.close()
            return False, "Registro de empréstimo original não encontrado no histórico."

        # Atualiza o status do item para Indisponível
        cur.execute("UPDATE items SET status='Indisponível' WHERE id=%s", (item_id,))

        # Cria o novo registro de "Confirmação" no histórico, salvando o caminho do anexo
        cur.execute("""
            INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, revenda, data_operacao, operation, termo_assinado_anexo)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 'Confirmação Empréstimo', %s)
        """, (
            item_id,
            logged_user,
            last_loan_details.get('usuario'),
            last_loan_details.get('cpf'),
            last_loan_details.get('cargo'),
            last_loan_details.get('center_cost'),
            last_loan_details.get('revenda'),
            datetime.now().strftime("%Y-%m-%d"),
            destination_path  # Salva o caminho do termo assinado
        ))

        conn.commit()
        cur.close()
        conn.close()
        return True, f"Empréstimo do item {item_id} confirmado com sucesso."


    def ret(self, pid, date_return, logged_user: str):
        item = self.find(pid)
        if not item:
            return False, "Item não encontrado."

        try:
            dt_return = datetime.strptime(date_return, "%d/%m/%Y")
        except ValueError:
            return False, "Data de devolução inválida (use dd/mm/aaaa)."
        
        if dt_return.date() > datetime.now().date():
            return False, "A data de devolução não pode ser no futuro."

        if item.get('date_issued') and dt_return.date() < item['date_issued']:
            data_formatada = format_date(str(item['date_issued']))
            return False, f"Data de devolução não pode ser anterior ao empréstimo ({data_formatada})."

        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        
        cur.execute("""
            SELECT usuario, cpf, cargo, center_cost, revenda FROM history
            WHERE item_id = %s AND operation = 'Empréstimo'
            ORDER BY data_operacao DESC LIMIT 1
        """, (pid,))
        last_loan_details = cur.fetchone() or {}

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
        cur = conn.cursor(pymysql.cursors.DictCursor)
        # ALTERADO: Adicionado filtro para não mostrar operações estornadas
        cur.execute("""
            SELECT
                h.id, h.item_id, h.operador,
                COALESCE(i.tipo, h.tipo) as tipo,
                COALESCE(i.brand, h.brand) as marca,
                COALESCE(i.model, h.model) as modelo,
                COALESCE(i.identificador, h.identificador) as identificador,
                COALESCE(i.nota_fiscal, h.nota_fiscal) as nota_fiscal,
                h.usuario, h.cpf, h.cargo, h.center_cost, h.revenda,
                h.data_operacao, h.operation
            FROM history h
            LEFT JOIN items i ON i.id = h.item_id
            WHERE h.is_reversed = 0
            ORDER BY h.data_operacao DESC, h.id DESC
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows


    def generate_monthly_report(self, ano, mes):
        """Relatório consolidado de um mês, usando Funções de Janela para maior precisão."""
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        
        # ALTERADO: Adicionado "WHERE is_reversed = 0" para ignorar operações estornadas
        sql = """
            WITH EventosOrdenados AS (
                SELECT
                    id, item_id, operation, data_operacao, operador, usuario,
                    cpf, cargo, center_cost, revenda,
                    LEAD(operation, 1) OVER (PARTITION BY item_id ORDER BY data_operacao, id) as proxima_operacao,
                    LEAD(data_operacao, 1) OVER (PARTITION BY item_id ORDER BY data_operacao, id) as proxima_data
                FROM
                    history
                WHERE is_reversed = 0
            ),
            RelatorioEmprestimos AS (
                SELECT
                    e.id as history_id,
                    e.item_id, e.operador, e.usuario, e.cpf, e.cargo, e.center_cost,
                    e.revenda, e.data_operacao, 'Empréstimo' AS operation_type,
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
            SELECT
                re.history_id, re.item_id, re.operador, re.usuario, re.cpf, re.cargo, re.center_cost, re.revenda,
                re.data_operacao, re.operation_type,
                COALESCE(i.tipo, h.tipo) AS tipo,
                COALESCE(i.brand, h.brand) AS brand,
                COALESCE(i.model, h.model) AS model,
                COALESCE(i.identificador, h.identificador) AS identificador,
                COALESCE(i.nota_fiscal, h.nota_fiscal) AS nota_fiscal,
                re.data_devolucao
            FROM RelatorioEmprestimos re
            LEFT JOIN items i ON i.id = re.item_id
            LEFT JOIN history h ON h.item_id = re.item_id AND h.id = (SELECT MAX(id) FROM history WHERE item_id = re.item_id)

            UNION ALL

            SELECT
                cad.id as history_id,
                cad.item_id, cad.operador, NULL, NULL, NULL, NULL, i.revenda,
                cad.data_operacao, 'Cadastro' AS operation_type,
                COALESCE(i.tipo, cad.tipo) AS tipo,
                COALESCE(i.brand, cad.brand) AS brand,
                COALESCE(i.model, cad.model) AS model,
                COALESCE(i.identificador, cad.identificador) AS identificador,
                COALESCE(i.nota_fiscal, cad.nota_fiscal) AS nota_fiscal,
                NULL AS data_devolucao
            FROM history cad
            LEFT JOIN items i ON i.id = cad.item_id
            WHERE cad.operation = 'Cadastro'
            AND cad.is_reversed = 0
            AND YEAR(cad.data_operacao) = %s AND MONTH(cad.data_operacao) = %s
            
            ORDER BY data_operacao, item_id;
        """
        cur.execute(sql, (ano, mes, ano, mes))
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    # ALTERADO: Lógica de estorno completamente refeita
    def reverse_history_entry(self, history_id: int, logged_user: str):
        """
        Estorna um lançamento: reverte o estado do item, marca a operação original
        como 'estornada' e cria um novo registro de 'Estorno' para auditoria.
        """
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        
        try:
            conn.start_transaction()

            cur.execute("SELECT * FROM history WHERE id = %s", (history_id,))
            entry_to_reverse = cur.fetchone()

            if not entry_to_reverse:
                return False, "Lançamento do histórico não encontrado."
            
            if entry_to_reverse['is_reversed']:
                return False, "Esta operação já foi estornada anteriormente."

            op = entry_to_reverse['operation']
            item_id = entry_to_reverse['item_id']

            # Reverte o estado do item no estoque
            if op == 'Empréstimo':
                cur.execute("UPDATE items SET status='Disponível', assigned_to=NULL, cpf=NULL, date_issued=NULL WHERE id = %s", (item_id,))
            
            elif op == 'Devolução':
                cur.execute("""
                    SELECT usuario, cpf, data_operacao FROM history
                    WHERE item_id = %s AND operation = 'Empréstimo' AND id < %s AND is_reversed = 0
                    ORDER BY data_operacao DESC, id DESC LIMIT 1
                """, (item_id, history_id))
                last_loan = cur.fetchone()

                if not last_loan:
                    return False, "Não foi possível encontrar o empréstimo original para reverter a devolução."

                cur.execute("UPDATE items SET status='Indisponível', assigned_to=%s, cpf=%s, date_issued=%s WHERE id = %s", 
                            (last_loan['usuario'], last_loan['cpf'], last_loan['data_operacao'], item_id))
            
            elif op == 'Cadastro':
                 # Em vez de deletar, apenas marca como inativo
                cur.execute("UPDATE items SET is_active = 0 WHERE id = %s", (item_id,))
            else:
                return False, f"Não é possível estornar uma operação do tipo '{op}'."

            # Marca a operação original como estornada
            cur.execute("UPDATE history SET is_reversed = 1 WHERE id = %s", (history_id,))
            
            # Cria o novo registro de 'Estorno' para auditoria
            cur.execute("""
                INSERT INTO history (item_id, operador, data_operacao, operation, usuario, cpf, cargo, center_cost, revenda)
                VALUES (%s, %s, %s, 'Estorno', %s, %s, %s, %s, %s)
            """, (
                item_id, logged_user, datetime.now().strftime("%Y-%m-%d"),
                entry_to_reverse.get('usuario'), entry_to_reverse.get('cpf'),
                entry_to_reverse.get('cargo'), entry_to_reverse.get('center_cost'),
                entry_to_reverse.get('revenda')
            ))
            
            conn.commit()
            return True, f"Operação '{op}' do item {item_id} estornada com sucesso."

        except pymysql.MySQLError as e:
            conn.rollback()
            return False, f"Erro ao estornar: {e}"
        finally:
            cur.close()
            conn.close()


    def generate_term(self, item_id, user):
        item = self.find(item_id)
        if not item: 
            return False, "Equipamento não encontrado."
        

        # Verifica por "Pendente", que é o estado correto para gerar um termo.
        if item["status"] != "Pendente" or item["assigned_to"] != user:
            return False, "Este equipamento não está pendente de empréstimo para este usuário."
        # ----------------------------------------

        revenda = item.get("revenda")
        modelo_path = TERMO_MODELOS.get(revenda)
        if not modelo_path or not os.path.exists(modelo_path): 
            return False, f"Modelo de termo não encontrado para {revenda}."
            
        safe_user_name = user.replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        saida_path = os.path.join(TERMS_DIR, f"termo_{item_id}_{safe_user_name}_{revenda}_{timestamp}.docx")
        doc = Document(modelo_path)
        
        substituicoes = {
            "{{nome}}": user, "{{data_hoje}}": datetime.now().strftime("%d/%m/%Y"),
            "{{cpf}}": format_cpf(item.get("cpf", "")), "{{data_cadastro}}": format_date(item.get("date_registered", "")),
            "{{data_emprestimo}}": format_date(item.get("date_issued", "")), "{{marca}}": f" {item.get('brand', '')}" if item.get("brand") else "",
            "{{modelo}}": f" {item.get('model', '')}" if item.get("model") else "", "{{tipo}}": item.get("tipo", ""),
            "{{identificador}}": f" {item.get('identificador', '')}" if item.get("identificador") else "",
            "{{nota_fiscal}}": f" {item.get('nota_fiscal', '')}" if item.get("nota_fiscal") else ""
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
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        issues = {day: 0 for day in days}
        returns = {day: 0 for day in days}

        # ALTERADO: Adicionado filtro para ignorar operações estornadas
        sql = """
            SELECT DAY(data_operacao) as dia, operation, COUNT(id) as total
            FROM history
            WHERE
                YEAR(data_operacao) = %s AND
                MONTH(data_operacao) = %s AND
                operation IN ('Empréstimo', 'Devolução') AND
                is_reversed = 0
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
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        registrations = {day: 0 for day in days}

        # ALTERADO: Adicionado filtro para ignorar operações estornadas
        sql = """
            SELECT DAY(data_operacao) as dia, COUNT(id) as total
            FROM history
            WHERE
                YEAR(data_operacao) = %s AND
                MONTH(data_operacao) = %s AND
                operation = 'Cadastro' AND
                is_reversed = 0
            GROUP BY dia
        """
        cur.execute(sql, (year, month))
        
        for row in cur.fetchall():
            registrations[row['dia']] = row['total']
        cur.close()
        conn.close()
        return list(registrations.keys()), list(registrations.values())