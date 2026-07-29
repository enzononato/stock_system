"""
Camada de acesso ao banco de dados para o inventário.
Adaptado do original para uso no backend FastAPI:
- Importações corrigidas para pacote app.db.*
- generate_term() e generate_and_initiate_return() retornam bytes (DOCX em memória)
  em vez de salvar em disco — o router controla o storage.
- confirm_loan() e confirm_return() recebem storage_key (str) em vez de path local.
- remove() e replace_peripheral() recebem storage_key (str) em vez de path local.
"""
import io
import calendar
import logging
from datetime import datetime
import pymysql
from docx import Document

from app.db.database_mysql import get_connection
from app.core.config import TERMO_MODELOS, TERMO_DEVOLUCAO_MODELOS
from app.db.utils import format_cpf, format_date

logger = logging.getLogger(__name__)


class InventoryDBManager:
    def __init__(self):
        self._create_tables()

    def _create_tables(self):
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS items (
            id INT AUTO_INCREMENT PRIMARY KEY,
            tipo VARCHAR(50), brand VARCHAR(100), model VARCHAR(100), identificador VARCHAR(100),
            nota_fiscal VARCHAR(50),
            status ENUM('Disponível','Indisponível','Pendente','Pendente Devolução') DEFAULT 'Disponível',
            assigned_to VARCHAR(100), cpf VARCHAR(20), revenda VARCHAR(100),
            dominio VARCHAR(50), host VARCHAR(100), endereco_fisico VARCHAR(150),
            cpu VARCHAR(100), ram VARCHAR(50), storage VARCHAR(50),
            sistema VARCHAR(100), licenca VARCHAR(100), anydesk VARCHAR(50),
            setor VARCHAR(100), ip VARCHAR(50), mac VARCHAR(50),
            fornecedor VARCHAR(150),
            potencia_nominal VARCHAR(50), autonomia_estimada VARCHAR(100), ip_snmp VARCHAR(50),
            codigo_patrimonial VARCHAR(100), responsavel VARCHAR(100), local_instalacao VARCHAR(150),
            poe ENUM('Sim','Não'), quantidade_portas VARCHAR(10),
            date_registered DATETIME NOT NULL, date_issued DATETIME,
            is_active TINYINT(1) DEFAULT 1
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS history (
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_id INT, peripheral_id INT,
            operador VARCHAR(100), usuario VARCHAR(100), cpf VARCHAR(20),
            cargo VARCHAR(100), center_cost VARCHAR(100), setor VARCHAR(100),
            fornecedor VARCHAR(150), revenda VARCHAR(100),
            data_operacao DATETIME,
            operation ENUM(
                'Cadastro','Empréstimo','Devolução','Edição','Exclusão','Estorno',
                'Confirmação Empréstimo','Confirmação Devolução',
                'Cadastro Periférico','Vínculo Periférico','Desvínculo Periférico','Substituição Periférico'
            ) DEFAULT 'Cadastro',
            is_reversed TINYINT(1) DEFAULT 0,
            details VARCHAR(255),
            tipo VARCHAR(50), brand VARCHAR(100), model VARCHAR(100), identificador VARCHAR(100),
            nota_fiscal VARCHAR(50), poe ENUM('Sim','Não'), quantidade_portas VARCHAR(10),
            operacao_anexo VARCHAR(255) NULL, termo_assinado_anexo VARCHAR(255) NULL,
            FOREIGN KEY(item_id) REFERENCES items(id) ON DELETE SET NULL
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS peripherals (
            id INT AUTO_INCREMENT PRIMARY KEY,
            tipo VARCHAR(50) NOT NULL, brand VARCHAR(100), model VARCHAR(100),
            identificador VARCHAR(100) UNIQUE,
            status ENUM('Disponível','Em Uso','Substituido') DEFAULT 'Disponível',
            motivo_substituicao VARCHAR(255),
            date_registered DATETIME NOT NULL,
            is_active TINYINT(1) DEFAULT 1
        )
        """)
        cur.execute("""
        CREATE TABLE IF NOT EXISTS equipment_peripherals (
            id INT AUTO_INCREMENT PRIMARY KEY,
            equipment_id INT NOT NULL, peripheral_id INT NOT NULL,
            FOREIGN KEY(equipment_id) REFERENCES items(id) ON DELETE CASCADE,
            FOREIGN KEY(peripheral_id) REFERENCES peripherals(id) ON DELETE CASCADE,
            UNIQUE(equipment_id, peripheral_id)
        )
        """)
        conn.commit()
        cur.close()
        conn.close()

    # ── Items ──────────────────────────────────────────────────────────────────

    def add_item(self, item_data: dict, logged_user: str):
        keys = ", ".join(item_data.keys())
        placeholders = ", ".join(["%s"] * len(item_data))
        sql = f"INSERT INTO items ({keys}) VALUES ({placeholders})"
        conn = get_connection()
        cur = conn.cursor()
        try:
            cur.execute(sql, list(item_data.values()))
            item_id = cur.lastrowid
            cur.execute(
                "INSERT INTO history (item_id, operador, data_operacao, operation) VALUES (%s,%s,%s,'Cadastro')",
                (item_id, logged_user, datetime.now()),
            )
            conn.commit()
            return True, item_id
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."
        finally:
            cur.close()
            conn.close()

    def update_item(self, item_id: int, item_data: dict, logged_user: str):
        sets = ", ".join([f"{k}=%s" for k in item_data.keys()])
        values = list(item_data.values()) + [item_id]
        conn = get_connection()
        cur = conn.cursor()
        cur.execute(f"UPDATE items SET {sets} WHERE id=%s", values)
        cur.execute(
            "INSERT INTO history (item_id, operador, data_operacao, operation) VALUES (%s,%s,%s,'Edição')",
            (item_id, logged_user, datetime.now()),
        )
        conn.commit()
        cur.close()
        conn.close()
        return True, f"Item {item_id} atualizado."

    def remove(self, item_id: int, logged_user: str, reason: str, storage_key: str = None):
        """Remove (soft-delete) um item. storage_key é a chave do anexo no storage backend."""
        item = self.find(item_id)
        if not item:
            return False, "ID não encontrado."
        if item["status"] != "Disponível":
            return False, "Não é possível remover produto emprestado."

        conn = get_connection()
        cur = conn.cursor()
        try:
            cur.execute("UPDATE items SET is_active=0 WHERE id=%s", (item_id,))
            cur.execute(
                """INSERT INTO history
                (item_id, operador, data_operacao, operation, details,
                 tipo, brand, model, identificador, nota_fiscal, fornecedor,
                 operacao_anexo, poe, quantidade_portas)
                VALUES (%s,%s,%s,'Exclusão',%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (
                    item_id, logged_user, datetime.now(), reason,
                    item.get("tipo"), item.get("brand"), item.get("model"),
                    item.get("identificador"), item.get("nota_fiscal"),
                    item.get("fornecedor"), storage_key,
                    item.get("poe"), item.get("quantidade_portas"),
                ),
            )
            conn.commit()
            return True, f"Aparelho {item_id} removido do estoque."
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."
        finally:
            cur.close()
            conn.close()

    def list_items(self):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("""
            SELECT i.*, COUNT(ep.peripheral_id) as peripheral_count
            FROM items i
            LEFT JOIN equipment_peripherals ep ON i.id = ep.equipment_id
            WHERE i.is_active = 1
            GROUP BY i.id
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    def find(self, item_id: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("SELECT * FROM items WHERE id=%s AND is_active=1", (item_id,))
        row = cur.fetchone()
        cur.close()
        conn.close()
        return row

    # ── Peripherals ────────────────────────────────────────────────────────────

    def add_peripheral(self, data: dict, logged_user: str):
        conn = get_connection()
        cur = conn.cursor()
        try:
            data["date_registered"] = datetime.now()
            keys = ", ".join(data.keys())
            placeholders = ", ".join(["%s"] * len(data))
            cur.execute(f"INSERT INTO peripherals ({keys}) VALUES ({placeholders})", list(data.values()))
            p_id = cur.lastrowid
            cur.execute(
                """INSERT INTO history (peripheral_id, operador, data_operacao, operation, tipo, brand, model, identificador)
                VALUES (%s,%s,%s,'Cadastro Periférico',%s,%s,%s,%s)""",
                (p_id, logged_user, datetime.now(), data.get("tipo"), data.get("brand"), data.get("model"), data.get("identificador")),
            )
            conn.commit()
            return True, f"Periférico '{data.get('tipo')}' cadastrado com ID {p_id}."
        except pymysql.MySQLError as e:
            conn.rollback()
            if e.args[0] == 1062:
                return False, "Já existe um periférico com este Identificador (Nº de Série)."
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."
        finally:
            cur.close()
            conn.close()

    def list_peripherals(self, status_filter="", type_filter="", include_inactive=False):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        where = []
        params = []
        if not include_inactive:
            where.append("is_active=1")
        if status_filter:
            where.append("status=%s")
            params.append(status_filter)
        if type_filter:
            where.append("tipo=%s")
            params.append(type_filter)
        query = "SELECT * FROM peripherals"
        if where:
            query += " WHERE " + " AND ".join(where)
        cur.execute(query, params)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    def list_peripherals_for_equipment(self, equipment_id: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute(
                """SELECT p.*, ep.id as link_id
                FROM peripherals p
                JOIN equipment_peripherals ep ON p.id = ep.peripheral_id
                WHERE ep.equipment_id = %s""",
                (equipment_id,),
            )
            return cur.fetchall()
        finally:
            cur.close()
            conn.close()

    def link_peripheral_to_equipment(self, equipment_id: int, peripheral_id: int, logged_user: str):
        conn = get_connection()
        cur = conn.cursor()
        try:
            cur.execute(
                "INSERT INTO equipment_peripherals (equipment_id, peripheral_id) VALUES (%s,%s)",
                (equipment_id, peripheral_id),
            )
            cur.execute("UPDATE peripherals SET status='Em Uso' WHERE id=%s", (peripheral_id,))
            cur.execute(
                """INSERT INTO history (item_id, peripheral_id, operador, data_operacao, operation)
                VALUES (%s,%s,%s,%s,'Vínculo Periférico')""",
                (equipment_id, peripheral_id, logged_user, datetime.now()),
            )
            conn.commit()
            return True, "Periférico vinculado com sucesso."
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro ao vincular periférico")
            return False, "Erro ao vincular periférico."
        finally:
            cur.close()
            conn.close()

    def unlink_peripheral_from_equipment(self, link_id: int, logged_user: str):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute("SELECT equipment_id, peripheral_id FROM equipment_peripherals WHERE id=%s", (link_id,))
            ids = cur.fetchone()
            if not ids:
                return False, "Vínculo não encontrado."
            cur.execute("DELETE FROM equipment_peripherals WHERE id=%s", (link_id,))
            cur.execute("UPDATE peripherals SET status='Disponível' WHERE id=%s", (ids["peripheral_id"],))
            cur.execute(
                """INSERT INTO history (item_id, peripheral_id, operador, data_operacao, operation)
                VALUES (%s,%s,%s,%s,'Desvínculo Periférico')""",
                (ids["equipment_id"], ids["peripheral_id"], logged_user, datetime.now()),
            )
            conn.commit()
            return True, "Periférico desvinculado com sucesso."
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro ao desvincular periférico")
            return False, "Erro ao desvincular periférico."
        finally:
            cur.close()
            conn.close()

    def replace_peripheral(
        self,
        equipment_id: int,
        old_peripheral_id: int,
        new_peripheral_id: int,
        reason: str,
        logged_user: str,
        storage_key: str = None,
    ):
        conn = get_connection()
        cur = conn.cursor()
        try:
            cur.execute(
                "UPDATE peripherals SET status='Substituido', is_active=0, motivo_substituicao=%s WHERE id=%s",
                (reason, old_peripheral_id),
            )
            cur.execute(
                "DELETE FROM equipment_peripherals WHERE equipment_id=%s AND peripheral_id=%s",
                (equipment_id, old_peripheral_id),
            )
            cur.execute(
                "INSERT INTO equipment_peripherals (equipment_id, peripheral_id) VALUES (%s,%s)",
                (equipment_id, new_peripheral_id),
            )
            cur.execute("UPDATE peripherals SET status='Em Uso' WHERE id=%s", (new_peripheral_id,))
            details_log = f"Substituído periférico ID {old_peripheral_id} por ID {new_peripheral_id}. Motivo: {reason}"
            cur.execute(
                """INSERT INTO history (item_id, peripheral_id, operador, data_operacao, operation, details, operacao_anexo)
                VALUES (%s,%s,%s,%s,'Substituição Periférico',%s,%s)""",
                (equipment_id, old_peripheral_id, logged_user, datetime.now(), details_log, storage_key),
            )
            conn.commit()
            return True, "Substituição realizada com sucesso."
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro ao substituir periférico")
            return False, "Erro ao substituir periférico."
        finally:
            cur.close()
            conn.close()

    # ── Loan workflow ──────────────────────────────────────────────────────────

    def issue(self, pid, user, cpf, center_cost, cargo, setor, revenda, date_issue, logged_user: str):
        item = self.find(pid)
        if not item:
            return False, "Item não encontrado."
        if item["status"] != "Disponível":
            return False, "Este item não está disponível para empréstimo."
        try:
            dt_issue = datetime.strptime(date_issue, "%d/%m/%Y")
            dt_issue = dt_issue.replace(
                hour=datetime.now().hour,
                minute=datetime.now().minute,
                second=datetime.now().second,
            )
        except ValueError:
            return False, "Data de empréstimo inválida (use dd/mm/aaaa)."
        if dt_issue.date() > datetime.now().date():
            return False, "A data de empréstimo não pode ser no futuro."
        if item.get("date_registered") and dt_issue.date() < item["date_registered"].date():
            return False, f"Data de empréstimo não pode ser anterior ao cadastro ({format_date(str(item['date_registered']))})."

        conn = get_connection()
        cur = conn.cursor()
        cur.execute(
            "UPDATE items SET status='Pendente', assigned_to=%s, cpf=%s, date_issued=%s, revenda=%s WHERE id=%s",
            (user, cpf, dt_issue, revenda, pid),
        )
        cur.execute(
            """INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, setor, revenda, fornecedor, data_operacao, operation)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'Empréstimo')""",
            (pid, logged_user, user, cpf, cargo, center_cost, setor, revenda, item.get("fornecedor"), dt_issue),
        )
        conn.commit()
        cur.close()
        conn.close()
        return True, f"Empréstimo do item {pid} para {user} iniciado. Status: Pendente."

    def confirm_loan(self, item_id: int, logged_user: str, storage_key: str):
        """Confirma empréstimo. storage_key é a chave do PDF assinado no storage backend."""
        item = self.find(item_id)
        if not item or item["status"] != "Pendente":
            return False, "Apenas itens com status 'Pendente' podem ser confirmados."

        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute(
                """SELECT id, usuario, cpf, cargo, center_cost, revenda FROM history
                WHERE item_id=%s AND operation='Empréstimo' AND is_reversed=0
                ORDER BY data_operacao DESC, id DESC LIMIT 1""",
                (item_id,),
            )
            last_loan = cur.fetchone()
            if not last_loan:
                raise Exception("Registro de empréstimo original não encontrado no histórico.")

            cur.execute("UPDATE items SET status='Indisponível' WHERE id=%s", (item_id,))
            for p in self.list_peripherals_for_equipment(item_id):
                cur.execute("UPDATE peripherals SET status='Em Uso' WHERE id=%s", (p["id"],))

            cur.execute(
                """INSERT INTO history (item_id, operador, usuario, cpf, cargo, center_cost, revenda, data_operacao, operation, termo_assinado_anexo)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,'Confirmação Empréstimo',%s)""",
                (
                    item_id, logged_user,
                    last_loan.get("usuario"), last_loan.get("cpf"),
                    last_loan.get("cargo"), last_loan.get("center_cost"),
                    last_loan.get("revenda"), datetime.now(), storage_key,
                ),
            )
            conn.commit()
            return True, f"Empréstimo do item {item_id} confirmado."
        except Exception:
            conn.rollback()
            logger.exception("Erro ao confirmar empréstimo")
            return False, "Erro ao confirmar empréstimo."
        finally:
            cur.close()
            conn.close()

    # ── Return workflow ────────────────────────────────────────────────────────

    def generate_return_term_bytes(self, item_id: int, logged_user: str):
        """
        Gera o termo de devolução em memória e inicia o processo (status → Pendente Devolução).
        Retorna (True, bytes_do_docx, filename) ou (False, msg_erro, None).
        """
        item = self.find(item_id)
        if not item or item["status"] != "Indisponível":
            return False, "Apenas itens com status 'Indisponível' podem ter termo de devolução gerado.", None

        user = item.get("assigned_to")
        if not user:
            return False, "Não foi possível encontrar o usuário associado.", None

        revenda = item.get("revenda")
        modelo_path = TERMO_DEVOLUCAO_MODELOS.get(revenda)
        if not modelo_path:
            import os
            if not os.path.exists(modelo_path or ""):
                return False, f"Modelo de termo de devolução não encontrado para {revenda}.", None

        linked_peripherals = self.list_peripherals_for_equipment(item_id)
        if linked_peripherals:
            peripherals_text = "\n".join(
                f"- {p['tipo']}: {p.get('brand','')} {p.get('model','')} (S/N: {p.get('identificador') or 'N/A'})"
                for p in linked_peripherals
            )
        else:
            peripherals_text = "Nenhum periférico adicional devolvido."

        detalhes_parts = [f"{item.get('tipo','')} {item.get('brand','')} {item.get('model','')}".strip()]
        for key, label in {"identificador": "S/N", "cpu": "CPU", "ram": "RAM",
                           "storage": "Armazenamento", "setor": "Setor", "ip": "IP", "mac": "MAC"}.items():
            if item.get(key):
                detalhes_parts.append(f"{label}: {item[key]}")
        detalhes_finais = " - ".join(detalhes_parts)

        substituicoes = {
            "{{nome}}": user or "",
            "{{data_hoje}}": datetime.now().strftime("%d/%m/%Y"),
            "{{cpf}}": format_cpf(item.get("cpf", "")),
            "{{data_emprestimo}}": format_date(item.get("date_issued", "")),
            "{{tipo}}": item.get("tipo", ""),
            "{{detalhes_equipamento}}": detalhes_finais,
            "{{perifericos}}": peripherals_text,
        }
        substituicoes = {k: (v if v is not None else "") for k, v in substituicoes.items()}

        try:
            doc = Document(modelo_path)
            for p in doc.paragraphs:
                for k, v in substituicoes.items():
                    if k in p.text:
                        p.text = p.text.replace(k, str(v))
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            for k, v in substituicoes.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, str(v))
            buf = io.BytesIO()
            doc.save(buf)
            doc_bytes = buf.getvalue()
        except Exception:
            logger.exception("Erro ao gerar documento")
            return False, "Erro ao gerar o documento.", None

        # Atualiza status e registra no histórico
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("UPDATE items SET status='Pendente Devolução' WHERE id=%s", (item_id,))
        cur.execute(
            "INSERT INTO history (item_id, operador, data_operacao, operation) VALUES (%s,%s,%s,'Devolução')",
            (item_id, logged_user, datetime.now()),
        )
        conn.commit()
        cur.close()
        conn.close()

        safe_user = user.replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"termo_devolucao_{item_id}_{safe_user}_{revenda}_{timestamp}.docx"
        return True, doc_bytes, filename

    def confirm_return(self, item_id: int, logged_user: str, storage_key: str):
        """Confirma devolução. storage_key é a chave do PDF assinado no storage backend."""
        item = self.find(item_id)
        if not item or item["status"] != "Pendente Devolução":
            return False, "Apenas itens com status 'Pendente Devolução' podem ser confirmados."

        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute(
                "UPDATE items SET status='Disponível', assigned_to=NULL, cpf=NULL, date_issued=NULL WHERE id=%s",
                (item_id,),
            )
            for p in self.list_peripherals_for_equipment(item_id):
                cur.execute("UPDATE peripherals SET status='Disponível' WHERE id=%s", (p["id"],))
            cur.execute(
                """INSERT INTO history (item_id, operador, data_operacao, operation, termo_assinado_anexo)
                VALUES (%s,%s,%s,'Confirmação Devolução',%s)""",
                (item_id, logged_user, datetime.now(), storage_key),
            )
            conn.commit()
            return True, f"Devolução do item {item_id} confirmada."
        except Exception:
            conn.rollback()
            logger.exception("Erro ao confirmar devolução")
            return False, "Erro ao confirmar devolução."
        finally:
            cur.close()
            conn.close()

    # ── Loan term generation ───────────────────────────────────────────────────

    def generate_loan_term_bytes(self, item_id: int):
        """
        Gera o termo de responsabilidade em memória.
        Retorna (True, bytes_do_docx, filename) ou (False, msg_erro, None).
        """
        item = self.find(item_id)
        if not item:
            return False, "Equipamento não encontrado.", None
        if item["status"] != "Pendente":
            return False, "Este equipamento não está pendente de empréstimo.", None

        linked_peripherals = self.list_peripherals_for_equipment(item_id)
        if linked_peripherals:
            peripherals_text = "\n".join(
                f"- {p['tipo']}: {p.get('brand','')} {p.get('model','')} (S/N: {p.get('identificador') or 'N/A'})"
                for p in linked_peripherals
            )
        else:
            peripherals_text = "Nenhum periférico adicional vinculado."

        revenda = item.get("revenda")
        modelo_path = TERMO_MODELOS.get(revenda)
        if not modelo_path:
            import os
            if not os.path.exists(modelo_path or ""):
                return False, f"Modelo de termo não encontrado para {revenda}.", None

        user = item.get("assigned_to", "")
        detalhes_parts = [f"{item.get('tipo','')} {item.get('brand','')} {item.get('model','')}".strip()]
        for key, label in {"identificador": "S/N", "cpu": "CPU", "ram": "RAM",
                           "storage": "Armazenamento", "setor": "Setor", "ip": "IP", "mac": "MAC"}.items():
            if item.get(key):
                detalhes_parts.append(f"{label}: {item[key]}")
        detalhes_finais = " - ".join(detalhes_parts)

        substituicoes = {
            "{{nome}}": user or "",
            "{{data_hoje}}": datetime.now().strftime("%d/%m/%Y"),
            "{{cpf}}": format_cpf(item.get("cpf", "")),
            "{{data_emprestimo}}": format_date(item.get("date_issued", "")),
            "{{detalhes_equipamento}}": detalhes_finais,
            "{{perifericos}}": peripherals_text,
        }
        substituicoes = {k: (v if v is not None else "") for k, v in substituicoes.items()}

        try:
            doc = Document(modelo_path)
            for p in doc.paragraphs:
                for k, v in substituicoes.items():
                    if k in p.text:
                        p.text = p.text.replace(k, str(v))
            for tabela in doc.tables:
                for linha in tabela.rows:
                    for celula in linha.cells:
                        for p in celula.paragraphs:
                            for k, v in substituicoes.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, str(v))
            buf = io.BytesIO()
            doc.save(buf)
            doc_bytes = buf.getvalue()
        except Exception:
            logger.exception("Erro ao gerar documento")
            return False, "Erro ao gerar o documento.", None

        safe_user = str(user).replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"termo_{item_id}_{safe_user}_{revenda}_{timestamp}.docx"
        return True, doc_bytes, filename

    # ── History ────────────────────────────────────────────────────────────────

    def list_history(self):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute("""
            SELECT h.id, h.item_id, h.operador, h.peripheral_id, h.details,
                COALESCE(i.tipo, p.tipo, h.tipo) as tipo,
                COALESCE(i.brand, p.brand, h.brand) as marca,
                COALESCE(i.model, p.model, h.model) as modelo,
                COALESCE(i.identificador, p.identificador, h.identificador) as identificador,
                COALESCE(i.nota_fiscal, h.nota_fiscal) as nota_fiscal,
                COALESCE(i.fornecedor, h.fornecedor) as fornecedor,
                h.usuario, h.cpf, h.cargo, h.center_cost, h.setor, h.revenda,
                h.data_operacao, h.operation
            FROM history h
            LEFT JOIN items i ON i.id = h.item_id
            LEFT JOIN peripherals p ON p.id = h.peripheral_id
            WHERE h.is_reversed = 0
            ORDER BY h.data_operacao DESC, h.id DESC
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    def generate_monthly_report(self, ano, mes):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        sql = """
            (SELECT h.id AS history_id, h.item_id, h.peripheral_id, h.operador,
                h.usuario, h.cpf, h.cargo, h.center_cost, h.setor,
                COALESCE(i.fornecedor, h.fornecedor) AS fornecedor, h.revenda, h.details,
                h.data_operacao AS data_emprestimo,
                (SELECT MIN(hc.data_operacao) FROM history hc WHERE hc.item_id=h.item_id AND hc.operation='Confirmação Empréstimo' AND hc.id>h.id AND hc.is_reversed=0) AS data_confirmacao,
                (SELECT MIN(hd.data_operacao) FROM history hd WHERE hd.item_id=h.item_id AND hd.operation='Devolução' AND hd.id>h.id AND hd.is_reversed=0) AS data_devolucao,
                'Empréstimo' AS operation_type,
                COALESCE(i.tipo, h.tipo) AS tipo, COALESCE(i.brand, h.brand) AS brand,
                COALESCE(i.model, h.model) AS model, COALESCE(i.identificador, h.identificador) AS identificador,
                COALESCE(i.nota_fiscal, h.nota_fiscal) AS nota_fiscal
            FROM history h LEFT JOIN items i ON i.id=h.item_id
            WHERE h.operation='Empréstimo' AND h.is_reversed=0 AND YEAR(h.data_operacao)=%s AND MONTH(h.data_operacao)=%s)
            UNION ALL
            (SELECT h.id, h.item_id, NULL, h.operador, NULL, NULL, NULL, NULL, NULL,
                i.fornecedor, i.revenda, h.details, h.data_operacao, NULL, NULL,
                'Cadastro', COALESCE(i.tipo,h.tipo), COALESCE(i.brand,h.brand),
                COALESCE(i.model,h.model), COALESCE(i.identificador,h.identificador), COALESCE(i.nota_fiscal,h.nota_fiscal)
            FROM history h LEFT JOIN items i ON i.id=h.item_id
            WHERE h.operation='Cadastro' AND h.is_reversed=0 AND YEAR(h.data_operacao)=%s AND MONTH(h.data_operacao)=%s)
            UNION ALL
            (SELECT h.id, h.item_id, h.peripheral_id, h.operador, NULL, NULL, NULL, NULL, NULL,
                NULL, NULL, h.details, h.data_operacao, NULL, NULL,
                h.operation, p.tipo, p.brand, p.model, p.identificador, NULL
            FROM history h LEFT JOIN peripherals p ON p.id=h.peripheral_id
            WHERE h.operation IN ('Cadastro Periférico','Vínculo Periférico','Desvínculo Periférico','Substituição Periférico')
            AND h.is_reversed=0 AND YEAR(h.data_operacao)=%s AND MONTH(h.data_operacao)=%s)
            UNION ALL
            (SELECT h.id, h.item_id, NULL, h.operador, NULL, NULL, NULL, NULL, NULL,
                h.fornecedor, h.revenda, h.details, h.data_operacao, NULL, NULL,
                'Exclusão', h.tipo, h.brand, h.model, h.identificador, h.nota_fiscal
            FROM history h
            WHERE h.operation='Exclusão' AND h.is_reversed=0 AND YEAR(h.data_operacao)=%s AND MONTH(h.data_operacao)=%s)
            ORDER BY data_emprestimo, item_id
        """
        cur.execute(sql, (ano, mes, ano, mes, ano, mes, ano, mes))
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return rows

    def reverse_history_entry(self, history_id: int, logged_user: str):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute("SELECT * FROM history WHERE id=%s", (history_id,))
            entry = cur.fetchone()
            if not entry:
                return False, "Lançamento não encontrado."
            if entry["is_reversed"]:
                return False, "Esta operação já foi estornada."

            op = entry["operation"]
            item_id = entry["item_id"]

            if op == "Confirmação Empréstimo":
                cur.execute("UPDATE items SET status='Pendente' WHERE id=%s", (item_id,))
            elif op == "Empréstimo":
                cur.execute("UPDATE items SET status='Disponível', assigned_to=NULL, cpf=NULL, date_issued=NULL WHERE id=%s", (item_id,))
            elif op == "Confirmação Devolução":
                cur.execute("UPDATE items SET status='Pendente Devolução' WHERE id=%s", (item_id,))
            elif op == "Devolução":
                cur.execute("""
                    SELECT usuario, cpf, data_operacao FROM history
                    WHERE item_id=%s AND operation IN ('Empréstimo','Confirmação Empréstimo') AND id<%s AND is_reversed=0
                    ORDER BY data_operacao DESC, id DESC LIMIT 1
                """, (item_id, history_id))
                last_loan = cur.fetchone()
                if not last_loan:
                    conn.rollback()
                    return False, "Não foi possível encontrar o empréstimo original."
                cur.execute(
                    "UPDATE items SET status='Indisponível', assigned_to=%s, cpf=%s, date_issued=%s WHERE id=%s",
                    (last_loan["usuario"], last_loan["cpf"], last_loan["data_operacao"], item_id),
                )
            elif op == "Cadastro":
                cur.execute("UPDATE items SET is_active=0 WHERE id=%s", (item_id,))
            else:
                conn.rollback()
                return False, f"Não é possível estornar uma operação do tipo '{op}'."

            cur.execute("UPDATE history SET is_reversed=1 WHERE id=%s", (history_id,))
            cur.execute(
                """INSERT INTO history (item_id, operador, data_operacao, operation, usuario, cpf, cargo, center_cost, revenda)
                VALUES (%s,%s,%s,'Estorno',%s,%s,%s,%s,%s)""",
                (
                    item_id, logged_user, datetime.now(),
                    entry.get("usuario"), entry.get("cpf"),
                    entry.get("cargo"), entry.get("center_cost"), entry.get("revenda"),
                ),
            )
            conn.commit()
            return True, f"Operação '{op}' do item {item_id} estornada com sucesso."
        except pymysql.MySQLError:
            conn.rollback()
            logger.exception("Erro ao estornar")
            return False, "Erro ao estornar a operação."
        finally:
            cur.close()
            conn.close()

    # ── Charts ─────────────────────────────────────────────────────────────────

    def get_issue_return_counts(self, year: int, month: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        issues = {d: 0 for d in days}
        returns = {d: 0 for d in days}
        cur.execute(
            """SELECT DAY(data_operacao) as dia, operation, COUNT(id) as total
            FROM history
            WHERE YEAR(data_operacao)=%s AND MONTH(data_operacao)=%s
            AND operation IN ('Empréstimo','Devolução') AND is_reversed=0
            GROUP BY dia, operation""",
            (year, month),
        )
        for row in cur.fetchall():
            if row["operation"] == "Empréstimo":
                issues[row["dia"]] = row["total"]
            elif row["operation"] == "Devolução":
                returns[row["dia"]] = row["total"]
        cur.close()
        conn.close()
        return days, list(issues.values()), list(returns.values())

    def get_registration_counts(self, year: int, month: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        num_days = calendar.monthrange(year, month)[1]
        days = list(range(1, num_days + 1))
        registrations = {d: 0 for d in days}
        cur.execute(
            """SELECT DAY(data_operacao) as dia, COUNT(id) as total
            FROM history
            WHERE YEAR(data_operacao)=%s AND MONTH(data_operacao)=%s
            AND operation='Cadastro' AND is_reversed=0
            GROUP BY dia""",
            (year, month),
        )
        for row in cur.fetchall():
            registrations[row["dia"]] = row["total"]
        cur.close()
        conn.close()
        return days, list(registrations.values())

    def get_active_signed_term_key(self, item_id: int):
        conn = get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        cur.execute(
            "SELECT termo_assinado_anexo FROM history "
            "WHERE item_id = %s AND operation = 'Confirmação Empréstimo' "
            "ORDER BY id DESC LIMIT 1",
            (item_id,),
        )
        row = cur.fetchone()
        cur.close()
        conn.close()
        return row["termo_assinado_anexo"] if row else None
