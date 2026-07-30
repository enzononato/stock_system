"""
Camada de acesso ao banco de dados da funcionalidade de Unidades.

`items.revenda` e `history.revenda` continuam ligados a `unidades.nome` POR
NOME (sem chave estrangeira, sem migração dos dados existentes — decisão já
tomada e não reaberta aqui). Isso tem uma consequência direta neste módulo:
renomear uma unidade precisa cascatear para as duas tabelas na MESMA
transação (ver `update_unidade`), senão itens e histórico "perdem" a unidade
a que pertenciam.

Segue o mesmo padrão de `InventoryDBManager`: tabela criada de forma
idempotente em `_create_tables()` (chamada pelo `__init__`), todos os
métodos usando `database_mysql.get_cursor()` (commit/rollback/close
garantidos), mensagens genéricas ao cliente em caso de erro de banco e
`logger.exception` no servidor — nunca a exceção crua do pymysql é
devolvida na resposta.
"""
import logging

import pymysql

from app.db.database_mysql import get_cursor

logger = logging.getLogger(__name__)

# As 4 chaves que `itens_por_status` sempre traz, mesmo com contagem zero —
# mantém o mesmo conjunto do ENUM `items.status` (ver InventoryDBManager).
_STATUS_ITEM = ("Disponível", "Indisponível", "Pendente", "Pendente Devolução")


class UnidadeDBManager:
    def __init__(self):
        self._create_tables()

    def _create_tables(self):
        with get_cursor(dict_cursor=False) as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS unidades (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nome VARCHAR(100) NOT NULL UNIQUE,
                razao_social VARCHAR(200) NOT NULL,
                cnpj VARCHAR(18) NOT NULL,
                endereco VARCHAR(255),
                cep VARCHAR(10),
                cidade VARCHAR(100),
                uf CHAR(2),
                is_active TINYINT(1) DEFAULT 1
            )
            """)

    # ── Leitura ──────────────────────────────────────────────────────────────

    def list_unidades(self, *, include_inactive: bool = False) -> list[dict]:
        where = "" if include_inactive else "WHERE is_active = 1"
        with get_cursor(commit=False) as cur:
            cur.execute(f"SELECT * FROM unidades {where} ORDER BY nome")
            return cur.fetchall()

    def get_unidade(self, unidade_id: int) -> dict | None:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM unidades WHERE id=%s", (unidade_id,))
            return cur.fetchone()

    def get_unidade_por_nome(self, nome: str) -> dict | None:
        with get_cursor(commit=False) as cur:
            cur.execute("SELECT * FROM unidades WHERE nome=%s", (nome,))
            return cur.fetchone()

    # ── Escrita ──────────────────────────────────────────────────────────────

    def add_unidade(self, data: dict) -> tuple[bool, str]:
        keys = ", ".join(data.keys())
        placeholders = ", ".join(["%s"] * len(data))
        sql = f"INSERT INTO unidades ({keys}) VALUES ({placeholders})"
        try:
            with get_cursor(dict_cursor=False) as cur:
                cur.execute(sql, list(data.values()))
            return True, f"Unidade '{data.get('nome')}' cadastrada com sucesso."
        except pymysql.MySQLError as e:
            if e.args and e.args[0] == 1062:
                return False, "Já existe uma unidade com este nome."
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."

    def update_unidade(self, unidade_id: int, data: dict) -> tuple[bool, str]:
        """Atualiza os dados cadastrais de uma unidade.

        Quando `data` contém `nome` diferente do atual, a mudança CASCATEIA
        para `items.revenda` e `history.revenda` na mesma transação — tudo
        dentro de um único `with get_cursor()`, então qualquer falha em
        qualquer etapa (inclusive o 1062 de nome duplicado) faz rollback de
        tudo, e nada fica aplicado pela metade.
        """
        try:
            with get_cursor() as cur:
                cur.execute("SELECT * FROM unidades WHERE id=%s", (unidade_id,))
                atual = cur.fetchone()
                if not atual:
                    return False, "Unidade não encontrada."

                novo_nome = data.get("nome")
                nome_mudou = novo_nome is not None and novo_nome != atual["nome"]

                sets = ", ".join(f"{k}=%s" for k in data.keys())
                valores = list(data.values()) + [unidade_id]
                cur.execute(f"UPDATE unidades SET {sets} WHERE id=%s", valores)

                itens_reapontados = 0
                historico_reapontado = 0
                if nome_mudou:
                    cur.execute(
                        "UPDATE items SET revenda=%s WHERE revenda=%s",
                        (novo_nome, atual["nome"]),
                    )
                    itens_reapontados = cur.rowcount
                    cur.execute(
                        "UPDATE history SET revenda=%s WHERE revenda=%s",
                        (novo_nome, atual["nome"]),
                    )
                    historico_reapontado = cur.rowcount

            if nome_mudou:
                return True, (
                    f"Unidade atualizada. Nome alterado de '{atual['nome']}' para "
                    f"'{novo_nome}': {itens_reapontados} item(ns) e "
                    f"{historico_reapontado} registro(s) de histórico reapontados."
                )
            return True, "Unidade atualizada com sucesso."
        except pymysql.MySQLError as e:
            if e.args and e.args[0] == 1062:
                return False, "Já existe uma unidade com este nome."
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."

    def deactivate_unidade(self, unidade_id: int) -> tuple[bool, str]:
        """Inativa (nunca apaga) uma unidade — recusa se houver itens ativos
        associados, para não deixar os seletores sem a opção e travar
        operações sobre esses itens."""
        try:
            with get_cursor() as cur:
                cur.execute("SELECT * FROM unidades WHERE id=%s", (unidade_id,))
                unidade = cur.fetchone()
                if not unidade:
                    return False, "Unidade não encontrada."
                if not unidade["is_active"]:
                    return False, "Esta unidade já está inativa."

                cur.execute(
                    "SELECT COUNT(*) as total FROM items WHERE revenda=%s AND is_active=1",
                    (unidade["nome"],),
                )
                total_itens = cur.fetchone()["total"]
                if total_itens > 0:
                    return False, (
                        f"Não é possível inativar: {total_itens} item(ns) ativo(s) "
                        "ainda associado(s) a esta unidade."
                    )

                cur.execute("UPDATE unidades SET is_active=0 WHERE id=%s", (unidade_id,))
            return True, f"Unidade '{unidade['nome']}' inativada com sucesso."
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."

    def reactivate_unidade(self, unidade_id: int) -> tuple[bool, str]:
        """Reverte a inativação (`is_active=1`) — dá ao Gestor um caminho de
        volta caso a unidade tenha sido inativada por engano.

        DECISÃO DE PRODUTO: recusa (não idempotente) se a unidade já estiver
        ativa, devolvendo mensagem clara em vez de tratar como sucesso
        silencioso. Escolhido por simetria com `deactivate_unidade`, que já
        recusa (mesma mensagem no padrão inverso) quando a unidade já está
        inativa — mantém o mesmo comportamento "recusa estado redundante" nas
        duas pontas em vez de um par assimétrico (uma tolerante, outra não)."""
        try:
            with get_cursor() as cur:
                cur.execute("SELECT * FROM unidades WHERE id=%s", (unidade_id,))
                unidade = cur.fetchone()
                if not unidade:
                    return False, "Unidade não encontrada."
                if unidade["is_active"]:
                    return False, "Esta unidade já está ativa."

                cur.execute("UPDATE unidades SET is_active=1 WHERE id=%s", (unidade_id,))
            return True, f"Unidade '{unidade['nome']}' reativada com sucesso."
        except pymysql.MySQLError:
            logger.exception("Erro de banco de dados")
            return False, "Erro ao processar a operação. Tente novamente."

    # ── Indicadores ──────────────────────────────────────────────────────────

    def get_indicadores(self, unidade_id: int) -> dict:
        """Indicadores agregados de uma unidade, todos filtrados por
        `revenda = <nome da unidade>` e ignorando histórico estornado
        (`is_reversed = 0`). Se a unidade não existir, devolve tudo zerado
        em vez de estourar — quem decide 404 é o router, que já checa a
        existência antes de chamar este método."""
        resultado = {
            "termos_emitidos": 0,
            "termos_confirmados": 0,
            "devolucoes_concluidas": 0,
            "itens_total": 0,
            "itens_por_status": {status: 0 for status in _STATUS_ITEM},
            "emprestimos_ativos": 0,
            "perifericos_vinculados": 0,
        }

        unidade = self.get_unidade(unidade_id)
        if not unidade:
            return resultado
        nome = unidade["nome"]

        with get_cursor(commit=False) as cur:
            # As três contagens de history abaixo usam JOIN com items (via item_id)
            # em vez de filtrar direto por history.revenda. Motivo: verificado em
            # InventoryDBManager (app/db/inventory_manager_db.py, fora da posse
            # deste módulo) que generate_return_term_bytes() e confirm_return()
            # NUNCA preenchem history.revenda ao inserir 'Devolução'/'Confirmação
            # Devolução' — só issue() e confirm_loan() preenchem essa coluna para
            # 'Empréstimo'/'Confirmação Empréstimo'. Filtrar 'Confirmação Devolução'
            # por history.revenda daria sempre zero. items.revenda é sempre
            # confiável (inclusive após rename, que cascateia para as duas
            # tabelas), então o JOIN por item_id é a forma correta e uniforme de
            # atribuir cada operação à unidade certa nos três casos.
            cur.execute(
                """SELECT COUNT(*) as total FROM history h
                JOIN items i ON i.id = h.item_id
                WHERE i.revenda=%s AND h.operation='Empréstimo' AND h.is_reversed=0""",
                (nome,),
            )
            resultado["termos_emitidos"] = cur.fetchone()["total"]

            cur.execute(
                """SELECT COUNT(*) as total FROM history h
                JOIN items i ON i.id = h.item_id
                WHERE i.revenda=%s AND h.operation='Confirmação Empréstimo' AND h.is_reversed=0""",
                (nome,),
            )
            resultado["termos_confirmados"] = cur.fetchone()["total"]

            cur.execute(
                """SELECT COUNT(*) as total FROM history h
                JOIN items i ON i.id = h.item_id
                WHERE i.revenda=%s AND h.operation='Confirmação Devolução' AND h.is_reversed=0""",
                (nome,),
            )
            resultado["devolucoes_concluidas"] = cur.fetchone()["total"]

            cur.execute(
                "SELECT COUNT(*) as total FROM items WHERE revenda=%s AND is_active=1",
                (nome,),
            )
            resultado["itens_total"] = cur.fetchone()["total"]

            cur.execute(
                "SELECT status, COUNT(*) as total FROM items "
                "WHERE revenda=%s AND is_active=1 GROUP BY status",
                (nome,),
            )
            for row in cur.fetchall():
                if row["status"] in resultado["itens_por_status"]:
                    resultado["itens_por_status"][row["status"]] = row["total"]
            resultado["emprestimos_ativos"] = resultado["itens_por_status"]["Indisponível"]

            cur.execute(
                """SELECT COUNT(*) as total FROM equipment_peripherals ep
                JOIN items i ON i.id = ep.equipment_id
                WHERE i.revenda=%s AND i.is_active=1""",
                (nome,),
            )
            resultado["perifericos_vinculados"] = cur.fetchone()["total"]

        return resultado


# Instância reaproveitada pelo processo. `__init__` roda `CREATE TABLE IF NOT
# EXISTS`, então instanciar a cada chamada — como a geração de termo fazia —
# dispara um DDL a cada operação de negócio. A criação é preguiçosa: o primeiro
# uso conecta, e o import do módulo continua sem tocar no banco, preservando a
# propriedade de o app subir com o MySQL indisponível.
_instancia: "UnidadeDBManager | None" = None


def get_unidade_manager() -> "UnidadeDBManager":
    """Devolve a instância compartilhada de UnidadeDBManager, criando-a na
    primeira chamada. Testes que precisem recriar o schema podem seguir
    instanciando `UnidadeDBManager()` diretamente."""
    global _instancia
    if _instancia is None:
        _instancia = UnidadeDBManager()
    return _instancia
