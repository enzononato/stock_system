"""Login por e-mail, Task 1: validação do e-mail corporativo e serviço de confirmação."""
import re

import pytest

from app.services.audit import AuditContext
from app.services.mail import limpar_outbox, outbox
from tests.fixtures_auth import ida_e_volta_migracao, linhas_auditoria

pytestmark = pytest.mark.integration

CTX = AuditContext(usuario_id=None, username=None, papel=None, ip="127.0.0.1")


@pytest.fixture(autouse=True)
def _outbox_limpo():
    limpar_outbox()
    yield
    limpar_outbox()


def _consultar(sql: str, params=()):
    from app.db.database_mysql import get_cursor

    with get_cursor(commit=False) as cur:
        cur.execute(sql, params)
        return list(cur.fetchall())


def _codigo_do_email() -> str:
    m = re.search(r"\b(\d{6})\b", outbox[-1].text)
    assert m, outbox[-1].text
    return m.group(1)


@pytest.fixture
def conta_antiga(user_manager, senha_padrao_teste) -> dict:
    """Conta sem e-mail (como as de produção antes da mudança)."""
    ok, msg = user_manager.add_user("joao", senha_padrao_teste, "Técnico")
    assert ok, msg
    return user_manager.get_user_by_username("joao")


def _envelhecer(usuario_id: int, executar_sql, segundos: int = 120) -> None:
    """Passa o intervalo de reenvio (60 s) com folga: o relógio do contêiner de teste
    às vezes se ajusta no meio da suíte."""
    executar_sql(
        "UPDATE email_confirmacoes SET criado_em = criado_em - INTERVAL %s SECOND WHERE usuario_id = %s",
        (segundos, usuario_id),
    )


def test_valida_dominio_normaliza():
    from app.schemas.validators import validar_email_corporativo

    assert validar_email_corporativo(" Joao.Silva@Revalle.COM.br ") == "joao.silva@revalle.com.br"


@pytest.mark.parametrize(
    "email",
    ["a@revalle.com.br.evil.com", "a@evilrevalle.com.br", "a@sub.revalle.com.br", "a@gmail.com", "semarroba"],
)
def test_recusa_dominios_parecidos(email):
    from app.schemas.validators import validar_email_corporativo

    with pytest.raises(ValueError, match=r"Use o seu e-mail @revalle.com.br\."):
        validar_email_corporativo(email)


def test_iniciar_envia_codigo_e_guarda_so_hash(conta_antiga):
    from app.services import confirmacao_email

    email = confirmacao_email.iniciar(conta_antiga, " Joao@Revalle.com.br ", CTX)

    assert email == "joao@revalle.com.br"
    assert len(outbox) == 1
    assert outbox[0].to == ["joao@revalle.com.br"]
    assert outbox[0].subject == "Código de confirmação — Patrimônio Revalle"
    assert "Vale por 15 minutos." in outbox[0].text
    codigo = _codigo_do_email()
    linhas = _consultar("SELECT * FROM email_confirmacoes WHERE usuario_id = %s", (conta_antiga["id"],))
    assert len(linhas) == 1
    assert linhas[0]["email"] == "joao@revalle.com.br"
    assert linhas[0]["codigo_hash"] != codigo
    assert codigo not in linhas[0]["codigo_hash"]
    assert len(linhas[0]["codigo_hash"]) == 64


def test_iniciar_recusa_dominio_de_fora(conta_antiga):
    from app.services import confirmacao_email

    with pytest.raises(confirmacao_email.DominioInvalido) as exc:
        confirmacao_email.iniciar(conta_antiga, "joao@gmail.com", CTX)
    assert exc.value.status_code == 422
    assert str(exc.value) == "Use o seu e-mail @revalle.com.br."
    assert outbox == []


def test_reenvio_antes_de_60s_recusado(conta_antiga, executar_sql):
    from app.services import confirmacao_email

    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo_antigo = _codigo_do_email()

    with pytest.raises(confirmacao_email.AguardeReenvio) as exc:
        confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    assert exc.value.status_code == 429
    assert str(exc.value) == "Aguarde 1 minuto para pedir outro código."
    assert len(outbox) == 1

    _envelhecer(conta_antiga["id"], executar_sql)
    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    assert len(outbox) == 2
    codigo_novo = _codigo_do_email()

    if codigo_antigo != codigo_novo:  # 1 em 1 milhão de serem iguais
        with pytest.raises(confirmacao_email.CodigoInvalido):
            confirmacao_email.confirmar(conta_antiga, codigo_antigo, CTX)
    assert confirmacao_email.confirmar(conta_antiga, codigo_novo, CTX) == "joao@revalle.com.br"


def test_email_de_outra_conta_recusado(conta_antiga, user_manager, senha_padrao_teste):
    from app.services import confirmacao_email

    ok, msg = user_manager.add_user("maria", senha_padrao_teste, "Técnico", "maria@revalle.com.br")
    assert ok, msg

    with pytest.raises(confirmacao_email.EmailEmUsoErro) as exc:
        confirmacao_email.iniciar(conta_antiga, "MARIA@revalle.com.br", CTX)
    assert exc.value.status_code == 409
    assert str(exc.value) == "Este e-mail já está em uso por outra conta."
    assert outbox == []


def test_confirmar_grava_email_confirmado(conta_antiga, user_manager):
    from app.services import confirmacao_email

    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo = _codigo_do_email()

    assert confirmacao_email.confirmar(conta_antiga, codigo, CTX) == "joao@revalle.com.br"

    atual = user_manager.get_user_by_id(conta_antiga["id"])
    assert atual["email"] == "joao@revalle.com.br"
    assert atual["email_confirmado"] == 1
    linhas = linhas_auditoria("usuario", "confirmar_email")
    assert len(linhas) == 1
    assert linhas[0]["entidade_id"] == str(conta_antiga["id"])
    assert linhas[0]["depois"] == {"email": "j***@revalle.com.br"}
    texto = str(linhas[0])
    assert "joao@revalle.com.br" not in texto
    assert codigo not in texto


def test_codigo_errado_5_vezes_esgota(conta_antiga):
    from app.services import confirmacao_email

    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo = _codigo_do_email()
    errado = f"{(int(codigo) + 1) % 10**6:06d}"

    for _ in range(4):
        with pytest.raises(confirmacao_email.CodigoInvalido) as exc:
            confirmacao_email.confirmar(conta_antiga, errado, CTX)
        assert exc.value.status_code == 400
        assert str(exc.value) == "Código inválido ou expirado."
    with pytest.raises(confirmacao_email.CodigoEsgotado) as exc:
        confirmacao_email.confirmar(conta_antiga, errado, CTX)
    assert exc.value.status_code == 400
    assert str(exc.value) == "Muitas tentativas. Entre de novo para receber outro código."

    with pytest.raises(confirmacao_email.ConfirmacaoErro):
        confirmacao_email.confirmar(conta_antiga, codigo, CTX)
    assert confirmacao_email_nao_gravou(conta_antiga)


def confirmacao_email_nao_gravou(usuario: dict) -> bool:
    from app.db.user_manager_db import UserDBManager

    atual = UserDBManager().get_user_by_id(usuario["id"])
    return not atual["email_confirmado"] and atual["email"] is None


def test_codigo_vencido_e_reuso_recusados(conta_antiga, executar_sql):
    from app.services import confirmacao_email

    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo = _codigo_do_email()
    executar_sql(
        "UPDATE email_confirmacoes SET expira_em = expira_em - INTERVAL 16 MINUTE WHERE usuario_id = %s",
        (conta_antiga["id"],),
    )
    with pytest.raises(confirmacao_email.CodigoInvalido):
        confirmacao_email.confirmar(conta_antiga, codigo, CTX)
    assert confirmacao_email_nao_gravou(conta_antiga)

    _envelhecer(conta_antiga["id"], executar_sql)
    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo = _codigo_do_email()
    assert confirmacao_email.confirmar(conta_antiga, codigo, CTX) == "joao@revalle.com.br"
    with pytest.raises(confirmacao_email.CodigoInvalido):
        confirmacao_email.confirmar(conta_antiga, codigo, CTX)


def test_codigo_vale_so_para_o_email_pedido(conta_antiga, user_manager, senha_padrao_teste, executar_sql):
    """O código é de uma conta só: outra conta não confirma com ele."""
    from app.services import confirmacao_email

    ok, _ = user_manager.add_user("ana", senha_padrao_teste, "Técnico")
    ana = user_manager.get_user_by_username("ana")
    confirmacao_email.iniciar(conta_antiga, "joao@revalle.com.br", CTX)
    codigo = _codigo_do_email()
    confirmacao_email.iniciar(ana, "ana@revalle.com.br", CTX)
    executar_sql("UPDATE email_confirmacoes SET codigo_hash = %s WHERE usuario_id = %s",
                 (confirmacao_email._hash_codigo(conta_antiga["id"], "joao@revalle.com.br", codigo), ana["id"]))

    with pytest.raises(confirmacao_email.CodigoInvalido):
        confirmacao_email.confirmar(ana, codigo, CTX)


def test_update_email_zera_confirmacao(conta_antiga, user_manager):
    user_manager.confirmar_email(conta_antiga["id"], "joao@revalle.com.br")
    assert user_manager.get_user_by_id(conta_antiga["id"])["email_confirmado"] == 1

    # Salvar o mesmo e-mail não tira a confirmação.
    ok, _ = user_manager.update_email(conta_antiga["id"], "joao@revalle.com.br")
    assert ok
    assert user_manager.get_user_by_id(conta_antiga["id"])["email_confirmado"] == 1

    ok, _ = user_manager.update_email(conta_antiga["id"], "joao.novo@revalle.com.br")
    assert ok
    atual = user_manager.get_user_by_id(conta_antiga["id"])
    assert atual["email"] == "joao.novo@revalle.com.br"
    assert atual["email_confirmado"] == 0


def test_confirmar_email_duplicado_levanta(conta_antiga, user_manager, senha_padrao_teste):
    from app.db.user_manager_db import EmailEmUso

    user_manager.add_user("maria", senha_padrao_teste, "Técnico", "maria@revalle.com.br")
    with pytest.raises(EmailEmUso):
        user_manager.confirmar_email(conta_antiga["id"], "maria@revalle.com.br")
    assert user_manager.email_em_uso_por_outro("maria@revalle.com.br", conta_antiga["id"])
    assert not user_manager.email_em_uso_por_outro("joao@revalle.com.br", conta_antiga["id"])


def test_get_user_para_login(conta_antiga, user_manager):
    user_manager.update_email(conta_antiga["id"], "joao@revalle.com.br")

    por_usuario = user_manager.get_user_para_login("joao")
    por_email = user_manager.get_user_para_login("JOAO@Revalle.com.br")
    assert por_usuario["id"] == por_email["id"] == conta_antiga["id"]
    for campo in ("id", "username", "password", "role", "email", "email_confirmado", "totp_ativo", "sessao_versao"):
        assert campo in por_email
    assert user_manager.get_user_para_login("ninguem@revalle.com.br") is None
    assert user_manager.get_user_para_login("ninguem") is None


def test_gerar_username_unico(conta_antiga, user_manager, senha_padrao_teste):
    user_manager.add_user("joao.silva", senha_padrao_teste, "Técnico")
    assert user_manager.gerar_username("joao.silva@revalle.com.br") == "joao.silva2"
    user_manager.add_user("joao.silva2", senha_padrao_teste, "Técnico")
    assert user_manager.gerar_username("Joao.Silva@revalle.com.br") == "joao.silva3"
    assert user_manager.gerar_username("maria@revalle.com.br") == "maria"


def test_migracao_0900_ida_e_volta(_test_schema):
    ida_e_volta_migracao(900)
    colunas = _consultar(
        "SELECT column_name AS nome FROM information_schema.columns "
        "WHERE table_schema = DATABASE() AND table_name = 'usuarios' AND column_name IN ('email_confirmado', 'senha_definida')"
    )
    assert len(colunas) == 2
    assert _consultar(
        "SELECT table_name FROM information_schema.tables "
        "WHERE table_schema = DATABASE() AND table_name = 'email_confirmacoes'"
    )
