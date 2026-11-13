"""
Zera o 2FA de um usuário (segredo e códigos de recuperação) e encerra as sessões dele.

Uso (a partir de backend/, rode só no servidor):
    python -m scripts.reset_2fa <usuario>

É a saída final para quem perdeu o autenticador e os códigos de recuperação sem que
outro Gestor possa zerar pela tela de Usuários. Fica registrado na auditoria como
ação do "sistema". No próximo login, um Gestor cadastra um novo autenticador; os
demais papéis voltam a entrar só com a senha. Não precisa da chave TOTP_ENCRYPTION_KEY.
"""
import argparse

from app.db.mfa_db import MfaDBManager
from app.db.user_manager_db import UserDBManager
from app.services.audit import AuditContext
from app.services.segundo_fator import zerar_2fa


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Zera o 2FA de um usuário do Patrimônio.")
    parser.add_argument("usuario", help="nome de usuário (login)")
    args = parser.parse_args(argv)

    usuario = UserDBManager().get_user_by_username(args.usuario)
    if usuario is None:
        print(f"Usuário '{args.usuario}' não encontrado.")
        return 1
    if not usuario.get("totp_ativo"):
        print(f"O usuário '{args.usuario}' não tem 2FA ativo. Nada a fazer.")
        return 0

    zerar_2fa(MfaDBManager(), usuario, AuditContext.sistema(), origem="script reset_2fa")
    print(
        f"2FA de '{args.usuario}' zerado e sessões encerradas. "
        "No próximo login, um Gestor cadastra um novo autenticador."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
