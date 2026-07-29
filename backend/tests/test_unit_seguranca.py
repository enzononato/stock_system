"""
Testes unitários puros para backend/app/core/security.py (hash de senha e JWT).

Importar app.core.security/app.core.config não faz nenhum I/O (Settings() só lê
variáveis de ambiente/.env via Pydantic) — por isso estes testes não precisam do banco
de teste nem do guard de ambiente, e sempre rodam.
"""
from datetime import datetime, timedelta

import pytest
from jose import jwt
from fastapi import HTTPException

from app.core.config import settings
from app.core.security import (
    hash_password,
    verify_password,
    create_access_token,
    create_refresh_token,
    decode_token,
)

pytestmark = pytest.mark.unit


class TestHashDeSenha:
    def test_hash_e_verificacao_bem_sucedida(self):
        hashed = hash_password("MinhaSenha#123")
        assert hashed != "MinhaSenha#123"
        assert verify_password("MinhaSenha#123", hashed) is True

    def test_verificacao_falha_com_senha_errada(self):
        hashed = hash_password("MinhaSenha#123")
        assert verify_password("SenhaErrada", hashed) is False


class TestTokens:
    def test_access_token_contem_claims_esperadas(self):
        # Usamos get_unverified_claims (só decodifica o payload, sem validar "sub") em vez
        # de jwt.decode()/decode_token() propositalmente: o objetivo aqui é só conferir que
        # create_access_token() monta as claims certas. A validação no decode tem um bug
        # próprio, caracterizado abaixo em testes dedicados.
        token = create_access_token({"sub": 1, "username": "fulano", "role": "Gestor"})
        payload = jwt.get_unverified_claims(token)
        assert payload["sub"] == 1
        assert payload["username"] == "fulano"
        assert payload["role"] == "Gestor"
        assert payload["type"] == "access"
        assert "exp" in payload

    def test_refresh_token_marca_type_refresh(self):
        token = create_refresh_token({"sub": 1, "username": "fulano", "role": "Gestor"})
        payload = jwt.get_unverified_claims(token)
        assert payload["type"] == "refresh"

    def test_decode_token_funciona_quando_sub_e_string(self):
        """Isola que decode_token() em si funciona bem — o problema (abaixo) é
        especificamente o tipo do valor colocado em "sub" por quem gera o token."""
        token = create_access_token({"sub": "2", "username": "ciclana", "role": "Técnico"})
        payload = decode_token(token, expected_type="access")
        assert payload["username"] == "ciclana"

    def test_decode_token_falha_com_sub_inteiro_como_o_login_real_gera(self):
        """
        BUG CONHECIDO (CRÍTICO): app/routers/auth.py::login() monta o token com
        `token_data = {"sub": user["id"], ...}`, e `user["id"]` é um INT nativo (coluna
        AUTO_INCREMENT do MySQL, lido via pymysql.cursors.DictCursor) — não uma string.
        `python-jose` valida no decode, por padrão (verify_sub=True), que o claim "sub"
        seja uma string (jose/jwt.py:_validate_sub); com um "sub" inteiro, ele levanta
        JWTClaimsError("Subject must be a string."), que decode_token() captura como
        JWTError genérico e converte em 401 "Token inválido ou expirado.".

        Resultado prático: TODO token de acesso emitido pelo fluxo de login real é
        rejeitado por decode_token() — ou seja, get_current_user()/gestor_only()/
        gestor_or_tecnico() falham para qualquer usuário, em qualquer rota autenticada,
        sempre. Ver a demonstração ponta a ponta via HTTP em
        test_auth.py::test_token_de_login_real_nao_autentica_em_nenhuma_rota.
        """
        token = create_access_token({"sub": 2, "username": "ciclana", "role": "Técnico"})
        with pytest.raises(HTTPException) as exc_info:
            decode_token(token, expected_type="access")
        assert exc_info.value.status_code == 401
        assert exc_info.value.detail == "Token inválido ou expirado."

    def test_decode_token_rejeita_type_incompativel(self):
        """Um refresh token não pode ser usado como access token (e vice-versa).

        Usa "sub" como string de propósito, para isolar esta checagem (type incompatível)
        do BUG CONHECIDO caracterizado acima (sub inteiro) — aqui queremos garantir que,
        mesmo corrigido o bug do sub, a checagem de "type" continua funcionando.
        """
        refresh = create_refresh_token({"sub": "1", "username": "fulano", "role": "Gestor"})
        with pytest.raises(HTTPException) as exc_info:
            decode_token(refresh, expected_type="access")
        assert exc_info.value.status_code == 401
        assert exc_info.value.detail == "Token inválido ou expirado."

    def test_decode_token_rejeita_assinatura_invalida(self):
        # "sub" como string: isola a checagem de assinatura do bug de sub inteiro.
        token_forjado = jwt.encode(
            {"sub": "1", "username": "invasor", "role": "Gestor", "type": "access",
             "exp": datetime.utcnow() + timedelta(minutes=5)},
            "chave-secreta-errada",
            algorithm=settings.JWT_ALGORITHM,
        )
        with pytest.raises(HTTPException) as exc_info:
            decode_token(token_forjado, expected_type="access")
        assert exc_info.value.status_code == 401
        assert exc_info.value.detail == "Token inválido ou expirado."

    def test_decode_token_rejeita_token_expirado(self):
        # "sub" como string: isola a checagem de expiração do bug de sub inteiro.
        token_expirado = jwt.encode(
            {"sub": "1", "username": "fulano", "role": "Gestor", "type": "access",
             "exp": datetime.utcnow() - timedelta(minutes=1)},
            settings.JWT_SECRET,
            algorithm=settings.JWT_ALGORITHM,
        )
        with pytest.raises(HTTPException) as exc_info:
            decode_token(token_expirado, expected_type="access")
        assert exc_info.value.status_code == 401
        assert exc_info.value.detail == "Token inválido ou expirado."

    def test_decode_token_rejeita_lixo(self):
        with pytest.raises(HTTPException) as exc_info:
            decode_token("isto-nao-e-um-jwt", expected_type="access")
        assert exc_info.value.status_code == 401
