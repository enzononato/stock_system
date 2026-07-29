from fastapi import FastAPI, Request
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from slowapi import _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded

from app.core.config import settings
from app.dependencies import limiter, get_inventory_db
from app.routers import auth, users, items, peripherals, loans, documents, history, reports, constants

app = FastAPI(
    title="Controle de Estoque Revalle",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    openapi_url="/api/openapi.json",
)

# Rate limiting (T2): estado do limiter + handler da exceção de 429.
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)


# Rótulos amigáveis para os campos que o usuário vê nos formulários. Sem isso a
# mensagem citaria o nome técnico da coluna ("nota_fiscal", "date_issue").
_ROTULOS_DE_CAMPO = {
    "cpf": "CPF",
    "nota_fiscal": "Nota fiscal",
    "mac": "MAC",
    "endereco_fisico": "Endereço físico (MAC)",
    "ip": "IP",
    "ip_snmp": "IP da placa SNMP",
    "date_registered": "Data de cadastro",
    "date_issue": "Data do empréstimo",
    "tipo": "Tipo",
    "brand": "Marca",
    "model": "Modelo",
    "revenda": "Revenda",
    "usuario": "Funcionário",
    "cargo": "Cargo",
    "setor": "Setor",
    "center_cost": "Centro de custo",
    "identificador": "Identificador",
    "password": "Senha",
    "username": "Usuário",
    "role": "Função",
    "new_password": "Nova senha",
    "reason": "Motivo",
}


@app.exception_handler(RequestValidationError)
def erro_de_validacao(request: Request, exc: RequestValidationError) -> JSONResponse:
    """Converte o 422 padrão do FastAPI num formato que o frontend saiba exibir.

    Por padrão o FastAPI responde 422 com `detail` sendo uma LISTA de objetos
    ({loc, msg, type}). Todo o frontend deste projeto lê `response.data.detail`
    esperando uma string e a joga direto num toast — com a lista, o usuário
    veria "[object Object]".

    Isso passou a importar quando as validações de domínio (CPF, nota fiscal,
    MAC, IP, datas) migraram dos routers para os schemas Pydantic: antes elas
    viravam HTTPException 400 com mensagem em português, agora são capturadas
    pelo Pydantic e virariam 422 com a lista. Aqui reduzimos a uma frase única,
    em português, citando o rótulo do campo — e mantemos a lista original em
    `errors` para quem estiver depurando ou consumindo a API programaticamente.
    """
    mensagens: list[str] = []
    for erro in exc.errors():
        # `loc` costuma ser ("body", "campo"); o último elemento textual é o campo.
        campo = next(
            (parte for parte in reversed(erro.get("loc", ())) if isinstance(parte, str) and parte != "body"),
            None,
        )
        msg = erro.get("msg", "valor inválido")
        # Pydantic prefixa mensagens de ValueError com "Value error, ".
        msg = msg.removeprefix("Value error, ")
        rotulo = _ROTULOS_DE_CAMPO.get(campo, campo)
        mensagens.append(f"{rotulo}: {msg}" if rotulo else msg)

    return JSONResponse(
        status_code=422,
        content={
            "detail": " | ".join(mensagens) or "Dados inválidos.",
            # jsonable_encoder é necessário: em Pydantic v2 os erros podem trazer
            # objetos de exceção em `ctx`, que json.dumps não serializa.
            "errors": jsonable_encoder(exc.errors()),
        },
    )

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(auth.router)
app.include_router(users.router)
app.include_router(items.router)
app.include_router(peripherals.router)
app.include_router(loans.router)
app.include_router(documents.router)
app.include_router(history.router)
app.include_router(reports.router)
app.include_router(constants.router)


@app.get("/api/health")
def health():
    """Healthcheck simples, sem tocar no banco: sempre responde 200 enquanto
    o processo estiver de pé (T11)."""
    return {"status": "ok"}


@app.get("/api/health/db")
def health_db():
    """Healthcheck que reporta o estado da conexão com o banco, sem derrubar
    o healthcheck principal caso o MySQL esteja indisponível (T11)."""
    try:
        get_inventory_db()
    except Exception as exc:
        return JSONResponse(
            status_code=503,
            content={"status": "unavailable", "database": "unreachable", "detail": str(exc)},
        )
    return {"status": "ok", "database": "reachable"}
