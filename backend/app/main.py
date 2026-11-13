import logging

from fastapi import FastAPI, Request
from fastapi.encoders import jsonable_encoder
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from slowapi import _rate_limit_exceeded_handler
from slowapi.errors import RateLimitExceeded

from app import jobs
from app.core.config import settings
from app.db.migrations import aplicar_pendentes, garantir_baseline
from app.dependencies import limiter, get_inventory_db
from app.services import scheduler
from app.routers import auth, users, items, peripherals, loans, documents, history, reports, constants, unidades, audit, colaboradores, locais
from app.routers import termos, integracoes_revsign

logger = logging.getLogger(__name__)

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
    # Campos de Unidade
    "nome": "Nome da unidade",
    "razao_social": "Razão social",
    "cnpj": "CNPJ",
    "endereco": "Endereço",
    "cep": "CEP",
    "cidade": "Cidade",
    "uf": "UF",
    # Campos de Colaborador
    "email": "E-mail",
    "telefone": "Telefone",
    "superior_id": "Superior",
    "unidade_id": "Unidade",
    "situacao": "Situação",
    # Campos de Local e vínculos do item
    "descricao": "Descrição",
    "ativo": "Ativo",
    "colaborador_id": "Colaborador",
    "local_id": "Local",
    "agendado_para": "Data agendada",
}

# "nome" vale "Nome da unidade" nas unidades; nas rotas de colaboradores e locais é só "Nome".
_ROTULOS_POR_ROTA = {
    "/api/colaboradores": {"nome": "Nome"},
    "/api/locais": {"nome": "Nome"},
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
    # Caminho da ROTA (ex.: "/api/colaboradores/{colaborador_id}") quando conhecido; senão o da requisição.
    caminho = getattr(request.scope.get("route"), "path", None) or request.url.path
    for erro in exc.errors():
        # `loc` costuma ser ("body", "campo"); o último elemento textual é o campo.
        campo = next(
            (parte for parte in reversed(erro.get("loc", ())) if isinstance(parte, str) and parte != "body"),
            None,
        )
        msg = erro.get("msg", "valor inválido")
        # Pydantic prefixa mensagens de ValueError com "Value error, ".
        msg = msg.removeprefix("Value error, ")
        rotulos = _ROTULOS_DE_CAMPO
        for prefixo, especificos in _ROTULOS_POR_ROTA.items():
            if caminho == prefixo or caminho.startswith(prefixo + "/"):
                rotulos = {**_ROTULOS_DE_CAMPO, **especificos}
        rotulo = rotulos.get(campo, campo)
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
app.include_router(unidades.router)
app.include_router(audit.router)
app.include_router(colaboradores.router)
app.include_router(locais.router)
app.include_router(termos.router)
app.include_router(integracoes_revsign.router)


@app.on_event("startup")
def aplicar_migracoes_no_boot() -> None:
    """Cria as tabelas legadas e aplica as migrações pendentes. Banco fora do ar
    não impede o boot (o erro fica no log e o /api/health/db reporta 503)."""
    try:
        garantir_baseline()
        aplicar_pendentes()
    except Exception:
        logger.exception("Falha ao aplicar as migrações do banco no boot.")


@app.on_event("startup")
async def iniciar_agendador() -> None:
    """Registra os jobs e liga o agendador (só com SCHEDULER_ENABLED). Nunca impede o boot."""
    try:
        jobs.registrar_jobs()
        scheduler.iniciar()
    except Exception:
        logger.exception("Falha ao iniciar o agendador de tarefas.")


@app.on_event("shutdown")
async def parar_agendador() -> None:
    try:
        await scheduler.parar()
    except Exception:
        logger.exception("Falha ao parar o agendador de tarefas.")


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
