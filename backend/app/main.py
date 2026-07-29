from fastapi import FastAPI
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
