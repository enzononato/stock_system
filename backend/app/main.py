from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.core.config import settings
from app.routers import auth, users, items, peripherals, loans, documents, history, reports, constants

app = FastAPI(
    title="Controle de Estoque Revalle",
    version="2.0.0",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
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
    return {"status": "ok"}
