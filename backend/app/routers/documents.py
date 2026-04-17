from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import Response, StreamingResponse
from datetime import datetime
import io

from app.db.inventory_manager_db import InventoryDBManager
from app.dependencies import gestor_or_tecnico, CurrentUser
from app.core.storage import storage

router = APIRouter(prefix="/api/documents", tags=["documents"])
inv = InventoryDBManager()


@router.post("/loan-term/{item_id}")
def generate_loan_term(item_id: int, current_user: CurrentUser = Depends(gestor_or_tecnico)):
    """Gera o termo de responsabilidade e retorna o DOCX para download."""
    ok, result, filename = inv.generate_loan_term_bytes(item_id)
    if not ok:
        raise HTTPException(status_code=400, detail=result)

    # Salva no storage e também retorna o arquivo direto
    storage.save("terms", filename, result)

    return Response(
        content=result,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/files/{category}/{filename}")
def download_file(
    category: str,
    filename: str,
    _: CurrentUser = Depends(gestor_or_tecnico),
):
    """Serve um arquivo armazenado (termo, PDF assinado, nota de remoção)."""
    # Validação de categoria para evitar path traversal
    allowed_categories = {
        "terms", "termos_assinados", "termos_devolucao",
        "termos_devolucao_assinados", "notas_remocao",
    }
    if category not in allowed_categories:
        raise HTTPException(status_code=400, detail="Categoria inválida.")
    if ".." in filename or "/" in filename:
        raise HTTPException(status_code=400, detail="Nome de arquivo inválido.")

    key = f"{category}/{filename}"
    if not storage.exists(key):
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")

    content = storage.read(key)
    # Determina o media type pelo sufixo
    if filename.endswith(".pdf"):
        media_type = "application/pdf"
    elif filename.endswith(".docx"):
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        media_type = "application/octet-stream"

    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
