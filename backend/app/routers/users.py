from fastapi import APIRouter, HTTPException, Depends

from app.schemas.users import UserCreate, UserResponse, PasswordUpdate
from app.db.user_manager_db import UserDBManager
from app.dependencies import gestor_only, get_user_db, CurrentUser

router = APIRouter(prefix="/api/users", tags=["users"])


@router.get("", response_model=list[UserResponse])
def list_users(_: CurrentUser = Depends(gestor_only), user_db: UserDBManager = Depends(get_user_db)):
    return user_db.get_all_users()


@router.post("", response_model=dict)
def create_user(
    body: UserCreate,
    _: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
):
    ok, msg = user_db.add_user(body.username, body.password, body.role)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.delete("/{user_id}")
def remove_user(
    user_id: int,
    current_user: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
):
    # T5: um Gestor não pode remover a si mesmo, nem remover o último
    # Gestor do sistema (o que causaria lockout — ninguém mais poderia
    # administrar usuários).
    if user_id == current_user.id:
        raise HTTPException(status_code=400, detail="Você não pode remover seu próprio usuário.")

    target = user_db.get_user_by_id(user_id)
    if not target:
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")

    if target.get("role") == "Gestor" and user_db.count_gestores() <= 1:
        raise HTTPException(status_code=400, detail="Não é possível remover o último Gestor do sistema.")

    ok, msg = user_db.remove_user(user_id)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.put("/{user_id}/password")
def update_password(
    user_id: int,
    body: PasswordUpdate,
    _: CurrentUser = Depends(gestor_only),
    user_db: UserDBManager = Depends(get_user_db),
):
    # 404 para usuário inexistente, coerente com remove_user e com o resto da API.
    # Antes, a troca de senha de um id inexistente respondia 200 "Senha alterada
    # com sucesso" sem ter alterado nada.
    if not user_db.get_user_by_id(user_id):
        raise HTTPException(status_code=404, detail="Usuário não encontrado.")

    ok, msg = user_db.update_password(user_id, body.new_password)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


# Observação: não existe hoje um endpoint de alteração de role de usuário
# (apenas criação, remoção e troca de senha). Caso um seja adicionado no
# futuro, aplique a mesma proteção usada em remove_user: um Gestor não pode
# rebaixar a si mesmo nem rebaixar o último Gestor do sistema (mesmo risco
# de lockout administrativo).
