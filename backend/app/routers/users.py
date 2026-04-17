from fastapi import APIRouter, HTTPException, Depends

from app.schemas.users import UserCreate, UserResponse, PasswordUpdate
from app.db.user_manager_db import UserDBManager
from app.dependencies import gestor_only, CurrentUser

router = APIRouter(prefix="/api/users", tags=["users"])
user_db = UserDBManager()


@router.get("", response_model=list[UserResponse])
def list_users(_: CurrentUser = Depends(gestor_only)):
    return user_db.get_all_users()


@router.post("", response_model=dict)
def create_user(body: UserCreate, _: CurrentUser = Depends(gestor_only)):
    ok, msg = user_db.add_user(body.username, body.password, body.role)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.delete("/{user_id}")
def remove_user(user_id: int, _: CurrentUser = Depends(gestor_only)):
    ok, msg = user_db.remove_user(user_id)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}


@router.put("/{user_id}/password")
def update_password(user_id: int, body: PasswordUpdate, _: CurrentUser = Depends(gestor_only)):
    ok, msg = user_db.update_password(user_id, body.new_password)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"detail": msg}
