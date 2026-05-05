from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from typing import List
import models, schemas, auth, database

router = APIRouter(prefix="/admin", tags=["admin"])

def check_admin(current_user: models.User = Depends(auth.get_current_user)):
    # Yalnızca admin kullanıcısı erişebilir
    if current_user.username != "admin":
        raise HTTPException(status_code=403, detail="Not authorized")
    return current_user

@router.get("/users", response_model=List[schemas.User])
def get_all_users(db: Session = Depends(database.get_db), admin: models.User = Depends(check_admin)):
    users = db.query(models.User).all()
    return users

@router.put("/users/{user_id}/password")
def update_user_password(user_id: int, payload: schemas.PasswordUpdate, db: Session = Depends(database.get_db), admin: models.User = Depends(check_admin)):
    user = db.query(models.User).filter(models.User.id == user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    
    user.hashed_password = auth.get_password_hash(payload.new_password)
    db.commit()
    auth.log_user_action(db, admin.id, "ADMIN_ACTION", f"Password changed for user {user.username}")
    return {"message": "Password updated successfully"}

@router.get("/logs")
def get_all_logs(db: Session = Depends(database.get_db), admin: models.User = Depends(check_admin)):
    logs = db.query(models.UserLog, models.User).join(models.User).order_by(models.UserLog.timestamp.desc()).limit(200).all()
    
    result = []
    for log, user in logs:
        result.append({
            "id": log.id,
            "username": user.username,
            "action": log.action,
            "timestamp": log.timestamp,
            "details": log.details
        })
    return result
