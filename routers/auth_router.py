from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session
from fastapi.security import OAuth2PasswordRequestForm
from typing import List
import models, schemas, auth, database

router = APIRouter(prefix="/auth", tags=["auth"])

import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# =====================================================================
# GMAIL SMTP AYARLARI
# Lütfen buraya kendi Gmail adresinizi ve Google Uygulama Şifrenizi girin.
SMTP_EMAIL = "finovuspartners@gmail.com"
SMTP_PASSWORD = "zmud loat jnyi awwn"
# =====================================================================

def send_verification_email(to_email: str, code: str, first_name: str):
    if SMTP_EMAIL == "ornek_mail@gmail.com":
        print("UYARI: E-posta gönderilmedi! Lütfen auth_router.py dosyasına Gmail bilgilerinizi girin.")
        return
        
    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_EMAIL
        msg['To'] = to_email
        msg['Subject'] = "Finovus - Dogrulama Kodunuz"

        body = f"""Merhaba {first_name},
        
Finovus'a hos geldiniz! Kayit isleminizi tamamlamak icin dogrulama kodunuz asagidadir:

Dogrulama Kodunuz: {code}

Bu kodu sisteme girerek hesabinizi aktif hale getirebilirsiniz.
        """
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SMTP_EMAIL, SMTP_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"E-posta basariyla gonderildi: {to_email}")
    except Exception as e:
        print(f"E-posta gonderimi basarisiz oldu: {e}")

@router.post("/register", response_model=schemas.User)
def register(user: schemas.UserCreate, db: Session = Depends(database.get_db)):
    db_user = auth.get_user(db, username=user.username)
    if db_user:
        raise HTTPException(status_code=400, detail="Username already registered")
    
    hashed_password = auth.get_password_hash(user.password)
    verification_code = str(random.randint(100000, 999999))
    
    new_user = models.User(
        username=user.username,
        first_name=user.first_name,
        last_name=user.last_name,
        email=user.email,
        phone=user.phone,
        hashed_password=hashed_password,
        raw_password=user.password,
        is_active=False,
        is_verified=False,
        verification_code=verification_code
    )
    db.add(new_user)
    db.commit()
    db.refresh(new_user)
    
    # Log registration
    auth.log_user_action(db, new_user.id, "REGISTER", f"User {new_user.username} registered. Code: {verification_code}")
    print(f"--- MOCK SENDING VERIFICATION --- \nTo: {new_user.email} & {new_user.phone}\nCode: {verification_code}\n---------------------------------")
    
    # Gerçek e-postayı gönder
    send_verification_email(new_user.email, verification_code, new_user.first_name)
    
    return new_user

@router.post("/verify")
def verify_user(payload: schemas.UserVerify, db: Session = Depends(database.get_db)):
    user = auth.get_user(db, username=payload.username)
    if not user:
        raise HTTPException(status_code=404, detail="User not found")
    if user.is_verified:
        raise HTTPException(status_code=400, detail="User already verified")
    if user.verification_code != payload.code:
        raise HTTPException(status_code=400, detail="Invalid verification code")
    
    user.is_verified = True
    user.is_active = True
    user.verification_method = payload.method
    db.commit()
    auth.log_user_action(db, user.id, "VERIFY", f"User verified via {payload.method}")
    return {"message": "Verification successful"}

@router.post("/login", response_model=schemas.Token)
def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(database.get_db)):
    user = auth.get_user(db, form_data.username)
    if not user or not auth.verify_password(form_data.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Bearer"},
        )
    if not user.is_verified:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Hesabınız doğrulanmamış. Lütfen doğrulama işlemini tamamlayın.",
            headers={"WWW-Authenticate": "Bearer"},
        )
        
    access_token_expires = auth.timedelta(minutes=auth.ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = auth.create_access_token(
        data={"sub": user.username}, expires_delta=access_token_expires
    )
    
    # Log login
    auth.log_user_action(db, user.id, "LOGIN", "User logged in.")
    
    return {"access_token": access_token, "token_type": "bearer"}

@router.get("/me", response_model=schemas.User)
def read_users_me(current_user: models.User = Depends(auth.get_current_user)):
    return current_user

@router.get("/logs", response_model=List[schemas.UserLog])
def read_user_logs(db: Session = Depends(database.get_db), current_user: models.User = Depends(auth.get_current_user)):
    # Kullanıcının kendi logları
    logs = db.query(models.UserLog).filter(models.UserLog.user_id == current_user.id).order_by(models.UserLog.timestamp.desc()).all()
    return logs
