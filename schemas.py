from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime

class UserBase(BaseModel):
    username: str
    first_name: Optional[str] = None
    last_name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None

class UserCreate(UserBase):
    first_name: str
    last_name: str
    email: str
    phone: str
    password: str

class User(UserBase):
    id: int
    is_active: bool
    is_verified: bool
    verification_method: Optional[str] = None
    raw_password: Optional[str] = None

    class Config:
        from_attributes = True

class UserVerify(BaseModel):
    username: str
    code: str
    method: str

class Token(BaseModel):
    access_token: str
    token_type: str

class TokenData(BaseModel):
    username: Optional[str] = None

class UserLog(BaseModel):
    id: int
    user_id: int
    action: str
    timestamp: datetime
    details: Optional[str] = None

    class Config:
        from_attributes = True

class PasswordUpdate(BaseModel):
    new_password: str

class AdminUserLog(BaseModel):
    id: int
    username: str
    action: str
    timestamp: datetime
    details: Optional[str] = None

    class Config:
        from_attributes = True
