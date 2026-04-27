from sqlalchemy import Boolean, Column, Integer, String, DateTime, ForeignKey
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    is_active = Column(Boolean, default=True)
    
    logs = relationship("UserLog", back_populates="user")

class UserLog(Base):
    __tablename__ = "user_logs"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"))
    action = Column(String) # e.g. "LOGIN", "UPLOAD_EXCEL", "REGISTER"
    timestamp = Column(DateTime, default=datetime.utcnow)
    details = Column(String, nullable=True)
    
    user = relationship("User", back_populates="logs")
