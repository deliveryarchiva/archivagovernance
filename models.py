from sqlalchemy import Column, Integer, String, Float, JSON, DateTime, Text
from sqlalchemy.sql import func
from database import Base


class ProjectSizing(Base):
    __tablename__ = "project_sizings"

    id              = Column(Integer, primary_key=True, index=True)
    crm_number      = Column(String, nullable=True)
    odl             = Column(String, nullable=True)
    cliente         = Column(String, nullable=False)
    titolo          = Column(String, nullable=False)
    answers         = Column(JSON, nullable=False)   # {"q1": "3", "q2": "1", ...}
    score           = Column(Float, nullable=False)
    taglia          = Column(String, nullable=False)  # xs | s | m | l | xl | xxl
    note            = Column(Text, nullable=True)
    created_at      = Column(DateTime(timezone=True), server_default=func.now())
    created_by      = Column(String, nullable=False)
    created_by_nome = Column(String, nullable=False)
