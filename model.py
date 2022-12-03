from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column
from sqlalchemy.types import Integer, Text, String,DateTime, Boolean


Base = declarative_base()

class Testing(Base):
    """testing db"""
    __tablename__ = "Testing"
    id = Column(Integer, primary_key=True, autoincrement="auto")
    empID = Column(String(5), nullable=False)
    empName = Column(Text, nullable=False)
    typeOfTest = Column(String(10), nullable=False)
    timeTested = Column(DateTime, nullable = False)
    result = Column(String(2))
    symptom = Column(Boolean)

class VisitorTesting(Base):
    __tablename__ = "visitorTesting"
    id = Column(Integer, primary_key=True, autoincrement="auto")
    visitorName = Column(String, nullable=False)
    timeTested = Column(DateTime, nullable = False)
    typeOfTest = Column(String(10), nullable=False)
    result = Column(String(2))
    symptom = Column(Boolean)