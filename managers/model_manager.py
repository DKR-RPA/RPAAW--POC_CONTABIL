from sqlalchemy import Column, Integer, Text, String, TIMESTAMP
from managers.config import config as c
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.dialects.postgresql import JSONB

Base = declarative_base()

class RPATable(Base):
    __tablename__ = c.l_table
    __table_args__ = {'schema': 'public'} 

    id = Column(Integer, primary_key=True)  
    nome_do_processo = Column(Text)
    status = Column(Integer) 
    task = Column(JSONB)
    data_inicio = Column(TIMESTAMP)
    data_fim = Column(TIMESTAMP)
    observacoes = Column(String(200)) 


class DashTable(Base):
    __tablename__ = 'dashboardmonitoramento'
    __table_args__ = {'schema': 'public'} 

    id = Column(Integer, primary_key=True)  
    nome_do_processo = Column(Text)
    status = Column(Integer) 
    task = Column(JSONB)
    data_inicio = Column(TIMESTAMP)
    data_fim = Column(TIMESTAMP)
    observacoes = Column(String(200))     
