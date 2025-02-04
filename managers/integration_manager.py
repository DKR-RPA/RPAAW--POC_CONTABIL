from sqlalchemy import create_engine
from contextlib import contextmanager
from sqlalchemy.orm import sessionmaker
from managers.config import config as c


@contextmanager
def open_postgres_conn(info:str):
    ''' Abre uma conexão com o Postgres e retorna o cursor e após a utilização, fecha a conexão '''
    
    try:
        engine = create_engine(c.p_str)    
        Session = sessionmaker(bind=engine)        
        session = Session()

        yield session
            
    except Exception as e:
        print(f'| POSTGRES | Abrindo conexão com o Postgres para ação: {info} | Erro: {e}')
        session.rollback()
        
    finally:
       session.close()

