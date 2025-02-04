import pytz
from datetime import datetime as dt
from managers.config import config as c
from managers.model_manager import RPATable, DashTable
from managers.integration_manager import open_postgres_conn


def get_date():
    ''' Função que retorna a data e hora atual no fuso horário de São Paulo no formato dd/mm/yyyy hh:mm'''

    date = dt.now(pytz.timezone('America/Sao_Paulo'))    
    return date


def start_log():
    ''' Alimenta as tabelas de log no banco de dados '''
    
    data = {
        "nome_do_processo": c.l_process,
        "status": None,
        "task": None, 
        "data_inicio": get_date(), 
        "data_fim": None,
        "observacoes": None
    }

    rpa_data = RPATable(**data)
    dash_data = DashTable(**data)
        
    with open_postgres_conn('LOG Inicial') as session:
        try:
            session.add(rpa_data)
            session.add(dash_data)
            session.commit()
            
            session.refresh(rpa_data)
            session.refresh(dash_data)
            
            return rpa_data.id, dash_data.id
            # return rpa_data.id, None
                
        except Exception as e:
            print(f'| START LOG | Iniciando Log no banco: {e}')
            return False


def update_log_entry(entry_id, updated_data, class_type):
    ''' Atualiza o log da execução no banco '''    
    
    with open_postgres_conn('LOG Final') as session:
        try:
            entry = session.query(class_type).filter_by(id=entry_id).one_or_none()
            
            if entry is None:
                print(f'| UPDATE LOG | Não identificado o registro com ID: {entry_id}')
                return False
            
            for key, value in updated_data.items():
                setattr(entry, key, value)
            
            session.commit()
            return True
                
        except:
            print(f'| UPDATE LOG | Falha em atualizar o registro da execução: {entry_id}')
            return False
        

def make_error_obj(error):
    ''' Retorna o objeto de erro '''     

    return {
        "nome_do_processo": c.l_process,
        "status": 500,
        "data_fim": get_date(),
        "observacoes": error
    }


def make_success_obj():
    ''' Retorna o objeto de sucesso '''     

    return {
        "nome_do_processo": c.l_process,
        "status": 200,
        "data_fim": get_date(),
        "observacoes": 'Robô executado com sucesso'
    }
