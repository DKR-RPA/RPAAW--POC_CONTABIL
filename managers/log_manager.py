import pytz, logging
from datetime import datetime as dt
from managers.config import config as c
from sqlalchemy.inspection import inspect
from managers.model_manager import RPATable, DashTable
from managers.integration_manager import open_postgres_conn

logging.basicConfig(level=logging.INFO, format='%(message)s')


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

    # Informação extra de controle da Bild
    extra_key_value = {"id_execucao": c.g_id_exec}  

    rpa_data = RPATable(**{**data, **extra_key_value})
    dash_data = DashTable(**data)
        
    with open_postgres_conn('LOG Inicial') as session:
        try:
            session.add(rpa_data)
            session.add(dash_data)
            session.commit()
            
            session.refresh(rpa_data)
            session.refresh(dash_data)
            
            return rpa_data.id, dash_data.id
                
        except :
            logging.info(f'| START LOG | Iniciando Log no banco')
            return False


def update_log_entry(entry_id, updated_data, class_type, dash_table_model=False):
    ''' Atualiza o log da execução no banco '''    
    
    with open_postgres_conn('Update Log') as session:
        try:
            entry = session.query(class_type).filter_by(id=entry_id).one_or_none()
            
            if entry is None:
                logging.info(f'| UPDATE LOG | Não identificado o registro com ID: {entry_id}')
                return False
            
            for key, value in updated_data.items():
                setattr(entry, key, value)

            # Pegando todas as informaçoes da linha da task que foi atualizada
            if dash_table_model:
                mapper = inspect(class_type)
                
                data_for_new_entry = {
                    column.key: getattr(entry, column.key) 
                    for column in mapper.attrs.values()  # Iterando entre as colunas mapeadas
                    if column.key not in ['id', 'id_execucao', 'ticket']  # Excluindo 'id', 'ticket' e 'id_execucao'
                }    

                new_entry = dash_table_model(**data_for_new_entry) # Populando o objeto com as novas informaçoes
                session.add(new_entry)  # Adicionando a nova linha na tabela
            
            session.commit()
            return True
                
        except:
            logging.info(f'| UPDATE LOG | Falha em atualizar o registro da execução: {entry_id}')
            return False
        
        
def make_error_obj(error):
    ''' Retorna o objeto de erro '''     

    return {
        # "id_execucao": c.g_id_exec,
        "status": 500,
        "data_fim": get_date(),
        "observacoes": error
    }


def make_success_obj():
    ''' Retorna o objeto do log com sucesso '''     

    return {
        "status": 200,
        "data_fim": get_date(),
        "observacoes": 'Robô executado com sucesso'
    }


def log(msg: str, step: str=None):
    ''' Faz o registro de um log '''

    if not step:
        logging.info(f'| {c.g_id_exec} |{msg}')

    else:
        logging.info(f'| {c.g_id_exec} |{step}|{msg}')


def log_erro(msg: str, step: str=None):
    ''' Faz o registro de um log de erro '''

    if not step:
        logging.info(f'[ERRO] | {c.g_id_exec} | {msg}')
    
    else:
        logging.info(f'[ERRO] | {c.g_id_exec} |{step}| {msg}')
   
    
def finish_log(error:bool=False):
    if error:
        info_data = make_error_obj('Erro: Falha durante execução do script')
        
    else:
        info_data = make_success_obj()        
        
    update_log_entry(c.rpa_data_id, info_data, RPATable)
    update_log_entry(c.dash_data_id, info_data, DashTable)      