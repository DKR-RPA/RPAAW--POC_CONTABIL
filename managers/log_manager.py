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
        "nome_do_processo": c.s_process,
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
                
        except:
            logging.info(f'| START LOG | Iniciando Log no banco')
            return False


def start_queue_task(task, ticket):
    ''' Cria os registros no banco com a data de início e task '''

    data = {
        "nome_do_processo": c.s_process,
        "status": None,
        "task": task, 
        "data_inicio": get_date(), 
        "data_fim": None,
        "observacoes": None,
        "id_execucao": c.g_id_exec, 
        "ticket": ticket
    }

    rpa_data = RPATable(**data)
        
    with open_postgres_conn('Criando linha da task') as session:
        try:
            session.add(rpa_data)
            session.commit()
            
            session.refresh(rpa_data)
            
            return rpa_data.id
                
        except:
            logging.info(f"Falha ao criar a task com un Requisicao: {task['Un Requisição']}")
            return False


def update_log_entry(entry_id, updated_data, class_type, dash_table_model=False):
    ''' Atualiza o log da execução no banco '''    
    
    with open_postgres_conn('LOG Final') as session:
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

def make_success_obj_task(ordem):
    ''' Retorna o objeto da task com sucesso '''     

    return {
        "id_execucao": c.g_id_exec,
        "status": 200,
        "data_fim": get_date(),
        "observacoes": f'Robô executado com sucesso: {ordem}'
    }