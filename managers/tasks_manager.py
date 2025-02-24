from managers.log_manager import log, log_erro
from managers.model_manager import RPATable
from managers.integration_manager import open_postgres_conn


def fetch_pending_tasks():
    ''' Busca tarefas pendentes no banco (status = 0) '''
    
    with open_postgres_conn('Buscando tasks no banco') as session:
        try:
            log('----> Procurando no banco tarefas para processar com status 0')
            tasks = (
                session.query(RPATable)
                .filter(RPATable.status == 0)                
                .all()
            )
            
            task_dicts = [task_to_dict(task) for task in tasks]
            return task_dicts
        
        except Exception as e:
            log_erro(f'Falha em trazer tarefas com status 0. Erro: {e}')
            return []
        

def task_to_dict(task):
    ''' Converte a instancia do RPATable em dicionario.'''

    return {
        'id': task.id,
        'nome_do_processo': task.nome_do_processo,
        'status': task.status,
        'task': task.task,
        'data_inicio': task.data_inicio,
        'data_fim': task.data_fim,
        'observacoes': task.observacoes,
        'id_execucao': task.id_execucao
    }        
