import requests
from managers.config import config as c
from managers.log_manager import  log, log_erro


def send_card(len_tasks, len_errors):
    ''' Monta e envia um card para o teams | Estrutura do infos '''
   
    log(f'----> Enviando card no Teams')
   
    payload = {
        'canal': c.t_area,
        'quantidadeEmpresa': len_tasks,
        'quantidadeErro': len_errors
    }   

    try:
        req = requests.post(c.t_url, json=payload)
        log(f'--> Status Code: {req.status_code}')
        log(f'--> Content: {req.content}')
        
    
    except Exception as e:
        log_erro(f'Falha ao enviar card no Teams. Erro: {e}')
