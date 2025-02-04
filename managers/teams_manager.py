import requests
from managers.config import config as c


def send_card(content):
    ''' Monta e envia um card para o teams | Estrutura do infos '''
   
    payload = {
        'area': c.t_area,
        'nomeRPA': c.t_rpa,
        'retornoRPA': content
    }   

    try:
        requests.post(c.t_url, json=payload)
    
    except:
        print('[E] Falha ao enviar card no Teams')
