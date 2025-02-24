import requests, os, pandas as pd
from datetime import datetime as dt
from managers.config import config as c
from managers.log_manager import log, log_erro

COLUMNS = ["Codigo", "Nome", "Status"]


def send_file():
    ''' Envia o arquivo para o Sharepoint '''
    
    log(f'----> Enviando excel no SharePoint:')
    
    headers = {'Authorization': f'Bearer {c.s_token}'}
    url = c.s_url
    form_data = {
        'fileName': c.folder_path_o.replace(f'{c.folder_name_o}/', ''),
        'folderName': c.s_rpa,
        'folderNameRPA': c.s_folder,
        'executionOutputFolder': dt.now().strftime("%d-%m-%y")
    }  

    with open(c.folder_path_o, 'rb') as f:
        try:
            files = {'file': f}
            req = requests.post(url, headers=headers, data=form_data, files=files)
            
            log(f'--> Status Code: {req.status_code}')
            log(f'--> Content: {req.content}')

        except Exception as e:
            log_erro(f'Falha ao enviar o arquivo para o Sharepoint. Erro: {e}')


def clear_folder():
    ''' Limpa a pasta de arquivos '''

    log(f'----> Limpando pasta: {c.folder_path_o}')
    
    if os.path.exists(c.folder_path_o) and os.path.isdir(c.folder_path_o):        
        for item in os.listdir(c.folder_path_o):
            item_path = os.path.join(c.folder_path_o, item)

            if os.path.isfile(item_path):
                os.remove(item_path)


def create_folder():
    ''' Cria a pasta onde o arquivo de output vai ser inserido '''

    log('----> Criando palsta de output')
    if not os.path.exists(c.folder_name_o):
        os.makedirs(c.folder_name_o)

    c.folder_path_o = f'{c.folder_name_o}/{c.g_file} - {c.zeev_ticket}.xlsx'


def initialize_sheet():
    ''' Inicializa ou sobrescreve uma planilha com as colunas fornecidas '''    
    
    log('----> Inicilização da planilha de retorno')
    
    df = pd.DataFrame(columns=COLUMNS) # Cria um DataFrame vazio com as colunas especificadas
    df.to_excel(c.folder_path_o, index=False)
    
    log(f"---->Planilha inicializada em: {c.folder_path_o}")