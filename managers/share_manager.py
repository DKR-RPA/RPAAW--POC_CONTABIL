import requests, os, logging
from datetime import datetime as dt
from managers.config import config as c

logging.basicConfig(level=logging.INFO, format='%(message)s')


def send_file(excel_file, file_name):
    ''' Envia o arquivo para o Sharepoint '''

    try:
        headers = {'Authorization': f'Bearer {c.s_token}'}
        url = c.s_url
        form_data = {
            'fileName': file_name,
            'folderName': c.s_rpa,
            'folderNameRPA': c.s_folder,
            'executionOutputFolder': dt.now().strftime("%d-%m-%y")
        }  

        with open(excel_file, 'rb') as f:
            files = {'file': f}
            r = requests.post(url, headers=headers, data=form_data, files=files)

        logging.info(f"Status: {r.status_code}")
        logging.info(f"Content: {r.content}")

    except Exception as e:
        logging.error(f'[E] Falha ao enviar o arquivo para o Sharepoint: {e}')


def clear_folder(folder_path):
    ''' Limpa a pasta de arquivos '''

    logging.error(f'Limpando pasta {folder_path}')
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)

            if os.path.isfile(item_path):
                os.remove(item_path)
