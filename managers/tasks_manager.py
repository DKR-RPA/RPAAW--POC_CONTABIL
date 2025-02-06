import logging, json
from managers.config import config as c
from managers.s3_manager import get_from_s3

logging.basicConfig(level=logging.INFO, format='%(message)s')


def fetch_pending_tasks():
    ''' Converte as informações da variavel de ambiente EXCEL_JSON para uma lista de dicionários '''

    try:
        local_file_path = get_from_s3()
        if not local_file_path:
            return []

        with open(local_file_path, 'r') as json_file:
            return json.load(json_file)

    except Exception as e:
        logging.info(f'Falha na leitura do JSON pelo EXCEL_JSON: {e}')
        return []   

