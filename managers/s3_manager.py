import boto3, io, logging
from managers.config import config as c
from managers.log_manager import get_date

logging.basicConfig(level=logging.INFO, format='%(message)s')


def send_to_s3(img_bytes: str, file_name: str):
    ''' Envia a print da tela para o S3 '''

    try:

    # ------- Criando objeto do cliente S3 -------            
        s3_client = boto3.client(
            's3',
            aws_access_key_id = c.a_key_id,
            aws_secret_access_key = c.a_secret
        )

    
    # ------- Convertendo screenshot -------
        screenshot_stream = io.BytesIO(img_bytes) # Convertendo para stream


    # ------- Upload da imagem -------  
        date = get_date()
        file_name = f"{date.strftime('%d-%m-%Y')}_{file_name}"
        s3_client.upload_fileobj(screenshot_stream, c.a_bucket, file_name)
    
    
    except:
        logging.info("| S3 | Falha enviando print para o bucket")


def get_from_s3():
    ''' Pega o JSON do S3 '''

    try:

    # ------- Criando objeto do cliente S3 -------            
        s3_client = boto3.client('s3')  

        file_key = c.g_excel.split('/')
        local_file_path = f'./input/{file_key[-1]}'
        file_key = f'{file_key[-2]}/{file_key[-1]}'

        logging.info(f"Bucket: {c.a_bucket}")
        logging.info(f"Key: {file_key}")
        logging.info(f"Local: {local_file_path}")

        try:
            # Download the file from S3 to local path
            s3_client.download_file(c.a_bucket, file_key, local_file_path)
            print(f"File downloaded successfully: {local_file_path}")
            return local_file_path
        
        except Exception as e:
            print(f"Error downloading file: {e}") 
    
    except:
        logging.info("| S3 | Falha pegando o json do bucket")