import boto3, io
from managers.config import config as c
from managers.log_manager import get_date


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
        print("| S3 | Falha enviando print para o bucket")
