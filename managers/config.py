from os import environ

class Config():    
    p_user = 'app_robos'
    p_pwd = r'9"w/y\7Lw&14'
    p_host = 'pdaw-autoworker01.cydh6wuum0lk.us-east-1.rds.amazonaws.com'
    p_db = 'postgres'
    p_port = '5432'

    p_str = f'postgresql://{p_user}:{p_pwd}@{p_host}:{p_port}/{p_db}'
    
    a_bucket = 'rpaaw24-processamentodohbkmega'
    a_key_id = ''#environ['AWS_ACCESS_KEY_ID']
    a_secret = ''#environ['AWS_SECRET_ACCESS_KEY']


    c_user = 'rpa.crb'#environ['USERNAME_MEGA_CRE_RPA']
    c_pwd = 'Di*Ffsp7wPA_6N6XmqxM'#environ['PASSWORD_MEGA_CRE_RPA']
    c_url = 'https://mega.bildvitta.com.br/'

    t_url = 'https://prod-115.westus.logic.azure.com:443/workflows/af36ceb7a95641d6ad962d5368f2ba1c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=h1DYbqGw-EmcAJv_bq1tDb3l3DqYMlw_fO9JPkuo608'
    t_area = 'cre'
    t_rpa = 'RPAAW24-PROCESSAMENTO DO HBK (MEGA)'

    l_process = "RPAAW24-PROCESSAMENTO DO HBK (MEGA)"
    l_table = 'rpaaw24processamentodohbkmega'

config = Config()