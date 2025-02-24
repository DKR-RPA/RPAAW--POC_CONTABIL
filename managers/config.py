from os import environ

class Config():    
    p_user = 'app_robos'
    p_pwd = r'9"w/y\7Lw&14'
    p_host = 'pdaw-autoworker01.cydh6wuum0lk.us-east-1.rds.amazonaws.com'
    p_db = 'postgres'
    p_port = '5432'

    p_str = f'postgresql://{p_user}:{p_pwd}@{p_host}:{p_port}/{p_db}'
    
    a_bucket = 'rpaaw34-poccontabil'
    a_key_id = ''#environ['AWS_ACCESS_KEY_ID']
    a_secret = ''#environ['AWS_SECRET_ACCESS_KEY']


    c_user = 'rpa.crb'#environ['USERNAME_MEGA_CRE_RPA']
    c_pwd = 'Di*Ffsp7wPA_6N6XmqxM'#environ['PASSWORD_MEGA_CRE_RPA']
    c_url = 'https://mega.bildvitta.com.br/'

    t_url = 'https://prod-29.westus.logic.azure.com:443/workflows/6735d8fca19c45cda97250d6c0c0c1a3/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=ZeCLu6RKb3K6eKRIuOfIzAZ5X3rZT_ovebmnKMbp_Mo'
    t_area = 'cont'
    
    s_url = 'https://x2drx2exzh.execute-api.us-east-1.amazonaws.com/api/v1/sharepoint/file'
    s_token = '6733187b-4651-5635-8cf0-c19f7f8a7d57'
    s_rpa = 'CONT'
    s_folder = 'RPAAW34-POC CONTABIL'

    l_process = 'RPAAW34-POC CONTABIL'
    l_table = 'rpaaw34poccontabil'

    g_file = 'OUTPUT_CHAMADO_REQ'
    g_ticket = ''#environ['ZEEV_TICKET']
    g_id_exec = ''#environ['ID_EXECUCAO']
    
    rpa_data_id = ''
    dash_data_id = ''
    folder_name_o = 'output_folder'
    folder_path_o = ''
    zeev_ticket = ''

config = Config()