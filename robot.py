import pyperclip, time, cv2, numpy as np, subprocess, io, logging
from PIL import Image
from managers.config import config as c
from managers.s3_manager import send_to_s3
from managers.log_manager import start_log, make_error_obj, update_log_entry, make_success_obj
from managers.model_manager import RPATable, DashTable
from managers.teams_manager import send_card
from managers.selenium_manager import close_leftover_webdriver_instances, start_selenium
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

import os
import shutil
from datetime import datetime, timedelta
import pandas as pd
import calendar

TEMP_METHOD = eval('cv2.TM_CCOEFF_NORMED')
CONFIDENCE = 0.8
TIMEOUT = 720
INTERVAL = 5
STR_CONCLUIDO = 'Processamento concluído'

# Variaveis de ambiente
base_path = "./Contabilidade/"
base_dir = "./Contabilidade/"

class ImageNotFoundException(Exception):
    """Exceção personalizada para quando a imagem não é encontrada dentro do timeout."""
    pass

logging.basicConfig(level=logging.INFO, format='%(message)s')


def login_portal_entrada(navigate):
    ''' Realiza o login no portal de entrada '''

    # -- Acesso portal de entrada --
    navigate.get(c.c_url)
    save_print(navigate, '01_Acesso_portal_entrada.png')
                          

    # -- Usuário --   
    logging.info('|--> Inserindo info do usuário')        
    user_input = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Editbox1"]')))
    user_input.clear()
    user_input.send_keys(c.c_user) 
    save_print(navigate, '01_1_Acesso_portal_User.png')
    time.sleep(3)
    
    
    # -- Senha --
    logging.info('|--> Inserindo info da senha')  
    password_input = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Editbox2"]')))
    password_input.clear()
    password_input.send_keys(c.c_pwd) 
    save_print(navigate, '01_2_Acesso_portal_Senha.png')
    time.sleep(3)


    # -- Botão de login --
    logging.info('|--> Click botão login')  
    btn_login = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="buttonLogOn"]')))
    btn_login.click()
    time.sleep(8) 

        
    # -- Botão Sistema Mega --
    navigate.switch_to.window(navigate.window_handles[-1]) # Indo para a proxima aba


def selecionar_sistema_mega(navigate):
    ''' Seleciona o sistema Mega e aguarda o carregamento da tela de login '''

    # ---- Menu do servidor ----
    logging.info('|--> Aguardando menu do servidor') 

    mega_button_img = "png_elements/13_button_mega.png"
    coords = find_element(mega_button_img, navigate, 'Login Portal de Entrada', f5=True)
    
    if coords:
        time.sleep(3)
    
    else:
        save_print(navigate, '03_ERRO_Acessando_servidor.png')
        raise ImageNotFoundException(f"Imagem '{mega_button_img}' não encontrada após {TIMEOUT} segundos.")

    # -- Realizando a ação de clicar no botão com o selenium --
    click(coords[0] + 32, coords[1] + 33)
    time.sleep(5)
    save_print(navigate, '03_Acesso_servidor.png')    


    # ---- Menu do Mega ----
    logging.info('|--> Aguardando menu do mega') 
    mega_logo_img = "png_elements/14_tela_login_mega.png"
    
    if not find_element(mega_logo_img, navigate, 'Login Portal de Entrada'): # Validação final
        save_print(navigate, '03_ERRO_Acessando_mega.png')
        raise ImageNotFoundException(f"Imagem '{mega_logo_img}' não encontrada após {TIMEOUT} segundos.")

    return True

                    
def login_portal_mega(navigate):
    ''' Realiza o login no portal do mega '''

    # ---- Realizar Login ----

    # -- Usuário --
    logging.info('|--> Inserindo info do usuário')
    ACTION.send_keys(c.c_user).perform()  
    # time.sleep(3) 
    # save_print(navigate, '05_Acesso_mega_User.png')

    # -- Tab --   
    logging.info('|--> Apertando TAB para senha')    
    ACTION.send_keys(Keys.TAB).perform()  
    # time.sleep(3) 
    # save_print(navigate, '05_1_Acesso_mega_Tab.png')

    # -- Senha --   
    logging.info('|--> Inserindo info da senha')  
    ACTION.send_keys(c.c_pwd).perform()
    # time.sleep(3) 
    # save_print(navigate, '05_2_Acesso_mega_Senha.png')

    # -- Enter --   
    logging.info('|--> Apertando ENTER para login')  
    ACTION.send_keys(Keys.ENTER).perform()
    # time.sleep(3) 
    # save_print(navigate, '05_3_Acesso_mega_Enter.png')


    # ---- Usuário já logado ----
    try:
        image_path = "png_elements/15_aviso_usuario_logado.png"
        if find_element(image_path, navigate, 'Buscando Popup Usuário Logado'): # Caso encontre o aviso de usuário já logado, seleciona a opção "Não" para não desconectar o usuário
            logging.info('|--> Manipulando Popup Usuário Logado')  
            save_print(navigate, '05_4_Acesso_mega_Popup.png')
            time.sleep(1)
            
            # -- Tab --
            logging.info('|--> Apertando TAB para não')  
            ACTION.send_keys(Keys.TAB).perform()
            time.sleep(2) 
            save_print(navigate, '05_4_1_Acesso_mega_Popup_Tab.png')

            # -- Tab --
            logging.info('|--> Apertando ENTER para não')  
            ACTION.send_keys(Keys.ENTER).perform()
            time.sleep(3) 
            save_print(navigate, '05_4_2_Acesso_mega_Popup_Enter.png')
    
    except:
        pass # Significa que não apareceu o Popup


    # ---- Menu do Mega ----
    logging.info('|--> Aguardando login no mega') 
    image_path = "png_elements/01_pagina_inicial_menu_user.png"  

    if not find_element(image_path, navigate, 'Login Portal do Mega'): # Validação final
        save_print(navigate, '05_ERRO_Realizando_login_portal_mega.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")

    return True


def trocar_empresa(navigate, codigo):
    ''' Realiza as ações para a troca de empresa para iniciar o processo de processamento '''


    # ---- Imagem do perfil ----   
    logging.info('|--> Procurando menu de perfil') 
   
    image_path = 'png_elements/01_pagina_inicial_menu_user.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da imagem do perfil') 

    if not coords:
        save_print(navigate, '07_ERRO_Procurando_menu_de_perfil.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
 
    click(coords[0] + 18, coords[1] + 17)   
    save_print(navigate, '07_Acessando_menu_de_perfil.png') 

    
    # ---- Trocar Empresa ----   
    logging.info('|--> Procurando trocar empresa') 

    image_path = 'png_elements/02_pagina_inicial_menu_user_trocar_empresa.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da imagem da opção Trocar de Empresa')

    if not coords:
        save_print(navigate, '07_ERRO_Procurando_menu_trocar_empresa.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
    
    # -- Realizando a ação de clicar no botão com o selenium --      
    click(coords[0] + 77, coords[1] + 18)
    time.sleep(1)
    save_print(navigate, '07_1_Acessando_menu_trocar_empresa.png')


    # ---- Alterar para Holding ----   
    logging.info('|--> Procurando empresa selecionada') 

    image_path = 'png_elements/03_tela_holding_lupa.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da Tela da empresa selecionada')

    if coords:
        time.sleep(1)
    
    else:
        save_print(navigate, '07_ERRO_Procurando_menu_empresa_selecionada.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
    
    # -- Realizando a ação de clicar no botão com o selenium --    
    click(coords[0] + 193, coords[1] + 14)
    save_print(navigate, '07_2_Acessando_menu_empresa_selecionada.png')

    # -- TAB --    
    logging.info('|--> Apertando TAB para selecionar empresa')
    ACTION.send_keys(Keys.TAB).perform()     
    save_print(navigate, '07_3_Acessando_menu_selecionar_empresa.png')
    time.sleep(1)


    # ---- Alterar critérios | Campo = Código, Tipo = "Igual a", Valor = 1 e Clica no botão selecionar ----
    logging.info('|--> Selecionando critérios') 

    # -- TAB --
    logging.info('|--> Apertando DOWN 6x')  
    for _ in range(6):
        ACTION.send_keys(Keys.ARROW_DOWN).perform()
        time.sleep(0.2) 
    
    save_print(navigate, '07_4_1_Selecionando_criterios_down_6x.png')
    time.sleep(1)

    # -- ENTER --
    logging.info('|--> Apertando ENTER') 
    ACTION.send_keys(Keys.ENTER).perform()    
    save_print(navigate, '07_4_2_Selecionando_criterios_enter.png')
    time.sleep(1)

    # -- TAB --
    logging.info('|--> Apertando TAB') 
    ACTION.send_keys(Keys.TAB).perform()
    save_print(navigate, '07_4_3_Selecionando_criterios_tab.png')
    time.sleep(1)

    # -- DELETE --
    logging.info('|--> Apertando BACKSPACE') 
    ACTION.send_keys(Keys.BACK_SPACE).perform()
    save_print(navigate, '07_4_4_Selecionando_criterios_backspace.png')
    time.sleep(1)

    # -- DELETE --
    logging.info('|--> Apertando DOWN 2x')  
    for _ in range(2):
        ACTION.send_keys(Keys.ARROW_DOWN).perform()
        time.sleep(0.2) 
    
    save_print(navigate, '07_4_5_Selecionando_criterios_down_2x.png')
    time.sleep(1)

    # -- ENTER --
    logging.info('|--> Apertando ENTER') 
    ACTION.send_keys(Keys.ENTER).perform()    
    save_print(navigate, '07_4_6_Selecionando_criterios_enter.png')
    time.sleep(1)

    # -- TAB --
    logging.info('|--> Apertando TAB') 
    ACTION.send_keys(Keys.TAB).perform()
    save_print(navigate, '07_4_7_Selecionando_criterios_tab.png')
    time.sleep(1)

    # -- Valor --   
    logging.info('|--> Inserindo valor 1')  
    ACTION.send_keys(codigo).perform()
    time.sleep(1) 
    save_print(navigate, '07_4_8_Selecionando_valor_1.png')

    # -- TAB --
    logging.info('|--> Apertando TAB') 
    ACTION.send_keys(Keys.TAB).perform()
    save_print(navigate, '07_4_9_Selecionando_criterios_tab.png')
    time.sleep(1)

    # -- ENTER --
    logging.info('|--> Apertando ENTER') 
    ACTION.send_keys(Keys.ENTER).perform()    
    save_print(navigate, '07_4_10_Selecionando_criterios_enter.png')

    if is_image_present(image_path, navigate, 'Validando se campo Consolidador carregou com sucesso', timeout=30):
        logging.info('|--> Campo Consolidador carregado com sucesso, clicando em Selecionar.')
        save_print(navigate, '07_4_11_validacao_consolidador.png')

        # -- ALT + S --
        logging.info('|--> Apertando ALT + S') # Simular o atalho alt + S para clicar no botão selecionar
        ACTION.key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).perform()    
        save_print(navigate, '07_4_11_Selecionando_criterios_alt_s.png')
        
        return True
        
    else:
        logging.warning('|--> Campo Consolidador não carregou dentro do tempo limite.')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")


def listar_arquivos_xlsx(diretorio):
    arquivos = []
    for root, _, files in os.walk(diretorio):  # Percorre diretórios e subdiretórios
        for f in files:
            if f.lower().strip().endswith(".xlsx"):
                arquivos.append(os.path.join(root, f))  # Salva o caminho completo
    return arquivos


def controle_pastas(base_dir):
    # Diretório base
    base_dir = os.path.expanduser(base_dir)
    poc_dir = os.path.join(base_dir, "POC")
    ano_atual = datetime.now().strftime("%Y")

    meses_portugues = {
        "January": "janeiro", "February": "fevereiro", "March": "março", 
        "April": "abril", "May": "maio", "June": "junho", 
        "July": "julho", "August": "agosto", "September": "setembro", 
        "October": "outubro", "November": "novembro", "December": "dezembro"
    }

    # Obtendo o nome do mês anterior em inglês
    mes_anterior_en = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%B")

    # Convertendo para português
    mes_anterior = meses_portugues.get(mes_anterior_en, mes_anterior_en)

    print(mes_anterior)

    # Caminhos das pastas
    ano_dir = os.path.join(poc_dir, ano_atual)
    mes_anterior_dir = os.path.join(ano_dir, mes_anterior)
    lancado_dir = os.path.join(mes_anterior_dir, "Lançado")

    # Criando as pastas, se não existirem
    for dir_path in [poc_dir, ano_dir, mes_anterior_dir, lancado_dir]:
        os.makedirs(dir_path, exist_ok=True)

    # Diretórios de saída
    diretorio_arquivos = mes_anterior_dir
    diretorio_lancado = lancado_dir

    # Verifica a existência de arquivos .xlsx na pasta POC e Mês
    arquivos_xlsx = set(listar_arquivos_xlsx(poc_dir) + listar_arquivos_xlsx(mes_anterior_dir))

    if not arquivos_xlsx:
        # Caso não existam, verificar na pasta do mês atual
        mes_atual = datetime.now().strftime("%B")
        mes_atual_dir = os.path.join(ano_dir, mes_atual)
        if not os.path.exists(mes_atual_dir):
            raise Exception(f"Business Exception: Nenhum arquivo .xlsx encontrado em {poc_dir} ou {mes_anterior_dir}")
    
    # Verifica se há um arquivo que contenha "POC_Composição" no nome
    arquivo_encontrado = None
    for arquivo in arquivos_xlsx:
        if "POC_Composição" in arquivo:
            arquivo_encontrado = arquivo
            print(f"O arquivo: {arquivo_encontrado}, localizado com sucesso.")
            
            # Move o arquivo encontrado para o diretório destino
            destino = os.path.join(diretorio_arquivos, os.path.basename(arquivo_encontrado))
            shutil.move(arquivo_encontrado, destino)
            break

    if not arquivo_encontrado:
        raise Exception("Business Exception: Nenhum arquivo contendo 'POC_Composição.xlsx' foi encontrado.")

    return True, destino

def tratar_competencia(competencia):
    """
    Separa e trata o mês e o ano da variável COMPETENCIA.
    - Mês: Primeira letra maiúscula, resto minúsculo.
    - Ano: Convertido para inteiro.
    """

    # Divide a string usando '/' como separador
    partes = competencia.split('/')

    if len(partes) != 2:
        raise Exception("Business Exception: Os dados no campo Competência estão fora do formata esperado: 'Dezembro/2024'.")  # Retorna erro caso o formato esteja incorreto

    mes = partes[0].strip()  # Remove espaços extras do mês
    ano = partes[1].strip()  # Remove espaços extras do ano

    # Tratamento do mês: primeira letra maiúscula e resto minúsculo
    mes_tratado = mes.capitalize()

    # Tratamento do ano: converte para inteiro (se possível)
    try:
        ano_tratado = int(ano)
    except ValueError:
        raise Exception("Business Exception: Os dados no campo Competência fora do formata esperado: 'Dezembro/2024'.")  # Retorna erro se o ano não for um número válido

    return True, mes_tratado, ano_tratado


def processar_excel(arquivo_encontrado):
    # Lendo o arquivo Excel
    df = pd.read_excel(arquivo_encontrado, dtype=str)  # Lendo tudo como string para evitar problemas

    # Tratando os nomes das colunas (tudo maiúsculo e sem espaços)
    df.columns = [col.strip().upper().replace(" ", "_") for col in df.columns]

    # Removendo espaços extras das células e tratando INCC
    if "INCC" in df.columns:
        df["INCC"] = df["INCC"].str.strip().str.upper().map({"S": True, "N": False})

    # Tratando os demais campos removendo espaços extras nas extremidades
    for col in df.columns:
        if df[col].dtype == "object":  # Apenas para colunas de texto
            df[col] = df[col].str.strip()

    return True, df


def obter_mes_e_ultimo_dia(mes_tratado, ano_tratado):
    # Cria um dicionário de meses em português
    meses = {
    "janeiro": "01",
    "fevereiro": "02",
    "março": "03",
    "abril": "04",
    "maio": "05",
    "junho": "06",
    "julho": "07",
    "agosto": "08",
    "setembro": "09",
    "outubro": "10",
    "novembro": "11",
    "dezembro": "12"
}
    
    # Converte o mês tratado para minúsculo e encontra o número do mês
    mes_numero = meses.get(mes_tratado.lower())
    
    if mes_numero is None:
        return "Mês inválido", None, None

    # Converte mes_numero para inteiro para passar para o calendar
    mes_numero_int = int(mes_numero)
    
    # Usando calendar.monthrange para obter o último dia do mês
    _, ultimo_dia = calendar.monthrange(ano_tratado, mes_numero_int)

    # Variáveis conforme solicitado:
    # 1. inicio = "01/mes_numero/ano_tratado"
    inicio = f"01/{mes_numero}/{ano_tratado}"
    
    # 2. fim = "ultimo_dia/mes_numero/ano_tratado"
    fim = f"{ultimo_dia}/{mes_numero}/{ano_tratado}"
    
    # 3. periodo = "01-mm-aaaa"
    periodo = inicio[-7:].replace("/", "-")
    
    # 4. ultimo_periodo_fechado = data com o ultimo dia do mês
    ultimo_periodo_fechado = datetime(ano_tratado, mes_numero_int, ultimo_dia)

    ultimo_periodo_fechado = ultimo_periodo_fechado.strftime("%d-%m-%Y")
    
    return True, mes_numero, ultimo_dia, inicio, fim, periodo, ultimo_periodo_fechado


def acesso_agendar_planilhas(navigate):
    ''' Faz a seleção do agendar planilha '''
    
    # ---- Menu de Pesquisar ----
    logging.info('|--> Procurando menu de busca') 
    image_path = 'png_elements/07_pagina_inicial_menu_lupa.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da imagem da lupa') 

    if not coords:
        save_print(navigate, '08_ERRO_Procurando_menu_de_busca.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
  
    click(coords[0] + 31, coords[1] + 43)    
    save_print(navigate, '08_Acessando_menu_de_busca.png') 


    # ---- Menu de Procurar----
    logging.info('|--> Label menu procurar') 
    image_path = 'png_elements/08_01_procurar.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da imagem da lupa') 

    if coords:
         # -- Inserindo info na busca --
        logging.info('|--> Preenchendo info na busca')
        ACTION.send_keys('Agendar planilhas').perform()  
        time.sleep(1) 
        save_print(navigate, '08_1_Preenchendo_Busca.png')

    else:
        save_print(navigate, '08_ERRO_Procurando_menu_de_busca.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
  
    # ---- Selecionar opcao Agendar planilhas----
    logging.info('|--> Selecionado opcao Agendar planilhas') 
    image_path = 'png_elements/08_pagina_inicial_menu_lupa_agendar_planilhas.png'
    coords = find_element(image_path, navigate, 'Reconhecimento da imagem da lupa') 

    if coords:
         # -- Selecionando opcao Agendra planilhas --
        click(coords[0] + 31, coords[1] + 43)    
        save_print(navigate, '08_Acessando_menu_de_busca.png') 

    else:
        save_print(navigate, '08_ERRO_Procurando_menu_de_busca.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")


def processar_empreendimentos(navigate, inicio, fim):
    ''' Realiza a ação de processar as informacoes do documento '''

    # ---- Tela Adendar por empreendimento ----
    logging.info('|--> Selecionando os empreendimentos') 
    image_path = 'png_elements/19_label_empreendimentos_blocos.png'
    coords = find_element(image_path, navigate, 'Selecionando os empreendimentos') 

    if coords:        
        # -- ALT + M --
        logging.info('|--> Apertando ALT + M') # Simular o atalho alt + M para clicar no botão marcar todos
        ACTION.key_down(Keys.ALT).send_keys('m').key_up(Keys.ALT).perform()    
        save_print(navigate, '19_01_Selecionando_criterios_alt_M.png')
        time.sleep(1) 

        # -- ALT + L --
        logging.info('|--> Apertando ALT + L') # Simular o atalho alt + L para clicar no botão alterar dados
        ACTION.key_down(Keys.ALT).send_keys('l').key_up(Keys.ALT).perform()    
        save_print(navigate, '19_02_Selecionando_criterios_alt_L.png')
        time.sleep(1)           

    else:
        save_print(navigate, '19_ERRO_Selecionar_empreendimento.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")  

    # ---- Pop Up dados da agenda ----
    logging.info('|--> Procurando pop up dados da agenda') 
    image_path = 'png_elements/20_dados_agenda.png'
    coords = find_element(image_path, navigate, 'Reconhecimento do pop up dados da agenda') 

    if coords:        
        # -- TAB --
        logging.info('|--> Apertando TAB') 
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, '20_01_Selecionando_criterios_tab.png')
        time.sleep(1)

        # -- TAB --
        logging.info('|--> Apertando TAB') 
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, '20_02_Selecionando_criterios_tab.png')
        time.sleep(1)

        # -- Valor --   
        logging.info('|--> Inserindo data inicio')  
        ACTION.send_keys(inicio).perform()
        time.sleep(1) 
        save_print(navigate, '20_03_Selecionando_data_inicio.png')

        # -- TAB --
        logging.info('|--> Apertando TAB') 
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, '20_04_Selecionando_criterios_tab.png')
        time.sleep(1)

        # -- Valor --   
        logging.info('|--> Inserindo data fim')  
        ACTION.send_keys(fim).perform()
        time.sleep(1) 
        save_print(navigate, '20_05_Selecionando_data_fim.png')

        # -- ALT + O --
        logging.info('|--> Apertando ALT + O') # Simular o atalho alt + O para clicar no botão OK
        ACTION.key_down(Keys.ALT).send_keys('o').key_up(Keys.ALT).perform()    
        save_print(navigate, '20_06_Selecionando_criterios_alt_o.png')
        time.sleep(1)

         # ---- Pop Up confirmacao ----
        logging.info('|--> Procurando pop up confirmacao') 
        image_path = 'png_elements/21_pop_up_confirmacao.png'
        coords = find_element(image_path, navigate, 'Reconhecimento do pop up confirmacao') 

        if coords:  

            # -- Enter --   
            logging.info('|--> Apertando ENTER para Sim do pop up de confirmacao')  
            ACTION.send_keys(Keys.ENTER).perform()
            time.sleep(3) 
            save_print(navigate, '20_07_otao_sim_pop_up_confirmacao.png')  

            # -- ALT + G --
            logging.info('|--> Apertando ALT + G') # Simular o atalho alt + O para clicar no botão Gerar
            ACTION.key_down(Keys.ALT).send_keys('g').key_up(Keys.ALT).perform()    
            save_print(navigate, '20_08_Selecionando_criterios_alt_G.png')
            time.sleep(1)

        # -- Enter --   
        logging.info('|--> Apertando ENTER para Sim do pop up de confirmacao')  
        ACTION.send_keys(Keys.ENTER).perform()
        time.sleep(3) 
        save_print(navigate, '20_09_Pop_up_gerando_planilha.png')     

        # -- Enter --   
        logging.info('|--> Apertando ENTER para OK do pop up de informacao')  
        ACTION.send_keys(Keys.ENTER).perform()
        time.sleep(3) 
        save_print(navigate, '20_09_Pop_up_planilha_gerada_com_sucesso.png')         

    else:
        save_print(navigate, '20_ERRO_Procurando_pop_up_dados_agenda.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")  

        return True



def click(coord_x, coord_y):
    ''' Realiza a ação do click via xdotool do Linux com base na coordenada fornecida '''

    subprocess.run(["xdotool", "mousemove", str(coord_x), str(coord_y)])
    time.sleep(0.5)  # Wait for a moment before clicking
    subprocess.run(["xdotool", "click", "1"])  # Left click


def find_element(image_path, navigate, stage, f5=False):
    ''' Espera até que a imagem seja encontrada na tela ou o timeout seja atingido. '''

    start_time = time.time()

    while time.time() - start_time < TIMEOUT:        
        try:
            time.sleep(INTERVAL)
            coords = get_coord(cv2.imread(image_path))
            
            if coords:
                return coords
    
        except Exception as e:
            logging.info(f"Falha na localização do elemento para: {stage}. Erro: {e}")

        if f5:
            navigate.refresh()

    raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.") # Retorna erro se não encontrar a imagem dentro do tempo limite


def is_image_present(image_path, navigate, stage, timeout=30, interval=1, f5=False):
    """
    Verifica se uma imagem está presente na tela dentro de um tempo limite.
 
    Args:
        image_path (str): Caminho da imagem para buscar na tela.
        navigate (selenium.webdriver): Instância do navegador Selenium (opcional, usado para refresh).
        stage (str): Descrição do estágio atual, para logs.
        timeout (int): Tempo máximo para aguardar pela imagem (em segundos).
        interval (float): Intervalo entre tentativas de busca (em segundos).
        f5 (bool): Se True, atualiza a página em cada tentativa.
 
    Returns:
        bool: True se a imagem foi encontrada, False caso contrário.
    """
 
    start_time = time.time()
 
    while time.time() - start_time < timeout:
        try:
            time.sleep(interval)
            if get_coord(cv2.imread(image_path)):
                logging.info(f"Imagem encontrada para: {stage}")
                return True  # Imagem encontrada
        except Exception as e:
            logging.info(f"Falha ao localizar imagem para: {stage}. Erro: {e}")
 
        if f5:
            navigate.refresh()
 
    logging.warning(f"Imagem '{image_path}' não encontrada dentro do timeout para: {stage}")
    return False  # Imagem não encontrada


def get_coord(template_img):
    ''' Retorna a coordenada exata da imagem de referencia na tela '''

    # ---- Tirando screenshot da tela ----
    process = subprocess.Popen(['import', '-window', 'root', 'png:-'], stdout=subprocess.PIPE)
    screenshot_data = process.stdout.read()
    process.stdout.close()
    image = Image.open(io.BytesIO(screenshot_data))
    screenshot_image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
    

    # ---- Realizando matchTemplate ----
    res = cv2.matchTemplate(screenshot_image, template_img, TEMP_METHOD)
    _, max_val, _, max_loc = cv2.minMaxLoc(res) 

    # ---- Retornando coordenada ----
    if max_val > CONFIDENCE:
        return max_loc
    
    else:
        return None


def execute_with_retry(func, *args, max_retries=3, **kwargs):
    ''' Executa uma função com o máximo de tentativas '''

    attempts = 0

    while attempts < max_retries:
        try:
            func(*args, **kwargs)  # Passa os parâmetros para a função
            return True
        
        except Exception as e:
            attempts += 1            
            logging.info(f"Error in {func.__name__}: {e}. Tentativa {attempts} / {max_retries}")
    
    logging.info(f"{func.__name__} falhou após {max_retries} tentativas.")
    return False


def save_print(navigate, file_name):
    ''' Realiza a screenshot da tela e envia para o S3 '''
    
    try:
        screenshot = navigate.get_screenshot_as_png() 
        time.sleep(0.5)
        send_to_s3(screenshot, file_name)
    
    except:
        logging.info("| PRINT | Falha capturando screenshot")  


def send_error(rpa_data_id, dash_data_id, error_msg):
    ''' Registra o erro no banco de dados '''

    error_data = make_error_obj(error_msg)
    update_log_entry(rpa_data_id, error_data, RPATable)
    update_log_entry(dash_data_id, error_data, DashTable)


def run_routine():
    global ACTION
    ''' Rotina que faz o disparo do processo dentro do sistema Mega, e ao final, realiza o envio da mensagem no teams com o status do processo '''

    try:
        # -------- Ativando LOG -------- 
        rpa_data_id, dash_data_id = start_log()
        
        # ---- Controle das pastas e arquivo: OC_Composicao.xlsx a ser processado ---
        try:
            logging.info('|----> Tentando localizar arquivo para processamento dos dados')
            success,  destino = controle_pastas(base_dir)
            if not success:
                raise Exception()
                
            logging.info('|----> [S] O arquivo POC_Composicao.xlsx localizado com sucesso')
                             
        except Exception as e:
            logging.info('|----> [E] O arquivo POC_Composicao.xlsx nao foi localizado\n', e)
            send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao localizar O arquivo POC_Composicao.xlsx')
            return
        
        # ---- Tratamento dos dados a serem processados ---
        try:
            logging.info('|----> Realizando tratamento do excel a ser processado')
            success, df = processar_excel(destino)
            if not success:
                raise Exception()
                
            logging.info('|----> [S] Realizando tratamento do excel com sucesso')
                             
        except Exception as e:
            logging.info('|----> [E] Realizando tratamento do excel\n', e)
            send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao realizar tratamento do excel')
            return


        # -------- Iniciando Selenium --------        
        with start_selenium() as selenium_context:
            logging.info('|----> Inciando navegador')
            navigate = selenium_context
            ACTION = ActionChains(navigate)
            

            # ---- Login no portal de entrada ---
            try:
                logging.info('|----> Tentando logar no portal de entrada')
                success = execute_with_retry(login_portal_entrada, navigate, max_retries=3)
                if not success:
                    raise Exception()
                
                logging.info('|----> [S] Login no portal de entrada')
                save_print(navigate, '02_Login_portal_entrada.png')
                 
            except Exception as e:
                logging.info('|----> [E] Login no portal de entrada\n', e)
                save_print(navigate, '01_ERRO_login_portal_entrada.png')
                send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante login no portal de entrada')
                return
            

            # ---- Login no portal de entrada ---
            try:
                logging.info('|----> Tentando acessar o sistema Mega')
                success = selecionar_sistema_mega(navigate)
                if not success:
                    raise Exception() 
                
                logging.info('|----> [S] Acesso ao sistema Mega')
                save_print(navigate, '04_Acesso_mega.png')
                 
            except Exception as e:
                logging.info('|----> [E] Acesso ao sistema Mega\n', e)
                save_print(navigate, '03_ERRO_acesso_mega.png')
                send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante acesso ao sistema Mega')
                return
            

            # ---- Login no portal do mega ---
            try:
                logging.info('|----> Tentando logar no portal do mega')
                success = login_portal_mega(navigate)
                if not success:
                    raise Exception()
                
                logging.info('|----> [S] Login no portal do mega')
                save_print(navigate, '06_Login_portal_mega.png')
                 
            except Exception as e:
                logging.info('|----> [E] Login no portal do mega\n', e)
                save_print(navigate, '05_ERRO_login_portal_mega.png')
                send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante login no portal do mega')
                return
            
            for idx, poc_composicao in enumerate(df.itertuples(index=True), start=1):  # O índice será acessado
                # Atribuindo valores às variáveis
                codigo = poc_composicao.CÓDIGO
                nome = poc_composicao.NOME
                diretorio = poc_composicao.DIRETÓRIO_RELATÓRIOS
                status = poc_composicao.STATUS
                competencia = poc_composicao.COMPETÊNCIA
                incc = poc_composicao.INCC
                estagio = poc_composicao.ESTÁGIO
                indice = poc_composicao.INDICE
                cadastro = poc_composicao.CADASTRO
                ajustes = poc_composicao.AJUSTES

                logging.info(f'Linha: {idx}  # Linha do DataFrame')
                logging.info(f'Código: {codigo}')
                logging.info(f'Nome: {nome}')
                logging.info(f'Diretório Relatórios: {diretorio}')
                logging.info(f'Status: {status}')
                logging.info(f'Competência: {competencia}')
                logging.info(f'INCC: {incc}')
                logging.info(f'Estágio: {estagio}')
                logging.info(f'Índice: {indice}')
                logging.info(f'Cadastro: {cadastro}')
                logging.info(f'Ajustes: {ajustes}')

                break

                # ---- Realizando tratamento nas competencias ---
                try:
                    logging.info('|----> Realizando tratamento nas competencias')
                    success, mes_tratado, ano_tratado = tratar_competencia(competencia)
                    if not success:
                        raise Exception()
                        
                    logging.info('|----> [S] Realizando tratamento nas competencias')
                                    
                except Exception as e:
                    logging.info('|----> [E] Realizando tratamento nas competencias\n', e)
                    send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao realizar tratamento nas competencias')
                    return


                # ---- Realizando tratamento das datas ---
                try:
                    logging.info('|----> Realizando tratamento nas datas')
                    success, mes_numero, ultimo_dia, inicio, fim, periodo, ultimo_periodo_fechado = obter_mes_e_ultimo_dia(mes_tratado, ano_tratado)
                    if not success:
                        raise Exception()
                        
                    logging.info('|----> [S] Realizando tratamento nas datas')
                                    
                except Exception as e:
                    logging.info('|----> [E] Realizando tratamento nas datas\n', e)
                    send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao realizar tratamento nas datas')
                    return


                # # ---- Trocando empresa ---
                try:
                    logging.info('|----> Trocando empresa')
                    success = trocar_empresa(navigate, codigo) 
                    if not success:
                        raise Exception()
                    
                    logging.info('|----> [S] Trocando empresa')
                    save_print(navigate, '08_Menu_apos_trocar_empresa.png')
                    
                except Exception as e:
                    logging.info('|----> [E] Trocando empresa\n', e)
                    save_print(navigate, '07_ERRO_trocando_empresa.png')
                    send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante seleção da empresa')
                    return
            

                # ---- Acessando retornos ---
                try:
                    logging.info('|----> Selecionando retornos pendentes')
                    success = acesso_agendar_planilhas(navigate)
                    if not success:
                        raise Exception()
                    
                    logging.info('|----> [S] Selecionando retornos pendentes')
                    save_print(navigate, '09_Selecao_retornos_pendentes.png')
                    
                except Exception as e:
                    logging.info('|----> [E] Selecionando retornos pendentes\n', e)
                    save_print(navigate, '08_ERRO_selecao_retornos_pendentes.png')
                    send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante seleção da pesquisa de retornos')
                    return
                

                # # ---- Processando empreendimentos ---
                try:
                    logging.info('|----> Processando empreendimentos')
                    success = processar_empreendimentos(navigate , inicio, fim)
                    if not success:
                        raise Exception()
                    
                    logging.info('|----> [S] Processando empreendimentos')
                    save_print(navigate, '11_empreendimento Processado.png')
                    
                except Exception as e:
                    logging.info('|----> [E] Processando empreendimentos\n', e)
                    save_print(navigate, '10_ERRO_processando_empreendimentos.png')
                    send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante processamento dos empreendimentos')
                    return
                          

            # ---- Enviando informações finais ---
            logging.info('|----> Enviando informações finais')
            send_card(str(content))
                
        close_leftover_webdriver_instances(vars())


        # -------- Finalizando Log --------
        success_data = make_success_obj()
        update_log_entry(rpa_data_id, success_data, RPATable)
        update_log_entry(dash_data_id, success_data, DashTable)        
    
    except:
        send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante o processamento do documento')


if __name__ == '__main__':
    run_routine()
    close_leftover_webdriver_instances(vars())
