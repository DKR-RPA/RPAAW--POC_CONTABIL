import pyperclip, time, cv2, numpy as np, subprocess, io, logging
from PIL import Image
from managers.config import config as c
from managers.s3_manager import send_to_s3
from managers.log_manager import start_log, start_queue_task, make_error_obj, update_log_entry, make_success_obj, make_success_obj_task
from managers.model_manager import RPATable, DashTable
from managers.share_manager import clear_folder, send_file
from managers.tasks_manager import fetch_pending_tasks
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
import time

TEMP_METHOD = eval('cv2.TM_CCOEFF_NORMED')
CONFIDENCE = 0.8
TIMEOUT = 30
INTERVAL = 5
STR_CONCLUIDO = 'Processamento concluído'


class ImageNotFoundException(Exception):
    """Exceção personalizada para quando a imagem não é encontrada dentro do timeout."""
    pass

logging.basicConfig(level=logging.INFO, format='%(message)s')
ZEEV_TICKET = c.g_ticket.split("/")[0]

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
    image_path = "png_elements/26_label_porla_mega.png"
    if find_element(image_path, navigate, 'Buscando label tela login mega'): 
           
            save_print(navigate, '05_Acesso_tela_login_mega_.png')
            time.sleep(1)
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
    logging.info(f'|--> Inserindo valor: {codigo}')  
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

    logging.info('|--> Validadndo a tela, verificando a existencia do campo Consolidador na tela.')

    image_path = 'png_elements/18_validacao_consolidador.png'
    coords = find_element(image_path, navigate, 'Validando se campo Consolidador carregou com sucesso')
    
    save_print(navigate, '07_4_11_validacao_consolidador.png')

    if coords:
        # -- ALT + S --
        time.sleep(3)
        logging.info('|--> Apertando ALT + S') # Simular o atalho alt + S para clicar no botão selecionar
        ACTION.key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).perform()    
        save_print(navigate, '07_4_11_Selecionando_criterios_alt_s.png')

        logging.info('|--> verificando a existencia de pop up erro.')

        try:
            image_path = 'png_elements/28_pop_up_erro_2.png'
            coords = find_element(image_path, navigate, 'Validando se ocorreu alguma erro no processamento de not valid', time_out=10)
        except:
            coords = ""
            logging.info('|--> Pop up de erro nao indentificado') 

        if coords:
            save_print(navigate, '07_4_12_validacao_consolidador.png')
            # -- ALT + O --
            time.sleep(1)
            logging.info('|--> Apertando ALT + O') # Simular o atalho alt + O para clicar no botão OK
            ACTION.key_down(Keys.ALT).send_keys('o').key_up(Keys.ALT).perform()    
            save_print(navigate, '07_4_11_Selecionando_criterios_alt_O.png')

            # ---- Fechar tela de trocar empresa ----
            logging.info('|--> Procurando botao de close') 
            image_path = 'png_elements/24_botao_close.png'
            coords = find_element(image_path, navigate, 'Reconhecimento do botao close') 

            if coords:  

                click(coords[0] + 10, coords[1] + 20)    
                save_print(navigate, '27_01_clicando_no_botao_close.png') 

                raise Exception("Erro - Ao inserir o Codigo da empresa/empreendimento o trocar empresa retornou vazio os campos.") 
            
        return True
        
    else:
        logging.warning('|--> Campo Consolidador não carregou dentro do tempo limite.')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")


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
        
    return True, inicio, fim


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

        return True

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

        try:
            image_path = 'png_elements/29_pop_up_informacao_.png'
            coords = find_element(image_path, navigate, 'Selecionando os empreendimentos')           
        except:
            coords =""

        if coords:

            # -- Enter --   
            save_print(navigate, '29_01_botao_ok_pop_up_informacao.png')  
            logging.info('|--> Apertando ENTER para OK do pop up de informacao')  
            ACTION.send_keys(Keys.ENTER).perform()
            time.sleep(1) 

            # ---- Fechar tela de agendas por empreendimento ----
            logging.info('|--> Procurando botao de close') 
            image_path = 'png_elements/24_botao_close.png'
            coords = find_element(image_path, navigate, 'Reconhecimento do botao close') 

            if coords:  

                click(coords[0] + 10, coords[1] + 20)    
                save_print(navigate, '27_01_clicando_no_botao_close.png') 

                raise Exception("Nao foi encontrado dados ao selecionar agendar a planilha.")          

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
        time.sleep(0.3)

        # -- TAB --
        logging.info('|--> Apertando TAB') 
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, '20_02_Selecionando_criterios_tab.png')
        time.sleep(0.3)

        # -- Valor --   
        logging.info('|--> Inserindo data inicio')  
        ACTION.send_keys(inicio).perform()
        time.sleep(0.3)
        save_print(navigate, '20_03_Selecionando_data_inicio.png')

        # -- TAB --
        logging.info('|--> Apertando TAB') 
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, '20_04_Selecionando_criterios_tab.png')
        time.sleep(0.3)

        # -- Valor --   
        logging.info('|--> Inserindo data fim')  
        ACTION.send_keys(fim).perform()
        time.sleep(0.3)
        save_print(navigate, '20_05_Selecionando_data_fim.png')

        # -- ALT + O --
        logging.info('|--> Apertando ALT + O') # Simular o atalho alt + O para clicar no botão OK
        ACTION.key_down(Keys.ALT).send_keys('o').key_up(Keys.ALT).perform()    
        save_print(navigate, '20_06_Selecionando_criterios_alt_o.png')
        time.sleep(1)
        
    else:
        save_print(navigate, '20_ERRO_Selecionar_dados_da_agenda.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")  

    # ---- Pop Up confirmacao ----
    logging.info('|--> Procurando pop up confirmacao') 
    image_path = 'png_elements/21_pop_up_confirmacao.png'
    coords = find_element(image_path, navigate, 'Reconhecimento do pop up confirmacao') 

    if coords:  

        # -- Enter --   
        logging.info('|--> Apertando ENTER para Sim do pop up de confirmacao')  
        ACTION.send_keys(Keys.ENTER).perform()
        time.sleep(1) 
        save_print(navigate, '20_07_otao_sim_pop_up_confirmacao.png')  

        # -- ALT + G --
        logging.info('|--> Apertando ALT + G') # Simular o atalho alt + O para clicar no botão Gerar
        ACTION.key_down(Keys.ALT).send_keys('g').key_up(Keys.ALT).perform()    
        save_print(navigate, '20_08_Selecionando_criterios_alt_G.png')
        time.sleep(1)

    
    # ---- Pop Up confirmacao deseja continuar----
    logging.info('|--> Procurando pop up confirmacao deseja continuar') 
    try:
        image_path = 'png_elements/22_pop_up_confirmacao_deseja_continuar.png'
        coords = find_element(image_path, navigate, 'Reconhecimento do pop up confirmacao deseja continuar', time_out=300) 
    except:
        coords = ""

    if coords:  

        # -- Enter --   
        logging.info('|--> Apertando ENTER para Sim do pop up de confirmacao')  
        ACTION.send_keys(Keys.ENTER).perform()
        time.sleep(1) 
        save_print(navigate, '22_01_botao_sim_pop_up_confirmacao_deseja_continuar.png')  

    else: 
        # ---- Pop Up erro ----
        logging.info('|--> Procurando pop up informando erro') 
        image_path = 'png_elements/27_pop_up_erro.png'
        coords = find_element(image_path, navigate, 'Reconhecimento do pop up erro') 

        if coords:  

            # -- Enter --   
            logging.info('|--> Apertando ENTER para OK do pop up erro')  
            ACTION.send_keys(Keys.ENTER).perform()
            time.sleep(1) 
            save_print(navigate, '27_01_botao_OK_pop_up_erro.png')  

            # ---- Fechar tela de agendas por empreendimento ----
            logging.info('|--> Procurando botao de close') 
            image_path = 'png_elements/24_botao_close.png'
            coords = find_element(image_path, navigate, 'Reconhecimento do botao close') 

            if coords:  

                click(coords[0] + 10, coords[1] + 20)    
                save_print(navigate, '27_01_clicando_no_botao_close.png') 

                raise Exception("Erro na geração de planilha.")  

    # ---- Pop Up informacao planilha gerada com ducesso ----
    logging.info('|--> Procurando pop up informacao planilha gerada com sucesso') 
    image_path = 'png_elements/23_pop_up_informacao_planilha_gerada_com_sucesso.png'
    coords = find_element(image_path, navigate, 'Reconhecimento do pop up informacao planilha gerada com sucesso', time_out=300) 

    if coords:  

        # -- Enter --   
        logging.info('|--> Apertando ENTER para SOK do pop up de informacao planilha gerada com sucesso')  
        ACTION.send_keys(Keys.ENTER).perform()
        time.sleep(1) 
        save_print(navigate, '23_01_botao_ok_pop_up_informacao_planilha_gerada_com_sucesso.png')

    # ---- Fechar tela de agendas por empreendimento ----
    logging.info('|--> Procurando botao de close') 
    image_path = 'png_elements/24_botao_close.png'
    coords = find_element(image_path, navigate, 'Reconhecimento do botao close') 

    if coords:  

        click(coords[0] + 10, coords[1] + 20)    
        save_print(navigate, '24_01_clicando_no_botao_close.png') 

    else:
        save_print(navigate, '24_ERRO_Procurando_botao_close.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")  

    # ---- Pop Up confirmcao grava as alteracoes efetuadas ----
    logging.info('|--> Procurando pop up de confirmacao grava as alteracoes efetuadas')    
    image_path = 'png_elements/25_pop_up_confirmacao_grava_alteracoes_efetuadas.png'
    coords = find_element(image_path, navigate, 'Reconhecimento do pop up de confirmacao grava as alteracoes efetuadas') 

    if coords:  

        # -- Enter --   
        logging.info('|--> Apertando ENTER para SOK do pop up de confirmacao grava as alteracoes efetuadas')    
        ACTION.send_keys(Keys.ENTER).perform()
        
        save_print(navigate, '25_01_botao_sim_pop_up_confirmacao_grava_alteracoes efetuadas.png')

    
        return True

    else:
        save_print(navigate, '25_ERRO_Procurando_botao_OK_pop_up_confirmacao_grava_alteracoes_efetuadas.png')
        raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.")
      

def click(coord_x, coord_y):
    ''' Realiza a ação do click via xdotool do Linux com base na coordenada fornecida '''

    subprocess.run(["xdotool", "mousemove", str(coord_x), str(coord_y)])
    time.sleep(0.5)  # Wait for a moment before clicking
    subprocess.run(["xdotool", "click", "1"])  # Left click


def find_element(image_path, navigate, stage, f5=False, time_out=TIMEOUT):
    ''' Espera até que a imagem seja encontrada na tela ou o timeout seja atingido. '''

    start_time = time.time()

    while time.time() - start_time < time_out:        
        try:
            # time.sleep(INTERVAL)
            coords = get_coord(cv2.imread(image_path))
            
            if coords:
                return coords
    
        except Exception as e:
            logging.info(f"Falha na localização do elemento para: {stage}. Erro: {e}")

        if f5:
            navigate.refresh()

    raise ImageNotFoundException(f"Imagem '{image_path}' não encontrada após {TIMEOUT} segundos.") # Retorna erro se não encontrar a imagem dentro do tempo limite


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


def initialize_sheet(columns):
    ''' Inicializa ou sobrescreve uma planilha com as colunas fornecidas '''
    
    # Cria um DataFrame vazio com as colunas especificadas
    df = pd.DataFrame(columns=columns)
    df.to_excel(CAMINHO_PLANILHA_RETORNO, index=False)
    logging.info(f"Planilha inicializada em: {CAMINHO_PLANILHA_RETORNO}")


    ''' Adiciona uma nova linha à planilha existente.'''

    if os.path.exists(CAMINHO_PLANILHA_RETORNO):
        df = pd.read_excel(CAMINHO_PLANILHA_RETORNO)
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        df.to_excel(CAMINHO_PLANILHA_RETORNO, index=False)
        logging.info(f"|----> Linha adicionada: {row_data}")
    
    else:
        logging.info("|----> Planilha não encontrada.")


def create_folder():
    global CAMINHO_PLANILHA_RETORNO
    ''' Cria a pasta onde o arquivo de output vai ser inserido '''

    folder_name_o = 'output'
    if not os.path.exists(folder_name_o):
        os.makedirs(folder_name_o)

    folder_name_i = 'input'
    if not os.path.exists(folder_name_i):
        os.makedirs(folder_name_i)

    CAMINHO_PLANILHA_RETORNO = f'output/{c.g_file} - {ZEEV_TICKET}.xlsx'


def run_routine():
    global ACTION
    ''' Rotina que faz o disparo do processo dentro do sistema Mega, e ao final, realiza o envio da mensagem no teams com o status do processo '''

    try:
        # -------- Ativando LOG -------- 
        rpa_data_id, dash_data_id = start_log()
        logging.info(f"|----> Zeev Ticket: {c.g_ticket}")  

        tasks = fetch_pending_tasks()   

        # -------- Interrompe a execução por falha no JSON -------- 
        if not tasks:
            logging.info('|----> Falha no JSON recebido')
            return
        
        logging.info(f'|--> Quantidade de tasks: {len(tasks)}')  

        # -------- Inicializa uma nova planilha de retorno -------- 
        logging.info('|----> Inicilização da planilha de retorno')
        
        columns = [
            "Codigo", "Nome", "Status"
        ]

        initialize_sheet(columns)                


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
            
            # ---- Iteração com o banco de dados ---
            for task in tasks:
        
                    # Atribuindo valores às variáveis
                    codigo = task['codigo']
                    nome = task['nome']
                    diretorio = task['diretorio_relatorio']
                    status = task['status']
                    competencia = ['competência']
                    incc = task['incc']
                    estagio = task['estagio']
                    indice = task['índice']
                    cadastro = task['cadastro']
                    ajustes = task['ajustes']


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

                     # ---- Linha que vai ser adicionada na planilha caso dê erro ---
                    error_obj = {
                        "Codigo": codigo,
                        "Nome": nome,
                        "Status": "Falha"
                        
                    }


                    arquivo_encontrado = os.path.abspath("./Contabilidade/POC/2025/janeiro/POC_Composição_Janeiro2025.xlsx")
                    print(arquivo_encontrado)  # Verifique se o caminho está correto

                    # Lendo o arquivo Excel
                    df = pd.read_excel(arquivo_encontrado, dtype=str)  # Lendo tudo como string para evitar problemas

                    codigo = df.iloc[7]["Código"]
                    competencia = df.iloc[7]["Competência"]

                    print(codigo)
                    print(competencia)

                    for index, row in df.iloc[200:].iterrows():
                            codigo = row["Código"]
                            competencia = row["Competência"]

                            print(f"Código: {codigo}, Competência: {competencia}")

                    # ---- Realizando tratamento nas competencias ---
                    try:
                        logging.info('|----> Realizando tratamento nas competencias')
                        success, mes_tratado, ano_tratado = tratar_competencia(competencia)
                        if not success:
                            raise Exception()
                            
                        logging.info('|----> [S] Realizando tratamento nas competencias')
                                        
                    except Exception as e:
                        logging.info('|----> [E] Realizando tratamento nas competencias\n, {e}')
                        send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao realizar tratamento nas competencias')
                        continue


                    # ---- Realizando tratamento das datas ---
                    try:
                        logging.info('|----> Realizando tratamento nas datas')
                        success, inicio, fim = obter_mes_e_ultimo_dia(mes_tratado, ano_tratado)
                        if not success:
                            raise Exception()
                            
                        logging.info('|----> [S] Realizando tratamento nas datas')
                                        
                    except Exception as e:
                        logging.info('|----> [E] Realizando tratamento nas datas\n, {e}')
                        send_error(rpa_data_id, dash_data_id, 'Erro: Falha ao realizar tratamento nas datas')
                        continue


                    # # ---- Trocando empresa ---
                    try:
                        logging.info('|----> Trocando empresa')
                        success = trocar_empresa(navigate, codigo) 
                        if not success:
                            raise Exception()
                        
                        logging.info('|----> [S] Trocando empresa')
                        save_print(navigate, '08_Menu_apos_trocar_empresa.png')
                        
                    except Exception as e:
                        logging.info('|----> [E] Trocando empresa\n, {e}')
                        save_print(navigate, '07_ERRO_trocando_empresa.png')
                        send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante seleção da empresa')
                        continue
                

                    # ---- Acessando agendar planilhas ---
                    try:
                        logging.info('|----> Selecionando agendar planilhas')
                        success = acesso_agendar_planilhas(navigate)
                        if not success:
                            raise Exception()
                        
                        logging.info('|----> [S] Selecionando agendar planilhas')
                        save_print(navigate, '09_Selecao_agendar_planilas.png')
                        
                    except Exception as e:
                        logging.info('|----> [E] Selecionando agendar planilhas\n, {e}')
                        save_print(navigate, '08_ERRO_selecao_agendar_planilhas.png')
                        send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante seleção de agendar planilhas')
                        continue
                    

                    # # ---- Processando empreendimentos ---
                    try:
                        logging.info('|----> Processando empreendimentos')
                        success = processar_empreendimentos(navigate , inicio, fim)
                        if not success:
                            raise Exception()
                        
                        logging.info('|----> [S] Processando empreendimentos')
                        save_print(navigate, '11_empreendimento Processado.png')
                        
                    except Exception as e:
                        logging.info('|----> [E] Processando empreendimentos\n, {e}')
                        save_print(navigate, '10_ERRO_processando_empreendimentos.png')
                        send_error(rpa_data_id, dash_data_id, 'Erro: Falha durante processamento dos empreendimentos')
                        continue
                          

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
