import time, cv2, numpy as np, subprocess, io, os, pandas as pd, calendar
from PIL import Image
from managers.config import config as c
from managers.s3_manager import send_to_s3
from managers.log_manager import start_log, make_error_obj, update_log_entry, log, log_erro, finish_log
from managers.model_manager import RPATable, DashTable
from managers.share_manager import clear_folder, send_file, create_folder, initialize_sheet
from managers.tasks_manager import fetch_pending_tasks
from managers.teams_manager import send_card
from managers.selenium_manager import close_leftover_webdriver_instances, start_selenium
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

TEMP_METHOD = eval('cv2.TM_CCOEFF_NORMED')
CONFIDENCE = 0.8
TIMEOUT = 60
INTERVAL = 5
ERRORS = 0


def login_portal_entrada(navigate, step):
    ''' Realiza o login no portal de entrada '''
        
    # -- Acesso portal de entrada --
    try:
        log('--> Acessando url do portal de entrada', step)
        navigate.get(c.c_url)
        save_print(navigate, f'_{step}_0_Acesso_portal_entrada.png')
    
    except Exception as e:
        log_erro(f'Falha ao acessar a url. Erro: {e}', step)
        return False
                          

    # -- Usuário --   
    try:
        log('--> Inserindo info do usuário', step)        
        user_input = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Editbox1"]')))
        user_input.clear()
        user_input.send_keys(c.c_user) 
        save_print(navigate, f'_{step}_1_Inserindo_info_do_usuario.png')
        time.sleep(2)
    
    except Exception as e:
        log_erro(f'Falha ao inserir usuario. Erro: {e}', step)
        return False
    
    
    # -- Senha --
    try:
        log('--> Inserindo info da senha', step)  
        password_input = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="Editbox2"]')))
        password_input.clear()
        password_input.send_keys(c.c_pwd) 
        save_print(navigate, f'_{step}_2_Acesso_portal_Senha.png')
        time.sleep(2)
        
    except Exception as e:
        log_erro(f'Falha ao inserir senha. Erro: {e}', step)
        return False


    # -- Botão de login --
    try:
        log('--> Click botão login', step)  
        btn_login = WebDriverWait(navigate, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="buttonLogOn"]')))
        btn_login.click()
        save_print(navigate, f'_{step}_3_Login_portal_entrada.png')
        time.sleep(8) 
    
    except Exception as e:
        log_erro(f'Falha ao clicar botão de login. Erro: {e}', step)
        return False

        
    # -- Trocando de aba --
    try:
        navigate.switch_to.window(navigate.window_handles[-1]) 
    
    except Exception as e:
        log_erro(f'Trocando de aba no browser. Erro: {e}', step)
        return False
    
    
    # -- Salvando screenshot --
    log('----> [S] Login no portal de entrada', step)   
    save_print(navigate, f'_{step}_4_Login_portal_entrada.png')
    
    return True


def selecionar_sistema_mega(navigate, step):
    ''' Seleciona o sistema Mega e aguarda o carregamento da tela de login '''
    time.sleep(60)
    # ---- Menu do servidor ----
    try:
        log('----> Acessando o sistema Mega', step)
        coords = find_element("png_elements/13_button_mega.png", navigate, 'Login Portal de Entrada', f5=True)        
        if not coords: raise Exception()
        
        click(coords[0] + 32, coords[1] + 33)
        save_print(navigate, f'_{step}_0_Acesso_servidor.png')    
    
    except Exception as e:
        log_erro(f'Falha durante acesso ao sistema Mega. Erro: {e}', step)
        return False


    # -- Salvando screenshot --
    log('----> [S] Acessando o sistema mega', step)
    save_print(navigate, f'_{step}_1_Login_portal_entrada.png')
    
    return True

                    
def login_portal_mega(navigate, step):
    ''' Realiza o login no portal do mega '''
    time.sleep(120)
    navigate.switch_to.window(navigate.window_handles[-1]) 
    # ---- Realizar Login ----
    label_mega_login_img = "png_elements/30_campo_usuario.png"
    coords = find_element(label_mega_login_img, navigate, 'Login Portal de Entrada', f5=True, time_out=30)

    if coords: 

        
        click(coords[0] + 50, coords[1] + 17)
        time.sleep(2.5)           
        
        # -- Usuário --
        try:
            log('--> Inserindo info do usuário', step)            
            ACTION.send_keys(c.c_user).perform()
            save_print(navigate, f'_{step}_0_Inserindo_info_do_usuario.png')
        
        except Exception as e:
            log_erro(f'Falha durante inserção do usuário. Erro: {e}', step)
            return False


        # -- Tab --   
        try:
            log('--> Apertando TAB para senha', step)   
            time.sleep(1) 
            ACTION.send_keys(Keys.TAB).perform() 
        
        except Exception as e:
            log_erro(f'Falha durante apertando TAB. Erro: {e}', step)
            return False


        # -- Senha --   
        try:
            log('--> Inserindo info da senha', step)  
            time.sleep(1)
            ACTION.send_keys(c.c_pwd).perform()
            save_print(navigate, f'_{step}_1_Inserindo_info_da_senha.png')

        except Exception as e:
            log_erro(f'Falha durante inserção da senha. Erro: {e}', step)
            return False
        
        
        # -- Enter --  
        try: 
            log('--> Apertando ENTER para login', step)  
            time.sleep(1) 
            ACTION.send_keys(Keys.ENTER).perform()
            save_print(navigate, f'_{step}_2_Aplicando_enter.png')

        except Exception as e:
            log_erro(f'Falha durante apertando ENTER. Erro: {e}', step)
            return False

    time.sleep(30)
    # ---- Usuário já logado ----
    try:
        if find_element("png_elements/15_aviso_usuario_logado.png", navigate, 'Buscando Popup Usuário Logado', 10): # Caso encontre o aviso de usuário já logado, seleciona a opção "Não" para não desconectar o usuário
            log('--> Manipulando Popup Usuário Logado', step)  
            save_print(navigate, f'_{step}_2_1_Identificando_popup_usuario_logado.png')

            
            # -- Tab --
            try:
                log('--> Apertando TAB para não', step)  
                time.sleep(1)
                ACTION.send_keys(Keys.TAB).perform()                 
                save_print(navigate, f'_{step}_2_2_Aplicando_tab.png')
                
            except Exception as e:
                log_erro(f'Falha durante apertando TAB no popup. Erro: {e}', step)
                return False


            # -- Tab --
            try:
                log('--> Apertando ENTER para não', step)  
                time.sleep(1) 
                ACTION.send_keys(Keys.ENTER).perform()                
                save_print(navigate, f'_{step}_2_3_Aplicando_enter.png')
                
            except Exception as e:
                log_erro(f'Falha durante apertando ENTER no popup. Erro: {e}', step)
                return False
    
    except:
        pass # Significa que não apareceu o Popup


    # ---- Menu do Mega ----
    try:
        log('--> Aguardando login no mega', step) # Validação final
        if not find_element("png_elements/01_pagina_inicial_menu_user.png"  , navigate, 'Login Portal do Mega'): raise Exception() 
    
    except Exception as e:
        log_erro(f'Falha durante validação da tela de login no mega. Erro: {e}', step)
        return False
    
    
    # -- Salvando screenshot --
    log('----> [S]  Inserindo info do sistema mega', step)
    save_print(navigate, f'_{step}_3_Login_sistema_mega.png')

    return True


def trocar_empresa(navigate, codigo, step):
    ''' Realiza as ações para a troca de empresa para iniciar o processo de processamento '''

    log(f'----> Trocando empresa: {codigo}', step)

    # ---- Imagem do perfil ----   
    try:
        log('--> Procurando menu de perfil', step) 
        coords = find_element('png_elements/01_pagina_inicial_menu_user.png', navigate, 'Reconhecimento da imagem do perfil') 
        if not coords: raise Exception()

        click(coords[0] + 18, coords[1] + 17)   
        save_print(navigate, f'_{codigo}__{step}_0_Acessando_menu_de_perfil.png') 
        
    except Exception as e:
        log_erro(f'Falha durante busca do menu de perfil. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_menu_de_perfil.png')
        return 

    
    # ---- Trocar Empresa ----   
    try:
        log('--> Procurando trocar empresa', step) 
        coords = find_element('png_elements/02_pagina_inicial_menu_user_trocar_empresa.png', navigate, 'Reconhecimento da imagem da opção Trocar de Empresa')
        if not coords: raise Exception()

        click(coords[0] + 77, coords[1] + 18)
        save_print(navigate, f'_{codigo}__{step}_1_Acessando_menu_trocar_empresa.png')
    
    except Exception as e:
        log_erro(f'Falha durante busca do menu de troca de empresa. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_menu_trocar_empresa.png')
        return 


    # ---- Alterar para Holding ----   
    try:
        log('--> Procurando empresa selecionada', step)
        coords = find_element('png_elements/03_tela_holding_lupa.png', navigate, 'Reconhecimento da Tela da empresa selecionada')
        if not coords: raise Exception()
                           
        click(coords[0] + 193, coords[1] + 14)
        save_print(navigate, f'_{codigo}__{step}_2_Acessando_menu_empresa_selecionada.png')
    
    except Exception as e:
        log_erro(f'Falha durante busca da empresa selecionada. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_empresa_selecionada.png')
        return 
    

    # -- TAB --    
    try:
        log('--> Apertando TAB para selecionar empresa', step)
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.5)
        save_print(navigate, f'_{codigo}__{step}_3_Acessando_menu_selecionar_empresa.png')
    
    except Exception as e:
        log_erro(f'Falha durante apertando TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_TAB.png')
        return 


    # ---- Alterar critérios | Campo = Código, Tipo = "Igual a", Valor = 1 e Clica no botão selecionar ----
    try:
        log('--> Selecionando critérios', step) 
        log('--> Apertando DOWN 6x', step)  
        for _ in range(6):
            ACTION.send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.2) 
        
        save_print(navigate, f'_{codigo}__{step}_4_1_Selecionando_criterios_down_6x.png')
        time.sleep(1)
    
    except Exception as e:
        log_erro(f'Falha durante apertando down 6x. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_down_6x.png')
        return 
    
    
    # -- ENTER --
    try:
        log('--> Apertando ENTER', step) 
        ACTION.send_keys(Keys.ENTER).perform()    
        time.sleep(0.5)
        save_print(navigate, f'_{codigo}__{step}_4_2_Selecionando_ENTER.png')        
        
    except Exception as e:
        log_erro(f'Falha durante apertando ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ENTER.png')
        return 


    # -- TAB --
    try:
        log('--> Apertando TAB', step) 
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.5)        
        
    except Exception as e:
        log_erro(f'Falha durante apertando TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_TAB.png')
        return


    # -- DELETE --
    try:
        log('--> Apertando BACKSPACE', step)  
        ACTION.send_keys(Keys.BACK_SPACE).perform()
        time.sleep(0.5)
        save_print(navigate, f'_{codigo}__{step}_4_3_Selecionando_criterios_BACKSPACE.png')        
    
    except Exception as e:
        log_erro(f'Falha durante apertando BACKSPACE. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_BACKSPACE.png')
        return


    # -- Down 2x --
    try:
        log('--> Apertando DOWN 2x', step)   
        for _ in range(2):
            ACTION.send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.2) 
        
    except Exception as e:
        log_erro(f'Falha durante apertando DOWN 2x. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_DOWN_2x.png')
        return


    # -- ENTER --
    try:
        log('--> Apertando ENTER', step)  
        ACTION.send_keys(Keys.ENTER).perform()  
        time.sleep(0.5)  
        save_print(navigate, f'_{codigo}__{step}_4_4_Selecionando_criterios_ENTER.png')
        
    
    except Exception as e:
        log_erro(f'Falha durante apertando ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ENTER.png')
        return


    # -- TAB --
    try:
        log('--> Apertando TAB', step)  
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.5) 
    
    except Exception as e:
        log_erro(f'Falha durante apertando TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_TAB.png')
        return


    # -- Valor --   
    try:
        log(f'--> Inserindo valor: {codigo}', step) 
        ACTION.send_keys(codigo).perform()
        time.sleep(1) 
        save_print(navigate, f'_{codigo}__{step}_4_5_Inserindo_codigo.png')
        
    except Exception as e:
        log_erro(f'Falha durante inserir codigo. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Inserindo_codigo.png')
        return


    # -- TAB --
    try:
        log('--> Apertando TAB', step)  
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.5)        
    
    except Exception as e:
        log_erro(f'Falha durante apertando TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_TAB.png')
        return
    

    # -- ENTER --
    try:
        log('--> Apertando ENTER', step)  
        ACTION.send_keys(Keys.ENTER).perform()  
        time.sleep(0.5)  
        save_print(navigate, f'_{codigo}__{step}_4_6_Selecionando_ENTER.png')
    
    except Exception as e:
        log_erro(f'Falha durante apertando ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ENTER.png')
        return
   

    # -- Validando tela --
    try:
        log('--> Validadndo a tela, verificando a existencia do campo Consolidador na tela.', step) 
        coords = find_element('png_elements/18_validacao_consolidador.png', navigate, 'Validando se campo Consolidador carregou com sucesso') 
        if not coords: raise Exception()
        
        save_print(navigate, f'_{codigo}__{step}_4_7_validacao_consolidador.png')
        time.sleep(1)
     
    except Exception as e:
        log_erro(f'Falha durante apertando ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ENTER.png')
        return   
        
    # -- ALT + S --
    try:        
        log('--> Apertando ALT + S', step)  # Simular o atalho alt + S para clicar no botão selecionar
        ACTION.key_down(Keys.ALT).send_keys('s').key_up(Keys.ALT).perform()    
        save_print(navigate, f'_{codigo}__{step}_4_8_Selecionando_criterios_alt_s.png')         
    
    except Exception as e:
        log_erro(f'Falha durante apertando ALT + S. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ALT_S.png')
        return


    # -- Verificando a existencia de pop up erro. --
    try:
        log('--> verificando a existencia de pop up erro.', step)
        coords = find_element( 'png_elements/28_pop_up_erro_2.png', navigate, 'Validando se ocorreu alguma erro na selecao da empresa', time_out=10)
    
    except:
        pass 
    
    if coords:            
                        
        # -- ALT + O --
        try:
            time.sleep(1)
            log('--> Apertando ALT + O', step)  # Simular o atalho alt + O para clicar no botão OK
            ACTION.key_down(Keys.ALT).send_keys('o').key_up(Keys.ALT).perform()  
            time.sleep(0.5)  
            save_print(navigate, f'_{codigo}__{step}_4_9_Selecionando_criterios_alt_O.png')
            
        except Exception as e:
            log_erro(f'Falha durante apertando ALT + O. Erro: {e}', step)
            save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_ALT_O.png')
            return


        # ---- Fechar tela de trocar empresa ----
        try:
            log('--> Procurando botao de close', step) 
            coords = find_element('png_elements/24_botao_close.png', navigate, 'Reconhecimento do botao close') 
            if not coords: raise Exception("Erro - Ao inserir o Codigo da empresa/empreendimento o trocar empresa retornou vazio os campos.")
            
            click(coords[0] + 10, coords[1] + 20)    
            save_print(navigate, f'_{codigo}__{step}_4_10_clicando_no_botao_close.png')                  
    
        except Exception as e:
            log_erro(f'Falha durante apertando close. Erro: {e}', step)
            save_print(navigate, f'_ERRO__{codigo}__{step}_Apertando_CLOSE.png')
            return
        
        
    # -- Salvando screenshot --
    log(f'----> [S] Trocando empresa: {codigo}', step)
    save_print(navigate, f'_{codigo}__{step}_5_Trocando_empresa.png')
    
    return True            
       

def tratar_competencia(competencia, step):
    ''' Separa e trata o mês e o ano da variável COMPETENCIA '''

    log('----> Realizando tratamento nas competencias', step)
    
    meses = {"janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6, "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12}
    
    partes = competencia.split('/') # Divide a string usando '/' como separador
    if len(partes) != 2:
        return "Business Exception: Os dados no campo Competência estão fora do formata esperado: 'Dezembro/2024'."  # Retorna erro caso o formato esteja incorreto    

    # Tratamento do ano: converte para inteiro (se possível)
    try:
        mes = partes[0].strip().lower()  # Remove espaços extras do mês
        ano = int(partes[1].strip())  # Remove espaços extras do ano
    
    except ValueError:
        return "Business Exception: Os dados no campo Competência fora do formata esperado: 'Dezembro/2024'."  # Retorna erro se o ano não for um número válido

    log(f'--> Competencia: Mes: {mes}, Ano: {ano}', step)
    
    mes_numero = meses.get(mes)
    if not mes_numero:
        return "Business Exception: Mês inserido inválido."  
    
    # Usando calendar.monthrange para obter o último dia do mês
    _, ultimo_dia = calendar.monthrange(ano, mes_numero)
    
    # Variáveis conforme solicitado:
    # 1. inicio = "01/mes_numero/ano_tratado"
    inicio = f"01/{mes_numero}/{ano}"
    
    # 2. fim = "ultimo_dia/mes_numero/ano_tratado"
    fim = f"{ultimo_dia}/{mes_numero}/{ano}"    
    log(f'--> Inicio: Mes: {inicio} | Fim: {fim}', step)
        
    return (inicio, fim)


def acesso_agendar_planilhas(navigate, codigo, step):
    ''' Faz a seleção do agendar planilha '''
    
    log(f'----> Selecionando agendar planilhas: {codigo}', step)
    
    
    # ---- Menu de Pesquisar ----
    try:
        log(f'--> Procurando menu de busca', step) 
        coords = find_element('png_elements/07_pagina_inicial_menu_lupa.png', navigate, 'Reconhecimento da imagem da lupa') 
        if not coords: raise Exception()
    
        click(coords[0] + 31, coords[1] + 43)    
        save_print(navigate, f'_{codigo}__{step}_0_Acessando_menu_pagina_inicial_menu_lupa.png')
    
    except Exception as e:
        log_erro(f'Falha durante reconhecimento da imagem da lupa. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_menu_pagina_inicial_menu_lupa.png')
        return 


    # ---- Validando Menu ----
    try:        
        if not find_element('png_elements/08_01_procurar.png', navigate, 'Reconhecimento da imagem da lupa'): raise Exception()
    
    except Exception as e:
        log_erro(f'Falha durante menu de busca. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_menu_de_busca.png')
        return 
    
    
    # ---- Menu de Procurar----
    try:
        log('-->  Preenchendo info na busca', step) 
        ACTION.send_keys('Agendar planilhas').perform()  
        time.sleep(0.5) 
        save_print(navigate, f'_{codigo}__{step}_1_Preenchendo_Busca.png')
            
    except Exception as e:
        log_erro(f'Falha durante preenchimento de info na busca. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Preenchendo_info_na_busca.png')
        return 
  
    # ---- Selecionar opcao Agendar planilhas----
    try:
        log('--> Selecionado opcao Agendar planilhas', step) 
        coords = find_element('png_elements/08_pagina_inicial_menu_lupa_agendar_planilhas.png', navigate, 'Reconhecimento da imagem da lupa')
        if not coords: raise Exception()
        
        click(coords[0] + 31, coords[1] + 43)
   
    except Exception as e:
        log_erro(f'Falha durante seleção de agendar planilhas. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_agendar_planilhas.png')
        return 
    

    # -- Salvando screenshot --
    log(f'----> [S] Selecionando agendar planilhas: {codigo}', step)
    save_print(navigate, f'_{codigo}__{step}_2_Selecao_agendar_planilas.png')
    
    return True       


def processar_empreendimentos(navigate, dado_tratado, codigo, step):
    ''' Realiza a ação de processar as informacoes do documento '''
    
    log(f'----> Processando empreendimentos: {codigo}', step)

    
    # ---- Tela Adendar por empreendimento ----
    try:
        if not find_element('png_elements/19_label_empreendimentos_blocos.png', navigate, 'Selecionando os empreendimentos'): raise Exception()
        
    except Exception as e:
        log_erro(f'Falha durante busca da tela Agendar por empreendimento. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Procurando_menu_agendar_por_empreendimento.png')
        return 

    
    # -- ALT + M --
    try:   
        log('--> Apertando ALT + M', step) # Simular o atalho alt + M para clicar no botão marcar todos
        ACTION.key_down(Keys.ALT).send_keys('m').key_up(Keys.ALT).perform()    
        time.sleep(0.5)
        save_print(navigate, f'_{codigo}__{step}_0_Selecionando_criterios_alt_M.png')         
    
    except Exception as e:
        log_erro(f'Falha durante selecionado ALT + M. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_criterios_alt_M.png')
        return 


    # -- ALT + L --
    try:
        log('--> Apertando ALT + L', step) # Simular o atalho alt + L para clicar no botão alterar dados
        ACTION.key_down(Keys.ALT).send_keys('l').key_up(Keys.ALT).perform()    
        time.sleep(0.5) 
        save_print(navigate, f'_{codigo}__{step}_1_Selecionando_criterios_alt_L.png')
    
    except Exception as e:
        log_erro(f'Falha durante selecionado ALT + L. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_criterios_alt_L.png')
        return 
         

    # -- Validando pop informacao --
    try:
        if not find_element('png_elements/29_pop_up_informacao_.png', navigate, 'Selecionando os empreendimentos'): raise Exception()
        save_print(navigate, f'_{codigo}__{step}_2_botao_ok_pop_up_informacao.png')  
    
    except Exception as e:
        log_erro(f'Falha durante validacao pop up informacao. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Validando_pop_up_informacao.png')
        return 
    

    # -- Enter --
    try:        
        log('--> Apertando ENTER para OK do pop up de informacao', step)  
        ACTION.send_keys(Keys.ENTER).perform()
    
    except Exception as e:
        log_erro(f'Falha durante selecionado ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_ENTER.png')
        return 

            
    # ---- Fechar tela de agendas por empreendimento ----
    try:
        log('--> Procurando botao de close', step) 
        coords = find_element('png_elements/24_botao_close.png', navigate, 'Reconhecimento do botao close')
        if not coords: raise Exception("Nao foi encontrado dados ao selecionar agendar a planilha.")

        click(coords[0] + 10, coords[1] + 20)    
        save_print(navigate, f'_{codigo}__{step}_3_clicando_no_botao_close.png') 
      
    except Exception as e:
        log_erro(f'Falha durante selecionado CLOSE. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_CLOSE.png')
        return


    # ---- Pop Up dados da agenda ----
    try:
        log('--> Procurando pop up dados da agenda', step) 
        image_path = 'png_elements/20_dados_agenda.png'
        if not find_element(image_path, navigate, 'Reconhecimento do pop up dados da agenda') : raise Exception()
    
    except Exception as e:
        log_erro(f'Falha durante selecionado validando_popuop. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Validando_popup.png')
        return

       
    # -- TAB --
    try:
        log('--> Apertando TAB 2x', step) 
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.5)
        ACTION.send_keys(Keys.TAB).perform()
        save_print(navigate, f'_{codigo}__{step}_4_Selecionando_criterios_tab.png')        
        
    except Exception as e:
        log_erro(f'Falha durante selecionado TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_TAB.png')
        return


    # -- Data Inicio --   
    try:
        log('--> Inserindo data inicio', step)  
        ACTION.send_keys(dado_tratado[0]).perform()
        time.sleep(0.3)
        save_print(navigate, f'_{codigo}__{step}_5_Selecionando_data_inicio.png')
    
    except Exception as e:
        log_erro(f'Falha durante selecionado data inicio. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_data_inicio.png')
        return

    
    # -- TAB --
    try:
        log('--> Apertando TAB', step) 
        ACTION.send_keys(Keys.TAB).perform()
        time.sleep(0.3)
    
    except Exception as e:
        log_erro(f'Falha durante selecionado TAB. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_TAB.png')
        return


    # -- Valor --   
    try:
        log('--> Inserindo data fim', step)  
        ACTION.send_keys(dado_tratado[1]).perform()
        time.sleep(0.3)
        save_print(navigate, f'_{codigo}__{step}_6_Selecionando_data_fim.png')
    
    except Exception as e:
        log_erro(f'Falha durante selecionado data fim. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_data_fim.png')
        return


    # -- ALT + O --
    try:
        log('--> Apertando ALT + O', step) # Simular o atalho alt + O para clicar no botão OK
        ACTION.key_down(Keys.ALT).send_keys('o').key_up(Keys.ALT).perform()    
        save_print(navigate, f'_{codigo}__{step}_7_Selecionando_criterios_alt_o.png')
    
    except Exception as e:
        log_erro(f'Falha durante selecionado OK. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Selecionando_OK.png')
        return
         

    # ---- Pop Up confirmacao ----
    try:
        log('--> Procurando pop up confirmacao', step) 
        if not find_element('png_elements/21_pop_up_confirmacao.png', navigate, 'Reconhecimento do pop up confirmacao') : raise Exception()
    
    except Exception as e:
        log_erro(f'Falha durante busca de pop confirmacao. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Buscando_pop_confirmacao.png')
        return

    
    # -- Enter --   
    try:
        log('--> Apertando ENTER para Sim do pop up de confirmacao', step)  
        ACTION.send_keys(Keys.ENTER).perform()
        save_print(navigate, f'_{codigo}__{step}_8_botao_sim_pop_up_confirmacao.png')  
    
    except Exception as e:
        log_erro(f'Falha durante clique de pop confirmacao. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Clicando_pop_confirmacao.png')
        return

    
    # -- ALT + G --
    try:
        log('--> Apertando ALT + G', step) # Simular o atalho alt + O para clicar no botão Gerar
        ACTION.key_down(Keys.ALT).send_keys('g').key_up(Keys.ALT).perform() 
        
    except Exception as e:
        log_erro(f'Falha durante selecionado ALT + G. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_selecionado_ALT_G.png')
        return

    
    # ---- Pop Up confirmacao deseja continuar ----
    try:
        log('--> Procurando pop up confirmacao deseja continuar', step) 
        coord = find_element('png_elements/22_pop_up_confirmacao_deseja_continuar.png', navigate, 'Reconhecimento do pop up confirmacao deseja continuar', time_out=300)
        
        if coord: 
            
            # -- Enter -- 
            try:                   
                log('--> Apertando ENTER para Sim do pop up de confirmacao', step)  
                ACTION.send_keys(Keys.ENTER).perform()
                save_print(navigate, f'_{codigo}__{step}_9_botao_sim_pop_up_confirmacao_deseja_continuar.png')
                
            except Exception as e:
                log_erro(f'Falha durante selecionado ENTER. Erro: {e}', step)
                save_print(navigate, f'_ERRO__{codigo}__{step}_selecionado_ENTER.png')
                return
    
    # ---- Pop Up erro ---- 
        else:                        
            log('--> Procurando pop up informando erro', step) 
            
            if find_element('png_elements/27_pop_up_erro.png', navigate, 'Reconhecimento do pop up erro'):               


                # -- Enter --   
                try:
                    log('|--> Apertando ENTER para OK do pop up erro', step)  
                    save_print(navigate, f'_{codigo}__{step}_9_botao_OK_pop_up_erro.png') 
                    ACTION.send_keys(Keys.ENTER).perform()                      
        
                except Exception as e:
                    log_erro(f'Falha durante selecionado ENTER. Erro: {e}', step)
                    save_print(navigate, f'_ERRO__{codigo}__{step}_selecionado_ENTER.png')
                    return


                # ---- Fechar tela de agendas por empreendimento ----
                try:
                    log('--> Procurando botao de close', step) 
                    coords = find_element('png_elements/24_botao_close.png', navigate, 'Reconhecimento do botao close') 
                    if  not coords: raise Exception()

                    click(coords[0] + 10, coords[1] + 20)    
                    save_print(navigate, f'_{codigo}__{step}_10_clicando_no_botao_close.png') 
                    return

                except Exception as e:
                    log_erro(f'Falha durante selecionado CLOSE. Erro: {e}', step)
                    save_print(navigate, f'_ERRO__{codigo}__{step}_selecionado_CLOSE.png')
                    return             
    
    except Exception as e:
        log_erro(f'Falha durante busca de popups. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_buscando_popups.png')
        return


    # ---- Pop Up informacao planilha gerada com ducesso ----
    try:
        log('--> Procurando pop up informacao planilha gerada com sucesso', step) 
        if not find_element('png_elements/23_pop_up_informacao_planilha_gerada_com_sucesso.png', navigate, 'Reconhecimento do pop up informacao planilha gerada com sucesso', time_out=300): raise Exception() 
    
    except Exception as e:
        log_erro(f'Falha durante busca de popup conbfirmacao. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_buscando_popup_confirmacao.png')
        return
    
    
    # -- Enter --   
    try:
        log('--> Apertando ENTER')  
        ACTION.send_keys(Keys.ENTER).perform()
        save_print(navigate, f'_{codigo}__{step}_10_Apertando_ENTER.png')
    
    except Exception as e:
        log_erro(f'Falha durante ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Durante_ENTER.png')
        return
    

    # ---- Fechar tela de agendas por empreendimento ----
    try:
        log('--> Procurando botao de close', step) 
        coords = find_element('png_elements/24_botao_close.png', navigate, 'Reconhecimento do botao close') 
        if not coords: raise Exception()

        click(coords[0] + 10, coords[1] + 20)    
        save_print(navigate, f'_{codigo}__{step}_11_clicando_no_botao_close.png') 

    except Exception as e:
        log_erro(f'Falha durante CLOSE. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Durante_CLOSE.png')
        return
    

    # ---- Pop Up confirmcao grava as alteracoes efetuadas ----
    try:
        log('--> Procurando pop up de confirmacao grava as alteracoes efetuadas', step)  
        if not find_element('png_elements/25_pop_up_confirmacao_grava_alteracoes_efetuadas.png', navigate, 'Reconhecimento do pop up de confirmacao grava as alteracoes efetuadas'): raise Exception()
        save_print(navigate, f'_{codigo}__{step}_12_pop_up_confirmacao_grava_alteracoes efetuadas.png')       
    
    except Exception as e:
        log_erro(f'Falha durante popup de confirmação. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Durante_popup_confirmacao.png')
        return
    
    
    # ---- ENTER ----
    try:
        log('--> Apertando ENTER para SOK do pop up de confirmacao grava as alteracoes efetuadas', step)    
        ACTION.send_keys(Keys.ENTER).perform()
    
    except Exception as e:
        log_erro(f'Falha durante ENTER. Erro: {e}', step)
        save_print(navigate, f'_ERRO__{codigo}__{step}_Durante_ENTER.png')
        return
    
    
    # -- Salvando screenshot --
    log(f'----> [S] Processando empreendimentos: {codigo}', step)
    save_print(navigate, f'_{codigo}__{step}_13_processando_empreendimentos.png')
    
    return True 
      

def click(coord_x, coord_y):
    ''' Realiza a ação do click via xdotool do Linux com base na coordenada fornecida '''

    subprocess.run(["xdotool", "mousemove", str(coord_x), str(coord_y)])
    time.sleep(0.5)
    subprocess.run(["xdotool", "click", "1"])  # Left click


def find_element(image_path, navigate, stage, f5=False, time_out=TIMEOUT):
    ''' Espera até que a imagem seja encontrada na tela ou o timeout seja atingido. '''

    start_time = time.time()

    while time.time() - start_time < time_out:        
        try:
            coords = get_coord(cv2.imread(image_path))
            
            if coords:
                return coords
    
        except Exception as e:
            log_erro(f"Falha na localização do elemento: {image_path} para: {stage}. Erro: {e}")

        if f5:
            navigate.refresh()

    return # Retorna erro se não encontrar a imagem dentro do tempo limite


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
            log_erro(f"Função: {func.__name__}. Tentativa {attempts} / {max_retries}. Erro: {e}")
    
    log_erro(f"Função: {func.__name__} falhou após {max_retries} tentativas.")
    return False


def save_print(navigate, file_name):
    ''' Realiza a screenshot da tela e envia para o S3 '''
    
    try:
        screenshot = navigate.get_screenshot_as_png() 
        time.sleep(0.5)
        send_to_s3(screenshot, file_name)
    
    except:
        log_erro(f" PRINT | Falha capturando screenshot com nome: {file_name}")  


def send_error(error_msg, rpa_data_id=False, dash_data_id=False):
    ''' Registra o erro no banco de dados '''

    error_data = make_error_obj(error_msg)
    
    if rpa_data_id: 
        update_log_entry(rpa_data_id, error_data, RPATable)
    
    else:
        update_log_entry(c.rpa_data_id, error_data, RPATable)
        
    if dash_data_id: 
        update_log_entry(c.dash_data_id, error_data, DashTable)


def exception_action(navigate, img_path, error_msg, rpa_data=False, dash_data_id=False, add_df=False, add_erro=False):
    ''' Realiza as ações necessárias após uma Exception  no código '''
    
    save_print(navigate, img_path)
    send_error(error_msg, rpa_data, dash_data_id)
    
    if add_df:
        save_task(add_df)
    
    if add_erro:
        global ERRORS
        
        ERRORS += 1


def save_task(row_data):
    ''' Adiciona uma nova linha à planilha existente.'''

    if os.path.exists(c.folder_path_o):
        df = pd.read_excel(c.folder_path_o)
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        df.to_excel(c.folder_path_o, index=False)
        log(f"=-=-> Linha adicionada na planilha.")
    
    else:
        log_erro("----> Planilha não encontrada.")


def prep_run():
    ''' Realiza todas as ações necessarias para o processo de execução '''
    global TASKS
    
    try:
        # -------- Ativando LOG -------- 
        c.rpa_data_id, c.dash_data_id = start_log()
        c.zeev_ticket = c.g_ticket.split("/")[0]
        log(f"----> Zeev Ticket: {c.g_ticket} <---- |")
        
        
        # -------- Buscando tasks --------
        TASKS = fetch_pending_tasks()   
        if not TASKS:
            log('----> Nenhuma tarefa para processar')
            return
        
        log(f'|--> Quantidade de tasks: {len(TASKS)}') 
        
        
        # -------- Planilha de retorno --------         
        initialize_sheet() 
        
        
        # -------- Criando pasta de output --------
        create_folder()
        
    except Exception as e:
        log_erro(f'Falha durante a preparação do ambiente para execução. Erro: {e}')
        

def finish_run():
    ''' Realiza todas as ações necessarias para finalizar o processo de execução '''
    global TASKS
    
    try:
        # -------- Limpando pasta de output --------
        clear_folder()
        
    except Exception as e:
        log_erro(f'Falha durante a finalizacao do ambiente para execução. Erro: {e}')


def run_routine():
    global ACTION, ERRORS 
    ''' Rotina que faz o disparo do processo dentro do sistema Mega, e ao final, realiza o envio da mensagem no teams com o status do processo '''

    try: 
        # -------- Iniciando Selenium --------        
        with start_selenium() as navigate:
            ACTION = ActionChains(navigate)
            
            # ---- Login no portal de entrada ---
            try:                
                if not execute_with_retry(login_portal_entrada, navigate, '1', max_retries=3): raise Exception()
                 
            except:
                exception_action(navigate, '1_ERRO_login_portal_entrada.png', 'Erro: Falha durante login no portal de entrada', dash_data_id=True)
                return
            

            # ---- Login no portal de entrada ---
            try:                
                if not selecionar_sistema_mega(navigate, '2'): raise Exception() 
                                 
            except:
                exception_action(navigate, '2_ERRO_acesso_mega.png', 'Erro: Falha durante seleção do sistema Mega', dash_data_id=True)
                return
            

            # ---- Login no portal do mega ---
            try:
                if not login_portal_mega(navigate, '3'): raise Exception()
                 
            except Exception as e:
                exception_action(navigate, '3_ERRO_acesso_mega.png', 'Erro: Falha durante acesso ao sistema Mega', dash_data_id=True)
                return
            
            # ---- Iteração com o banco de dados ---
            for task in TASKS:
        
                # Atribuindo valores às variáveis               
                nome = task['task']['nome']
                incc = task['task']['incc']
                status = task['task']['status']
                codigo = task['task']['codigo']
                estagio = task['task']['estagio']
                ajustes = task['task']['ajustes']
                cadastro = task['task']['cadastro']                
                diretorio = task['task']['diretorio_relatorio']
                competencia = task['task']['competencia']
                rpa_data_id = task['id']
                
                # Linha que vai ser adicionada na planilha caso dê erro
                error_obj = {"Codigo": codigo, "Nome": nome, "Status": "Falha"}
                
                log(f'----> Iniciando processamento\nCódigo: {codigo}\nNome: {nome}\nDiretório Relatórios: {diretorio}\nStatus: {status}\nCompetência: {competencia}\nINCC: {incc}\nEstágio: {estagio}\nCadastro: {cadastro}\nAjustes: {ajustes}')              


                # ---- Realizando tratamento na competencia ---
                try:                    
                    dado_tratado = tratar_competencia(competencia, '4') 
                    if isinstance(dado_tratado, str): 
                        exception_action(navigate, '4_ERRO_Trantando_competencia.png', dado_tratado, rpa_data_id, add_erro=True)
                        continue                    
                    
                    if not dado_tratado: raise Exception()
                                    
                except:
                    exception_action(navigate, '4_ERRO_Trantando_competencia.png', 'Erro: Falha ao realizar tratamento na competencia', rpa_data_id, add_df=error_obj, add_erro=True)
                    continue                


                # ---- Trocando empresa ---
                try:
                    if not trocar_empresa(navigate, codigo, '5'): raise Exception()
                    
                except:
                    exception_action(navigate, '5_ERRO_Trocando_empresa.png', 'Erro: Falha ao trocar empresa', rpa_data_id, add_df=error_obj, add_erro=True)
                    continue 
            

                # ---- Acessando agendar planilhas ---
                try:                    
                    if not acesso_agendar_planilhas(navigate, codigo, '6'): raise Exception()
                    
                except:
                    exception_action(navigate, '6_ERRO_Acessando_agendar_planilha.png', 'Erro: Falha ao Acessar agendar planilha', rpa_data_id, add_df=error_obj, add_erro=True)
                    continue
                

                # # ---- Processando empreendimentos ---
                try:
                    if not processar_empreendimentos(navigate, dado_tratado, codigo, '7'): raise Exception()
                    save_task({"Codigo": codigo, "Nome": nome, "Status": "Planilha gerada com sucesso"})
                    update_log_entry(rpa_data_id, 'Planilha gerada com sucesso', RPATable)
                    
                except:
                    exception_action(navigate, '7_ERRO_Processar_empreendimento.png', 'Erro: Falha ao Processar Empreendimento', rpa_data_id, add_df=error_obj, add_erro=True)
                    continue
                          

            # ---- Enviando informações finais ---
            log('----> Enviando informações finais')
            send_card(len(TASKS), ERRORS)
            send_file()
                
        close_leftover_webdriver_instances(vars())


        # -------- Finalizando Log --------        
        finish_log()
    
    except:        
        finish_log(error=True)


if __name__ == '__main__':
    prep_run()
    run_routine()
    finish_run()
    close_leftover_webdriver_instances(vars())
