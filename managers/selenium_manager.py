import os, shutil
from selenium import webdriver
from contextlib import contextmanager
from managers.log_manager import log
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service


def get_current_webdriver_path():
    ''' Cria uma pasta com o driver e disponiobiliza esse driver para a aplicação '''

    default_webdriver_path = GeckoDriverManager().install() 
    tmp_folder = './webdrivers/'
    
    if not os.path.exists('./webdrivers'):
        os.mkdir(tmp_folder)
        shutil.copy(default_webdriver_path, tmp_folder)
        os.system(f'chmod +rwx {tmp_folder}/geckodriver') # Dando permissão para conseguir acessar

    return tmp_folder


def close_leftover_webdriver_instances(vars):
    ''' Finaliza todas as instancias do webdriver que por algum motivo não foram finalizadas '''

    try:
        for name in list(vars.keys()):
            try:
                if str(type(vars[name])) == "<class 'selenium.webdriver.firefox.webdriver.WebDriver'>":
                    vars[name].close()

            except Exception:
                pass

    except Exception:
        pass
    

@contextmanager
def start_selenium(): 
    ''' Inicia uma nova instancia do Chrome para ser utilizado com Selenium utilizando o with para garantir que o webdriver será finalizado '''
    
    try:
        target_webdriver_folder_path = get_current_webdriver_path()
        target_webdriver_executable_path = os.path.join(target_webdriver_folder_path, 'geckodriver')

        opts = Options()    
        opts.add_argument("--no-sandbox")
        service = Service(executable_path=target_webdriver_executable_path)
        navigate = webdriver.Chrome(options=opts, service=service) 
        navigate.set_window_size(1382, 736)
        
        log('|----> Inciando navegador')
        yield navigate # Mantem o driver ativo até o fim do with

    except Exception as e:
        log(f'| SELENIUM | Iniciando Selenium | Erro: {e}')        

    finally:        
        try: # Finaliza o webdriver
            if navigate is not None:
                navigate.close()
                navigate.quit()

        except Exception as e:
            log('| SELENIUM | Finalizando Selenium, forçando close/quit')   
     