import os
import logging
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

logger = logging.getLogger(__name__)

def abrir_chrome(url, pasta_download):
    logger.info("Garantindo que a pasta exista")
    os.makedirs(pasta_download, exist_ok=True)

    ######## Não necessário nesse momento, setei para instalar automatico ##########
    '''logger.info("Caminho para o ChromeDriver (ajuste se necessário)")
    caminho_chromedriver = "chrome_drive\chromedriver.exe"'''
    ################################################################################
    
    logger.info("Configurando as preferências para o Chrome")
    prefs = {
        "pluguns.always_open_pdf_externally": True,
        "download.default_directory": os.path.abspath(pasta_download),
        "safebrowsing.enable":True,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "intl.accept_languages": 'pt-BR'
    }

    logger.info("Configurando opções do Chrome")
    options = Options()
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")
    # options.add_argument("--headless")  # descomente se quiser rodar sem abrir navegador

    logger.info("Criando instância do navegador")
    #service = Service(caminho_chromedriver)
    service = Service(ChromeDriverManager().install())
    chrome = webdriver.Chrome(service=service, options=options)

    logger.info("Acessando o site")
    chrome.get(url)

    sleep(5)
    return chrome