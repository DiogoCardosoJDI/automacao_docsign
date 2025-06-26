import logging
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

logger = logging.getLogger(__name__)

def gerar_documento_1(chrome, dados):

    try:
        timeout = 30
        logger.info("Preenchendo Fornecedor")
        seu_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[1]/input')))
        seu_nome.send_keys(dados['Unnamed: 1'].strip())
        seu_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[2]/input')))
        seu_email.send_keys(dados['Unnamed: 2'].strip())

        logger.info("Preenchendo Analista Observador")
        obs_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[3]/input')))
        obs_nome.send_keys(dados['Unnamed: 3'].strip())
        obs_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[4]/input')))
        obs_email.send_keys(dados['Unnamed: 4'].strip())

        logger.info("Preenchendo Fornecedor Observador")
        forn_obs_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[5]/input')))
        forn_obs_nome.send_keys(dados['Unnamed: 5'].strip())
        forn_obs_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[6]/input')))
        forn_obs_email.send_keys(dados['Unnamed: 6'].strip())

        logger.info("Preenchendo Gestor1")
        gest1_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[7]/input')))
        gest1_nome.send_keys(dados['Unnamed: 7'].strip())
        gest1_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[8]/input')))
        gest1_email.send_keys(dados['Unnamed: 8'].strip())

        logger.info("Preenchendo Gestor2")
        gest2_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[9]/input')))
        gest2_nome.send_keys(dados['Unnamed: 9'].strip())
        gest2_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[10]/input')))
        gest2_email.send_keys(dados['Unnamed: 10'].strip())

        logger.info("Preenchendo Gestor3")
        gest3_nome = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[11]/input')))
        gest3_nome.send_keys(dados['Unnamed: 11'].strip())
        gest3_email = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[1]/div/div[12]/input')))
        gest3_email.send_keys(dados['Unnamed: 12'].strip())
        
        logger.info("Clicando no botão iniciar")
        botao_comecar_assinatura = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/form/div[2]/div/div/button')))
        botao_comecar_assinatura.click()

        logger.info("Verificando se existe popup na tela e clicando para fechar")
        if verificar_elemento_existe(chrome, '//*[@id="ModalContainer"]/div[2]/div[2]/div/div/div[3]/div/div[2]/div[2]/button[2]'):
            clicar_ok = WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ModalContainer"]/div[2]/div[2]/div/div/div[3]/div/div[2]/div[2]/button[2]')))
            clicar_ok.click()

        sleep(5)

        return chrome

    except Exception as e:
        logger.error(f"Elemento não encontrado após {timeout} segundos. Erro: {e}")
        return None

def gerar_documento_2(chrome, dados):
    try:
        logger.info("Iniciando segunda parte do preencimento")
        campos = chrome.find_elements(By.XPATH, "//*[starts-with(@id, 'tab-form-element')]")
        ref_valor = 13
        cont = 0
        passou = False
        pular_linhas = False
        pular_proximas = False
        assina = True

        for i, campo in enumerate(campos):
            try:
                tag_name = campo.tag_name.lower()

                if ref_valor == 16 and not passou:
                    campo.clear()
                    campo.send_keys("")
                    passou = True
                    continue

                elif ref_valor in [45, 52, 59]:
                    ref_valor += 1
                    valor_input = str(dados[f'Unnamed: {str(ref_valor)}']).strip()
                    if valor_input == 'nan':
                        valor_input = ""
                
                elif ref_valor == 19:
                    valor_input = dados[f'Unnamed: {str(ref_valor)}'].strftime('%d/%m/%Y').strip()

                elif ref_valor == 22 and not pular_linhas:
                    pular_linhas = True
                    continue

                elif pular_linhas and cont < 11:
                    campo.clear()
                    campo.send_keys("")
                    cont += 1
                    continue
                elif pular_proximas and ref_valor >= 35 and ref_valor <= 37:
                    if ref_valor < 38:                   
                        ref_valor += 1
                        continue
                    elif pular_proximas:
                        pular_proximas = False
                elif pular_proximas and ref_valor >= 39:
                    if ref_valor < 60:                   
                        ref_valor += 1
                        continue

                elif ref_valor == 32:
                    ref_valor += 1
                    valor_input = str(dados[f'Unnamed: {str(ref_valor)}']).strip()
                    if valor_input == 'nan':
                        valor_input = ""

                elif ref_valor in [22, 34] and assina:
                    logger.info("Proceurando botões 'Rubricar'")
                    wait = WebDriverWait(chrome, 10)
                    actions = ActionChains(chrome)
                    wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//div[contains(text(), 'Rubricar')]]")))
                    botoes_rubricar = chrome.find_elements(By.XPATH, "//button[.//div[contains(text(), 'Rubricar')]]")
                    
                    logger.info(f"Encontrados {len(botoes_rubricar)} botões 'Rubricar'.")
                    primeiro = True
                    for y, botao in enumerate(botoes_rubricar, start=1):
                        try: 
                            if botao.is_displayed() and botao.is_enabled():
                                if ref_valor == 22 and assina:
                                    actions.move_to_element(botao).click().perform()
                                    logger.info(f"Botão {y} clicado com sucesso.")
                                    sleep(0.5)
                                    break
                                elif ref_valor == 34 and assina:
                                    if primeiro:
                                        primeiro = False
                                        continue
                                    actions.move_to_element(botao).click().perform()
                                    logger.info(f"Botão {y} clicado com sucesso.")
                                    sleep(0.5)
                            else:
                                logger.warning(f"Botão {y} não está visível ou habilitado.")
                        except Exception as e:
                            logger.error(f"Erro ao clicar no botão {i}: {e}")

                    assina = False  

                elif ref_valor == 27:
                    ref_valor += 1
                    valor_input = dados[f'Unnamed: {str(ref_valor)}'].strftime('%d/%m/%Y').strip()

                else:
                    valor_input = str(dados[f'Unnamed: {str(ref_valor)}']).strip()
                    if valor_input == 'nan':
                        valor_input = ""
                    assina = True

                logger.info("Verificando se o campo é <select> e realizando tratamento no texto para correta seleção")
                if tag_name == "select":
                    if ref_valor == 30:
                        if int(valor_input) > 1:
                            valor_input = f'{valor_input.strip()} Meses'
                        else:
                            valor_input = f'{valor_input.strip()} Mês'
                    elif ref_valor == 31:
                        if int(valor_input) > 1:
                            valor_input = f'{valor_input.strip()} dias'
                        else:
                            valor_input = f'{valor_input.strip()} dia'
                    elif ref_valor in [34, 38]:
                            valor_input = f'{valor_input.strip().upper()}'
                            if ref_valor in [34, 38] and valor_input == 'NÃO':
                                pular_proximas = True
                            else:
                                pular_proximas = False
                    elif ref_valor == 60:
                        valor_input = valor_input.replace(' / ' , '/').replace('  ', ' ')
                    try:
                        select = Select(campo)
                        select.select_by_visible_text(valor_input)
                        logger.info(f"Campo select {i} preenchido com '{valor_input}'")
                    except Exception as e:
                        logger.error(f"Erro ao selecionar opção '{valor_input}' no campo {i}: {e}")
                else:
                    campo.clear()
                    if ref_valor in [35, 36, 40, 41, 47, 48, 54, 55]:
                        if not pd.isna(dados[f'Unnamed: {str(ref_valor)}']):
                            valor_input = dados[f'Unnamed: {str(ref_valor)}'].strftime('%d/%m/%Y').strip()
                            campo.send_keys(valor_input)
                    else:
                        campo.send_keys(valor_input)
                    logger.info(f"Campo input {i} preenchido com '{valor_input}'")

                ref_valor += 1
                if ref_valor == 63 and assina:
                    logger.info("Procurando botões 'Assinar'")
                    wait = WebDriverWait(chrome, 10)
                    actions = ActionChains(chrome)
                    wait.until(EC.presence_of_element_located((By.XPATH, "//button[.//div[contains(text(), 'Assinar')]]")))
                    botao_assinar = chrome.find_elements(By.XPATH, "//button[.//div[contains(text(), 'Assinar')]]")
                    logger.info(f"Encontrados {len(botao_assinar)} botões 'Assinar'.")
                    primeiro = True
                    for t, botao in enumerate(botao_assinar, start=1):
                        try: 
                            if botao.is_displayed() and botao.is_enabled():
                                if ref_valor == 63:
                                    actions.move_to_element(botao).click().perform()
                                    logger.info(f"Botão {t} clicado com sucesso.")
                                    sleep(0.5)
                                    break
                        except Exception as e:
                            logger.error(f"Erro ao clicar no botão {i}: {e}")
                          
                    assina = False
                elif ref_valor > 64:
                    break

            except Exception as e:
                logger.warning(f"Erro ao preencher campo {i}: {e}")

        logger.info("Clicando no botão concluir")
        btn_concluir = chrome.find_elements(By.XPATH, '//*[@id="end-of-document-btn-finish"]')
        #btn_concluir.click()

        logger.info("Fechando o chrome")
        chrome.quit()
        return chrome

    except Exception as e:
        logger.error(f"Erro: {e}")
        return None

def verificar_elemento_existe(chrome, xpath):
    try:
        timeout = 30
        logger.info("Tentando localizar o elemento")
        WebDriverWait(chrome, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        logger.info("Elemento encontrado, retornando 'True'")
        return True  
    except NoSuchElementException:
        logger.error("Elemento não encontrado, retornando 'False'")
        return False  