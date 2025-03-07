import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import logging
import os
from config import vars_map

# Configuração de logging
# log_file_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\log\arquivo_de_log.txt"
# os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
# logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')



def access_website(driver, url,logger):
    """Acessa o site especificado e espera que o elemento 'btn-large' esteja visível."""
    logger.info(f"Acessando o site: {url}")
    driver.get(url)
    WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-large')))
    logger.info(f"Site {url} acessado com sucesso!")

def start_challenge(driver, logger):
    """Clica no botão para iniciar o desafio."""
    logger.info("Iniciando desafio")
    driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
    logger.info("Desafio iniciado com sucesso!")

def capture_form_xpaths(logger):
    """Define os XPaths dos campos do formulário em um dicionário."""
    logger.info("Capturando XPaths dos campos do formulário")
    locators = {
        'first_name': '//input[@ng-reflect-name="labelFirstName"]',
        'last_name': '//input[@ng-reflect-name="labelLastName"]',
        'company_name': '//input[@ng-reflect-name="labelCompanyName"]',
        'role_in_company': '//input[@ng-reflect-name="labelRole"]',
        'address': '//input[@ng-reflect-name="labelAddress"]',
        'email': '//input[@ng-reflect-name="labelEmail"]',
        'phone_number': '//input[@ng-reflect-name="labelPhone"]',
        'submit': '//input[@value="Submit"]'
    }
    logger.debug(f"XPaths capturados: {locators}")
    return locators

def fill_form_data(driver, data,logger):
    """Preenche o formulário com dados do DataFrame."""
    logger.info("Preenchendo formulário com dados do arquivo Excel")
    locators = capture_form_xpaths(logger)
    for index, row in data.iterrows():
        try:
            driver.find_element(By.XPATH, locators['first_name']).send_keys(row['RAZÃO SOCIAL'])
            driver.find_element(By.XPATH, locators['last_name']).send_keys(row['SITUAÇÃO CADASTRAL'])
            driver.find_element(By.XPATH, locators['company_name']).send_keys(row['NOME FANTASIA'])
            driver.find_element(By.XPATH, locators['role_in_company']).send_keys(row['DESCRIÇÃO MATRIZ FILIAL'])
            driver.find_element(By.XPATH, locators['address']).send_keys(row['ENDEREÇO'])
            driver.find_element(By.XPATH, locators['email']).send_keys(row['E-MAIL'])
            driver.find_element(By.XPATH, locators['phone_number']).send_keys(row['TELEFONE + DDD'])
            driver.find_element(By.XPATH, locators['submit']).click()
            logger.info(f"Linha {index + 1} inserida com sucesso!")
        except Exception as e:
            logger.error(f"Erro ao inserir dados da linha {index + 1}: {e}")
            break #para a execução do for caso encontre um erro.

def capture_execution_time(driver,logger):
    """Captura o tempo de execução exibido no site."""
    logger.info("Capturando tempo de execução")
    wdw = WebDriverWait(driver, 40)
    message = wdw.until(EC.visibility_of_element_located((By.CLASS_NAME, 'message2')))
    execution_time = message.text
    logger.info(f"Tempo de execução capturado: {execution_time}")
    return execution_time

def take_success_screenshot(driver, image_path):
    """Tira um screenshot da página."""
    logging.info(f"Tirando screenshot da página e salvando em {image_path}")
    driver.save_screenshot(image_path)
    logging.info(f"Screenshot salvo com sucesso em {image_path}")

def initialize_browser(logger):
    """Inicializa o navegador Chrome."""
    logger.info("Inicializando navegador Chrome")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    logger.info("Navegador Chrome inicializado com sucesso!")
    return driver

def close_browser(driver,logger):
    """Fecha o navegador Chrome."""
    logger.info("Fechando navegador Chrome")
    driver.quit()
    logger.info("Navegador Chrome fechado com sucesso!")

def rpa_challenge(logger,df):
    """Função principal para orquestrar a execução do script."""
    image_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\img"
    try:
        data = df
        logger.info(f"Arquivo carregado com sucesso!")
    except Exception as e:
        logger.error(f"Erro ao carregar o arquivo Excel")
        return

    driver = initialize_browser(logger)
    try:
        access_website(driver, vars_map['DEFAUT_RPACHALLENGE_URL'],logger)
        #start_challenge(driver,logger)
        fill_form_data(driver, data,logger)
        #execution_time = capture_execution_time(driver,logger)
        #logger.info(f"Tempo de execução: {execution_time}")
        #take_success_screenshot(driver, image_path,logger)
        logger.info("Execução finalizada com sucesso!")
    except Exception as e:
        logger.error(f"Processo rpa_challenge")
    finally:
        close_browser(driver,logger)

if __name__ == "__main__":
    rpa_challenge()