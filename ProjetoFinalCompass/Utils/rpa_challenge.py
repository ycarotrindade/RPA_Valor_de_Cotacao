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

# Configuração de logging
log_file_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\log\arquivo_de_log.txt"
os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')



def access_website(driver, url):
    """Acessa o site especificado e espera que o elemento 'btn-large' esteja visível."""
    logging.info(f"Acessando o site: {url}")
    driver.get(url)
    WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.CLASS_NAME, 'btn-large')))
    logging.info(f"Site {url} acessado com sucesso!")

def start_challenge(driver):
    """Clica no botão para iniciar o desafio."""
    logging.info("Iniciando desafio")
    driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
    logging.info("Desafio iniciado com sucesso!")

def capture_form_xpaths():
    """Define os XPaths dos campos do formulário em um dicionário."""
    logging.info("Capturando XPaths dos campos do formulário")
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
    logging.debug(f"XPaths capturados: {locators}")
    return locators

def fill_form_data(driver, data):
    """Preenche o formulário com dados do DataFrame."""
    logging.info("Preenchendo formulário com dados do arquivo Excel")
    locators = capture_form_xpaths()
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
            logging.info(f"Linha {index + 1} inserida com sucesso!")
        except Exception as e:
            logging.error(f"Erro ao inserir dados da linha {index + 1}: {e}")
            break #para a execução do for caso encontre um erro.

def capture_execution_time(driver):
    """Captura o tempo de execução exibido no site."""
    logging.info("Capturando tempo de execução")
    wdw = WebDriverWait(driver, 40)
    message = wdw.until(EC.visibility_of_element_located((By.CLASS_NAME, 'message2')))
    execution_time = message.text
    logging.info(f"Tempo de execução capturado: {execution_time}")
    return execution_time

def take_success_screenshot(driver, image_path):
    """Tira um screenshot da página."""
    logging.info(f"Tirando screenshot da página e salvando em {image_path}")
    driver.save_screenshot(image_path)
    logging.info(f"Screenshot salvo com sucesso em {image_path}")

def initialize_browser():
    """Inicializa o navegador Chrome."""
    logging.info("Inicializando navegador Chrome")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    logging.info("Navegador Chrome inicializado com sucesso!")
    return driver

def close_browser(driver):
    """Fecha o navegador Chrome."""
    logging.info("Fechando navegador Chrome")
    driver.quit()
    logging.info("Navegador Chrome fechado com sucesso!")

def rpa_challenge():
    """Função principal para orquestrar a execução do script."""
    excel_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processar\dados_completos.xlsx"
    image_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\img"
    logging.info(f"Iniciando execução do script com arquivo Excel: {excel_path}")
    try:
        data = pd.read_excel(excel_path)
        logging.info(f"Arquivo {excel_path} carregado com sucesso!")
    except Exception as e:
        logging.error(f"Erro ao carregar o arquivo Excel: {e}")
        return

    driver = initialize_browser()
    try:
        access_website(driver, "https://www.rpachallenge.com/")
        start_challenge(driver)
        fill_form_data(driver, data)
        execution_time = capture_execution_time(driver)
        logging.info(f"Tempo de execução: {execution_time}")
        take_success_screenshot(driver, image_path)
        logging.info("Execução finalizada com sucesso!")
    except Exception as e:
        logging.error(f"Erro durante a execução: {e}")
    finally:
        close_browser(driver)

if __name__ == "__main__":
    rpa_challenge()