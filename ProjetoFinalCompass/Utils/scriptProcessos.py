from dotenv import load_dotenv
import os
from botcity.web import WebBot, By, element_as_select
from botcity.maestro import BotMaestroSDK
import pandas as pd
import traceback
from .helper_functions import *
from .IntegratedLogger import *

load_dotenv(override=True)

def catchJadlogPrice(bot:WebBot,maestro:BotMaestroSDK,df:pd.DataFrame,logger:IntegratedLogger):
    sucess_items = 0
    try:
        # Definição de Constantes
        if maestro is None:
            DEFAULT_URL_JADLOG = os.getenv('DEFAULT_URL_JADLOG')
            BASE_LOG_PATH = os.getenv('BASE_LOG_PATH')
            ACTIVITY_LABEL = None
            PICKUP_VALUE = 50
            ORIGIN_CEP = 38182428
        else:
            execution = maestro.get_execution(task_id=maestro.task_id)
            DEFAULT_URL_JADLOG = execution.parameters.get('DEFAULT_URL_JADLOG')
            BASE_LOG_PATH = execution.parameters.get('BASE_LOG_PATH')
            ACTIVITY_LABEL = os.getenv('ACTIVITY_LABEL')
            PICKUP_VALUE = execution.parameters.get('PICKUP_VALUE')
            PICKUP_VALUE = 50 if len(str(PICKUP_VALUE)) == 0 else PICKUP_VALUE
            ORIGIN_CEP = execution.parameters.get('ORIGIN_CEP')
            ORIGIN_CEP = 38182428 if len(ORIGIN_CEP) == 0 else ORIGIN_CEP
        
        # Main
        logger.info('-'*10 + " Início - catchJadlogPrice " + '-'*10)
        # Essa parte considera que nenhum navegador está aberto
        logger.info('Abre o site de simulação do Jadlog')
        bot.browse(DEFAULT_URL_JADLOG)
        
        logger.info('Verifica se o site abriu corretamente')
        element = bot.find_element('#origem')
        if element is None:
            raise Exception('Página não carregada corretamente')
        
        logger.debug('Reduz o dataframe original para trabalhar somente com as informações necessárias')
        df_filtered = df[['TIPO DE SERVIÇO JADLOG','DIMENSÕES CAIXA (altura x largura x comprimento cm)','PESO DO PRODUTO','CEP']]
        logger.debug(df_filtered.columns.__repr__())
        
        logger.info('Tenta preencher informações dos campos')
        for index, serie in df.iterrows():
            try:
                logger.debug('Inserindo tipo de serviço jedlog')
                jadlog_service_select = element_as_select(bot.find_element('#modalidade'))
                jadlog_value = get_jadlog_value(serie['TIPO DE SERVIÇO JADLOG'])
                jadlog_service_select.select_by_value(jadlog_value)
                
                height, width, lenght = serie['DIMENSÕES CAIXA (altura x largura x comprimento cm)'].split(' x ')
                logger.debug('Inserindo dimensões do pacote')
                logger.debug(f'height = {height} | width = {width} | length = {lenght}')
                width_input = bot.find_element('#valLargura')
                width_input.clear()
                width_input.send_keys(width)
                
                height_input = bot.find_element('#valAltura')
                height_input.clear()
                height_input.send_keys(height)
                
                length_input = bot.find_element('#valComprimento')
                length_input.clear()
                length_input.send_keys(lenght)
                
                logger.debug('Inserindo peso do produto')
                weight_input = bot.find_element('#peso')
                weight_input.clear()
                weight_input.send_keys(serie['PESO DO PRODUTO'])
                
                logger.debug('Inserindo CEP de destino')
                bot.find_element('#destino').send_keys(serie['CEP'])
                
                logger.debug('Inserindo CEP de origem')
                bot.find_element('#origem').send_keys(ORIGIN_CEP)
                
                logger.debug(f'Inserindo valor de coleta | {PICKUP_VALUE}')
                pickup_input = bot.find_element('#valor_coleta')
                pickup_input.clear()
                pickup_input.send_keys(PICKUP_VALUE)
                
                logger.debug('Inserindo valor do pedido')
                
                package_value_input = bot.find_element('#valor_mercadoria')
                package_value_input.clear()
                package_value_input.send_keys(serie['VALOR DO PEDIDO'])
                
                logger.debug('Clica no botão calcular e pega o valor de cotação')
                bot.find_element('//input[@value="Simular"]',By.XPATH).click()
                
                # Delay necessário senão o valor de cotação pego vai ser o anterior
                bot.wait(1000)
                quotation = bot.find_element('//span[contains(text(),"R$")]',By.XPATH).get_attribute('innerText').replace('R$ ','').replace('.',',')
                
                df.loc[index,'VALOR COTAÇÃO JADLOG'] = quotation
                sucess_items += 1
            except:
                logger.error(process_name='Inserindo valores no site JadLog',bot=bot)
        
        bot.stop_browser()
    except:
        logger.error("Execução de CatchJadlogPrice",bot)
    finally:
        return df, sucess_items