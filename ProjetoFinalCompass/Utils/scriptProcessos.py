from dotenv import load_dotenv
import os
from botcity.web import WebBot, By, element_as_select
from botcity.maestro import BotMaestroSDK
import pandas as pd
from .helper_functions import *
from .IntegratedLogger import *
from config import vars_map


load_dotenv(override=True)

def catchJadlogPrice(bot:WebBot,maestro:BotMaestroSDK,df_filtered:pd.DataFrame,df_output:pd.DataFrame,logger:IntegratedLogger):
    '''Acessa o site do jadlog e pega a cotação da entrega
    
    # Argumentos
        bot: `WebBot`
            Objeto WebBot usado pelo botcity
        
        maestro: `BotMaestroSDK`
            Objeto maestro usado pelo botcity
        
        df_filtered: `pd.DataFrame`
            Dataframe onde contém só as informações necessárias para a realização do processo
        
        df_output: `pd.DataFrame`
            Dataframe de saída que vai ser editado
        
        logger: `IntegratedLogger`
            Logger usado para gerar arquivos de log
        
        # Retorno
            Dataframe de saída editado
    
    '''
    
    
    try:
        # Definição de Constantes
        DEFAULT_URL_JADLOG = vars_map['DEFAULT_URL_JADLOG']
        PICKUP_VALUE = vars_map['PICKUP_VALUE']
        ORIGIN_CEP = vars_map['ORIGIN_CEP']
        
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
        df_filtered = df_filtered[['CNPJ','TIPO DE SERVIÇO JADLOG','DIMENSÕES CAIXA (altura x largura x comprimento cm)','PESO DO PRODUTO','CEP','VALOR DO PEDIDO']]
        logger.debug(df_filtered.columns.__repr__())
        
        logger.info('Tenta preencher informações dos campos')
        for index, serie in df_filtered.iterrows():
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
                destino_element = bot.find_element('#destino')
                destino_element.clear()
                destino_element.send_keys(serie['CEP'])
                
                logger.debug('Inserindo CEP de origem')
                origin_element = bot.find_element('#origem')
                origin_element.clear()
                origin_element.send_keys(ORIGIN_CEP)
                
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
                
                logger.debug('Inserindo valor de cotação')
                quotation = bot.find_element('//span[contains(text(),"R$")]',By.XPATH).get_attribute('innerText').replace('R$ ','').replace('.',',')
                df_output.loc[df_output['CNPJ'] == serie['CNPJ'],'VALOR COTAÇÃO JADLOG'] = f'R$ {quotation}'
            except:
                df_output.loc[df_output['CNPJ'] == serie['CNPJ'],'STATUS'] = f'Falha cotação jadlog'
                logger.error(process_name='Inserindo valores no site JadLog')
                continue
        bot.stop_browser()
    except:
        logger.error("Execução de CatchJadlogPrice",bot)
    finally:
        return df_output