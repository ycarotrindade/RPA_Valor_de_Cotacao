# Importa o webbot
from botcity.web import WebBot, Browser, By

# Importa o maestro
from botcity.maestro import *

# Importa outras bibliotecas necessárias
from Utils import *
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import os
import pandas as pd
from config import vars_map

# Define algumas variáveis globais
load_dotenv(override=True)

IS_MAESTRO_CONNECTED = vars_map['IS_MAESTRO_CONNECTED']
ACTIVITY_LABEL = vars_map['ACTIVITY_LABEL']
BASE_LOG_PATH = vars_map

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = not IS_MAESTRO_CONNECTED


def main():
    global BASE_LOG_PATH
    # Verifica se o maestro está conectado e seleciona de acordo
    BASE_LOG_PATH = vars_map['BASE_LOG_PATH']
    DEFAULT_PROCESSADOS_PATH = vars_map['DEFAULT_PROCESSADOS_PATH']
    DEFAULT_PROCESSAR_PATH = vars_map['DEFAULT_PROCESSAR_PATH']
    maestro = vars_map['DEFAULT_MAESTRO']
    execution = vars_map['DEFAULT_EXECUTION']
    
    bot = WebBot()

    bot.headless = False

    # Chrome como browser padrão
    bot.browser = Browser.CHROME

    # Informando webdriver
    bot.driver_path = ChromeDriverManager().install()
    
    # Iniciando logger 
    logger = IntegratedLogger(maestro=maestro,filepath=BASE_LOG_PATH,activity_label=ACTIVITY_LABEL)
    try:
        logger.info('='*10 + " Início do Processo: RPA VALOR COTAÇÃO " + "="*10)
        df = open_excel_file_to_dataframe(os.path.join(DEFAULT_PROCESSAR_PATH,'Planilha de Entrada Grupos.xlsx'),logger)
        df_output = create_output_dataframe(df,logger)
        api_data, df_output = api_data_lookup(df_output,logger)
        api_data = make_endereco(api_data)
        df_output, df_correios, df_jadlog = make_jadlog_correios_dataframes(df_output,api_data,logger)
        rpa_challenge(logger,api_data)
        df_filtered, df_output = interaction_df_correios(df_filtered=df_correios,df_output=df_output,bot=bot,logger=logger)
        df_output = catchJadlogPrice(bot=bot,maestro=maestro,df_filtered=df_jadlog,df_output=df_output,logger=logger)
        save_df_output_to_excel(DEFAULT_PROCESSADOS_PATH,df_output,logger)
    except:
        logger.error('Execução RPA_Valor_Cotação')
    
    
    
    if IS_MAESTRO_CONNECTED:
        maestro.finish_task(
            task_id=execution.task_id,
            status=AutomationTaskFinishStatus.SUCCESS,
            message="Task Finished OK.",
            total_items=0,
            processed_items=0,
            failed_items=0
        )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
