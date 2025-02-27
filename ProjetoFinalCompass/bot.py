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

# Define algumas variáveis globais
load_dotenv(override=True)

IS_MAESTRO_CONNECTED = eval(os.getenv('IS_MAESTRO_CONNECTED'))
ACTIVITY_LABEL = os.getenv('ACTIVITY_LABEL')
BASE_LOG_PATH = os.getenv('BASE_LOG_PATH') if not IS_MAESTRO_CONNECTED else ''

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = not IS_MAESTRO_CONNECTED


def main():
    global BASE_LOG_PATH
    # Verifica se o maestro está conectado e seleciona de acordo
    if IS_MAESTRO_CONNECTED:
        maestro = BotMaestroSDK.from_sys_args()
        execution = maestro.get_execution()
        BASE_LOG_PATH = execution.parameters.get('BASE_LOG_PATH')
        DEFAULT_PROCESSAR_PATH = execution.parameters.get('DEFAULT_PROCESSAR_PATH')
        DEFAULT_PROCESSADOS_PATH = execution.parameters.get('DEFAULT_PROCESSADOS_PATH')
        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")
    else:
        maestro = None
        DEFAULT_PROCESSAR_PATH = os.getenv('DEFAULT_PROCESSAR_PATH')
        DEFAULT_PROCESSADOS_PATH = os.getenv('DEFAULT_PROCESSADOS_PATH')

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
        df_filtered, empty_cells = clean_df_if_null(df,logger)
        df_output = write_if_null_output(df_output,empty_cells,logger)
        dados = pd.read_csv(r'ProjetoFinalCompass\Processar\dados_processados.csv')
        df_filtered = pd.merge(df_filtered,dados,on='CNPJ',how='left')
        df_filtered = df_filtered.drop(['Email'],axis=1)
        df_filtered, empty_cells = clean_df_if_null(df_filtered,logger)
        df_output = write_if_null_output(df_output,empty_cells,logger)
        df_filtered, df_output = interaction_df_correios(df_filtered=df_filtered,df_output=df_output,bot=bot,logger=logger)
        df_filtered.to_excel(r'ProjetoFinalCompass\Processados\teste.xlsx',index=False)
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
