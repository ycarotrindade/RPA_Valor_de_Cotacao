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

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = not IS_MAESTRO_CONNECTED


def main():
    # Verifica se o maestro está conectado e seleciona de acordo
    maestro = vars_map['DEFAULT_MAESTRO']
    execution = vars_map['DEFAULT_EXECUTION']
    bot = vars_map['DEFAULT_BOT']
    
    # Iniciando logger 
    logger = IntegratedLogger(maestro=maestro,filepath=vars_map['BASE_LOG_PATH'],activity_label=vars_map['ACTIVITY_LABEL'])

    try:
        logger.info('='*10 + " Início do Processo: RPA VALOR COTAÇÃO " + "="*10)
        
        # Cria um DataFrame com os dados de entrada do arquivo Excel 
        df = open_excel_file_to_dataframe(os.path.join(vars_map['DEFAULT_PROCESSAR_PATH'],'Planilha de Entrada Grupos.xlsx'), logger)
        # Cria um DataFrame para receber os dados de saída
        df_output = create_output_dataframe(df, logger)

        # Faz requisição no site Brasil API e retorna os resultados nos DataFrames de saída e para as cotações
        api_data, df_output = api_data_lookup(df_output, logger)

        # Cria uma coluna única para o endereço
        api_data = make_endereco(api_data, logger)
        # Cria DataFrames para trabalhar com as cotações de Correios e JadLog, e preenche dados no DataFrame de saída
        df_output, df_correios, df_jadlog = make_jadlog_correios_dataframes(df_output,api_data, logger)

        # Realiza o preenchimento das informações no site RPA Challenge
        rpa_challenge(df=api_data, logger=logger)

        # Faz a cotação dos preços no site do Correios e retorna os resultados para o DataFrame de saída
        df_output = interaction_df_correios(df_filtered=df_correios,df_output=df_output,bot=bot,logger=logger)
        # Faz a cotação dos preços no site da JadLog e retorna os resultados para o DataFrame de saída
        df_output = catchJadlogPrice(bot=bot,maestro=maestro,df_filtered=df_jadlog,df_output=df_output,logger=logger)

        # Salva o DataFrame de saída em um arquivo Excel (.xlsx)
        output_file = save_df_output_to_excel(vars_map['DEFAULT_PROCESSADOS_PATH'],df_output, logger)
        # Faz a comparação entre as cotações do Correios e JadLog e preenche de verde a célula de menor valor
        compare_quotation(df_output,output_file, logger)

        # Faz o envio de e-mails da finalização do processo com arquivo Excel dos dados de saída 
        send_emails(output_file)
        
    except:
        logger.error('Execução RPA_Valor_Cotação')
        if IS_MAESTRO_CONNECTED:
            maestro.finish_task(
                task_id=execution.task_id,
                status=AutomationTaskFinishStatus.FAILED,
                message="Task finished with errors",
                total_items=0,
                processed_items=0,
                failed_items=0
            )
    else:
        total_tasks, total_finished, total_errors = calc_finish_task(df_output)
        if IS_MAESTRO_CONNECTED:
            maestro.post_artifact(
                task_id=execution.task_id,
                artifact_name=os.path.basename(output_file),
                filepath=output_file
            )
            maestro.finish_task(
                task_id=execution.task_id,
                status=AutomationTaskFinishStatus.SUCCESS,
                message="Task Finished OK.",
                total_items=total_tasks,
                processed_items=total_finished,
                failed_items=total_errors
            )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
