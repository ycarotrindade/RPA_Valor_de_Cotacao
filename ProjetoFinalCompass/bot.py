# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Import other libraries
from Utils import *
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import os

load_dotenv(override=True)

IS_MAESTRO_CONNECTED = eval(os.getenv('IS_MAESTRO_CONNECTED'))
ACTIVITY_LABEL = os.getenv('ACTIVITY_LABEL')
BASE_LOG_PATH = os.getenv('BASE_LOG_PATH') if not IS_MAESTRO_CONNECTED else ''

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = not IS_MAESTRO_CONNECTED


def main():
    global BASE_LOG_PATH
    if IS_MAESTRO_CONNECTED:
        maestro = BotMaestroSDK.from_sys_args()
        execution = maestro.get_execution()
        BASE_LOG_PATH = execution.parameters.get('BASE_LOG_PATH')
        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")
    else:
        maestro = None

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Setting default browser to Chrome
    bot.browser = Browser.CHROME

    # Setting Webdriver
    bot.driver_path = ChromeDriverManager().install()
    
    processed_items = 0
    
    dados = {
        'TIPO DE SERVIÇO JADLOG':['JADLOG Expresso','JADLOG Econômico','JADLOG Econômico'],
        'DIMENSÕES CAIXA (altura x largura x comprimento cm)':['36 x 28 x 28','24 x 16 x 8','36 x 28 x 28'],
        'PESO DO PRODUTO':["29","15","11"],
        'CEP': ['41940000','70310500','25569900'],
        'VALOR DO PEDIDO':['872,5','295,3','945,4'],
        'VALOR COTAÇÃO JADLOG':[None,None,None]
    }
    
    df = pd.DataFrame(dados)
    logger = IntegratedLogger(maestro=maestro,filepath=BASE_LOG_PATH,activity_label=ACTIVITY_LABEL)
    df_copy, sucess_items = catchJadlogPrice(bot=bot, maestro=maestro, df=df,logger=logger)
    
    
    
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
