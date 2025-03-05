import os
from botcity.web import WebBot, Browser, By
from botcity.maestro import *
from dotenv import load_dotenv

load_dotenv(override=True)

IS_MAESTRO_CONNECTED = eval(os.getenv('IS_MAESTRO_CONNECTED'))
if IS_MAESTRO_CONNECTED:
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    BASE_LOG_PATH = execution.parameters.get('BASE_LOG_PATH')
    DEFAULT_PROCESSAR_PATH = execution.parameters.get('DEFAULT_PROCESSAR_PATH')
    DEFAULT_PROCESSADOS_PATH = execution.parameters.get('DEFAULT_PROCESSADOS_PATH')
    DEFAULT_CORREIOS_URL = execution.parameters.get('DEFAULT_CORREIOS_URL')
    DEFAULT_BRASILAPI_URL = execution.parameters.get('DEFAULT_BRASILAPI_URL')
    ORIGIN_CEP = execution.parameters.get('ORIGIN_CEP')
    DEFAULT_URL_JADLOG = execution.parameters.get('DEFAULT_URL_JADLOG')
    PICKUP_VALUE = execution.parameters.get('PICKUP_VALUE')
else:
    maestro = None
    execution = None
    BASE_LOG_PATH = os.getenv('BASE_LOG_PATH')
    DEFAULT_PROCESSAR_PATH = os.getenv('DEFAULT_PROCESSAR_PATH')
    DEFAULT_PROCESSADOS_PATH = os.getenv('DEFAULT_PROCESSADOS_PATH')
    DEFAULT_CORREIOS_URL = os.getenv('DEFAULT_CORREIOS_URL')
    DEFAULT_BRASILAPI_URL = os.getenv('DEFAULT_BRASILAPI_URL')
    ORIGIN_CEP = os.getenv('ORIGIN_CEP')
    DEFAULT_URL_JADLOG = os.getenv('DEFAULT_URL_JADLOG')
    PICKUP_VALUE = os.getenv('PICKUP_VALUE')


vars_map = {
    'IS_MAESTRO_CONNECTED':IS_MAESTRO_CONNECTED,
    'ACTIVITY_LABEL':os.getenv('ACTIVITY_LABEL'),
    'BASE_LOG_PATH':BASE_LOG_PATH,
    'DEFAULT_PROCESSAR_PATH':DEFAULT_PROCESSAR_PATH,
    'DEFAULT_PROCESSADOS_PATH':DEFAULT_PROCESSADOS_PATH,
    'DEFAULT_CORREIOS_URL':DEFAULT_CORREIOS_URL,
    'DEFAULT_BRASILAPI_URL':DEFAULT_BRASILAPI_URL,
    'DEFAULT_MAESTRO':maestro,
    'DEFAULT_EXECUTION':execution,
    'ORIGIN_CEP':ORIGIN_CEP,
    'PICKUP_VALUE':PICKUP_VALUE,
    'DEFAULT_URL_JADLOG':DEFAULT_URL_JADLOG
}