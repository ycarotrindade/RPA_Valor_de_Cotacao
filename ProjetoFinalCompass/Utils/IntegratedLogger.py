from botcity import maestro
import logging
import os
from datetime import datetime
from botcity.web import WebBot
import traceback
import sys
from PIL import ImageGrab
from .email_functions import send_error_email

class IntegratedLogger:
    '''Classe que integra logs locais e logs do botcity.
    
    # Atributos
    
        * maestro: `BotMaestroSDK`    
            Objeto maestro criado pelo botcity, caso seja None o programa só opera com logs locais.
            
        * filepath: `PathLike`
            Caminho base onde os logs serão armazenados.
        
        * image_filepath: `PathLike`
            Caminho onde serão armazenadas as imagens dos erros, por padrão é o arquivo raiz de logs.
        
        * activity_label: `str`
            Label usado pelo Maestro para identificar qual a automação que está operando.
        
        * dev_logger: `Logger`
            Logger usado para criar devlogs (logs específicos para o desenvolver), esse log está em nível de debug e pega todas as linhas de erros.
        
        * client_logger: `Logger`
            Logger usado para criar clientlogs (logs específicos para clientes), esse log está em nível de info e pega apenas a última linha do erro.
        
        * datetime_format: `str`
            Formato de data e hora usado dentro dos logs.
        
        * datetime_file_format: `str`
            Formato de data e hora usado no nome de arquivos, já que alguns caracteres não são permitidos na hora de criar arquivos.
        

    
    '''
    def __init__(self,maestro:maestro.BotMaestroSDK, filepath:os.PathLike, activity_label:str):
        self.maestro = maestro
        self.filepath = filepath
        self.image_filepath = filepath
        self.activity_label = activity_label
        self.dev_logger = logging.getLogger('dev_logger')
        self.client_logger = logging.getLogger('client_logger')
        self.datetime_format = '(%d-%m-%Y_%H:%M:%S)'
        self.datetime_file_format = '%d-%m-%Y_%H-%M-%S'
        self.__inital_configs()
        
    
    def __inital_configs(self):
        
        '''Configurações de loggers, não deve ser chamada fora da classe'''
        
        self.filepath = os.path.join(self.filepath,datetime.now().strftime('%d-%m-%Y'))
        self.image_filepath = os.path.join(self.filepath,'Errors')
        self.filepath = os.path.join(self.filepath,'Files')
        os.makedirs(self.filepath,exist_ok=True)
        os.makedirs(self.image_filepath,exist_ok=True)
        
        self.dev_logger.setLevel(logging.DEBUG)
        dev_logger_stream_handler = logging.StreamHandler()
        dev_logger_stream_handler.setLevel(logging.DEBUG)
        dev_logger_stream_handler.setFormatter(logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',datefmt=self.datetime_format))
        dev_logger_file_handler = logging.FileHandler(filename=os.path.join(self.filepath,f'devlog_{datetime.now().strftime(self.datetime_file_format)}.log'),mode='a',encoding='utf-8')
        dev_logger_file_handler.setLevel(logging.DEBUG)
        dev_logger_file_handler.setFormatter(logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',datefmt=self.datetime_format))
        self.dev_logger.handlers = [dev_logger_stream_handler,dev_logger_file_handler]
        
        self.client_logger.setLevel(logging.INFO)
        client_logger_file_handler = logging.FileHandler(filename=os.path.join(self.filepath,f'log_{datetime.now().strftime(self.datetime_file_format)}.log'),mode='a',encoding='utf-8')
        client_logger_file_handler.setLevel(logging.INFO)
        client_logger_file_handler.setFormatter(logging.Formatter(fmt='%(asctime)s - %(levelname)s - %(message)s',datefmt=self.datetime_format))
        self.client_logger.addHandler(client_logger_file_handler)
        
        

    def info(self,msg:str):
        '''Insere uma mensagem de level INFO no log.
        
        # Parâmetros
        
            * msg: `str`
                Mensagem a ser inserida no log.
        
        '''
        msg_list = msg.splitlines()
        list(map(lambda message:self.dev_logger.info(message),msg_list))
        list(map(lambda message:self.client_logger.info(message),msg_list))
        if self.maestro is not None:
            self.maestro.new_log_entry(
                activity_label=self.activity_label,
                values={
                    'Datetime':datetime.now().strftime(self.datetime_format),
                    'Level':'INFO',
                    'Message':msg_list[-1]
                }
            )

    def debug(self,msg:str):
        '''Insere uma mensagem de level DEBUG no log.
        
        # Parâmetros
        
            * msg: `str`
                Mensagem a ser inserida no log.
        
        '''
        msg_list = msg.splitlines()
        list(map(lambda message:self.dev_logger.debug(message),msg_list))
        if self.maestro is not None:
            self.maestro.new_log_entry(
                activity_label=self.activity_label,
                values={
                    'Datetime':datetime.now().strftime(self.datetime_format),
                    'Level':'DEBUG',
                    'Message':msg_list[-1]
                }
            )

    def warning(self,process_name:str):
        '''Insere uma mensagem de level WARNING no log e captura a tela, a mensagen é capturada automaticamente.
        
        # Parâmetros
            
            * process_name: `str`
                Nome do processo que estava sendo executado antes do erro.
            
        
        '''
        msg_list = traceback.format_exc().splitlines()
        etype, value, _ = sys.exc_info()
        msg_reduced = traceback.format_exception_only(etype,value)
        list(map(lambda message:self.dev_logger.warning(message),msg_list))
        self.client_logger.warning(msg_reduced)
        image_filepath = os.path.join(self.image_filepath,f'{datetime.now().strftime(self.datetime_file_format)}_RPA_{process_name}.jpg')
        ImageGrab.grab().save(image_filepath)
        send_error_email(process_name,msg_reduced,image_filepath)
        if self.maestro is not None:
            self.maestro.new_log_entry(
                activity_label=self.activity_label,
                values={
                    'Datetime':datetime.now().strftime(self.datetime_format),
                    'Level':'WARNING',
                    'Message':msg_reduced
                }
            )
            self.maestro.error(
                task_id=self.maestro.get_execution().task_id,
                exception=Exception(msg_reduced),
                screenshot=image_filepath
            )
        

    def error(self,process_name:str):
        '''Insere uma mensagem de level ERROR no log e captura a tela, a mensagem é capturada automaticamente.
        
        # Parâmetros
            
            * process_name: `str`
                Nome do processo que estava sendo executado antes do erro.
        
        '''
        msg_list = traceback.format_exc().splitlines()
        etype, value, _ = sys.exc_info()
        msg_reduced = traceback.format_exception_only(etype,value)
        list(map(lambda message:self.dev_logger.error(message),msg_list))
        self.client_logger.error(msg_reduced)
        image_filepath = os.path.join(self.image_filepath,f'{datetime.now().strftime(self.datetime_file_format)}_RPA_{process_name}.jpg')
        ImageGrab.grab().save(image_filepath)
        send_error_email(process_name,msg_reduced,image_filepath)
        if self.maestro is not None:
            self.maestro.new_log_entry(
                activity_label=self.activity_label,
                values={
                    'Datetime':datetime.now().strftime(self.datetime_format),
                    'Level':'INFO',
                    'Message':msg_reduced
                }
            )
            self.maestro.error(
                task_id=self.maestro.get_execution().task_id,
                exception=Exception(msg_reduced),
                screenshot=image_filepath
            )