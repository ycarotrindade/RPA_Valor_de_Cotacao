import smtplib
from email.message import EmailMessage
from datetime import datetime
import openpyxl
import sys
import os
import pyautogui
from botcity.maestro import *
# from botcity.web import BotWeb
# from botcity.web.browsers.chrome import default_options


# Global settings
EMAIL = 'exemplo.compass@gmail.com'
PASSWORD_FILE = r"C:\RPA\bot-mail-2\bot-mail\password.txt"
PROCESS_NAME = "E-mail de finalização da execução"
EXCEL_FILE_PATH = r"C:\RPA\bot-mail-2\bot-mail\rede\E-mails para Notificações - RPA.xlsx"
EXIT_SHEET = r"C:\RPA\bot-mail-2\bot-mail\rede\PlanilhaSaida\cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx"
LOG_DIR = r"C:\RPA\bot-mail-2\bot-mail\rede\Logs"

# Run the script with an environment configured for UTF-8
sys.stdout.reconfigure(encoding='utf-8') 

# Make sure the log folder exists
os.makedirs(LOG_DIR, exist_ok=True)

# Maestro SDK settingd
BotMaestroSDK.RAISE_NOT_CONNECTED = False

# Screenshot with PyAutoGui
def capture_screenshot(filename="error_screenshot.png"):
    """ Capture a screenshot and save it in the logs directory """
    try:
        # Add timestamp to the filename to prevent overwriting
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_filename = f"{timestamp}_{filename}"
        screenshot_path = os.path.join(LOG_DIR, screenshot_filename)

        # Capture the screenshot
        screenshot = pyautogui.screenshot()
        screenshot.save(screenshot_path)

        print(f"INFO - Screenshot salva em: {screenshot_path}")
    except Exception as e:
        print(f"ERROR - Erro ao capturar screenshot: {e}")
        capture_screenshot()  # Capture the screenshot if an error occurs
        return []

# Screenshot with With BotWeb
# def capture_screenshot(filename="error_screenshot.png"): 
#     """ Capture a screenshot and save it in the logs directory """
#     try:
#         bot = BotWeb()
#         bot.configure(default_options(headless=True))
#         bot.start_browser()
#         screenshot_path = os.path.join(LOG_DIR, filename)
#         bot.screenshot(screenshot_path)
#         bot.stop_browser()
#         print(f"INFO - Screenshot salva em: {screenshot_path}")
#     except Exception as e:
#         print(f"ERROR - Erro ao capturar screenshot: {e}")

def get_current_timestamp():
    """ Retorna a data e hora atual formatadas. """
    now = datetime.now()
    return now.strftime("%d%m%Y"), now.strftime("%H%M")

def read_password():
    """Reads the password from the file """
    try:
        with open(PASSWORD_FILE) as file:
            return file.readline().strip()
    except FileNotFoundError:
        print("ERROR - Arquivo de senha não encontrado.")
        capture_screenshot()  # Capture the screenshot if an error occurs
        return []
        

def read_emails_from_excel():
    """ Reads the emails from the spreadsheet and returns a list. """
    try:
        workbook = openpyxl.load_workbook(EXCEL_FILE_PATH)
        sheet = workbook.active
        emails = []

        for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
            for cell in row:
                if cell.value:
                    emails.append(cell.value)

        return emails
    except Exception as e:
        print(f"ERROR - Erro ao ler emails do Excel: {e}")
        capture_screenshot()
        return []

def send_emails():
    """ Sends emails according to the list extracted from Excel """
    formatted_date, formatted_time = get_current_timestamp()
    email_password = read_password()

    if not email_password:
        return

    email_body = (f"O processo RPA {PROCESS_NAME} foi executado com sucesso "
                  f"na data {formatted_date} - {formatted_time}")
    
    try:
        with open(EXIT_SHEET, "rb") as content_file:
            content = content_file.read()
    except FileNotFoundError:
        print("ERROR - Planilha de saída não encontrada.")
        capture_screenshot()
        return

    emails = read_emails_from_excel()

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            # smtp.set_debuglevel(1)
            smtp.ehlo('localhost') # Set a compatible hostname
            smtp.login(EMAIL, email_password)

            for email in emails:
                msg = EmailMessage()
                msg['Subject'] = f"RPA {PROCESS_NAME} - {formatted_date} - {formatted_time}"
                msg['From'] = EMAIL
                msg['To'] = email
                msg.set_content(email_body)
                msg.add_attachment(content, maintype='application', subtype='xlsx', filename='resultado.xlsx')

                smtp.send_message(msg)
                print(f"INFO - E-mail enviado para: {email}")

        print("INFO - Todos os e-mails foram enviados com sucesso!")
    
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar e-mails: {e}")
        capture_screenshot()

def main():
    """ Runs the main flow """
    try:
        maestro = BotMaestroSDK.from_sys_args()
        execution = maestro.get_execution()

        print(f"Task ID: {execution.task_id}")
        print(f"Parâmetros: {execution.parameters}")

        send_emails()
    
    except Exception as e:
        print(f"Erro inesperado: {e}")
        capture_screenshot()

if __name__ == '__main__':
    main()
