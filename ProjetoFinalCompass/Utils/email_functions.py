import smtplib
from email.message import EmailMessage
from datetime import datetime
import openpyxl
import os
from botcity.maestro import *
from config import vars_map


# Global settings
EMAIL = vars_map['EMAIL_USERNAME']
PROCESS_NAME = "E-mail de finalização da execução"
EXCEL_FILE_PATH = vars_map['DEFAULT_EMAILS_FILE']

def get_current_timestamp():
    """ Retorna a data e hora atual formatadas. """
    now = datetime.now()
    return now.strftime("%d%m%Y"), now.strftime("%H%M")
        
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
        print(f"ERROR - read_emails_from_excel")
        return []

def send_emails(output_file):
    """ Sends emails according to the list extracted from Excel """
    formatted_date, formatted_time = get_current_timestamp()
    email_password = vars_map['EMAIL_PASSWORD']

    if not email_password:
        return

    email_body = (f"O processo RPA {PROCESS_NAME} foi executado com sucesso "
                  f"na data {formatted_date} - {formatted_time}")
    
    try:
        with open(output_file, "rb") as content_file:
            content = content_file.read()
    except FileNotFoundError:
        print("ERROR - Anexo planilha de saída")
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
                print(f"E-mail enviado para: {email}")

        print("Todos os e-mails foram enviados com sucesso!")
    
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar e-mails")

# Send error notification email
def send_error_email(process_name, error_message, screenshot_path):
    formatted_date = datetime.now().strftime("%d%m%Y")
    formatted_time = datetime.now().strftime("%H%M")
    email_password = vars_map['EMAIL_PASSWORD']
    if not email_password:
        return

    subject = f"Erro - RPA [{process_name}] {formatted_date} ⏰ {formatted_time}"
    body = (f"Foi encontrado ERRO durante a execução do processo RPA [PROCESS NAME], "
            f"na data {formatted_date} às {formatted_time}, na tarefa {process_name}.\n\n"
            f"Detalhes do erro:\n{error_message}")
    emails = read_emails_from_excel()

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL, email_password)
            for email in emails:
                msg = EmailMessage()
                msg['Subject'] = subject
                msg['From'] = EMAIL
                msg['To'] = email
                msg.set_content(body)

                if screenshot_path:
                    with open(screenshot_path, "rb") as img:
                        msg.add_attachment(img.read(), maintype='image', subtype='png',
                                           filename=os.path.basename(screenshot_path))

                smtp.send_message(msg)
                print(f"INFO - E-mail de erro enviado para: {email}")

    except Exception as e:
        raise(f"Erro ao enviar e-mail de notificação: {e}")
