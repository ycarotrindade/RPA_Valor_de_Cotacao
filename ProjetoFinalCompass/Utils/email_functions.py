import smtplib
from email.message import EmailMessage
from datetime import datetime
import openpyxl
import sys
import os
from botcity.maestro import *
from config import vars_map
from Utils.IntegratedLogger import IntegratedLogger


# Global settings
EMAIL = vars_map['EMAIL_USERNAME']
PROCESS_NAME = "E-mail de finalização da execução"
EXCEL_FILE_PATH = vars_map['DEFAULT_EMAILS_FILE']

def get_current_timestamp():
    """ Retorna a data e hora atual formatadas. """
    now = datetime.now()
    return now.strftime("%d%m%Y"), now.strftime("%H%M")
        
def read_emails_from_excel(logger:IntegratedLogger):
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
        logger.error(f"read_emails_from_excel")
        return []

def send_emails(output_file,logger:IntegratedLogger):
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
        logger.error(process_name="Anexo planilha de saída")
        return

    emails = read_emails_from_excel(logger)

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
                logger.info(f"E-mail enviado para: {email}")

        logger.info("Todos os e-mails foram enviados com sucesso!")
    
    except smtplib.SMTPException as e:
        logger.error(f"Erro ao enviar e-mails")
