# Defining constants
EMAIL = 'exemplo.compass@gmail.com'
PASSWORD = r'C:\RPA\bot-mail\password.txt'
PROCESS_NAME = "E-mail de finalização da execução"
EXCEL_FILE_PATH = r"C:\RPA\bot-mail\rede\E-mails para Notificações - RPA.xlsx"
EXIT_SHEET = r"C:\RPA\bot-mail\rede\PlanilhaSaida\cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx"

import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
import openpyxl

# Get current date and time
now = datetime.now()
print(f"Data e hora atual: {now}")

# Format the date and time in the desired format
formatted_date = now.strftime("%d%m%Y")
formatted_time = now.strftime("%H%M")
print(f"Data formatada: {formatted_date}")
print(f"Hora formatada: {formatted_time}")

with open(PASSWORD) as p: 
    password = p.readlines()
    p.close()

print(f"Conteúdo do arquivo de senha: {password}")

email_password = password[0]
print(f"Senha do e-mail: {email_password}")

subject = f"RPA {PROCESS_NAME} - {formatted_date} - {formatted_time}"
print(f"Assunto do e-mail: {subject}")

# Opening the Excel file
wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
sheet = wb.active
print(f"Arquivo Excel aberto: {EXCEL_FILE_PATH}")

# List to store emails
emails = []

# Iterating over all cells in column A
for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
    for cell in row:
        email = cell.value
        if email:
            emails.append(email)

print(f"E-mails coletados: {emails}")

email_body = f"O processo RPA {PROCESS_NAME} foi executado com sucesso na data {formatted_date} - {formatted_time}"
print(f"Corpo do e-mail: {email_body}")

with open(EXIT_SHEET, "rb") as content_file:
    content = content_file.read()

# Cria um loop for para enviar um email por vez
for email in emails:
    msg = EmailMessage()
    msg['subject'] = subject
    msg['From'] = EMAIL
    msg['To'] = email
    msg.set_content(email_body)
    msg.add_attachment(content, maintype='application', subtype='xlsx', filename='cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx')
    print(f"Enviando e-mail para: {email}")

    # Modifica a porta de comunicação, adiciona uma etapa de segurança a conexão
    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.set_debuglevel(1)
        smtp.starttls()
        smtp.local_hostname = "localhost"
        smtp.login(EMAIL, email_password)
        smtp.send_message(msg)
        print("E-mail enviado com sucesso!")
