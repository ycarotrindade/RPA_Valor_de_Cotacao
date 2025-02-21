"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/desktop/
"""
# Defining constants
EMAIL = 'exemplo.compass@gmail.com'
PASSWORD = 'password.txt'
PROCESS_NAME = "E-mail de finalização da execução"
EXCEL_FILE_PATH = r"C:\RPA\bot-mail-2\bot-mail\rede\E-mails para Notificações - RPA.xlsx"
EXIT_SHEET = r"C:\RPA\bot-mail-2\bot-mail\rede\PlanilhaSaida\cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx"

import os
import smtplib
from email.message import EmailMessage
from datetime import datetime
import openpyxl

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *


# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    # Get current date and time
    now = datetime.now()

    # Format the date and time in the desired format
    formatted_date = now.strftime("%d%m%Y")
    formatted_time = now.strftime("%H%M")

    with open (PASSWORD) as p: 
        password = p.readlines()
        p.close()

    email_password = password[0]
    email_password

    msg = EmailMessage()

    subject = f"RPA {PROCESS_NAME} - {formatted_date} - {formatted_time}"
    msg['subject'] = subject

    msg['From'] = EMAIL

    # Opening the Excel file
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb.active

    # List to store emails
    emails = []

    # Iterating over all cells in column A
    for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
        for cell in row:
            email = cell.value
            if email:
                emails.append(email)
                
    emails_str = ";".join(emails)   
    msg['To'] = emails_str

    email_body = f"O processo RPA {PROCESS_NAME} foi executado com sucesso na data {formatted_date} - {formatted_time}"
    msg.set_content(email_body)

    with open (EXIT_SHEET, "rb") as content_file:
        content = content_file.read()   
        msg.add_attachment(content, maintype='application', subtype='xlsx', filename='cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.set_debuglevel(1)
        smtp.local_hostname = "localhost"
        smtp.login(EMAIL, email_password)
        smtp.send_message(msg)
    
    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()