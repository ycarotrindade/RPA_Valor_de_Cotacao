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
https://documentation.botcity.dev/tutorials/python-automations/web/
"""

# Defining constants
CHROMEDRIVER_PATH = r"resources\chromedriver-win64\chromedriver.exe"
EXCEL_FILE_PATH = r"C:\RPA\bot-email\resources\rede\E-mails para Notificações - RPA.xlsx"
PROCESS_NAME = "E-mail de finalização da execução"
ATTACHMENT_PATH = r"C:\RPA\bot-email\resources\cnpjs_dd-mm-aaaa_hh-mm-ss.xlsx"
USERNAME = "tamirys.silva.pb@compasso.com.br"

# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Import getpass to type email password
import getpass

# Import openpyxl to open the Excel file
import openpyxl

# Import datetime to dynamically insert dates
from datetime import datetime

# not_found function declared
def not_found(label):
    print(f"Element not found: {label}")

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

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = CHROMEDRIVER_PATH

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Opens the Outlook website.
    bot.browse("https://outlook.office.com/mail/")

    # Searching for element 'username'
    if not bot.find("username", matching=0.97, waiting_time=10000):
        not_found("username")
    bot.click_relative(150, 14)

    bot.paste(USERNAME)

    # Searching for element 'username-email-next'
    if not bot.find("username-email-next", matching=0.97, waiting_time=10000):
        not_found("username-email-next")
    bot.click()

    # Searching for element 'password'
    if not bot.find("password", matching=0.97, waiting_time=10000):
        not_found("password")
    bot.click_relative(104, 13)

    password = getpass.getpass(prompt="Enter your password: ")

    bot.kb_type(password)

    # Searching for element 'enter-email'
    if not bot.find("enter-email", matching=0.97, waiting_time=10000):
        not_found("enter-email")
    bot.click()

    # Pause for Authenticator
    bot.wait(10000)

    # Searching and selecting for element 'not-show-message-again'
    if not bot.find("not-show-message-again", matching=0.97, waiting_time=10000):
        not_found("not-show-message-again")
    bot.click_relative(30, 33)

    # Searching for element 'yes-message' after Authenticator
    if not bot.find("yes-message", matching=0.97, waiting_time=10000):
        not_found("yes-message")
    bot.click()

    # Searching for element 'new-email' to begin to write a new e-mail
    if not bot.find("new-email", matching=0.97, waiting_time=10000):
        not_found("new-email")
    bot.click()

    # Opening the Excel file
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb.active

    # Searching for element 'recipients'
    if not bot.find("recipients", matching=0.97, waiting_time=10000):
        not_found("recipients")
    bot.click_relative(105, 26)

    # Pasting receipients
    # Iterating over all cells in column A (A1, A2, A3,... until the last filled cell)
    for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
        for cell in row:
            email = cell.value
            if email:
                #print(f"Inserindo e-mail: {email}")

                # Paste the email into the "recipients" field
                bot.paste(email)

                # After each email, press the TAB key to separate the addresses
                bot.kb_type(";")

    # Get current date and time
    now = datetime.now()

    # Format the date and time in the desired format
    formatted_date = now.strftime("%d%m%Y")
    formatted_time = now.strftime("%H%M")

    # Searching for element 'add-subject'
    if not bot.find("add-subject", matching=0.97, waiting_time=10000):
        not_found("add-subject")
    bot.click_relative(266, 17)

    # Email subject
    subject = f"RPA {PROCESS_NAME} - {formatted_date} - {formatted_time}"
    bot.paste(subject)

    # Pressing TAB to enter the email body field
    bot.kb_type("\t")

    # Email body
    email_body = f"O processo RPA {PROCESS_NAME} foi executado com sucesso na data {formatted_date} - {formatted_time}"
    bot.paste(email_body)

    # Attaching the spreadsheet
    # Searching for element 'insert-att'
    if not bot.find("insert-att", matching=0.97, waiting_time=10000):
        not_found("insert-att")
    bot.click()

    # Searching for element 'attach'
    if not bot.find("attach", matching=0.97, waiting_time=10000):
        not_found("attach")
    bot.click_relative(167, 15)

    # Searching for element 'attach-2'
    if not bot.find("attach-2", matching=0.97, waiting_time=10000):
        not_found("attach-2")
    bot.click()
    bot.wait(4000)
    
    # bot.kb_type(ATTACHMENT_PATH)
    # bot.kb_type("\n")

    input() # Somente para pausar o programa até fazer um input

    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )


if __name__ == '__main__':
    main()
