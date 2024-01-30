" BOT WHATSAPP AUTOMESSAGER SHOOTER V1.0 "

# Libraries
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

# Start WhatsApp Web on browser
webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Load excel spreadsheet and data
workbook = openpyxl.load_workbook('contacts.xlsx')
page_contacts = workbook['Sheet1']

for linha in page_contacts.iter_rows(min_row=2):
    
    name = linha[0].value
    number = linha[1].value    

    # Edit message as needed
    # Example 01
    message = f'Olá {name}, tudo bem? lore epsium lore epsium link https://www.loreepsium.com'

    # Create personalized WhatsApp links and send messages to each customer based on data
    try:
        link_whatsapp_message = f'https://web.whatsapp.com/send?phone={number}&text={quote(message)}'
        webbrowser.open(link_whatsapp_message)
        sleep(10)
        arrow = pyautogui.locateCenterOnScreen('arrow.png')
        sleep(2)
        pyautogui.click(arrow[0],arrow[1])
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    # Errors
    except:
        print(f'Não foi possível enviar menssagem para {name}')
        with open('errors.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{name},{number}{os.linesep}')
    
