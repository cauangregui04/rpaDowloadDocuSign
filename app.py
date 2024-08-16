from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
import pyautogui
import pandas as pd
import os

global o

i = 0
o = 0

carga = pd.read_excel("cargaDocu.xlsx", dtype=str)

servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
prefs = {'download.prompt_for_download': True,
         'download.directory_upgrade': True,
         'plugins.always_open_pdf_externally': False
        }
options.add_experimental_option('prefs', prefs)
navegador = webdriver.Chrome(service=servico, options=options)

navegador.get("https://account.docusign.com/")

navegador.maximize_window()
time.sleep(2)
navegador.find_element('xpath',
                       '/html/body/div/div/div[2]/div/main/div/div/div[2]/form/div[1]/div[2]/div/input').send_keys(
    "email")
navegador.find_element('xpath', '/html/body/div/div/div[2]/div/main/div/div/div[2]/form/div[2]/button').click()
time.sleep(2)
navegador.find_element('xpath',
                       '/html/body/div/div/div[2]/div/main/div/div/div[2]/form/div[1]/div[2]/div/input').send_keys(
    "senha")
navegador.find_element('xpath', '/html/body/div/div/div[2]/div/main/div/div/div[2]/form/div[2]/button').click()
time.sleep(80)

def docuRPA():
    global o
    ID = str(carga.iloc[i, 0])
    navegador.get(f"https://apps.docusign.com/api/send/api/accounts/518dfdcb-512a-479a-a818-f1c9c0b24beb/envelopes/{ID}/documents/combined?escape_non_ascii_filenames=true&language=pt_BR&certificate=true&shared_user_id=e9cfd55c-506d-47e0-a912-193a180a872a")
    time.sleep(9)
    pyautogui.write(f"{ID}.pdf")
    time.sleep(3)

    pyautogui.press("enter")
    time.sleep(9)
    print(f"Arquivo {i + 1} de {len(carga)} baixado!")
    if o == 30:
        navegador.refresh()
        o = 0
    else:
        pass

while i < len(carga):
    docuRPA()
    i += 1
    o += 1

input("========Processo Concluído========")
