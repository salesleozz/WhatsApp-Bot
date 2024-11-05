import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

workbook = openpyxl.load_workbook('Contatos.xlsx') #nome do arquivo
pagina_clientes = workbook['Telefones'] #nome da planilha

for linha in pagina_clientes.iter_rows(min_row=1):
    telefone = linha[0].value
    
    mensagem = f'Boa dia, gostaria de saber como funciona serviço de voces' #mensagem que vai ser enviada para o telefone

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        pyautogui.press('enter')
        sleep(10)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {telefone}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{telefone}{os.linesep}')