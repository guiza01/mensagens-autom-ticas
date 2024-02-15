import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

sleep(5)

workbook = openpyxl.load_workbook('clientes.xlsx')
paginaCliente = workbook['Planilha1']

for linha in paginaCliente.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá, {nome} seu boleto vence dia {vencimento.strftime("%d/%m/%Y")}. Por favor, me pague!! Meu pix é (87) 99665-9985...'
    

    # padrão de link https://web.whatsapp.com/send?phone=&text

    try:
        linkMensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(linkMensagem)
        sleep(7)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(1.5)
        pyautogui.click(seta[0], seta[1])
        sleep(1.5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    except:
        print(f'Não foi possivel enviar a mensagem para a pessoa {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)