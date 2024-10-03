"""


AUTOMAÇÃO DE MENSAGENS NO WHATSAPP - BOT ENVIA MENSAGENS COM VENCIMENTO DE BOLETOS

"""

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import pywhatkit as kit

webbrowser.open('https://web.whatsapp.com/')
sleep(10)

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

# Iterando sob a planilha e os contatos
for linha in pagina_clientes.iter_rows(min_row=2):
    
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá, {nome} , sou Ricardo, assistente virtual da RS Seguros, venho aqui dizer que seu boleto com vencimento em {vencimento} está pronto. Segue abaixo o anexo, para pagamento. Atenciosamente, RS Seguros'
 

    try:
        link_msg_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_msg_wpp)
        sleep(10)
        # Simular o envio do anexo
        #pyautogui.hotkey('ctrl', 'o')  # Abre o diálogo de abrir arquivo
        #sleep(2)  # Espera o diálogo abrir
        # Selecionar a barra de endereços
        
        pyautogui.click(x=1361, y=715) 
        sleep(2)
        pyautogui.click(x=1407, y=467)  
        sleep(2)
        pyautogui.hotkey('ctrl', 'l')  # Foca na barra de endereços
        sleep(2)
        # Digitar o caminho do arquivo
        pyautogui.doubleClick(x=962, y=392)  # Ajuste as coordenadas para clicar no ícone de anexo
        sleep(2)
        pyautogui.press('enter')  # Pressiona Enter para anexar o arquivo
        sleep(2)

        # Pressionar Enter novamente para enviar o anexo
        pyautogui.press('enter')
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
 