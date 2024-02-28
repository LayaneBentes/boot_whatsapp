import openpyxl 
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui




webbrowser.open('https://web.whatsapp.com/')
sleep(20)


planilha = openpyxl.load_workbook('teste.xlsx') #ler planilha
alunos = planilha['Planilha1'] #planilha onde se encontra os dados

for linha in alunos.iter_rows(min_row=2):

    nome = linha[0].value
    contato = linha[1].value

    mensagem = f'Olá {nome} selacione sua opção:\n1-nota da redação\n2-Nota subjetiva'
    
    sleep(5)
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={contato}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        botao_seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(botao_seta[0], botao_seta[1])
        sleep(10)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f"não foi possivel mandar a mensagem.")
        with open("erros.csv","a", newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{contato}')

# input('')

