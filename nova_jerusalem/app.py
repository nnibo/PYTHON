import openpyxl
import pyperclip
import pyautogui
#PASSOS

#1- entrar na planilha

workbook = openpyxl.load_workbook('nova_jerusalem.xlsx')
pagina_inicial = workbook['Planilha1']


#2- clicar na informacao necessaria
#3- copiar a informacao
#4- clicar 2 vezes no texto do canva
#5- colar o a informacao copiada do excel, no canva

for linha in pagina_inicial.iter_rows(min_row=2):
    #PRIMEIRA ETIQUETA
    numero_crianca1 = linha[0].value
    pyperclip.copy(numero_crianca1)
    pyautogui.click(842,285, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    nome1 = linha[1].value
    pyperclip.copy(nome1)
    pyautogui.click(909,348, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    idade1 = linha[3].value
    pyperclip.copy(idade1)
    pyautogui.click(834,446, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    roupa1 = linha[4].value
    pyperclip.copy(roupa1)
    pyautogui.click(909,447, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    calcado1 = linha[5].value
    pyperclip.copy(calcado1)
    pyautogui.click(983,447, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    # SEGUNDA ETIQUETA
    numero_crianca2 = linha[0].value
    pyperclip.copy(numero_crianca2)
    pyautogui.click(840,614, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    nome2 = linha[1].value
    pyperclip.copy(nome2)
    pyautogui.click(913,680, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    idade2 = linha[3].value
    pyperclip.copy(idade2)
    pyautogui.click(833,777, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    roupa2 = linha[4].value
    pyperclip.copy(roupa2)
    pyautogui.click(905,784, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    calcado2 = linha[5].value
    pyperclip.copy(calcado2)
    pyautogui.click(981,781, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    

#6- clicar em duplicar pagina
    pyautogui.click(1133,552, duration = 1)

#7- repetir ate 198

# PyAutoGui(automacao de clicks e teclado)
# Openpyxl (leitura e automacao do excel)
# pip install mouseinfo
# python from mouseinfo import mouseInfo
# mouseInfo() para abrir a ferramenta da posicao do mouse



