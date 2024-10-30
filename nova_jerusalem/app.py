import openpyxl
import pyperclip
import pyautogui
#PASSOS

#1- entrar na planilha

workbook = openpyxl.load_workbook('nova_jerusalem.xlsx')
pagina_inicial = workbook['Planilha1']


#2- clicar na informacao necessaria
#3- copiar a informacao

for linha in pagina_inicial.iter_rows(min_row=2):
    numero_crianca = linha[0].value
    pyperclip.copy(numero_crianca)
    pyautogui.click(1550, 492, duration = 1, clicks = 2)
    pyautogui.hotkey('ctrl','v')

    #nome = linha[1].value
    #sexo = linha[2].value
    #idade = linha[3].value
    #roupa = linha[4].value
    #calcado = linha[5].value


#4- clicar 2 vezes no texto do canva
#5- colar o a informacao copiada do excel, no canva
#6- clicar em duplicar pagina
#7- repetir ate 198

# PyAutoGui(automacao de clicks e teclado)
# Openpyxl (leitura e automacao do excel)



