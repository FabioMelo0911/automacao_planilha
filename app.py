#ler dados da planilha
#inserir cada celula da planilha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
# Seleciona a aba 'Vendas'  
worksheet = workbook['vendas']

# Define os cabeçalhos das colunas
for linha in worksheet.iter_rows(min_row=2):
    #cliente
    pyautogui.click(862,54,duration=1.5)
    pyautogui.typewrite(linha[0].value)
    #produto
    pyautogui.click(860,80,duration=1.5)
    pyautogui.typewrite(linha[1].value)
    #quantidade
    pyautogui.click(881,106,duration=1.5)
    pyautogui.typewrite(str(linha[2].value))
    #categoria
    pyautogui.click(941,133,duration=1.5)
    pyautogui.typewrite(linha[3].value)
    #salvar
    pyautogui.click(822,161,duration=1.5)
    pyautogui.typewrite(linha[3].value)
    print(linha[0].value)
    print(linha[1].value)
    print(linha[2].value)
    print(linha[3].value)