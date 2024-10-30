# ler dados da planilha 
# inserir cada célula de cada linha em um campo do sistema

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
paginas_de_vendas = workbook['vendas']

for linha in paginas_de_vendas.iter_rows(min_row=2):
    # nome
    pyautogui.click(881,499,duration=1.5)
    pyautogui.write(linha[0].value)
    # produto
    pyautogui.click(888,526,duration=1.5)
    pyautogui.write(linha[1].value)
    # quantidade
    pyautogui.click(884,551,duration=1.5)
    pyautogui.write(str(linha[2].value))
    # categoria
    pyautogui.click(946,578,duration=1.5)
    pyautogui.write(linha[3].value)
    # salvar
    pyautogui.click(815,606,duration=1.5)
    # ok
    pyautogui.click(818,567,duration=1.5)
    