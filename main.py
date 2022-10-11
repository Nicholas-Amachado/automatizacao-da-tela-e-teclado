import time

import pyautogui
import pyperclip
import pandas as pd

pyautogui.PAUSE = 2
pyautogui.alert('irá começar, por favor NÃO mexa no mouse e nem no teclado!!')
time.sleep(3)


"""""""""
pyautogui.click(x=919, y=1061)
pyautogui.click(x=1616, y=182)
pyautogui.write('wallpaper engine')
pyautogui.click(x=1754, y=241, clicks=2)
pyautogui.click(x=1616, y=182)
pyautogui.write("wallpaper64.exe")
pyautogui.click(x=1754, y=241)
pyautogui.click(x=1670, y=29)
pyautogui.click(x=1950, y=10)
"""""

# Agora vai começar a parte web da automatização
pyautogui.PAUSE = 2
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
time.sleep(2)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(1)
pyautogui.click(x=406, y=383, clicks=2)
time.sleep(2)
pyautogui.click(x=389, y=503)
pyautogui.click(x=1675, y=239)
pyautogui.click(x=1327, y=755)
time.sleep(3)


tabela = pd.read_excel(r'C:\Users\nicho\OneDrive\Documentos\Vendas - Dez.xlsx')
faturamento = tabela["Valor Final"].sum()
print(f'o valor do faturamento total é de {faturamento}.')
quantidade = tabela["Quantidade"].sum()
print(f'total de quantides vendidas é {quantidade}')
VU = tabela["Valor Unitário"].sum()
print(f'o valor unitário é de {VU}')


pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.press("enter")
pyautogui.click(x=77, y=248)
time.sleep(2)
pyautogui.write("nicholasalexandre100@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write('Relatorio de Vendas')
pyautogui.press("tab")
corpo = f"""Prezados,

A seguir o relatório de Vendas Atualizado
Somado o valor da Quntidade de produtos Vendidos,
o Valor Unitário e o total de vendas.
faturamneto das vendas de R${faturamento:,.2f} .
Quantidade de Vendas R${quantidade:,} .
Valor de unitario de R${VU:,.2f} .

Obrigado pela Atencao
Qualquer Duvida, Entrar em contato
"""


pyautogui.write(corpo)
pyautogui.hotkey('ctrl', 'enter')




