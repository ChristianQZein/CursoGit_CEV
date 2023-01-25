# Este projeto tem como foco a abertura de um drive do google utilizando a biblioteca pyautogui e pyperclip, download do arquivo csv, abertura do arquivo com pandas, tratamento da tabela e encaminhamento via g-mail de um e-mail para a diretoria com os dados do faturamento obitido.

# O código foi escrito no Jupyter Notebook

pip install pyautogui
pip install pyperclip

import pyautogui as pa
import pyperclip as pc
import pandas as pd
import time

# Passo 1 -> Abrir o Drive

pa.PAUSE = 1 

pa.hotkey("ctrl", "t")
pc.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pa.hotkey("ctrl", "v")
pa.press("enter")

# utilização do time para que a página tenha tempo de carregar

time.sleep(5)

# Passo 2 -> Navegar até a página do arquivo

pa.click(x=450, y=347, clicks=2)
time.sleep(2)

# Passo 3 -> Exportar o Relatório

pa.click(x=462, y=409)
pa.click(x=1662, y=191)
pa.click(x=1450, y=696)
time.sleep(5)
pa.hotkey("enter")

# Passo 4 -> Calcular os indicadores 

# Abrir o arquivo csv com pandas

tabela = pc.read_excel(r"C:\Users\chrii\Downloads\Vendas - Dez.xlsx")

display(tabela)

faturamento = tabela["Valor Unitário"].sum()
quantidade = tabela["Quantidade"].sum()

print(faturamento)
print(quantidade)


# Passo 5 -> Enviar email para a diretoria

pa.hotkey("ctrl", "t")
pc.copy("https://mail.google.com/mail/u/0/#inbox")
pa.hotkey("ctrl","v")
pa.press("enter")
time.sleep(5)

texto = f"""
Prezados,
Bom dia

O faturamento de ontem foi de: R${faturamento:,;2f}
A quantidade de produtos foi de: {quantidade:,;2f}

Atenciosamente
Christian Q. Zein"""

# Comandos para confeccionar o email

pa.click(x=73, y=204)
pyautogui.click(x=73, y=204)

# preencher informações do email
pc.copy("the.christianzein@gmail.com")
pa.hotkey("ctrl", "v")
pa.click(x=1255, y=494)
pc.copy("Relatório de Vendas")
pa.hotkey("ctrl", "v")
pa.click(x=1238, y=611)
pc.copy(texto)
pa.hotkey("ctrl", "v")
pa.click(x=1179, y=979)
