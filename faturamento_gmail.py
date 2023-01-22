#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyautogui
import time
import pyperclip
import pandas as pd

df = pd.read_excel(r'C:\Users\victo\OneDrive\Documentos\Exportar\Vendas - Dez.xlsx')
faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

#Abrindo o Gmail
pyautogui.press("winleft")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.alert("Atenção o programa vai começar a rodar, aperte OK e não mexa em nada!")
pyautogui.hotkey('ctrl','t')

link = "https://mail.google.com/mail/u/0/#inbox"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
time.sleep(5)

#ESCREVENDO O EMAIL
pyautogui.click(x=194, y=208)
pyautogui.write('victortrabalhoti@gmail.com')
pyautogui.press("tab")
pyautogui.press("tab")
assunto = "Relatorio de Vendas de Ontem"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
pyautogui.press("tab")
texto = f"""
Prezados, bom dia!

O Faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {qtde_produtos:,}

Cordialmente,
Victor Nunes Pinho"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl','v')
pyautogui.hotkey('ctrl','enter')

