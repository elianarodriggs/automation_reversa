# Logistics Reversal Automation Bot

This project is a bot designed to automate the process of handling reverse logistics on the web using Selenium WebDriver. 
The bot performs tasks such as logging in, canceling collections, and proceeding with collections based on data from an Excel file.

# Features
• Automates login process to the MadeiraMadeira admin portal.

• Handles reverse logistics tasks including canceling and proceeding with collections.

• Uses Selenium WebDriver for browser automation.

• Provides a Tkinter-based GUI for user input.

# Requirements
  • Python 3.x
  
  • Selenium
  
  • pandas
  
  • tkinter
  
  • Chrome WebDriver

# Installation
  1. Clone the repository:

```
git clone https://github.com/yourusername/logistics-reversal-bot.git
cd logistics-reversal-bot

```
2. Install the required packages:

```
pip install selenium pandas tkinter

```
3. Download the ![Chrome WebDriver](https://developer.chrome.com/docs/chromedriver/downloads?hl=pt-br) and place it in your PATH or in the project directory.

# Usage
1. Ensure you have the Excel file (file_name.xlsx) in the specified directory.

Run the script:
```
python bot.py

```

2. A Tkinter window will pop up for you to enter your login credentials and a code.
   
4. The bot will log in to the MadeiraMadeira admin portal and start processing the tasks based on the Excel file.


## Code Overview

# Import Libraries

```
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox

```

URLs Definition

```
portal_admin = "https://portal.....com.br/admin/auth/index"
portal_lr = "https://portal.....com.br/admin/sac/index"
portal_log_reversa = "https://portal.....com.br/admin/coletas-reversas/index"

```

Tkinter Login Window

```
def janela_interacao():
    def entrar():
        login = usuario_entry.get()
        senha = senha_entry.get()
        codigo = codigo_entry.get()
        if not (login and senha and codigo):
            messagebox.showinfo("Erro", "Favor digitar login, senha e código!")
        else:
            print('login:', login, 'senha:', senha, 'código:', codigo)
            root.destroy()

    root = tk.Tk()
    root.title("Login Portal Fornecedor")

    tk.Label(root, text="Usuário", width=20).grid(row=0, column=0)
    usuario_entry = tk.Entry(root, width=30)
    usuario_entry.grid(row=0, column=1)

    tk.Label(root, text="Senha", width=20).grid(row=1, column=0)
    senha_entry = tk.Entry(root, width=30, show="*")
    senha_entry.grid(row=1, column=1)

    tk.Label(root, text="Código", width=20).grid(row=2, column=0)
    codigo_entry = tk.Entry(root, width=30)
    codigo_entry.grid(row=2, column=1)

    entrar_button = tk.Button(root, text="Entrar", width=20, command=entrar)
    entrar_button.grid(row=3, column=1)

    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')

    root.mainloop()

```

# Selenium Automation

Initialize WebDriver

```
browser = webdriver.Chrome()
browser.get(portal_admin)
browser.maximize_window()
wait = WebDriverWait(browser, 5, poll_frequency=1)
```

Login Process

```
element = wait.until(EC.element_to_be_clickable((By.ID, 'salvar')))
login_input = browser.find_element(By.NAME, 'email').send_keys(login)
senha_input = browser.find_element(By.NAME, 'password').send_keys(senha)
bt = browser.find_element(By.NAME, 'submit').click()
```
Read Excel File

```
ocr_reversa = pd.read_excel(r'caminho_arquivo\file.xlsx')
cancelar_coleta = ocr_reversa.drop(ocr_reversa.loc[(ocr_reversa['Posição Final'] != 'Cancelar coleta')].index)
prosseguir_coleta = ocr_reversa.drop(ocr_reversa.loc[(ocr_reversa['Posição Final'] != 'Prosseguir coleta')].index)

```

Cancel Collection

```

if not cancelar_coleta.empty:
    for index, row in cancelar_coleta.iterrows():
        browser.get(portal_lr)
        browser.maximize_window()
        time.sleep(3)
        # Continue with actions to cancel the collection...

```
Proceed with Collection

```

for index, row in prosseguir_coleta.iterrows():
    browser.get(portal_log_reversa)
    browser.maximize_window()
    time.sleep(3)
    # Continue with actions to proceed with the collection...

```
Proceed with Collection

```

for index, row in prosseguir_coleta.iterrows():
    browser.get(portal_log_reversa)
    browser.maximize_window()
    time.sleep(3)
    # Continue with actions to proceed with the collection...

```


