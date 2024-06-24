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


# Definição das URLs dos portais utilizados
portal_admin = "https://portal.madeiramadeira.com.br/admin/auth/index" #apenas para logar no admin
portal_lr = "https://portal.madeiramadeira.com.br/admin/sac/index" #para cancelamento de coleta
portal_log_reversa = "https://portal.madeiramadeira.com.br/admin/coletas-reversas/index" #para prosseguir com a coleta


# Função para interagir com a janela de login
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

    # Aumentando o tamanho dos labels e campos de entrada
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

    # Centralizando a janela na tela
    root.update_idletasks()  # Atualiza o gerenciador de layout
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'+{x}+{y}')  # Define a posição da janela

    root.mainloop()

janela_interacao()

# Função de interação com a janela de login para obter as credenciais
senha, login, codigo = janela_interacao()
print("login: ", login, " - senha: ", senha, "- codigo: ", codigo)
print("Aguardando WebDriver")

# Inicialização do WebDriver do Chrome
browser = webdriver.Chrome()
#browser = webdriver.Chrome(os.path.join(dirname, "chromedriver.exe"))
print("WebDriver Iniciado")
browser.get(portal_admin)
browser.maximize_window()
wait = WebDriverWait(browser, 5, poll_frequency=1)

# INSERÇÃO DAS CREDENCIAIS
element = wait.until(EC.element_to_be_clickable((By.ID, 'salvar')))
login_input = browser.find_element(By.NAME, 'email').send_keys(login)
senha_input = browser.find_element(By.NAME, 'password').send_keys(senha)
bt = browser.find_element(By.NAME, 'submit').click()
time.sleep(5)
senha_input = browser.find_element(By.NAME, 'password').send_keys(senha)
code_input = browser.find_element(By.NAME, 'code').send_keys(codigo)
bt = browser.find_element(By.NAME, 'submit').click()
time.sleep(10)

# AGUARDA PÁGINA CARREGAR
element = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="table_admin_fornecedores_filter"]/button')))
browser.get(portal_admin)
time.sleep(2)
last_pv = 0
last_nf = 0


time.sleep(10)
# Leitura do arquivo Excel contendo ocorrências de reversa
ocr_reversa = pd.read_excel(r'caminho_arquivo\Ocr_Reversa.xlsx')

#dataframes que separaram cancelamento de coleta e prosseguir com coleta
cancelar_coleta = ocr_reversa.drop(ocr_reversa.loc[(ocr_reversa['Posição Final'] != 'Cancelar coleta')] .index)
prosseguir_coleta = ocr_reversa.drop(ocr_reversa.loc[(ocr_reversa['Posição Final'] != 'Prosseguir coleta')] .index)


# Iteração sobre os DataFrames
# Cancelamento de coleta
if not cancelar_coleta.empty:
    for index, row in cancelar_coleta.iterrows(): #cancelar coleta
        print("Index: " + str(index) + " Pedido: " + row['PV'])

        # Abrir o navegador e acessar o portal de cancelamento de coleta
        browser.get(portal_lr)
        browser.maximize_window()
        #browser.set_window_rect(x=0, y=0, width=1500, height=1000)
        # browser.maximize_window()
        time.sleep(3)

        # Abrir o portal de logística reversa e inserir o número do pedido para pesquisa
        pendencias_lr = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div[1]/div/div[1]/a[2]')
        pendencias_lr.click()
        time.sleep(5)

        # Colocar o número do pedido no campo de pesquisa
        pedido = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div[3]/div[2]/label/input')
        pedido.send_keys(row['PV'])
        time.sleep(3)
        pedido.send_keys(Keys.ENTER)
        time.sleep(5)
        try:
            # Clicar nas ações disponíveis para o pedido
            acoes = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div[3]/table/tbody/tr/td[16]/div/a')
            acoes.click()
            time.sleep(3)
            # Consultar o histórico do pedido
            cons_historico = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div[3]/table/tbody/tr/td[16]/div/ul/li[4]/a')
            cons_historico.click()
            time.sleep(3)
            # Colar texto no chat relacionado ao cancelamento da coleta
            descricao = browser.find_element(By.XPATH,'/html/body/div[2]/div[11]/div[3]/form/textarea')
            time.sleep(5)
            descricao.send_keys(row['Responsável Ativo'] + " " +  str(row['Controle do Ativo']) + " " +  row['Observação'] + " - " + " Coleta cancelada nf refaturada")
            time.sleep(2)
            descricao.send_keys(Keys.ENTER)
            # Sair do chat
            sair_descricao = browser.find_element(By.XPATH,'/html/body/div[2]/div[11]/div[1]/button')
            sair_descricao.click()
            time.sleep(5)
            # Finalizar a ação de cancelamento
            acoes = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div[3]/table/tbody/tr/td[16]/div/a')
            acoes.click()
            time.sleep(3)
            finalizar = browser.find_element(By.XPATH,'/html/body/div[2]/div[1]/div/section/div/div[3]/table/tbody/tr/td[16]/div/ul/li[5]/a')
            finalizar.click()
            time.sleep(3)
            # Confirmar o cancelamento
            sim = browser.find_element(By.XPATH,'/html/body/div[2]/div[20]/div/div/div[3]/button[2]')
            sim.click()
        except NoSuchElementException:
            print(row['PV'] + ' ' + 'pedido não encontrado')


# Prosseguir com a coleta
for index, row in prosseguir_coleta.iterrows(): #prosseguir coleta
    print("Index: " + str(index) + " Pedido: " + row['PV'])

    browser.get(portal_log_reversa)
    browser.maximize_window()
    time.sleep(3)
    try:
        # Acessar a opção de problemas de coleta
        problema_coleta = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div/div/div[1]/div[1]/div/a[6]')
        problema_coleta.click()
        time.sleep(5)

        # Inserir o número do pedido para pesquisa
        pedido = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div/div/div[2]/div[2]/label/input')
        pedido.send_keys(row['PV'])
        time.sleep(8)
        pedido.send_keys(Keys.ENTER)

        # Acessar as ações disponíveis para o pedido
        acoes = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div/div/div[2]/table/tbody/tr[1]/td[13]/div/a')
        acoes.click()
        time.sleep(3)

        # Visualizar o problema relacionado à coleta
        visualizar_problema = browser.find_element(By.XPATH, '/html/body/div[2]/div[1]/div/section/div/div/div/div[2]/table/tbody/tr/td[13]/div/ul/li[1]/a')
        visualizar_problema.click()
        time.sleep(3)

        # Responder ao problema
        problema_resposta = browser.find_element(By.XPATH,'/html/body/div[2]/div[2]/div/div[2]/form/div[1]/div/textarea')
        problema_resposta.send_keys(row['Observação'])
        problema_resposta.send_keys(Keys.ENTER)
        time.sleep(3)

        # Enviar a resposta
        enviar_resposta = browser.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/form/div[2]/button')
        enviar_resposta.click()
        time.sleep(3)

        # Reativar a coleta
        reativar_coleta = browser.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div[2]/form/div[2]/a[1]')
        reativar_coleta.click()
    except NoSuchElementException:
        print(row['PV'] + ' ' + 'pedido não encontrado')
