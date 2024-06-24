import os.path
import time
import tkinter as tk
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime

def login():# Cria a janela principal
    global root
    root = tk.Tk()
    root.title("Login Zabbix")

    # Cria os widgets da interface
    label_usuario = tk.Label(root, text="Usuário:")
    label_usuario.pack()
    entry_usuario = tk.Entry(root)
    entry_usuario.pack()

    label_senha = tk.Label(root, text="Senha:")
    label_senha.pack()
    entry_senha = tk.Entry(root, show="*")  # A senha é mostrada como asteriscos
    entry_senha.pack()

    button_login = tk.Button(root, text="Login", command=lambda: tentarLogin(entry_usuario.get(), entry_senha.get()))
    button_login.pack()

    root.mainloop()    # Inicia o loop de eventos da interface

def tentarLogin(usuario,senha):

    global drives
    if login_zabbix(usuario, senha):
        root.destroy()  # Fecha a janela de login
        extrairCSv(usuario)
    else:
        print("Erro durante o login. Tentando novamente...")
        root.destroy()  # Fecha a janela de login
        login()

def login_zabbix(usuario, senha):
    global driver

    # Configurações do WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--window-position=1920,0")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    driver = webdriver.Chrome(options=options)        # Inicializa o driver do Chrome

    # URL para login no Zabbix
    url = ("https://zabbix-client01.compwire.com.br/zabbix.php?show=1&name=&severities%5B2%5D=2&severities%5B3%5D=3&severities%5B4%5D=4&severities%5B5%5D=5&inventory%5B0%5D%5Bfield%5D=type&inventory%5B0%5D%5Bvalue%5D=&evaltype=0&tags%5B0%5D%5Btag%5D=&tags%5B0%5D%5Boperator%5D=0&tags%5B0%5D%5Bvalue%5D=&show_tags=3&tag_name_format=0&tag_priority=&show_opdata=0&show_timeline=1&filter_name=Dayle&filter_show_counter=0&filter_custom_time=0&sort=clock&sortorder=DESC&age_state=0&show_suppressed=0&unacknowledged=0&compact_view=0&details=0&highlight_row=0&action=problem.view&groupids%5B%5D=21&groupids%5B%5D=28&groupids%5B%5D=31&groupids%5B%5D=41&groupids%5B%5D=138&groupids%5B%5D=236")
    driver.get(url)

    time.sleep(5)  # Espera um pouco para a página carregar

    driver.find_element(By.ID, "login").click() # Clica no botão de login

    time.sleep(2)  # Espera um pouco para a página carregar

    usuario_field = driver.find_element(By.ID, "name")  # Preenche os campos de usuário e senha
    usuario_field.clear()  # Limpa o campo de usuário
    usuario_field.send_keys(usuario)  # Insere o usuário

    senha_field = driver.find_element(By.ID, "password")
    senha_field.clear()  # Limpa o campo de senha
    senha_field.send_keys(senha)  # Insere a senha

    driver.find_element(By.ID, "enter").click()# Clica no botão de login

    time.sleep(20)# Aguarde o login ser realizado (pode precisar de ajustes dependendo da página de login do Zabbix)

    try:
        driver.find_element(By.ID, "export_csv")
        return True

    except Exception as e:
        print('Erro durante o login:', e)
        return False

def extrairCSv(usuario):
    global driver

    caminhoZabbix = fr'C:\Users\{usuario}\Downloads\zbx_problems_export.csv'  # caminho para o export do zabbix

    if os.path.exists(caminhoZabbix):  # excluir o export do zabbix
        os.remove(caminhoZabbix)

    driver.find_element(By.NAME, "filter_apply").click()  # Applay

    time.sleep(2)

    driver.find_element(By.ID, "export_csv").click() # export

    time.sleep(4)    # espera 4s

    excel(usuario, caminhoZabbix)    # chama o excel

    driver.quit()     # sai do imput

def excel(usuario,caminhoZabbix):
    df = pd.read_csv(caminhoZabbix)    #alimenta o excel com o csv exportado pelo zabbix

    dataFormatada = datetime.now().strftime('%d-%m')     #formata a data em dia/mes para criar a pasta e arquivo excel

    mesformatado = int(datetime.now().strftime('%m'))   #formato mes como numero para interar sobre lista com messes

    meses=['Janeiro', 'Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro', 'Dezembro'] # lista de meses 

    mes='' # mes atual
    if mesformatado in range(1,len(meses)+1): # verificar em qual mes estamos 
        mes = meses[mesformatado - 1] 

        # verifica se a pasta existe se não ela cria a pasta
    if not os.path.exists(fr"C:\Users\{usuario}\OneDrive - compwire.com.br\Equipe CSC - Compwire\Daily\{mes}"):
        os.makedirs(fr"C:\Users\{usuario}\OneDrive - compwire.com.br\Equipe CSC - Compwire\Daily\{mes}")

    #verifica se a pasta existe se não ela cria a pasta
    if not os.path.exists(fr"C:\Users\{usuario}\OneDrive - compwire.com.br\Equipe CSC - Compwire\Daily\{mes}\{dataFormatada}"):
        os.makedirs(fr"C:\Users\{usuario}\OneDrive - compwire.com.br\Equipe CSC - Compwire\Daily\{mes}\{dataFormatada}")

    arquivo = f"Daily - Criticos - Backlog - Zabbix {dataFormatada}.xlsx"     #da nome para o arquivo que sera criado

    caminhoArquivo=fr"C:\Users\{usuario}\OneDrive - compwire.com.br\Equipe CSC - Compwire\Daily\Junho\{dataFormatada}\{str(arquivo)}"     # passa o caminho e o nome do arquivo

    time.sleep(10)     #tempo de espera para

    #exclui colunas não utilizadas
    df = df.drop(columns=['Recovery time', 'Tags'],errors='ignore')

    #escolhendo colunas que terao os nomes alterados
    alterar = {'Duration': 'Aging','Ack': 'Hoje','Time': 'Clientes','Status': 'Dias anteriores'}

    df.rename(columns=alterar, inplace=True)    #alltera o nome das colunas

    df = df[['Clientes', 'Host', 'Problem', 'Severity', 'Aging', 'Dias anteriores', 'Hoje']]    #ordena ar colunas

    #apaga o conteudo das colunas
    df['Clientes'] = np.nan
    df['Dias anteriores'] = np.nan
    df['Hoje'] = np.nan

    #cria o cliente sme e coloca a informação de report
    df['Dias anteriores'] = df['Host'].apply(lambda x: 'Sera enviado um report diario diretamente para a SME (período da manha).' if 'sme' in str(x) else '')

    #cria o clientes filtrando pelo host
    df['Clientes'] = df['Host'].apply(lambda x: 'SERPRO' if 'SERPRO' in str(x) else ('SME' if 'sme' in str(x) else('MPMS' if 'mpms' in str(x) else('MPMS' if 'MPMS' in str(x) else ('SANEPAR' if 'sanepar' in str(x) else'')))))

    prioridade = {'SME':5,'SANEPAR':4, 'MPMS':3,'SERPRO':2, '':1} # crio numero de prioridade
    df['Prioridade'] = df['Clientes'].map(prioridade) # add prioridades

    df_sorted = df.sort_values(by='Prioridade', ascending=True) # ordena por prioridade

    df_sorted=df_sorted.drop(columns='Prioridade') #exclui a coluna prioridade

    if not os.path.exists(arquivo):    #verifica se existe o arquivo antes de salvar

        df_sorted.to_excel(caminhoArquivo, index=False, sheet_name='Zabbix')  # criar o excel 
        wb = load_workbook(caminhoArquivo) #abre o arquivo
        ws = wb.active 
        tab_range = ws.dimensions # tamanho da tabela
        tab = Table(displayName="Zabbix", ref=tab_range) # cria a tabelas
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)  # style da tabela
        tab.tableStyleInfo = style
        ws.add_table(tab)
        wb.save(caminhoArquivo)
        os.startfile(caminhoArquivo)

login()