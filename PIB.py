# %%
#SCRIPT QUE ENTRA NO SITE DO IBGE E BAIXA OS DADOS DO PIB DIARIAMENTE
#CRIADOR: RAFAELA BERNARDES RABELO
#BIBLIOTECA PARA IMPORTAR: pip install schedule selenium pandas openpyxl requests beautifulsoup4
import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import Workbook, load_workbook
import requests
from bs4 import BeautifulSoup

driver = webdriver.Chrome() 
page = requests.get('https://sidra.ibge.gov.br/tabela/5932')
driver.get("https://sidra.ibge.gov.br/tabela/5932")
driver.maximize_window()
soup = BeautifulSoup(page.text, 'html.parser')

print("Site abriu com sucesso")
time.sleep(3) 

# CLICAR NO BOTÃO DE SELECIONAR APENAS ÍNDICE GERAL
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div[1]/div[4]/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[4]/div/div/div/button')
button_pesquisar.click()
print("Botão 'SELECIONAR' clicado.")
time.sleep(3)

# CLICAR NO BOTÃO DE SELECIONAR APENAS ÍNDICE GERAL
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div[1]/div[4]/div[2]/div/div[2]/div[2]/div/div[2]/div/div/div/div/div[3]/div/div/div/button')
button_pesquisar.click()
print("Botão 'SELECIONAR' clicado.")
time.sleep(3)

# CLICAR NO BOTÃO DE TIRAR TODOS SELECIONADOS
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div[1]/div[4]/div[3]/div/div[2]/div[2]/div/div[1]/div[1]/div/button[1]')
button_pesquisar.click()
print("Botão 'RETIRAR' clicado.")
time.sleep(3)

# CLICAR NO BOTÃO DE SELECIONAR APENAS ÍNDICE GERAL
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div[1]/div[4]/div[4]/div/div[2]/div[2]/div/div[1]/div[1]/div/button[1]')
button_pesquisar.click()
print("Botão 'SELECIONAR' clicado.")
time.sleep(3)



# CLICAR NO BOTÃO DE DOWNLOAD
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div[1]/div[5]/div[2]/div/div[2]/button[2]')
button_pesquisar.click()
print("Botão 'DOWNLOAD' clicado.")
time.sleep(3)

# CLICAR NO BOTÃO DE DOWNLOAD
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[7]/div/div/div[2]/div/div/div[2]/a')
button_pesquisar.click()
print("Botão 'DOWNLOAD' clicado.")
time.sleep(30)

# Identificar o arquivo mais recente na pasta de downloads
downloads_path = r"C:\Users\rafaela.rabelo\Downloads"
arquivos = os.listdir(downloads_path)
arquivos_csv = [arquivo for arquivo in arquivos if arquivo.endswith(".xlsx")]
arquivo_mais_recente = max(arquivos_csv, key=lambda x: os.path.getctime(os.path.join(downloads_path, x)))

# Caminho completo para o arquivo mais recente
caminho_arquivo = os.path.join(downloads_path, arquivo_mais_recente)

# Caminho de destino do arquivo Excel
caminho_destino = r"C:\Users\rafaela.rabelo\Downloads\CDI\PIB.xlsx"

# ABRIR O ARQUIVO E SALVAR NO EXCEL
file = pd.read_excel(caminho_arquivo)
file.to_excel(caminho_destino, sheet_name="PIB", index=False)

# Carregar o arquivo Excel novamente (opcional, se necessário para processamento adicional)
wb = load_workbook(caminho_destino)

# Salvar o arquivo Excel (opcional, se necessário para processamento adicional)
wb.save(caminho_destino)
wb.close()

print("PIB TRANSFERIDO para", caminho_destino)



