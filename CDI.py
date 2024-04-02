# %%
#SCRIPT QUE ENTRA NO SITE DO BACEN E BAIXA OS DADOS DO CDI DIARIAMENTE
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
import schedule

driver = webdriver.Chrome() 
page = requests.get('https://sistema.muve.delivery/Home/Login')
driver.get("https://www3.bcb.gov.br/novoselic/pesquisa-taxa-apurada.jsp")
driver.maximize_window()
soup = BeautifulSoup(page.text, 'html.parser')

print("Site abriu com sucesso")
time.sleep(3) 

datainicial = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/form/div[1]/div[1]/selic-datepicker/input')
datainicial.click()
print("Elemento encontrado")
time.sleep(3) 

# APAGA O VALOR QUE EXISTE E COLOCA UM NOVO
datainicial.clear()
datainicial.send_keys("01/01/2016") 
print("Data inicial definida para 01/01/2000")
time.sleep(6)

# CLICAR NO BOTÃO DE CONSULTAR
button_pesquisar = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/button[3]')
button_pesquisar.click()
print("Botão 'Pesquisar' clicado.")
time.sleep(10)

# CLICAR NO BOTÃO EXPORTAR
exportar_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/selic-exporter/div/ul/li[1]/p')
exportar_button.click()
time.sleep(2) 
print("Botão 'Exportar' clicado.")

# BAIXAR ARQUIVO
opcao_csv = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/selic-exporter/div/ul/li[1]/ul/li[3]/a')
opcao_csv.click()

print("Arquivo baixado.")

# Aguardar o download ser concluído
time.sleep(10)

# Identificar o arquivo mais recente na pasta de downloads
downloads_path = r"C:\Users\rafaela.rabelo\Downloads"
arquivos = os.listdir(downloads_path)
arquivos_csv = [arquivo for arquivo in arquivos if arquivo.endswith(".csv")]
arquivo_mais_recente = max(arquivos_csv, key=lambda x: os.path.getctime(os.path.join(downloads_path, x)))

# Caminho completo para o arquivo mais recente
caminho_arquivo = os.path.join(downloads_path, arquivo_mais_recente)

# Caminho de destino do arquivo Excel
caminho_destino = r"C:\Users\rafaela.rabelo\Downloads\CDI\CDI.xlsx"

# ABRIR O ARQUIVO E SALVAR NO EXCEL
file = pd.read_table(caminho_arquivo, sep=";")
file.to_excel(caminho_destino, sheet_name="CDI", index=False)

# Carregar o arquivo Excel novamente
wb = load_workbook(caminho_destino)

# Salvar o arquivo Excel
wb.save(caminho_destino)
wb.close()

print("CDI TRANSFERIDO para", caminho_destino)


