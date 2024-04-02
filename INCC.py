# %%
#SCRIPT QUE ENTRA NO SITE DA FGV E BAIXA OS DADOS DO INCC DIARIAMENTE
#CRIADOR: RAFAELA BERNARDES RABELO
#BIBLIOTECA PARA IMPORTAR: pip install schedule selenium pandas openpyxl requests beautifulsoup4
import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook, load_workbook
import pandas as pd

chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--no-sandbox")

driver = webdriver.Chrome(options=chrome_options)

page_url = "https://extra-ibre.fgv.br/autenticacao_produtos_licenciados/?ReturnUrl=%2fautenticacao_produtos_licenciados%2flista-produtos.aspx"
driver.get(page_url)

print("Site abriu com sucesso")
time.sleep(3) 

# CLICAR NO BOTÃO DE FECHAR DÚVIDAS
button_serieinstitucional = driver.find_element(By.XPATH, '/html/body/main/div/div/form/div[3]/div/div/div[2]/div[3]/a[2]')
button_serieinstitucional.click()
print("Botão 'series institucionais' clicado.")
time.sleep(6)

datainicial = driver.find_element(By.ID, 'txtBuscarSeries')
datainicial.click()
print("Elemento encontrado")
time.sleep(3)

# APAGA O VALOR QUE EXISTE E COLOCA UM NOVO
datainicial.clear()
datainicial.send_keys("INCC todos os itens") 
print("Digitado INCC")
time.sleep(2)

# CLICAR NO BOTÃO DE CONSULTAR
button_ok = driver.find_element(By.ID, 'butBuscarSeries')
button_ok.click()
print("Botão 'OK' clicado.")
time.sleep(10)

# CLICAR NO BOTÃO DE FECHAR DÚVIDAS
button_selecionartodos = driver.find_element(By.ID, 'btnSelecionarTodas')
button_selecionartodos.click()  
print("Botão 'selecionar todos' clicado.")
time.sleep(6)

# CLICAR NO BOTÃO DE FECHAR DÚVIDAS
button_visualizar = driver.find_element(By.ID, 'butBuscarSeriesOK')
button_visualizar.click()
print("Botão 'ok' clicado.")
time.sleep(6)

# CLICAR NO BOTÃO DE FECHAR DÚVIDAS
button_serieihistorica = driver.find_element(By.ID, 'cphConsulta_rbtSerieHistorica')
button_serieihistorica.click()
print("Botão 'serie histórica' clicado.")
time.sleep(6)

# CLICAR NO BOTÃO DE FECHAR DÚVIDAS
button_visualizar = driver.find_element(By.ID, 'cphConsulta_butVisualizarResultado')
button_visualizar.click()
print("Botão 'OK' clicado.")
time.sleep(40)


# Depois de abrir a página, encontre o iframe
iframe = driver.find_element(By.ID, 'cphConsulta_ifrVisualizaConsulta')

# Mude o contexto para o iframe
driver.switch_to.frame(iframe)

# Agora você está dentro do iframe e pode interagir com os elementos dentro dele
link_csv = driver.find_element(By.ID, 'lbtSalvarCSV')

# Faça as operações desejadas (por exemplo, clique nos links)
link_csv.click()
# Volte para o contexto padrão fora do iframe
driver.switch_to.default_content()
print("Aqruivo baixado")
time.sleep(30)

# Identificar o arquivo mais recente na pasta de downloads
downloads_path = r"C:\Users\rafaela.rabelo\Downloads"
arquivos = os.listdir(downloads_path)
arquivos_csv = [arquivo for arquivo in arquivos if arquivo.endswith(".csv")]
arquivo_mais_recente = max(arquivos_csv, key=lambda x: os.path.getctime(os.path.join(downloads_path, x)))

# Caminho completo para o arquivo mais recente
caminho_arquivo = os.path.join(downloads_path, arquivo_mais_recente)

# Caminho de destino do arquivo Excel
caminho_destino = r"C:\Users\rafaela.rabelo\Downloads\CDI\INCC.xlsx"

# ABRIR O ARQUIVO E SALVAR NO EXCEL
file = pd.read_table(caminho_arquivo, sep=";")
file.to_excel(caminho_destino, sheet_name="INCC", index=False)

# Carregar o arquivo Excel novamente
wb = load_workbook(caminho_destino)

# Salvar o arquivo Excel
wb.save(caminho_destino)
wb.close()

print("INCC TRANSFERIDO para", caminho_destino)
