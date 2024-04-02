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
# Configurações do WebDriver
driver = webdriver.Chrome() 
page_url = 'https://www3.bcb.gov.br/sgspub/consultarvalores/consultarValoresSeries.do?method=consultarSeries&series=7461'
driver.get(page_url)
time.sleep(2)  # Aguardar um curto período para a página carregar

# Encontrar a tabela
tabela_desempenho = driver.find_element(By.XPATH, '/html/body/center/form/table[2]')
linhas_tabela = tabela_desempenho.find_elements(By.XPATH, './tbody/tr')

# Variável para armazenar dados
dados = []

# Iterar sobre as linhas da tabela
for i, linha in enumerate(linhas_tabela, start=1):
    # Encontrar todas as colunas na linha
    colunas = linha.find_elements(By.XPATH, './/td')

    # Verificar se existem pelo menos duas colunas
    if len(colunas) >= 2:
        # Extrair dados de cada coluna
        data_7461 = colunas[0].text
        var_percentual = colunas[1].text

        # Adicionar os dados à lista
        dados.append({
            "Data": data_7461,
            "7461": var_percentual
        })
    else:
        print("Número insuficiente de colunas nesta linha, pulando para a próxima")

# Fechar o navegador
driver.quit()

# Exemplo de como salvar os dados diretamente em um arquivo Excel na pasta desejada
excel_file_path = r'C:\Users\rafaela.rabelo\Downloads\CDI\INCC - % mão de obra.xlsx'
df = pd.DataFrame(dados)
df.to_excel(excel_file_path, index=False)

# Imprimir mensagem informando onde o arquivo Excel foi salvo localmente
print(f"Excel salvo em: {excel_file_path}")