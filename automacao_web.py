# BAIXE O ARQUIVO "Produtos.xlsx" E COLOQUE NO MESMO LOCAL DO CÓDIGO, PARA O MESMO FUNCIONAR CORRETAMENTE

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

#  PEGANDO COTAÇÃO DO DÓLAR
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)
navegador.get("https://www.google.com/")
navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys("cotacao dolar" + Keys.ENTER)
cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
navegador.find_element(By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[3]/div/div[3]/div[1]/div').click()

#  PEGANDO COTAÇÃO DO EURO
navegador.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys("cotacao euro" + Keys.ENTER)
cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
navegador.quit()

#  PEGANDO COTAÇÃO DO OURO
navegador = webdriver.Chrome()
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value')
cotacao_ouro = cotacao_ouro.replace(",", ".")
navegador.quit()

# ATUALIZANDO OS DADOS DA TABELA EXCEL COM AS COTAÇÕES ATUAIS
try:
    tabela = pd.read_excel("Produtos.xlsx")
except:
    print("ERRO! Verifique se o arquivo excel está no mesmo caminho do código")
else:
    tabela.loc[tabela['Moeda'] == "Dólar", "Cotação"] = float(cotacao_dolar)
    tabela.loc[tabela['Moeda'] == "Euro", "Cotação"] = float(cotacao_euro)
    tabela.loc[tabela['Moeda'] == "Ouro", "Cotação"] = float(cotacao_ouro)

    tabela["Preço de compra"] = tabela["Preço Original"]  * tabela["Cotação"]
    tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

    # GERANDO UMA NOVA TABELA EXCEL COM OS VALORES ATUALIZADOS
    tabela.to_excel("Novos Produtos.xlsx", index=False)
    print("tabela atualizada")
