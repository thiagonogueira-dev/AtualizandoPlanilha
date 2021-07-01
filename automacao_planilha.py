# Navegador utilizado para a automação: Google Chrome #
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

navegador = webdriver.Chrome()
navegador.get("https://www.google.com")

# Atualizando a cotação do Dólar
navegador.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys("Dólar")
navegador.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys(Keys.ENTER)
cotacao_dolar = float(navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value"))
cotacao_dolar = f'{cotacao_dolar:_.3f}'  # Ajustando o valor do Dólar com 3 casas decimais
print(cotacao_dolar)

# Atualizando a cotação do Euro
navegador.get("https://www.google.com")
navegador.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys("Euro")
navegador.find_element_by_xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input").send_keys(Keys.ENTER)
cotacao_euro = float(navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value"))
cotacao_euro = f'{cotacao_euro:_.3f}'   # Ajustando o valor do Euro com 3 casas decimais
print(cotacao_euro)

# Atualizando a cotação do Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje")
cotacao_ouro = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute("value")
cotacao_ouro = float(f'{cotacao_ouro}'.replace(",", "."))  # Ajustando o preço do Ouro para o padrão do Python
print(cotacao_ouro)


# Importando a base de dados e atualizando a cotação
tabela = pd.read_excel("Produtos.xlsx")

tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)

# Atualizando o preço da coluna 'Preço Base Reais'
tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]

# Atualizando o preço da coluna 'Preço Final'
tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]

# Salvando a tabela atualizada
tabela.to_excel("./Produtos_Atualizada.xlsx", index=False)

# Fechando o navegador
navegador.close()
