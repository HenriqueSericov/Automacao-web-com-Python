# Importando as Bibliotecas
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Atribuindo o webdriver á variavel navegador
navegador = webdriver.Chrome()

# Pegando a cotação do Dólar
navegador.get("https://www.google.com/search?q=cota%C3%A7%C3%A3o+do+dolar&oq=cota%C3%A7%C3%A3o+do+dolar+&aqs=chrome..69i57j0i433i457i512j0i402l2j0i512l6.10087j1j7&sourceid=chrome&ie=UTF-8/")
cotacao_dolar = navegador.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

# Pegando a cotação do Euro
navegador.get("https://www.google.com/search?q=cota%C3%A7%5Dao+do+euro&sxsrf=AJOqlzUkzoIHEngEugZGmKtzVt2JCUv0dA%3A1673814107500&ei=W2DEY47XHcjY1sQPooetwAU&ved=0ahUKEwjO8K-6s8r8AhVIrJUCHaJDC1gQ4dUDCA8&uact=5&oq=cota%C3%A7%5Dao+do+euro&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAzIPCAAQgAQQsQMQChBGEIICMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKMgcIABCABBAKOgcIIxDqAhAnOgwIABDqAhC0AhBDGAE6BAgAEEM6CwguEIAEELEDEIMBOgsILhCABBDHARDRAzoLCAAQgAQQsQMQgwE6BQguEIAEOggILhCxAxCDAToJCCMQJxBGEIICOgQIIxAnOggILhCDARCxAzoFCAAQgAQ6BwgAELEDEEM6CAgAEIAEELEDOgUIABCSAzoMCAAQHhDxBBDJAxAKOgkIABAeEPEEEAo6DAgjELECECcQRhCCAjoKCAAQgAQQsQMQCjoNCAAQgAQQsQMQgwEQCjoMCAAQsQMQQxBGEIICOg0IABCABBCxAxDJAxAKSgQIQRgASgQIRhgBULwJWNkjYIImaAFwAXgAgAF1iAHFDZIBBDEuMTWYAQCgAQGwARTAAQHaAQYIARABGAE&sclient=gws-wiz-serp")
cotacao_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')

# Pegando cotação do Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje#:~:text=O%20valor%20do%20grama%20do,%C3%A9%20de%20car%C3%A1ter%20exclusivamente%20informativo.")
cotacao_ouro = navegador.find_element('xpath','//*[@id="comercial"]').get_attribute('value')
# Substituindo a , do valor por um . para que ele possa ser convertido para float
cotacao_ouro = cotacao_ouro.replace(",",".")

# Importando Tabela
tabela = pd.read_excel('Produtos.xlsx')

# Atualizando valor da cotação
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"]=float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"]=float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"]=float(cotacao_ouro)

# Atualizando preço de compra e venda baseando-se na cotação 
tabela["Preço de Compra"] = tabela["Cotação"]*tabela["Preço Original"]
tabela["Preço de Venda"] = tabela["Preço de Compra"]*tabela["Margem"]
# Formatando valor de venda na tabela 
tabela["Preço de Venda"] = tabela["Preço de Venda"].map("R${:.2f}".format)

# Exportando tabela Atualizada
tabela.to_excel('Produtos Novos.xlsx', index=False)

print(tabela)

