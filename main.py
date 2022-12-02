import requests
from bs4 import BeautifulSoup
import pandas as pd
lista = []
listatitulos = []
listaprecos = []
listadesc = []

r = requests.get("file:///C:/Users/Aluno%20INRB/Desktop/Linguagem%20de%20marca%C3%A7%C3%A3o/Projeto%20Integrador%20-%20LIMA/Home.html")
soup = BeautifulSoup(r.content, "html.parser")

s = soup.find('div', class_="quadradoscate")
tituloproduto = s.find_all('a')
for c in tituloproduto:
    listatitulos.append(c.text)
precoprodutos = s.find_all('h3')
for c in precoprodutos:
    listaprecos.append(c.text)
descricaoprodutos = s.find_all('p')
for c in descricaoprodutos:
    listadesc.append(c.text)


s = soup.find('div', class_="quadradoscate2")
tituloproduto = s.find_all('a')
for c in tituloproduto:
    listatitulos.append(c.text)
precoprodutos = s.find_all('h3')
for c in precoprodutos:
    listaprecos.append(c.text)
descricaoprodutos = s.find_all('p')
for c in descricaoprodutos:
    listadesc.append(c.text)


s = soup.find('div', class_="quadradoscate3")
tituloproduto = s.find_all('a')
for c in tituloproduto:
    listatitulos.append(c.text)
precoprodutos = s.find_all('h3')
for c in precoprodutos:
    listaprecos.append(c.text)
descricaoprodutos = s.find_all('p')
for c in descricaoprodutos:
    listadesc.append(c.text)

listalegal = {"Nome produto": listatitulos,
              "Preço produto": listaprecos,
              "Descrição": listadesc}
listateste = pd.DataFrame(listalegal)
listateste.to_excel("ProjetoIntegrador.xlsx")