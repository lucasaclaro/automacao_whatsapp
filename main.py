# Bibliotecas necessárias para abrir o Excel no Python
import pandas as pd
import openpyxl

# Biblioteca para formatação de texto URL
import urllib
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
# Necessário para habilitar a função de apertar a tecla ENTER
from selenium.webdriver.common.keys import Keys
from time import sleep


contatos = pd.read_excel('telefones.xlsx')

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

navegador.get('https://web.whatsapp.com/')
sleep(1)

while len(navegador.find_elements(By.XPATH, '//*[@id="side"]/header/div[1]')) < 1:
    sleep(10)

for i, mensagem in enumerate(contatos['Mensagem']):
    pessoa = contatos.loc[i, 'Pessoa']
    numero = contatos.loc[i, 'Número']
    # Método para personalizar a mensagem. Isso é necessário porque o texto deve ser importado para o formato URL.
    texto = urllib.parse.quote(f'Oi, {pessoa}! {mensagem}')
    link = f'https://web.whatsapp.com/send?phone={numero}&text={texto}'
    navegador.get(link)
    sleep(10)
    # Apertar ENTER
    navegador.find_elements(By.XPATH, '// *[ @ id = "main"] / footer / div[1] / div / span[2] / div / div[2] / div[1] / div / div / p / span').send_keys(Keys.ENTER)
    sleep(10)




input('')