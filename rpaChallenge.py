from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

site = "http://rpachallenge.com/" # variável para o site
arquivo = 'C:\challenge.xlsx' # variável para o arquivo de excel do desafio
df_registros = pd.read_excel(arquivo) # variável para ler os dados do arquivo de excel
print(df_registros.info()) # mostrando os dados do arquivo de excel

driver = webdriver.Firefox(executable_path=r"C:\geckodriver.exe") # abrindo o navegador Firefox
print("Iniciando nosso robô...\n")
print("Acessando site...")
driver.get(site) # abrindo o site rpa challenge conforme variável site

print('Starting...')
driver.find_element('xpath', '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click() # clicando no botão Start
time.sleep(2)

# looping para ler cada item do arquivo de excel challenge.xlsx
for i, r in df_registros.iterrows():        
    role = r['Role in Company']
    email = r['Email']
    first_name = r['First Name']
    last_name = r['Last Name '] # necessário colocar um espaço ao final do nome visto que o arquivo foi baixado desta forma no RPA Challenge
    phone = r['Phone Number']
    company = r['Company Name']
    address = r['Address']

    print(first_name)
    
    # buscando o input "Role in Company"
    print("Writing role...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelRole']")
    textbox.clear()
    textbox.send_keys(role)
    
    # buscando o input "Email"
    print("Writing email...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelEmail']")
    textbox.clear()
    textbox.send_keys(email)

    # buscando o input "First Name"
    print("Writing first name...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelFirstName']")
    textbox.clear()
    textbox.send_keys(first_name)

    # buscando o input "Last Name"
    print("Writing last name...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelLastName']")
    textbox.clear()
    textbox.send_keys(last_name)
    #time.sleep(0.5)

    # buscando o input "Phone Number"
    print("Writing phone...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelPhone']")
    textbox.clear()
    textbox.send_keys(phone)

    # buscando o input "Address"
    print("Writing address...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelAddress']")
    textbox.clear()
    textbox.send_keys(address)
    
    # buscando o input "Company Name"
    print("Writing company...")
    textbox = driver.find_element("xpath", "//input[@ng-reflect-name='labelCompanyName']")
    textbox.clear()
    textbox.send_keys(company)

    # buscando o botão "Submit"
    print('Submit...')
    botao = driver.find_element("xpath", "//input[@type='submit']")
    botao.click()
    time.sleep(0.5)

print('Processo finalizado')
