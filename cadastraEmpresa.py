import traceback
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from docx import Document
import openpyxl
import unicodedata
import re

# Solicita as credenciais do usuário
emailLogin = 'guilherme.loureiro@setuptecnologia.com.br'
senhaLogin = 'Racewin@1406'
print("Acessando Acessorias Agora")

# Caminho para o driver do Edge
driver_path = 'msedgedriver.exe'

# Configura as opções do Edge
edge_options = Options()
edge_options.headless = False  # Executa o Edge em modo não headless

# Inicializa o navegador Edge
service = Service(driver_path)
edge_driver = webdriver.Edge(service=service, options=edge_options)

# Abre o link desejado
url = f"https://app.acessorias.com/sysmain.php?m=4"
edge_driver.get(url)

url = 'https://app.acessorias.com/sysmain.php?m=4'
edge_driver.get(url)

# Carrega o arquivo .xlsx
workbook = openpyxl.load_workbook('empresas.xlsx')

# Seleciona a planilha
sheet = workbook['Empresas']

for row in sheet.iter_rows(min_row=5, min_col=1, max_col=11):
    
    IdEmpresa = row[0].value          # Coluna 1: Código do Sistema Contábil (ID)
    NomeFantasia = row[1].value       # Coluna 2: Nome Fantasia
    RazaoSocial = row[2].value        # Coluna 3: Razão Social
    Cnpj = row[3].value               # Coluna 4: CNPJ/CPF/CEI/CAEPF
    InscEst = row[4].value            # Coluna 5: Insc. Estadual
    InscUf = row[5].value             # Coluna 6: Insc. Estadual UF
    NomeCtto = row[6].value           # Coluna 7: Nome do Contato
    EmailCtto = row[7].value          # Coluna 8: E-mail do Contato
    Regime = row[8].value             # Coluna 9: Regime Tributário
    Apelido = row[10].value           # Coluna 11: Apelido e-Contínuo
    
    print(f"ID Empresa: {IdEmpresa}")
    print(f"Nome Fantasia: {NomeFantasia}")
    print(f"Razão Social: {RazaoSocial}")
    print(f"CNPJ/CPF/CEI/CAEPF: {Cnpj}")
    print(f"Inscrição Estadual: {InscEst}")
    print(f"Inscrição Estadual UF: {InscUf}")
    print(f"Nome do Contato: {NomeCtto}")
    print(f"E-mail do Contato: {EmailCtto}")
    print(f"Regime Tributário: {Regime}")
    print(f"Apelido e-Contínuo: {Apelido}")
    print("-" * 50)  # Separador para facilitar a leitura dos dados

    teste=input('breakpoint')
    # Lógica de login
    try:
        # Espera o campo de e-mail aparecer
        email_input = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'mailAC'))
        )
        # Insere o e-mail no campo
        email_input.send_keys(emailLogin)
    except:
        print("Erro ao inserir o e-mail no campo de login:")
        traceback.print_exc()

    try:
        # Espera o campo de senha aparecer
        senha_input = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'passAC'))
        )
        # Insere a senha no campo
        senha_input.send_keys(senhaLogin)
    except:
        print("Erro ao inserir a senha no campo de senha:")
        traceback.print_exc()

    try:
        # Espera o botão de login aparecer e ser clicável
        login_button = WebDriverWait(edge_driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="site-corpo"]/section[1]/div/form/div[2]/button'))
        )
        # Clique no botão de login
        login_button.click()
    except:
        print("Erro ao clicar no botão de login:")
        traceback.print_exc()

    time.sleep(3)

    url = 'https://app.acessorias.com/sysmain.php?m=4'
    edge_driver.get(url)

    #Clicar no botão nova empresa
    try:
        novaEmpresaButton = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="main-container"]/div[2]/div[2]/div/div/form[1]/div[1]/div[3]/button[2]'))
        )
        novaEmpresaButton.click()
    except Exception as e:
        print("Erro ao clicar no botão nova empresa")
        traceback.print_exc()

    try:
        # Espera o campo de senha aparecer
        cnpjInput = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'field_EmpCNPJ'))
        )
        
        # Insere a senha no campo
        cnpjInput.send_keys(Cnpj)
    except:
        print("Erro ao inserir o CNPJ na empresa")
        traceback.print_exc()
        
    try:
        # Espera o campo de senha aparecer
        cnpjInput = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'field_EmpCNPJ'))
        )
        
        # Insere a senha no campo
        cnpjInput.send_keys(Cnpj)
    except:
        print("Erro ao inserir o CNPJ na empresa")
        traceback.print_exc()

    teste=input('breakpoint')
    
    