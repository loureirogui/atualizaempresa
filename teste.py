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
import random

# Solicita as credenciais do usuário
emailLogin = 'guilherme.loureiro@setuptecnologia.com.br'
senhaLogin = 'Racewin@1406'
print("Acessando Acessorias Agora")

# Caminho para o driver do Edge
driver_path = 'msedgedriver.exe'

# Configura as opções do Edge
edge_options = Options()
edge_options.add_argument('--log-level=3')  # Isso define o nível de log para 'fatal', suprimindo a maioria das mensagens de erro.
edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Exclui certos logs.

# Inicializa o navegador Edge
service = Service(driver_path)
edge_driver = webdriver.Edge(service=service, options=edge_options)
edge_driver.set_window_size(1200, 800)
# Abre o link desejado
url = f"https://app.acessorias.com/sysmain.php?m=4"
edge_driver.get(url)

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

# Carrega o arquivo .xlsx
workbook = openpyxl.load_workbook('empresas.xlsx')

# Seleciona a planilha
sheet = workbook['Empresas']

numeros_possiveis = list(range(60000, 65001))

# Função para obter um ID aleatório caso o do cliente esteja em uso
def obter_id_aleatorio():
    if numeros_possiveis:  # Verifica se ainda há números disponíveis
        return numeros_possiveis.pop(random.randint(0, len(numeros_possiveis) - 1))
    else:
        raise ValueError("Todos os IDs possíveis já foram usados.")

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
    Apelido = row[9].value           # Coluna 11: Apelido e-Contínuo

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

    #inserir cnpj
    try:
        # Espera o campo aparecer
        cnpjInput = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'field_EmpCNPJ'))
        )
        
        # Insere o cnpj no campo
        cnpjInput.send_keys(Cnpj)
    except:
        print("Erro ao inserir o CNPJ na empresa")
        traceback.print_exc()

    try:
        # Procurar o botão buscar
        buscarButton = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'btCNPJ'))
        )
        
        # Insere a senha no campo
        buscarButton.click()
    except:
        print()
        
    # Regime tributário
    try:
        # Encontra o seletor <select> pelo nome
        selectRegime = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'field_EmpRegID'))
        )

        # Cria uma instância de Select com o elemento encontrado
        select = Select(selectRegime)
        
        # Normaliza o texto do regime para comparação
        regime_normalized = unicodedata.normalize('NFKD', Regime).encode('ascii', 'ignore').decode('utf-8').upper()

        # Itera através das opções do select
        found = False
        for option in select.options:
            # Normaliza o texto da opção
            option_text_normalized = unicodedata.normalize('NFKD', option.text).encode('ascii', 'ignore').decode('utf-8').upper()
            if option_text_normalized == regime_normalized:
                select.select_by_visible_text(option.text)
                found = True
                break

        if not found:
            print(f"A empresa de CNPJ: {Cnpj} foi cadastrada sem Regime tributário pois o '{Regime}' não está de acordo com o apontado na planilha")

    except Exception as e:
        print(f"Erro ao selecionar o regime tributário: {e}")

    try:
        # Procurar o botão buscar
        buscarButton = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.ID, 'btCNPJ'))
        )
        
        # Insere a senha no campo
        buscarButton.click()
    except:
        print()
        
    
    try:
        # Espera o campo de ID da empresa aparecer
        idInput = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.NAME, 'EmpNewID'))
        )
        idInput.clear()
        
        # Insere o ID da empresa no campo
        idInput.send_keys(IdEmpresa)
        idInput.send_keys(Keys.TAB)

    except Exception as e:
        print("Erro ao processar o ID da empresa.")
        traceback.print_exc()
    
    # Espera o alerta aparecer indicando que o ID já está em uso
    try:
        alert_element = WebDriverWait(edge_driver, 5).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, '.swal2-popup.swal2-modal.swal2-show'))
        )

        # Verifica se o alerta contém a mensagem esperada
        if alert_element:
            okButton = WebDriverWait(edge_driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div[6]/div/div[3]/button[1]'))
            )
            okButton.click()
            idAleatorio = obter_id_aleatorio()
            idInput.clear()
            idInput.send_keys(idAleatorio)
            print(f"A empresa de CNPJ: {Cnpj} terá um id aleatorio de número {idAleatorio} pois o id {IdEmpresa} da planilha já está em uso por outro CNPJ.")
        else:
            print("Nenhum alerta de erro encontrado.")
    except Exception as e:
        print("Erro ao processar o ID da empresa.")
        traceback.print_exc()
    
    time.sleep(1)
    if InscEst and InscUf:
        try:
            # Espera o campo de IE da empresa aparecer
            ieInput = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.ID, 'EmpNewIE'))
            )
            
            # Insere o ID da empresa no campo
            ieInput.send_keys(InscEst)

        except Exception as e:
            print("Erro ao processar a inscrição estadual da empresa.")
            traceback.print_exc()
        
        try:
            # Encontra o seletor <select> pelo nome
            selectIEUF = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.NAME, 'EmpIEUF'))
            )

            # Cria uma instância de Select com o elemento encontrado
            select = Select(selectIEUF)
            
            # Normaliza o texto do InscUf para letras minúsculas
            normalized_insc_uf = InscUf.lower()
            
            # Normaliza o texto das opções para comparação
            for option in select.options:
                # Normaliza o texto da opção para letras minúsculas
                normalized_option_text = option.text.lower()
                if normalized_option_text == normalized_insc_uf:
                    select.select_by_visible_text(option.text)
                    break
            else:
                print('Não encontrei a UF correspondente.')
                print(f"UF informada: {InscUf}")
            
        except Exception as e:
            print(f"Empresa sem IEUF cadastrada")
        
        
        # Clicar no botão para add a IE
        try:
            addIEButton = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="EmpIEz"]/div[1]/div[3]/a'))
            )
            
            addIEButton.click()

        except Exception as e:
            print("Erro ao clicar no botão para add IE")
            traceback.print_exc()

    if Apelido:    
        #inserir cnpj
        try:
            # Espera o campo aparecer
            ApelidoEcont = WebDriverWait(edge_driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="EmpApelido"]'))
            )
            
            # Insere o cnpj no campo
            ApelidoEcont.send_keys(Apelido)
        except:
            print("Erro ao inserir o Apelido Econtínuo")
            traceback.print_exc()
    time.sleep(1)
    if NomeCtto and EmailCtto:
        # Separar os nomes e e-mails por '/'
        nomes_contatos = NomeCtto.split('/')
        emails_contatos = EmailCtto.split('/')

        # Iterar sobre cada nome e e-mail para adicionar como contato
        for i, (nome, email) in enumerate(zip(nomes_contatos, emails_contatos)):
            nome = nome.strip()  # Remove espaços em branco no início e no fim
            email = email.strip()

            try:
                # Localiza o ícone de adicionar contato
                    # Clicar no botão de adicionar novo contato se houver mais de um
                # Inserir nome do contato
                nomeCttoInput = WebDriverWait(edge_driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, 'CttNome_0'))
                )
                nomeCttoInput.clear()
                nomeCttoInput.send_keys(nome)

                # Inserir email do contato
                emailCttoInput = WebDriverWait(edge_driver, 10).until(
                    EC.visibility_of_element_located((By.NAME, 'CttEMail_0'))
                )
                emailCttoInput.clear()
                emailCttoInput.send_keys(email)

                time.sleep(1)
                try:
                    marcarDPTO = WebDriverWait(edge_driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, '//*[@id="dptoCtt_New_0"]/div[1]/div[1]/span/a[1]'))
                    )
                    marcarDPTO.click()
                except Exception as e:
                    print("Erro ao marcar todos os dptos")
                    traceback.print_exc()

                # Salvar contato usando a função JavaScript diretamente
                try:
                    edge_driver.execute_script("addCtt('0', true);")
                    time.sleep(1)
                except Exception as e:
                    print(f"Erro ao salvar o contato {nome}")
                    traceback.print_exc()

            except Exception as e:
                print(f"Erro ao inserir o contato {nome}: {e}")
                traceback.print_exc()

        

    teste=input('breakpoint')
    
edge_driver.quit()
    