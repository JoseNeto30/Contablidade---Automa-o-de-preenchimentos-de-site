# 1 - Navegar até o site : https://contabilidade-devaprender.netlify.app/
# 2 - Digitar e-mail
# 3 - Digitar Senha
# 4 - Clica em Entrar
# 5 - Clicar em cada campo e preencher com a informação extraida da planilha
# 6 - Clicar em Cadastrar
# 7 - Repetir passo 5 e 6 e 7

from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

driver = webdriver.Chrome()

# 1 - Navegar até o site : https://contabilidade-devaprender.netlify.app/
driver.get('https://contabilidade-devaprender.netlify.app/')
sleep(5)
email = driver.find_element(By.XPATH, "//input[@id='email']")
sleep(2)
email.send_keys('adm2@contabilidade.com')

# 3 - Digitar Senha
senha = driver.find_element(By.XPATH, "//input[@id='senha']")
sleep(2)
senha.send_keys('contabilidade123')

# 4 - Clica em Entrar
botao_entrar = driver.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(2)
botao_entrar.click()

# 5 - Extrair as informações da planilha
empresas = openpyxl.load_workbook('./empresas.xlsx')
paginas_empresas = empresas['dados empresas']

for linha in paginas_empresas.iter_rows(min_row=2, values_only=True):
    nome_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha

    driver.find_element(By.ID, 'nomeEmpresa').send_keys(nome_empresa)
    sleep(1)

    driver.find_element(By.ID, 'emailEmpresa').send_keys(email)
    sleep(1)

    driver.find_element(By.ID, 'telefoneEmpresa').send_keys(telefone)
    sleep(1)

    driver.find_element(By.ID, 'enderecoEmpresa').send_keys(endereco)
    sleep(1)

    driver.find_element(By.ID, 'cnpj').send_keys(cnpj)
    sleep(1)

    driver.find_element(By.ID, 'areaAtuacao').send_keys(area_atuacao)
    sleep(1)

    driver.find_element(By.ID, 'numeroFuncionarios').send_keys(
        quantidade_de_funcionarios)
    sleep(1)

    driver.find_element(By.ID, 'dataFundacao').send_keys(data_fundacao)
    sleep(1)

    driver.find_element(By.ID, 'Cadastrar').click()
    sleep(3)

# 6 - Clicar em cada campo e preencher com a informação da planilha.
# 7 - Clicar em Cadastrar
# 8 - Repetir passo 5 e 6 e 7
