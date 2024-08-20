Este código utiliza a biblioteca Selenium para automatizar o preenchimento e emissão de notas fiscais com base em dados fornecidos em uma planilha Excel. O código realiza login em um sistema web, preenche um formulário com as informações de cada cliente e emite a nota fiscal correspondente.

Configuração e Inicialização
Importação das Bibliotecas: O código importa as bibliotecas necessárias, incluindo selenium para automação do navegador e pandas para manipulação de dados.

Configuração do Navegador: São definidas opções para o navegador Chrome, incluindo configurações para o diretório de download e preferências de segurança.

Inicialização do WebDriver: O WebDriver do Chrome é configurado e iniciado com as opções definidas. O caminho atual do diretório é armazenado para referência.

Acesso e Login no Sistema
Carregamento da Página de Login: O código abre uma página HTML local contendo o formulário de login.

Preenchimento do Login e Senha: Insere as credenciais de login (e-mail e senha) no formulário e clica no botão de login para acessar o sistema.

Processamento e Emissão de Notas Fiscais
Importação da Planilha: O código lê a planilha NotasEmitir.xlsx, que contém os dados dos clientes para os quais as notas fiscais devem ser emitidas.

Preenchimento do Formulário: Para cada linha da planilha, o código preenche os campos do formulário de nota fiscal com as informações do cliente, incluindo:

Nome/Razão Social
Endereço
Bairro
Município
CEP
UF
CPF/CNPJ
Inscrição Estadual
Descrição
Quantidade
Valor Unitário
Valor Total
Emissão da Nota Fiscal: Após preencher os campos, o código clica no botão para emitir a nota fiscal.

Recarregamento da Página: A página é recarregada após cada emissão para limpar o formulário e preparar para o próximo cliente.

Finalização
Fechamento do Navegador: Após completar o preenchimento e emissão das notas fiscais para todos os clientes, o navegador é fechado.

Código:

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
import os

# Abrir o navegador
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\euric\Documents\Códigos Py\Selenium",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=options)
caminho = os.getcwd()

# entrar na página de login 
arquivo = caminho + r"\login.html"
navegador.get(arquivo)

# preencher o login e a senha
navegador.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys("lira@gmail.com")
navegador.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys("123456")
# clicar no botao de fazer login
navegador.find_element(By.XPATH, '/html/body/div/form/button').click()

# importar a base de clientes
tabela = pd.read_excel("NotasEmitir.xlsx") 

# Rodar o processo de emissão de nota fiscal para cada cliente da tabela

for linha in tabela.index:
    # preencher os dados da NF
    
    # nome/razao social
    navegador.find_element(By.NAME, 'nome').send_keys(tabela.loc[linha, "Cliente"])

    # endereco
    navegador.find_element(By.NAME, 'endereco').send_keys(tabela.loc[linha, "Endereço"])

    # bairro
    navegador.find_element(By.NAME, 'bairro').send_keys(tabela.loc[linha, "Bairro"])

    # municipio
    navegador.find_element(By.NAME, 'municipio').send_keys(tabela.loc[linha, "Municipio"])

    # cep
    navegador.find_element(By.NAME, 'cep').send_keys(str(tabela.loc[linha, "CEP"]))
    
    # UF
    navegador.find_element(By.NAME, 'uf').send_keys(tabela.loc[linha, "UF"])
    
    # CPF/CNPJ
    navegador.find_element(By.NAME, 'cnpj').send_keys(str(tabela.loc[linha, "CPF/CNPJ"]))

    # Inscricao estadual
    navegador.find_element(By.NAME, 'inscricao').send_keys(str(tabela.loc[linha, "Inscricao Estadual"]))

    # descrição
    texto = tabela.loc[linha, "Descrição"]
    navegador.find_element(By.NAME, 'descricao').send_keys(texto)

    # quantidade
    navegador.find_element(By.NAME, 'quantidade').send_keys(str(tabela.loc[linha, "Quantidade"]))

    # valor unitario
    navegador.find_element(By.NAME, 'valor_unitario').send_keys(str(tabela.loc[linha, "Valor Unitario"]))

    # valor total
    navegador.find_element(By.NAME, 'total').send_keys(str(tabela.loc[linha, "Valor Total"]))
    
    # clicar em emitir nota fiscal
    navegador.find_element(By.CLASS_NAME, 'registerbtn').click()
    
    # recarregar a página para limpar o formulário
    navegador.refresh()

# fechar navegador
navegador.quit()

