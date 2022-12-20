from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl
import time


linkDasEmpresas = []
hrefList = []
dados = []
listaGeral = []
driver = webdriver.Chrome()

book = openpyxl.load_workbook('planilha de ju.xlsx')
dados_empresas = book['planilhaJu']


def scrollDown():
    # scroll down
    driver.execute_script(
        'window.scrollTo(0, document.body.scrollHeight);')
    time.sleep(1.5)
    # scrollup
    driver.execute_script('window.scrollTo(0, document.body.scrollTop);')


def pegarTodosOsLinks():
    # pega todos os links disponiveis no site e joga numa lista
    linkList = driver.find_elements(By.TAG_NAME, 'a')
    for item in linkList:
        x = item.get_attribute('href')
        hrefList.append(x)
        time.sleep(0.3)


def pegarLinksDasEmpresas():#filtra todos os links do site e seleciona apenas os links das empresas e manda eles pra uma lista
    
    for i in range(110, 130):
        print(hrefList[i])
        linkDasEmpresas.append(hrefList[i])
        time.sleep(0.3)
    
    hrefList.clear()#limpa a lista de todos os links do site para que a execuçao na proxima pagina nao seja danificada






def pegarTelefone():#pega todos os telefones da e joga para a lista
    try:
        try:
            telefone = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[5]/div/div[2]/p[2]').get_attribute('innerHTML')

            dados.append(telefone)
        except:
            telefone = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[4]/div/div[2]/p[2]').get_attribute('innerHTML')
            dados.append(telefone)
    except:
        dados.append('telefone nao encontrado')


def pegarQtdeClientes():#pega a quantidade de clientes e joga para lista
    try:
        try:
            cliente = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[3]/div/div[2]/ul/li[2]/div[2]/p').get_attribute('innerHTML')

            dados.append(cliente)
        except:
            cliente = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[2]/div/div[2]/ul/li[2]/div[2]/p').get_attribute('innerHTML')

            dados.append(cliente)
    except:
        dados.append('qnt de clientes nao foi encontrada')
    


def pegarNomeDaEmpresa():#pega o nome da empresa e joga para lista
    nomeDaEmpresa = driver.find_element(
        By.CLASS_NAME, 'case27-primary-text').text
    dados.append(nomeDaEmpresa)

def pegarEndereco():#pega o nome da empresa e joga para list
    try:
        endereco = driver.find_element(
            By.CLASS_NAME, "map-block-address").get_attribute('innerText')
        dados.append(endereco)
    except:
        dados.append('endereço nao encontrado')



def pegarColaboradores():#pega a quantidade de colaboradores e joga para lista
    try:
        try:
            colaboradores = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[2]/div/div[2]/ul/li[3]/div[2]/p').get_attribute('innerHTML')
            dados.append(colaboradores)

        except:
            colaboradores = driver.find_element(
                By.XPATH, '/html/body/div[1]/div[2]/div[4]/section/div/div/div[1]/div/div[3]/div/div[2]/ul/li[3]/div[2]/p').get_attribute('innerHTML')
            dados.append(colaboradores)
    except:
        dados.append('colaboradores nao encontrados')




for i in range(14):
    driver.get('https://ecossistema.pe/busca/?pg={}&sort=latest'.format(i))
    scrollDown()
    time.sleep(2)
    pegarTodosOsLinks()
    time.sleep(3)
    pegarLinksDasEmpresas()
    time.sleep(3)
    for item in linkDasEmpresas:
        driver.get('{}'.format(item))
        time.sleep(3)
        pegarNomeDaEmpresa()
        time.sleep(1)
        pegarColaboradores()
        time.sleep(1)
        pegarQtdeClientes()
        time.sleep(1)
        pegarTelefone()
        time.sleep(1)
        pegarEndereco()
        time.sleep(1)
        listaGeral.append(dados)
        print('dados atualizados')
        print(dados)
        time.sleep(0.5)
        dados_empresas.append(
            [dados[0], dados[1], dados[2], dados[3], dados[4]])
        dados.clear()
        time.sleep(0.3)
    linkDasEmpresas.clear()
    book.save('planilha de ju.xlsx')


time.sleep(3)


book.save('planilha de ju.xlsx')
