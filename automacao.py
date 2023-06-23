### Caso a extraçao ocorra fora da data prevista na tabela Calendario semanal
### Inserir na linha 5 a ultima data de quando as bases foram disponibilizadas de acordo a tabela, 
### Caso faça na data correta, apenas descomentar a linha 6
from datetime import date
dateNow = date(2023,8,30)
#dateNow = date.today()

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import datetime
from time import sleep
import pandas as pd
import selenium
import os
import shutil
import zipfile
from os import listdir
from os.path import isfile, join
import sys
import traceback
import csv

import cx_Oracle

from Dependencias.functionScrap import *


#Conexão com o Banco
oracle_pass = os.environ.get('ORACLEPASS')
oracle_user = os.environ.get('ORACLEUSER')
conn = cx_Oracle.connect(f'{oracle_user}/{oracle_pass}xx_high')
cursor = conn.cursor()


userSystem = os.getlogin()
options = Options()
#options.add_argument('window-size=400,800')
options.add_experimental_option("detach",True)
#options.add_argument('headless')
global navegador 
navegador = webdriver.Chrome(executable_path=r'C:\Scripts\chromedriver.exe',options=options)
mouse = webdriver.ActionChains(navegador)
navegador.implicitly_wait(2)

print("Execução Script ppppp")

## ler arquivo responsavel por escrever as datas corretas nos arquivos
dateTable = pd.DataFrame(pd.read_excel(r'.\\Dependencias\\calendarioNielsinFormatado.xlsx'))
dateTableW = dateTable['Liberação da base yyyy Semanal'] 
########

xxxxxxLOGIN = os.environ.get("xxLOGIN")
xxxxxxPASS = os.environ.get("xxPASS")

navegador.get('https://xxxxxxxxxx')

listaPastas = ['cccc',
               'Pizzas (vvvvvv)',
               'Industralizados (bbbbbb)',
               'zzzz'] # lista de pastas que serão criadas para organização da extração

listaPastasExtracao = ['//*[text()="zzzzz - zzzz"]',
                       '//*[text()="zzzzz - cccc"]',
                       '//*[text()="zzzzz - vvvvvv"]',
                       '//*[text()="zzzzz - bbbbbb"]']# lista com endereços xpath das pastas a serem abertas na extração

listasSubPastasExtracao = ['//*[@id="N10050_item"]',
                           '//*[@id="N10164_item"]',
                           '//*[@id="N1011F_item"]',
                           '//*[@id="N100AC_item"]']# lista com endereços xpath das subpastas a serem abertas na extração


listaNomeArquivos = [ 'F1121304C.zip','F1123B04_SFF.zip',
                     'F19123B24C.zip','F19123B24_SFF.zip',
                     'F1123B19C.zip','F198123B19_SFF.zip',
                     'F200J123C.zip','F21230JB03_SFF.zip',
                     'F012301C.zip','F0012301_SFF.zip']# nome dos arquivos esperados no download



listaElementosParaDownload = [['yyyy CIP - oo Cong (F119JB04C)','SCT CIP - oo Con (F119JB04_SFF)'],
                              ['yyyy CIP - cccc (F006JB01C)','yyyy - cccc (F006JB01_SFF)'],
                              ['F223JB03C.zip','SCT - vvvvvv (F200JB03_SFF)'],
                              ['F198JB19C.zip','SCT CIP - Indu2323ri32ado de Carne (F198JB19_SFF)',
                              'SCT CIP - Industrializado de Carne (F1923JB24C)','SCT CIP - Indu23do de23ne (F192324_SFF)']]# elementos buscados na tela para poder executar o click e baixar o arquivo, aqui se trata de uma matriz

listaEnderecoOrganizacao = [ "C:\\Users\\{}\\Downloads\\ppppp\\zzzz\\".format(userSystem),
                            "C:\\Users\\{}\\Downloads\\ppppp\\Industralizados (bbbbbb)\\".format(userSystem),
                          "C:\\Users\\{}\\Downloads\\ppppp\\Industralizados (bbbbbb)\\".format(userSystem),
                          "C:\\Users\\{}\\Downloads\\ppppp\\Pizzas (vvvvvv)\\".format(userSystem),
                          "C:\\Users\\{}\\Downloads\\ppppp\\cccc\\".format(userSystem)] # respectiva pasta onde os arquivos vão ser organizados


#Função para tirar os arquivos dentro de uma pasta zipada    
def unzip_file(file_name):
    try:
        with zipfile.ZipFile(file_name, 'r') as zip_ref:
            zip_ref.extractall('temp')
        new_file_name = [join('temp', f) for f in listdir('temp') if isfile(join('temp', f))][0]
        return new_file_name
    except Exception as e:
        print("Exception (unzip_file): {}".format(str(e)))
        print(traceback.format_exc())
        sys.exit(1)

#funcao para para pegar só os arquivos qe precisa ser deszipados
def extracao():
    arquivos = os.listdir()
    print(arquivos)
    for zip in arquivos:
        unzip_file(zip)
        
def pathCreate():
    os.makedirs('./ppppp/')
    for nome in range(0,4):
        os.mkdir('./ppppp/{}'.format(listaPastas[nome]))

from selenium.common.exceptions import NoSuchElementException
def enterToFrame(valor):  # Função responsavel por navegar entre os Iframe da página
    navegador.switch_to.default_content() # ele sempre volta ao iframe principal, e navega a um iframe geral para depois pular para frame correto, sendo o da esquerda com icones de navegação, e o da direita para download dos arquivos
    match valor:
        case 'principal':
            sleep(2)
            navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ogdp_frameset"]'))
        case 'esquerda':
            try:
                navegador.switch_to.default_content()
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ogdp_frameset"]'))
                navegador.switch_to.frame(navegador.find_element(By.XPATH,"//frame[@name='ws_left']"))
            except NoSuchElementException:
                navegador.switch_to.default_content()
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ogdp_frameset"]'))
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ws_frameset"]/frame[1]'))
                
        case 'direita':
            try:
                navegador.switch_to.default_content()
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ogdp_frameset"]'))
                navegador.switch_to.frame(navegador.find_element(By.XPATH,"//frame[@name='ws_right']"))
            except NoSuchElementException:
                navegador.switch_to.default_content()
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ogdp_frameset"]'))
                navegador.switch_to.frame(navegador.find_element(By.XPATH,'//*[@id="ws_frameset"]/frame[2]'))
                
        case _:
            pass
        
def elementDownload(listaElementosParaDownload): # Função para executar o click nos elementos da lista e executar o download
    
    enterToFrame('direita')
    
    textosLinks = navegador.find_elements(By.TAG_NAME,'tr')
    sleep(2)
    for item in textosLinks:
        
        qualquerVariavel = item.text
    
        if qualquerVariavel.find(listaElementosParaDownload) != -1: # primeiramente os elementos da respectiva pagina é aberto, logo é armazenado em uma variavel
            item.find_element(By.XPATH,'td/a').click()              # jogado em um for para navega entre todos o valores dessa lista criado, para quando o nome especifico
            print(item.text)                                        #do arquivo necessário for encontrado, executar o click
     
       
## na pagina onde é digitado o login e senha, esses comandos são responsaveis por navegar entre eles enviando os dados
sleep(2)
botao = findPath("//*[@id='form19']/div[2]/input",navegador)
sendKey("//*[@id='input27']",navegador,xxxxxxLOGIN)
mouse.move_to_element(botao).perform()
mouse.click().perform()

sleep(2)
botao = findPath("//*[@value='Verificar']",navegador)
sendKey('//*[@type="password"]',navegador,xxxxxxPASS)
mouse.move_to_element(botao).perform()
mouse.click().perform()



# Nas abas da página, aqui o driver navega em mycontent > databases
enterToFrame('')
botao = findPath("//a[text()='My Content']",navegador)
sleep(0.1)
mouse.move_to_element(botao).perform()
sleep(0.1)
mouse.perform()
botao = findPath("//a[text()='Databases']",navegador)
sleep(0.1)
mouse.move_to_element(botao).perform()
sleep(0.1)
mouse.click().perform()

for local in range(0,4):
    enterToFrame('esquerda')
    # Entra na pasta específica da extração
    botao1 = findPath(listaPastasExtracao[local],navegador)
    mouse.move_to_element(botao1).perform()
    mouse.click().perform()
    sleep(2)
    
    ## Entra em yyyy
    botao2 = findPath(listasSubPastasExtracao[local],navegador)
    mouse.move_to_element(botao2).perform()
    mouse.click().perform()
    sleep(2)
    
    #for linha in range(0,len(listaElementosParaDownload)):
    for coluna in range(0,len(listaElementosParaDownload[local])):
        print(listaElementosParaDownload[local][coluna])
        elementDownload(listaElementosParaDownload[local][coluna])
        print('Download iniciado')
        sleep(10)
    # Fecha pasta aberta 
    enterToFrame('esquerda')   
    mouse.move_to_element(botao1).perform()
    mouse.click().perform()
    
## trecho responsavel por pegar o valor da data da extraçao, e colocar o valor correto nos arquivos referente a tabela CalendarioNielsen.xlsx
valueList = 0


for data in range(0, len(dateTable)):
    #print(dateTable.iloc[data])
    if dateNow == dateTable['Liberação da base yyyy Semanal'].iloc[data]:
        valorAEscrever = dateTable['Dt. Final Domingo'].iloc[data].strftime('%d.%m')
        semanaEx = dateTable['Semana'].iloc[data].replace('W','S')
        print(dateTable['Liberação da base yyyy Semanal'].iloc[data])
        
print(valorAEscrever, ' Data a ser escrita na semana"', semanaEx)

 ## Trecho responsável por esperar o download dos arquivos
os.chdir(r"C:\\Users\\{}\\Downloads".format(userSystem))## Definido onde todo o código vai execurtar, como se estivesse dentro da pasta downloads
lista = os.listdir('.') ## lista tudo na pasta referida e armazena em lista
lista = os.listdir(r'C:\\Users\\{}\\Downloads'.format(userSystem))
var = 0
while var == 0: 
    if all(valores in lista for valores in listaNomeArquivos): 
        print(listaNomeArquivos)
        print(lista)

        var = 1
        print('Todos os arquivos presentes.')
        #Com todos os aquivos presentes dentro de Download, tira os arquivos dentro da pasta zip e bota na pasta temp
        
        #Precisa Excluir a past ja existente com os arquivos ja deszipados para fazer uma nova
        if os.path.isdir(r'C:\\Users\\{}\\Downloads\\temp'.format(userSystem)):
            shutil.rmtree(r'C:\\Users\\{}\\Downloads\\temp'.format(userSystem))
        
        extracao()
        
    else:
        sleep(200) 
        print('Aguardando download...')
        lista = os.listdir(r'C:\\Users\\{}\\Downloads'.format(userSystem))

# Apaga qualquer pasta existente no diretorio chamada ppppp e cria outra
dir = './ppppp'

print("Criando pastas")
try:
    if os.path.isdir(dir):
        shutil.rmtree(dir)
        pathCreate()
    else:
        pathCreate()
except PermissionError:
        pathCreate()

# Coloca no nome dos arquivos especificos a data relativa a coluna da extração
origem = r"C:\\Users\\{}\\Downloads\\".format(userSystem)

for source in listaNomeArquivos:
    print(source)
    if source in lista: ## Encontra arquivo especifico e adiciona a data para seu nome

        old_name = r"C:\\Users\\{}\\Downloads\\{}".format(userSystem,source)
        new_name = r"C:\\Users\\{}\\Downloads\\{} - {} {}.zip".format(userSystem,source.replace('.zip',''),semanaEx, valorAEscrever)  
        os.rename(old_name,
                  new_name) 
        new_name = "{} - {} {}.zip".format(userSystem,source.replace('.zip',''),semanaEx, valorAEscrever)   
        #sleep(2)
        print(new_name)

print("Organizando arquivos")
## Bloco responsável por organizar os arquivos de acordo a suas respectivas pastas para o ftp
arquivo = 0
endereco = 0
i=0
origem = r"C:\\Users\\{}\\Downloads\\".format(userSystem)
limite =  len(listaNomeArquivos)
while arquivo < limite:
    try:
        arquivo1 = "{} - {} {}.zip".format(listaNomeArquivos[arquivo].replace('.zip',''),semanaEx, valorAEscrever)
        arquivo2 =  "{} - {} {}.zip".format(listaNomeArquivos[arquivo+1].replace('.zip',''),semanaEx, valorAEscrever)  
        if (arquivo1) or (arquivo2) in lista:
            
            destino = r"{}".format(listaEnderecoOrganizacao[endereco])

            print(destino)
            
            new_name = "{} - {} {}.zip".format(listaNomeArquivos[arquivo].replace('.zip',''),semanaEx, valorAEscrever)  
            try:
                shutil.move(origem + new_name ,destino + new_name)
                print(new_name, ' movido para ', destino)
            
            except:
                print('Arquivo ', new_name ,' não encontrado na pasta download')
                pass
        
            new_name = "{} - {} {}.zip".format(listaNomeArquivos[arquivo+1].replace('.zip',''),semanaEx, valorAEscrever)   
            try:
                shutil.move(origem + new_name ,destino + new_name)
                print(new_name, ' movido para ', destino)
            
            except:
                print('Arquivo ', new_name ,' não encontrado na pasta download')
                pass
            
            lista = os.listdir(r'C:\\Users\\{}\\Downloads'.format(userSystem))
            endereco += 1
            arquivo += 2
            sleep(1)
            
    except IndexError: ### :) confia
        pass
    sleep(2)
    i += 2
