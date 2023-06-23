import os
from time import sleep
from selenium.webdriver.common.by import By
import selenium
import shutil

def extracao():
    arquivos = os.listdir()
    print(arquivos)
    for zip in arquivos:
        pass
        #unzip_file(zip)

## clickPath serve para clicar em elementos da tela
def clickPath(path,navegador): # função para clicar em elemento xpath
    
    try:
        navegador.find_element(By.XPATH,path).click()
        
    except selenium.common.exceptions.NoSuchElementException:
        print('erro click')
        clickPath(path,navegador)
        
    except selenium.common.exceptions.ElementNotInteractableException:
        print('Elemento não interavel')  
        clickPath(path,navegador)  
   
## sendKey é a função ultilizada para enviar caracteres ou executar botoes na tela,como seta, enter etc     
def sendKey(path,navegador,key): # função para enviar string a tela
    
    try:
        navegador.find_element(By.XPATH,path).send_keys(key)
    
    except selenium.common.exceptions.NoSuchElementException:
        print("tentou escrever")
        sendKey(path,navegador,key)

    except selenium.common.exceptions.ElementNotInteractableException:
        print('not interactable')
        sendKey(path,navegador,key)

## findPath, usado para encontrar xpath na tela podendo o retornar a uma variavel, com tratamento prévio caso não apareça 
def findPath(path,navegador): # função para ver se path esta existente na tela
    
    try:
        return navegador.find_element(By.XPATH,path)
    
    except selenium.common.exceptions.NoSuchElementException:
       findPath(path,navegador) 
       print('Elemento não encontrado')
       sleep(2)
 
 
## Função responsavel por criar pasta em destino que foi passado na variavel DESTINO
## Caso haja uma pasta já criada com o mesmo nome, ele apagará a antiga e vai criar uma nova
## Atente que para ultilizar, o destino tem que ter o nome da pasta a ser criada no final da str a ser passada
## destino = f"C:\\Users\\t14\\Aquiles"     <- no ex a pasta aquiles nâo existia e foi criada 
def deleteAndCreateNewPath(destino):

    try: 
        if os.path.isdir(destino):
            shutil.rmtree(destino)
            os.mkdir(destino)

        else:
            os.mkdir(destino)
    
    except:
        print("Algo ocorreu, pasta não criada")