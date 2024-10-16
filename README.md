#Demanda Preciso de uma automatização para PJE (processo judicial eletrônico)
#Basicamente, preciso extrair informações de determinada classe processual como nome,cpf,numero do processo,dentro outras informações
#Após isso, preciso que essas informações sejam lançadas em uma planilha excel
#Aguardo contato


#Passo 1- Entrar no site (https://pje-consulta-publica.tjmg.jus.br/)
#Passo 2- Clicar no campo OAB (Digitar o numero do OAB - Selecionar o Estado - Entrar em cada 1 dos processos e extrair (Numero do adv, Numero do Processo, Nome dos participantes))
#Passo 3 - Salvar dados em planilhas
#Passo 4 - Repetir até finalizar todos os processos daquele adv


from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.select import Select
import openpyxl
numero_oab=259155
planilha_dados_consulta=openpyxl.load_workbook('dados da consulta.xlsx')
planilha_processos=planilha_dados_consulta['processos']

#Passo 1- Entrar no site (https://pje-consulta-publica.tjmg.jus.br/)
driver=webdriver.Chrome()
driver.get('https://pje-consulta-publica.tjmg.jus.br/')
sleep(5)

#Passo 2- Clicar no campo OAB (Digitar o numero do OAB 
campo_numero_oab=driver.find_element(By.XPATH,"//input[@id='fPP:Decoration:numeroOAB']")
sleep(2)
campo_numero_oab.click()
sleep(1)
campo_numero_oab.send_keys(numero_oab)
#Passo 3 - Cliclar no campo Seleção do Estado SP
campo_estado_oab=driver.find_element(By.XPATH,"//select[@id='fPP:Decoration:estadoComboOAB']")
sleep(1)
opcoes_uf=Select(campo_estado_oab)
sleep(1)
opcoes_uf.select_by_visible_text('SP')
sleep(1)
#Passo 4 - Clicar no Botão de Pesquisar
botao_pesquisar=driver.find_element(By.XPATH,"//input[@id='fPP:searchProcessos']")
sleep(1)
botao_pesquisar.click()
sleep(5)
#Passo 5 - Clilar nos Botões dos Processos
links_abrir_processo=driver.find_elements(By.XPATH,"//a[@title='Ver Detalhes']")

#Passo 6 - Entrar em cada um dos processos
for link in links_abrir_processo:
    janela_principal=driver.current_window_handle
    link.click()
    sleep(5)
    janelas_abertas = driver.window_handles
    for janela in janelas_abertas:
        if janela not in janela_principal:
            driver.switch_to.window(janela)
            sleep(5)
            numero_processo = driver.find_elements(By.XPATH, '//div[@class="propertyView "]//div[@class="col-sm-12 "]')[0]
            participantes=driver.find_elements(By.XPATH,'//tbody[contains(@id,"processoPartesPoloAtivoResumidoList:tb")]//span[@class="text-bold"]')
            #Passo 9 guardar mais que um participante (se houver mais que um)
            lista_participantes=[]
            for participante in participantes:
                lista_participantes.append(participante.text)
            #Passo 8 guardar um participante(se houver apenas um)
            if len(lista_participantes)==1:
                planilha_processos.append([numero_oab,numero_processo.text,lista_participantes[0]])
            else:
                planilha_processos.append([numero_oab,numero_processo.text,','.join(lista_participantes)])
                # Passo 7 - Salvar dados em uma planilha
            planilha_dados_consulta.save('dados da consulta.xlsx')
            driver.close()
            #Passo 8 guardar um participante(se houver apenas um)
            #Passo 9 guardar mais que um participante (se houver mais que um)
# Passo 7 - Salvar dados em uma planilha
    driver._switch_to.window(janela_principal)
