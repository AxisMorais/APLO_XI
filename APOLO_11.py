import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
from bs4 import BeautifulSoup
import time
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
import pyautogui

pyautogui.position()

# Leitura do arquivo Excel e criação do DataFrame
diretorio = 'C:/Users/pres00310855/Desktop/Integracao/Equipamentos_SMASSAC.xlsx'
armazenador = pd.read_excel(diretorio)
data_frame = pd.DataFrame(armazenador)

# Dicionário de mapeamento de valores
mapeamento_valores = {
    'Central de Abastecimento da Agricultura Familiar e Urbana': '53',
    'Centro Especializado de Atendimento à Mulher': '54',
    'Centro de Referênciay': '55',
    'Conselho Municipal de Assistência Social': '70',
    'Conselho Tutelar': '56',
    'Feira Coberta': '58',
    'Mercado Distrital': '59',
    'Núcleo de Atendimento às Medidas Socioeducativas e Protetivas': '60',
    'Plantão do Conselho': '61',
    'Refeitório Popular': '62',
    'Restaurante Popular': '73',
    'Abrigo': '74',
    'Albergue': '75',
    'Banheiro Público': '69',
    'Centro de Apoio Comunitário': '68',
    'Ciame': '66',
    'Horta Comunitária': '67',
    'República': '63',
    'Unidade de Acolhimento Infantil': '64',
    'Unidade de Acolhimento Institucional': '65'
}

# CRIACAO DE UMA VARIAVEL PARA CONTAR QUANTAS UNIDADES FORAM VINCULADAS
cadastrados = 0

# CRIACAO DE UMA VARIAVEL PARA CONTAR QUANTAS UNIDADES NÃO FORAM VINCULADAS
naoCadastrados = 0

# CASO OCORRA ALGUMA QUEBRA O ALGORITIMO RETOMA A CONTAGEM INICIANDO NO ULTIMO PONTO DE PARADA WHILE
pontoDeParadaWhile = 0

# VARIAVEL QUE IDENTIFICA O INICIO DA EXECUÇAO PARA FINS DE ANALISE DE PERFORMANCE
start = time.time()

# ENQUANTO O PONTO DE PARA FOR DIFERENTE DE 2030, CONTINUE TENTANDO REALIZAR O CADASTRO

#  INICIO DA CRONOMETRAGEM DA EXECUCAO EM SEGUNDOS
start = time.time()
while (pontoDeParadaWhile != 88):  # A planilha inicia na linha 2, temos um total de 90 linhas (90 - 2 = 88)

    for x in range(pontoDeParadaWhile, 89):  # 88 linhas + a linha zero = 89 linhas a percorrer no for

        # INDICADOR DE QUEBRA MOSTRANDO ONDE QUEBROU E ONDE DEVE RETORNAR:
        print(pontoDeParadaWhile, "Ao quebrar retorne a partir do: ", pontoDeParadaWhile)

        # CONFIGURANDO O TAMANHO DA JANELA 1000 POR 1000
        options = Options()
        options.add_argument('window-size=1000,1000')

        # INSERINDO  AS CONFIGURAÇOES DE TAMANHO PELO OPTIONS NA VARIAVEL NAVEGADOR
        navegador = webdriver.Chrome(options=options)
        navegador.get('http://cic.pbh')

        try:
            # ISERIR LOGIN PARA # ACESSAR O SITE DO CIC:
            navegador.find_element(By.NAME, 'josso_username').send_keys('thiago.conegundes')
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERIR A SENHA
            navegador.find_element(By.NAME, 'josso_password').send_keys('Th1505@')
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NO BOTAO PARA ACESSAR SISTEMA
            navegador.find_element(By.CLASS_NAME, "botao").click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO GERAL
            navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/a').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO UNIDADE
            navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/ul/li[2]/a').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICAR NA OPCAO NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="novo"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        # -------------------------------------------------------------------------
        # PEGANDO OS DADOS DO DATA FRAME
        valorLinha = pontoDeParadaWhile

        nome = data_frame.at[valorLinha, 'UNIDADE']

        tipo = data_frame.at[valorLinha, 'TIPO_UNIDADE']

        log = data_frame.at[valorLinha, 'TIPO_LOGRADOURO']

        nomeLog = data_frame.at[valorLinha, 'NOME_LOGRADOURO']

        bairro = data_frame.at[valorLinha, 'BAIRRO']

        numlog = data_frame.at[valorLinha, 'NUMERO']

        time.sleep(5)

        # ---------------------------------------------------------------------------------

        try:
            # INSERIR A UNIDADE NO CAMPO NOME UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-nome"]').send_keys(nome)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(3)

        try:
            # CLICAR NA OPCAO TITULARIDADE
            navegador.find_element('xpath', '//*[@id="mestre-titularidade"]/option[4]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(3)

        try:
            # CLICAR NA OPCAO TIPO DE UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-tipo_unidade"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(5)

        try:
            # INSERIR O TIPO DE UNIDADE
            navegador.find_element('xpath', '//*[@id="mestre-tipo_unidade"]').send_keys(tipo)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        time.sleep(5)

        try:
            # CLICAR NA OPCAO LOGRADOURO
            navegador.find_element('xpath', '//*[@id="aba-endereco"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(5)

        try:
            # CLICAR NO BOTAO NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="novo"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # CLICANDO NO BOTAO PARA CADASTRAR NOVA UNIDADE
            navegador.find_element('xpath', '//*[@id="detalhe-1-vinculado"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(2)

        try:
            # ENTRANDO NO I-FRAME PARA INSERIR DOS DADOS
            navegador.switch_to.frame('ifDlgEndereco')

            tipoLogradouro = 'RUA'
            # INSERINDO O LOGRADOURO
            if log == tipoLogradouro:
                # ISERINDO O TIPO DE LOGRADOURO "RUA"
                # Seleciona a opção RUA
                navegador_II = navegador.find_element(By.ID, 'tipoLogradouro')
                time.sleep(2)
                seletor = Select(navegador_II)
                seletor.select_by_value('RUA')
                navegador.find_element(By.XPATH, '//*[@id="logradouro"]').send_keys(nomeLog)
            else:
                # ISERINDO O TIPO DE LOGRADOURO AVENIDA
                navegador_II = navegador.find_element(By.ID, 'tipoLogradouro')
                time.sleep(2)
                seletor = Select(navegador_II)
                seletor.select_by_value('AVE')
                navegador.find_element(By.XPATH, '//*[@id="logradouro"]').send_keys(nomeLog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # TRANSFORMANDO O LOGRADOURO EM NUMERO
            numlog = int(numlog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERINDO O NUMERO INICIAL DO LOGRADOURO NO IFRAME
            navegador.find_element('xpath', '//*[@id="numeroInicial"]').send_keys(numlog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # PRENCHENDO O NUMERO FINAL NO IFRAME
            navegador.find_element('xpath', '//*[@id="numeroFinal"]').send_keys(numlog)
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            # INSERINDO O BAIRRO
            navegador.find_element('xpath', '//*[@id="bairro"]').send_keys(bairro)

        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break
        try:
            # CLICAR NA OPCAO PESQUISAR
            navegador.find_element('xpath', '//*[@id="pesquisar"]').click()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)


        # ESPERANDO 2 SEGUNDOS PARA CARREGAR A PÁGINA
        time.sleep(2)

        try:
            #CLICANDO NA TABELA ONDE APRESENTA O RESULTADO DO ENDEREÇO APRESENTADO PELO SISTEMA
            navegador.find_element('xpath', '//*[@id="conteudo"]/table/tbody/tr[2]').click()
            time.sleep(2)
        except:
            pontoDeParadaWhile = pontoDeParadaWhile + 1
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            navegador.quit()
            break

        try:
            # FINALIZANDO O IFRAME
            navegador.switch_to.default_content()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        time.sleep(2)

        try:
            #CLICANDO NA OPCAO PRINCIPAL
            navegador.find_element('xpath', '// *[ @ id = "detalhe-1-principal"]').click()

        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            break

        try:
            pontoDeParadaWhile = pontoDeParadaWhile + 1
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            pontoDeParadaWhile = pontoDeParadaWhile + 1
            break


        time.sleep(3)

        try:
            # CLICANDO NA OPCAO GRAVAR
            navegador.find_element('xpath', '//*[@id="gravar"]').click()
            time.sleep(3)
            navegador.quit()
        except:
            print("houve uma quebra nesse ponto: ", pontoDeParadaWhile)
            navegador.quit()
            break


end = time.time()

# APRESENTANÇÃO DO RELATÓRIO TEMPO DE EXECUÇÃO E QUATIDADE DE VINCULADOS
print('Tempo total em segundos: ')
print(end - start)
print('--------------')
print('Quantidade de unidade cadastradas:', cadastrados)


