# Lembrar:
# Instalar programa python em python.org
# Instalar um editor de códigos, como VSCode (recomendo) e PyCharm
# Baixar o Chrome driver e extrair o arquivo .exe para a mesma pasta onde está
# salvo este programa (main.py)
# Link para baixar o Chrome Driver: https://chromedriver.chromium.org/downloads
# Instalar o Selenium: escrever no terminal "pip install selenium"

import time
from datetime import datetime
from selenium import webdriver

# Laço de repetição
while True:
    email = 'duvidaecontato@gmail.com' # Local para inserir o e-mail
    senha = '156324iT'                 # Local para inserir a senha

    # Caminho para o python executar o Chrome Driver
    DRIVER_PATH = 'C:/Users/Carlos/Desktop/WebScraper/chromedriver.exe'

    # Executando o navegador
    navegador = webdriver.Chrome(executable_path=DRIVER_PATH)

    # Setando uma página da web
    navegador.get('https://prenotami.esteri.it/')

    # Tela 1: insere e-mail e senha e clica no botão "Avanti"
    try:
        login = navegador.find_element("xpath",'//*[@id="login-email"]').send_keys(email)
        tela1 = 'ok'
    except:
        tela1 = 'erro'
        print('Erro ao inserir e-mail')

    try:
        password = navegador.find_element("xpath",'//*[@id="login-password"]').send_keys(senha)
        tela1 = 'ok'
    except:
        tela1 = 'erro'
        print('Erro ao inserir senha')

    try:
        submit = navegador.find_element("xpath",'//*[@id="login-form"]/button').click()
        tela1 = 'ok'
    except:
        tela1 = 'erro'
        print('Erro ao clicar no botão "Avanti"')

    # Tela 2: Clicando na aba "Reservar"
    if tela1 == 'ok':
        try:
            submit = navegador.find_element("xpath",'//*[@id="advanced"]/span').click()
            tela1 = 'ok'
        except:
            tela2 = 'erro'
        
    #Tela 3: Clicando no botão "Reservar"
    if tela2 == 'ok':
        try:
            tela3 = 'ok'
        except:
            tela3 = 'erro'
            print('Erro ao clicar no botão Reservar')


    '''if resultado == 'sem vagas':
        navegador.quit()

    else:
        alarme = 'disparar'

        # Chamando a data e a hora
        data_hora = datetime.now()
        texto_data_hora = data_hora.strftime('%d/%m/%Y %H:%M')
        print('Acesso dia ' + texto_data_hora + ' Horas')'''

    # Intervalo de 60 segundos entre uma execução e outra
    time.sleep(30)
    navegador.quit()
    time.sleep(30)

