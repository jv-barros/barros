# new automate to xmenu 
# automacao funcionando 
# alterar nomenclatura de arquivo 

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from pathlib import Path
import os
import time
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from datetime import datetime, timedelta
from dotenv import load_dotenv
import logging
from selenium.webdriver.common.action_chains import ActionChains

# new function to close browser
def close_browser(navegador):
    navegador.quit()

def caminho_ultimo_arquivo_baixado():
    # Define o caminho relativo ao script para a pasta 'downloads'
    current_directory = os.path.dirname(os.path.realpath(__file__))
    caminho = os.path.join(current_directory, 'downloads')
    # Cria a pasta se ela não existir
    if not os.path.exists(caminho):
        os.makedirs(caminho)

    lista_arquivos = os.listdir(caminho)
    if not lista_arquivos:
        return None  # Retorna None se não houver arquivos

    lista_datas = [(os.path.join(caminho, arquivo), os.path.getmtime(os.path.join(caminho, arquivo))) for arquivo in lista_arquivos]
    lista_datas.sort(key=lambda x: x[1], reverse=True)
    return lista_datas[0][0]  # Retorna o caminho do arquivo mais recente

def ultimo_arquivo_baixado():
    current_directory = os.path.dirname(os.path.realpath(__file__))
    download_path = os.path.join(current_directory, 'downloads')
    lista_arquivos = os.listdir(download_path)
    
    if not lista_arquivos:  # Verifica se a lista de arquivos está vazia
        logging.warning("Nenhum arquivo encontrado na pasta downloads.")
        return None  # Retorna None ou você pode escolher retornar uma mensagem específica
    
    lista_datas = [(Path(download_path, arquivo).stat().st_mtime, arquivo) for arquivo in lista_arquivos]
    lista_datas.sort(reverse=True)
    
    if lista_datas:  # Verifica se a lista_datas não está vazia após o sort
        return lista_datas[0][-1]
    else:
        logging.error("Erro ao tentar classificar os arquivos por data.")
        return None

def setup_browser():
    # Defina o caminho de downloads relativo ao diretório onde o script está sendo executado
    current_directory = os.path.dirname(os.path.realpath(__file__))
    download_path = os.path.join(current_directory, 'downloads')
    
    # Crie a pasta de downloads se ela não existir
    if not os.path.exists(download_path):
        os.makedirs(download_path)
    
    options = Options()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
        "safebrowsing.enabled": "false"
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("detach", True)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return navegador

def wait_and_find_element(navegador, by, value):
    """
    Wait for an element to be present and return it.

    Args:
    navegador: WebDriver instance to interact with.
    by (By): The type of strategy to locate the element (e.g., By.XPATH, By.ID).
    value (str): The value of the locator (e.g., the actual XPath or ID).

    Returns:
    WebElement: The found element.

    Raises:
    TimeoutException: If the element is not found within the specified timeout.
    """
    try:
        logging.info(f"Waiting for element by {by} with value '{value}'...")
        element = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((by, value)))
        logging.info("Element found.")
        return element
    except TimeoutException:
        logging.error(f"Failed to find element by {by} with value '{value}'.")
        raise TimeoutException(f"Element by {by} with value '{value}' could not be found within the timeout period.")

def excluir_arquivo(local_file_path):
    nome_arquivo = os.path.basename(local_file_path)  # Extrai apenas o nome do arquivo
    try:
        os.remove(local_file_path)
        logging.info(f"Arquivo '{nome_arquivo}' excluído com sucesso do sistema local.")
    except OSError as e:
        logging.error(f"Erro ao excluir o arquivo '{nome_arquivo}': {e}")
    logging.info("--------------------------------------------------")

def upload_to_sharepoint(sp_folder, local_file_path, sharepoint_file_name):
    nome_arquivo = os.path.basename(local_file_path)
    print(f"Iniciando o upload do arquivo '{nome_arquivo}' para o SharePoint.")
    
    with open(local_file_path, 'rb') as file_input:
        sp_folder.upload_file(file_input, sharepoint_file_name)
    
    print(f"Arquivo '{nome_arquivo}' enviado como '{sharepoint_file_name}' no SharePoint.")
    print("--------------------------------------------------")

def fechar_pop_up_xmenu(navegador):
    try:
        pop_up_close_button = WebDriverWait(navegador, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="onesignal-slidedown-cancel-button"]'))
        )
        pop_up_close_button.click()
        logging.info("Pop-up fechado com sucesso.")
    except:
        logging.info("Pop-up não encontrado ou já fechado.")

def baixar_e_upload_relatorio_xmenu(navegador):
        navegador.get("https://portal.netcontroll.com.br/#/auth/login")
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "email")))
        navegador.find_element(By.ID, 'email').send_keys("felipe@orionbi.com.br")
        navegador.find_element(By.ID, 'password').send_keys("Orion123!")
        navegador.find_element(By.XPATH, '//*[@id="login-form"]/footer/button').click()
        logging.info("Dados de login inseridos e submetidos.")


        time.sleep(2)
        # click to select unit
        navegador.find_element(By.XPATH, '//*[@id="GridLoginPartnerSelector"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[1]').click()



        # click to close alert 
        time.sleep(20)
        navegador.find_element(By.XPATH, '/html/body/div[4]/div/div/div[2]/button[2]').click()

        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="content"]/div[2]/div[2]/div/div/div[2]/div/a[8]/span/i')))
        navegador.find_element(By.XPATH, '//*[@id="content"]/div[2]/div[2]/div/div/div[2]/div/a[8]/span/i').click()
        logging.info("Navegação para a página do relatório iniciada.")

        # find page "Relatórios" here 


        WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gridContainer"]/div/div[6]/div/div/div/div/table/tbody/tr[19]/td[3]')))
        navegador.find_element(By.XPATH, '//*[@id="gridContainer"]/div/div[6]/div/div/div/div/table/tbody/tr[19]/td[3]').click()
        navegador.find_element(By.XPATH, '//*[@id="gridContainer"]/div/div[6]/div/div/div/div/table/tbody/tr[19]/td[3]').click()
        logging.info("Relatório selecionado.")

        # find "301 - Entr. Mercadorias [AC]"

        data_fim_formatada = datetime.now().strftime("%d/%m/%Y")
        navegador.find_element(By.XPATH, '//*[@id="content"]/app-relatorio-entrada-mercadoria-periodo/div/div/app-date-interval-bs/span/div[1]/div/app-bs-date-picker[1]/dx-date-box/div/div/div[1]/input').send_keys("01/01/2023")
        navegador.find_element(By.XPATH, '//*[@id="content"]/app-relatorio-entrada-mercadoria-periodo/div/div/app-date-interval-bs/span/div[1]/div/app-bs-date-picker[2]/dx-date-box/div/div/div[1]/input').send_keys(data_fim_formatada)
        navegador.find_element(By.XPATH, '//*[@id="content"]/app-relatorio-entrada-mercadoria-periodo/div/div/app-date-interval-bs/span/div[1]/div/button').click()
        logging.info("Datas configuradas e relatório solicitado.")

        # filter to data inserted 
        # erro no drag and drop - corrigir

        time.sleep(5)

        # Encontre o elemento de origem e destino usando XPaths fornecidos
        source_element1 = navegador.find_element(By.XPATH, '//*[@id="GridRelatorioEntradaMercadoriaPeriodo"]/div/div[4]/div/div/div[1]/div/div/div/div[2]/div')
        target_element1 = navegador.find_element(By.XPATH, '//*[@id="GridRelatorioEntradaMercadoriaPeriodo"]/div/div[5]/div/table/tbody/tr[1]')

        # Crie uma cadeia de ações
        actions = ActionChains(navegador)
        actions.drag_and_drop(source_element1, target_element1).perform()
        print("Elemento1 arrastado e solto com sucesso.")

        time.sleep(15) 

            # drag and drop second element
            # actions = ActionChains(navegador)

        source_element2 = navegador.find_element(By.XPATH, '//*[@id="GridRelatorioEntradaMercadoriaPeriodo"]/div/div[4]/div/div/div[1]/div/div/div/div')
        target_element2 = navegador.find_element(By.XPATH, '//*[@id="GridRelatorioEntradaMercadoriaPeriodo"]/div/div[5]/div/table/tbody/tr[1]')

        actions.drag_and_drop(source_element2, target_element2).perform()
        print("Elemento2 arrastado e solto com sucesso.")
            

            # Clicar no botão 3 pontos para abrir menu 
        navegador.find_element(By.XPATH, '//*[@id="GridRelatorioEntradaMercadoriaPeriodo"]/div/div[4]/div/div/div[3]/div[9]').click()
        time.sleep(5)

            # Clicar em "Exportar todos os dados" utilizando uma busca mais específica por classe e texto
        exportar_todos_os_dados = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'dx-datagrid-export-button')]//span[contains(text(), 'Exportar todos os dados')]"))
        )
        exportar_todos_os_dados.click()
        time.sleep(10)


        # to rename file 
        # Directory where files are downloaded
        # download_directory = "C://Users/lucas/Downloads/"


        new_file_name = f"301 - Entr. Mercadorias [AC].xlsx"
        local_file_path = caminho_ultimo_arquivo_baixado()
        upload_to_sharepoint(sp_folder_cantucci, local_file_path, new_file_name)
        excluir_arquivo(local_file_path)

try:
    load_dotenv()

    #Configuração de Logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Configuração de autenticação Sharepoint
    authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='felipe@orionbi.com.br', password='Bart123!').GetCookies()
    site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)

    # Acesso à biblioteca do SharePoint
    # path alterado - ok 
    sp_folder_cantucci = site.Folder('Documentos Compartilhados/Power BI/Cantucci/Xmenu')


    data_atual = datetime.now()
    data_fim_formatada = data_atual.strftime("%d/%m/%Y")
    navegador = setup_browser() 
    local_file_path = caminho_ultimo_arquivo_baixado()

    #Call functions

    baixar_e_upload_relatorio_xmenu(navegador)

    close_browser(navegador)

except:

    print("Erro de navegacao.")
    print("Reiniciando navegacao.")

    load_dotenv()

    #Configuração de Logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    # Configuração de autenticação Sharepoint
    authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='felipe@orionbi.com.br', password='Bart123!').GetCookies()
    site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)

    # Acesso à biblioteca do SharePoint
    # path alterado - ok 
    sp_folder_cantucci = site.Folder('Documentos Compartilhados/Power BI/Cantucci/Xmenu')


    data_atual = datetime.now()
    data_fim_formatada = data_atual.strftime("%d/%m/%Y")
    navegador = setup_browser()

    #Call functions

    close_browser(navegador)
    baixar_e_upload_relatorio_xmenu(navegador)
