from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
from selenium import webdriver
import os
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from pathlib import Path
from datetime import date, timedelta



# Configuração de autenticação SharePoint
authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='joaovitor@orionbi.com.br', password='V)346720147023ay').GetCookies()
site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)

# Acesso à biblioteca do SharePoint
sp_folder_superquadra_entradas = site.Folder('Documentos Compartilhados/Power BI/Superquadra/Mané (sem data)')

#upload to server and github 
def setup_browser():
    options = Options()
    options.add_experimental_option("detach", True)
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return navegador

def close_browser(navegador):
    navegador.quit()

def wait_and_find_element(navegador, by, value):
    return WebDriverWait(navegador, 10).until(EC.presence_of_element_located((by, value)))

#login is fine 
def login_zigpay_spq(navegador, url, cpf, senha):
    navegador.get(url)
    # Modifique os seletores de acordo com a estrutura da página
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div[1]/form/div[2]/div[1]/div/input').send_keys(cpf)
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div[1]/form/div[2]/div[2]/div/input').send_keys(senha)
    time.sleep(2)
    navegador.find_element('xpath', '//*[@id="zig-popup-anchor"]/div/div[1]/form/div[3]/button').click()
    time.sleep(2)
    baixar_relatorio_spq_mane(navegador)    


# def selecionar_dia_calendario(navegador, dia):
#     # Base XPATH for the calendar days
#     base_xpath = '/html/body/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div['
#     dia_str = str(dia)

#     for week in range(1, 7):  # There can be up to 6 weeks in a month
#         for day in range(1, 8):  # There are 7 days in a week
#             dia_xpath = f"{base_xpath}{week}]/div[{day}]/button"
#             try:
#                 # Find the day button element
#                 dia_elemento = navegador.find_element(By.XPATH, dia_xpath)
#                 # Check if the button's text matches the desired day
#                 if dia_elemento.text == dia_str:
#                     print(f"Selecionando dia {dia} no calendário...")
#                     dia_elemento.click()
#                     return True
#             except Exception as e:
#                 print(f"Erro ao tentar selecionar o dia {dia_str} no calendário: {e}")
#                 continue  # Continue to the next day if an error occurs

#     return False  # Return False if the day could not be selected

# def selecionar_data(navegador, data):
#     # Formata a data para obter o dia
#     dia = data.day
#     print(f"Selecionando data: {data} (dia: {dia})")

#     # Abre o calendário da data inicial
#     print("Abrindo calendário da data inicial...")
#     navegador.find_element(By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/div/div[1]/div[1]/div/input').click()

#     # Aguarda até que o calendário esteja visível
#     WebDriverWait(navegador, 10).until(
#         EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]'))
#     )

#     # Seleciona o dia especificado na data inicial
#     if not selecionar_dia_calendario(navegador, dia):
#         print(f"Não foi possível selecionar o dia {dia} no calendário da data inicial.")
#         return False

#     # Fecha o calendário da data inicial
#     print("Fechando calendário da data inicial...")
#     navegador.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/button[2]').click()

#     # Abre o calendário da data final
#     print("Abrindo calendário da data final...")
#     navegador.find_element(By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/div/div[1]/div[2]/div/input').click()

#     # Aguarda até que o calendário esteja visível
#     WebDriverWait(navegador, 10).until(
#         EC.visibility_of_element_located((By.XPATH, '/html/body/div[2]/div[3]/div/div[1]/div/div[2]/div[2]/div/div[1]'))
#     )

#     # Seleciona o dia especificado na data final
#     if not selecionar_dia_calendario(navegador, dia):
#         print(f"Não foi possível selecionar o dia {dia} no calendário da data final.")
#         return False

#     # Fecha o calendário da data final
#     print("Fechando calendário da data final...")
#     navegador.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/button[2]').click()

#     return True

def baixar_relatorio_spq_mane(navegador):
    print("Acessando a página de relatórios...")
    navegador.get("https://multiloja.zigpay.com.br/places/1d02dc84-e124-42e2-81f2-ba83233080a2/productSold")
    time.sleep(5)

    button1 = wait_and_find_element(navegador, By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div[2]/div[2]/button[1]')
    button1.click()
    time.sleep(1)  # Adjust the sleep time as needed

    button2 = wait_and_find_element(navegador, By.XPATH, '/html/body/div/div/div/div/div/div[2]/div/div[2]/div[2]/button[3]')
    button2.click()
    time.sleep(1)  # Adjust the sleep time as needed

    # Obtenha o dia atual
    hoje = date.today()
    # print(f"Baixando relatório para a data: {hoje.strftime('%d/%m/%Y')}")
    print(f"Baixando relatório de ontem.")


    # Seleciona o dia atual no calendário
    # if not selecionar_data(navegador, hoje):
    #     print(f"Falha ao selecionar a data: {hoje}")
    #     return

    # Iniciar o download do relatório
    print("Iniciando download do relatório...")
    time.sleep(2)
    navegador.find_element(By.XPATH, '//*[@id="more-header"]').click()
    WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="menu-list-grow"]/li'))
    ).click()

    print("Download do relatório concluído.")
   
def upload_to_sharepoint_superquadra(sp_folder, local_file_path, sharepoint_file_name):
    # Correcting the file path and the file input handling
    with open(local_file_path, 'rb') as file_input:
        sp_folder.upload_file(file_input, sharepoint_file_name)
    print(f"Arquivo '{local_file_path}' enviado como '{sharepoint_file_name}' no SharePoint.")

def excluir_arquivo(local_file_path):
    try:
        os.remove(local_file_path)
        print(f"Arquivo '{local_file_path}' excluído com sucesso.")
    except OSError as e:
        print(f"Erro ao excluir o arquivo '{local_file_path}': {e}")

def caminho_ultimo_arquivo_baixado(caminho, new_name):
    # List all files in the directory, exclude directories
    lista_arquivos = [f for f in os.listdir(caminho) if os.path.isfile(os.path.join(caminho, f))]
    
    if not lista_arquivos:
        raise FileNotFoundError("No files found in the specified directory.")
    
    # Create a list of tuples with file paths and their modification times
    lista_datas = [(os.path.join(caminho, arquivo), Path(os.path.join(caminho, arquivo)).stat().st_mtime) for arquivo in lista_arquivos]
    
    # Sort the list by modification time in descending order
    lista_datas.sort(key=lambda x: x[1], reverse=True)
    
    # Get the most recently modified file
    ultimo_arquivo = lista_datas[0][0]
    
    # Define the new file path
    novo_caminho = os.path.join(caminho, new_name)
    
    # Rename the file
    os.rename(ultimo_arquivo, novo_caminho)
    
    return novo_caminho

def get_yesterday_date_dd_mm_yyyy():
    # Get current date
    current_date = datetime.now()
    
    # Subtract one day to get yesterday's date
    yesterday_date = current_date - timedelta(days=1)
    
    # Format date as DD-MM-YYYY
    formatted_date = yesterday_date.strftime("%d-%m-%Y")
    
    return formatted_date



def execute_script():
    navegador = setup_browser()
    login_zigpay_spq(navegador, "https://multiloja.zigpay.com.br/login", "superquadra.mane", "B3550@=")
    baixar_relatorio_spq_mane(navegador)
    time.sleep(30)

    # Get the current date formatted as DD-MM-YYYY
    current_date = get_yesterday_date_dd_mm_yyyy()

    # Directory where files are downloaded
    download_directory = "C://Users/lucas/Downloads/"
    new_file_name = f"Mane sem data {current_date}.xlsx"
    
    try:
        # Get and rename the latest downloaded file
        local_file_path = caminho_ultimo_arquivo_baixado(download_directory, new_file_name)
        
        # Upload the renamed file to SharePoint
        upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, new_file_name)
        
        # Delete the local file after uploading
        excluir_arquivo(local_file_path)
    except Exception as e:
        print(f"An error occurred: {e}")
    
    close_browser(navegador)



# Execute the script
execute_script()


# demand finished 
#imports send email 
import smtplib
import email.message
import datetime
#send class function 
class EmailConf:
    def __init__(self):
        self.email1 = 'jvcbcarvalho@gmail.com'
        self.senha = 'lhkl neir pmrt tkzx'

    def enviaemail(self, destinatario, dest_copia, assunto, corpo):
        hora_atual = datetime.datetime.now().time()
        manha_inicio = datetime.time(6, 0, 0)  # 6:00 - morning 
        tarde_inicio = datetime.time(16, 0, 0)  # 16:00 - afternoon  
        noite_inicio = datetime.time(18, 0, 0)  # 23:00 - afternoon 

        if hora_atual < manha_inicio:
            saudacao = "Boa noite"
        elif hora_atual < tarde_inicio:
            saudacao = "Bom dia"
        elif hora_atual < noite_inicio:
            saudacao = "Boa tarde"
        else:
            saudacao = "Boa noite"
        try:
            corpo_email = f"""
            <p>{saudacao}!</p>
            <p></p>
            <p></p>
            <p>{corpo}</p>
            <p>--<br>
            João Barros<br>
            Full Stack Developer<br>
            <br>
            Contato - 61 9 9252 5843<br>
            Email: jvcbcarvalho@gmail.com</p>
            <img src="https://i.imgur.com/KLUHTR8.png heigt="200" width="300" alt="João Barros">
            """

            msg = email.message.Message()
            msg['Subject'] = assunto
            msg['From'] = self.email1
            msg['To'] = destinatario
            msg['Cc'] = dest_copia
            msg.add_header('Content-Type', 'text/html')
            msg.set_payload(corpo_email)

            with smtplib.SMTP('smtp.gmail.com', 587) as s:
                s.starttls()  # Habilita o TLS
                s.login(self.email1, self.senha)

                destinatarios = destinatario.split(', ') + dest_copia.split(', ')

                s.sendmail(msg['From'], destinatarios, msg.as_string().encode('utf-8'))
                print('E-mail enviado com sucesso!')

        except smtplib.SMTPAuthenticationError:
            print('Erro de autenticação. Verifique o endereço de e-mail e senha.')
        except smtplib.SMTPException as e:
            print(f'Erro ao enviar o e-mail: {e}')

# Example usage
email_sender = EmailConf()
email_sender.enviaemail('jvcbcarvalho@gmail.com', 'felipe@orionbi.com.br', 'Alerta de execução de automação', 'Confirmação de execução.')
