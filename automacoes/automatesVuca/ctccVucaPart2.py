from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from pathlib import Path
import os
import time
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from datetime import datetime, timedelta

# Configuração de autenticação Sharepoint
authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='felipe@orionbi.com.br', password='Bart123!').GetCookies()
site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)


data_atual = datetime.now()
data_fim_formatada = data_atual.strftime("%d/%m/%Y")
# main functions
def caminho_ultimo_arquivo_baixado(caminho):
    lista_arquivos = os.listdir(caminho)
    lista_datas = [(os.path.join(caminho, arquivo), Path(caminho, arquivo).stat().st_mtime) for arquivo in lista_arquivos]
    lista_datas.sort(key=lambda x: x[1], reverse=True)
    return lista_datas[0][0]

def setup_browser():
    options = Options()
    options.add_experimental_option("detach", True)
    navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return navegador

def wait_and_find_element(navegador, by, value):
    return WebDriverWait(navegador, 10).until(EC.presence_of_element_located((by, value)))

def login(navegador, url, cpf, senha):
    navegador.get(url)
    # Modifique os seletores de acordo com a estrutura da página
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[1]/dd/input').send_keys(cpf)
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[2]/dd/input').send_keys(senha)
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[3]/dd/button').click()

def ultimo_arquivo_baixado(caminho):
    lista_arquivos = os.listdir(caminho)
    lista_datas = [(Path(caminho, arquivo).stat().st_mtime, arquivo) for arquivo in lista_arquivos]
    lista_datas.sort(reverse=True)
    return lista_datas[0][-1]

def upload_to_sharepoint_superquadra(sp_folder, local_file_path, sharepoint_file_name):
    with open(local_file_path, 'rb') as file_input:
        sp_folder.upload_file(file_input, sharepoint_file_name)
    print(f"Arquivo '{local_file_path}' enviado como '{sharepoint_file_name}' no SharePoint.")

def excluir_arquivo(local_file_path):
    try:
        os.remove(local_file_path)
        print(f"Arquivo '{local_file_path}' excluído com sucesso.")
    except OSError as e:
        print(f"Erro ao excluir o arquivo '{local_file_path}': {e}")

def baixar_produtos(navegador, data_inicio, data_fim):
    url_base = "https://cantucci.vucasolution.com.br/retaguarda/pg_relatorios_produtosdevenda.php?csv=1"
    parametros = f"&datahora_inicio={data_inicio.strftime('%d%%2F%m%%2F%Y+00%%3A00')}&datahora_fim={data_fim.strftime('%d%%2F%m%%2F%Y+23%%3A59')}&data_rapido={data_inicio.strftime('%d%%2F%m%%2F%Y+00%%3A00')}&unidades%5B%5D=895&unidades%5B%5D=1018&unidades%5B%5D=1019&unidades%5B%5D=1108&unidades%5B%5D=1724&unidades%5B%5D=1740&unidades%5B%5D=1780&unidades%5B%5D=1781&listagem=&turno=&tipo=recebido"
    navegador.get(url_base + parametros)
    # Aguarda um tempo para garantir que o download inicie.
    time.sleep(5)


def handle_report_download_and_upload(
    download_function,
    upload_folder,
    upload_filename,
    delete_after_upload=True
):
    try:
        download_function(navegador)
        print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
        local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
        upload_to_sharepoint_superquadra(upload_folder, local_file_path, upload_filename)
    except Exception as e:
        print(f"An error occurred during {download_function.__name__}: {e}")
    finally:
        if delete_after_upload and local_file_path:
            excluir_arquivo(local_file_path)

# new function to close browser
def close_browser(navegador):
    navegador.quit()

# main variables
navegador = setup_browser()
login(navegador, "https://cantucci.vucasolution.com.br/retaguarda/", "02966597119", "029")


local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")

# Acesso à biblioteca do SharePoint
sp_folder_cantucci = site.Folder('Documentos Compartilhados/Power BI/Cantucci')

# Acesso à biblioteca do SharePoint
sp_folder_cantucci_entradas = site.Folder('Documentos Compartilhados/Power BI/Cantucci/Entradas')

# Acesso à biblioteca do SharePoint
sp_folder_cantucci_comandas = site.Folder('Documentos Compartilhados/Power BI/Cantucci/Comandas')

sp_folder_cantucci_produtos = site.Folder('Documentos Compartilhados/Power BI/Cantucci/Produtos')

# Exemplo de como usar a função modificada para o mês atual
hoje = datetime.now()
primeiro_dia_mes =  hoje.replace(day=1)
data_atual = primeiro_dia_mes

while data_atual <= hoje:
    baixar_produtos(navegador, data_atual, data_atual)
    
    # Espera um momento para o download ser concluído
    time.sleep(2)  # Ajuste conforme necessário
    
    local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
    print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
    
    nome_arquivo_sharepoint = data_atual.strftime('%Y-%m-%d') + ".csv"
    upload_to_sharepoint_superquadra(sp_folder_cantucci_produtos, local_file_path, nome_arquivo_sharepoint)
    
    excluir_arquivo(local_file_path)
    
    data_atual += timedelta(days=1)

# part 01 of the vuca reports 

def baixar_relatorio_caixas_salao(navegador):

    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_caixa.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[1]/dd/input').send_keys("01/01/2024")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[6]/dd/button/i').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(10)

def baixar_relatorio_caixas_delivery(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_caixadelivery.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[1]/dd/input').send_keys("01/01/2024")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[6]/dd/button/i').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(10)

def baixar_relatorio_comandas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_relatorios_comandas_conferencia.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[1]/dd/input').send_keys("01/01/2024")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[3]/dl[4]/dd/button[1]').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(60)

def baixar_relatorio_b2b(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_relatorios_b2b.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[1]/dd/input').send_keys("01/01/2024")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[2]/dl[5]/dd/button').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(10)

def baixar_relatorio_cmc_carnes_vermelhas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=6")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_queijos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=19")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_secos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=14")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_hortifruti(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=10")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_vinhos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=3")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_condimentos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=7")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_nao_alcoolicos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2024&data_fim={data_fim_formatada}&&id_custodemateriaprima=4")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_carnes_brancas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=5")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)
    
#call functions
try:
    handle_report_download_and_upload(
        baixar_relatorio_caixas_salao, sp_folder_cantucci, "Faturamento Loja - Vuca.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_caixas_delivery, sp_folder_cantucci, "Faturamento Delivery - Vuca.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_comandas, sp_folder_cantucci_comandas, "2024.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_b2b, sp_folder_cantucci, "B2B.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_carnes_vermelhas, sp_folder_cantucci_entradas, "CARNES VERMELHAS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_queijos, sp_folder_cantucci_entradas, "QUEIJOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_secos, sp_folder_cantucci_entradas, "SECOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_hortifruti, sp_folder_cantucci_entradas, "HORTIFRUTI.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_vinhos, sp_folder_cantucci_entradas, "VINHOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_condimentos, sp_folder_cantucci_entradas, "CONDIMENTOS, MOLHOS E OLEOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_nao_alcoolicos, sp_folder_cantucci_entradas, "BEBIDAS NÃO ALCOOLICAS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_carnes_brancas, sp_folder_cantucci_entradas, "CARNES BRANCAS.csv"
    )   
finally:
    close_browser(navegador)