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
authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='joaovitor@orionbi.com.br', password='V)346720147023ay').GetCookies()
site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)


data_atual = datetime.now()
data_fim_formatada = data_atual.strftime("%d/%m/%Y")

def caminho_ultimo_arquivo_baixado(caminho):
    lista_arquivos = os.listdir(caminho)
    if not lista_arquivos:  # Check if the list is empty
        raise FileNotFoundError(f"No files found in the directory: {caminho}")
    
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

def baixar_relatorio_caixas_salao(navegador):

    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_caixa.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div/dl[1]/dd/input').send_keys("01/01/2022")
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
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[3]/form/div/dl[1]/dd/input').send_keys("01/01/2022")
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
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[1]/dd/input').send_keys("01/01/2022")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[1]/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[2]/form/div[2]/dl[5]/dd/button').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(10)

def baixar_relatorio_cmc_carnes_vermelhas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=6")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_queijos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=19")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_secos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=14")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_hortifruti(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=10")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_vinhos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=3")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_condimentos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=7")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_nao_alcoolicos(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=4")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_carnes_brancas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=5")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_pescados(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=13")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_embalagem(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=18")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_ovos_e_laticinios(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=12")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_confeitaria(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=8")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_limpeza(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=15")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_cervejas(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=1")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_destilados(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=2")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_brownie(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=11")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_interno(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=17")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_descartaveis(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=9")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_desvinculados(navegador):
    navegador.get("https://cantucci.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&&id_custodemateriaprima=16")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

# def baixar_relatorio_xmenu(navegador): #em fase de testes
    navegador.get("https://portal.netcontroll.com.br/#/auth/login")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="email"]').send_keys("gerencia.cantucci@gmail.com")
    navegador.find_element('xpath', '//*[@id="password"]').send_keys("M@nuella19")
    navegador.find_element('xpath', '//*[@id="login-form"]/footer/button').click()
    time.sleep(2)
    navegador.find_element('xpath', '//*[@id="content"]/div[2]/div[2]/div/div/div[2]/div/a[8]/span/i').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="gridContainer"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[35]/td[3]').click()
    navegador.find_element('xpath', '//*[@id="gridContainer"]/div/div[6]/div/div/div[1]/div/table/tbody/tr[35]/td[3]').click()
    navegador.find_element('xpath', '//*[@id="content"]/app-relatorio-fechamento-caixa/div/div/app-date-interval-bs/span/div[1]/div/app-bs-date-picker[1]/dx-date-box/div/div/div[1]/input').send_keys("01/01/2023")
    navegador.find_element('xpath', '//*[@id="content"]/app-relatorio-fechamento-caixa/div/div/app-date-interval-bs/span/div[1]/div/app-bs-date-picker[2]/dx-date-box/div/div/div[1]/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="content"]/app-relatorio-fechamento-caixa/div/div/app-date-interval-bs/span/div[1]/div/button').click()
    navegador.find_element('xpath', '//*[@id="GridRelatorioFechamentoCaixa"]/div/div[4]/div/div/div[3]/div[9]/div/div[1]/i').click() # Clicar no botão com os "três pontinhos"


    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="dx-9d3db346-8e99-d263-414b-bbc376003e2f"]/div[1]/div/div[1]/div[2]/div[6]').click()
    time.sleep(10)
    navegador.find_element('xpath', '//*[@id="content"]/app-relatorio-fechamento-caixa/div/div/app-date-interval-bs/span/div[1]/div/button').click()

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


# new function to close browser
def close_browser(navegador):
    navegador.quit()

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
# report with error 
    # handle_report_download_and_upload(
    #     baixar_relatorio_b2b, sp_folder_cantucci, "B2B.csv"
    # )

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

    handle_report_download_and_upload(
        baixar_relatorio_cmc_pescados, sp_folder_cantucci_entradas, "PESCADOS E FRUTOS DO MAR.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_embalagem, sp_folder_cantucci_entradas, "EMBALAGEM.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_ovos_e_laticinios, sp_folder_cantucci_entradas, "OVOS E LATICINIOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_confeitaria, sp_folder_cantucci_entradas, "CONFEITARIA.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_limpeza, sp_folder_cantucci_entradas, "MATERIAL DE LIMPEZA.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_cervejas, sp_folder_cantucci_entradas, "ALCOOLICOS - CERVEJAS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_destilados, sp_folder_cantucci_entradas, "ALCOOLICOS - DESTILADOS.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_brownie, sp_folder_cantucci_entradas, "MR. BROWNIE.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_interno, sp_folder_cantucci_entradas, "USO INTERNO.csv"
    )

    handle_report_download_and_upload(
        baixar_relatorio_cmc_descartaveis, sp_folder_cantucci_entradas, "DESCARTAVEIS.csv"
    )

    # handle_report_download_and_upload(
    #     baixar_relatorio_cmc_desvinculados, sp_folder_cantucci_entradas, "DESVINCULADOS E EXCLUIDOS.csv"
    # )
finally:
    close_browser(navegador)


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
            <p>-> Execução finalizada </p>
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
