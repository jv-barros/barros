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
import datetime
from datetime import datetime, timedelta

# Configuração de autenticação Sharepoint
authcookie = Office365('https://orionbusinessintelligence.sharepoint.com', username='felipe@orionbi.com.br', password='Bart123!').GetCookies()
site = Site('https://orionbusinessintelligence.sharepoint.com/sites/Grupo3v', version=Version.v365, authcookie=authcookie)

data_atual = datetime.now()
data_fim_formatada = data_atual.strftime("%d/%m/%Y")

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

def login_vuca(navegador, url, cpf, senha):
    navegador.get(url)
    # Modifique os seletores de acordo com a estrutura da página
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[1]/dd/input').send_keys(cpf)
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[2]/dd/input').send_keys(senha)
    wait_and_find_element(navegador, By.XPATH, '/html/body/section/div[1]/form/dl[3]/dd/button').click()

def baixar_relatorio_caixas_salao(navegador):

    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_caixa.php")
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
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_caixadelivery.php")
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
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_relatorios_comandas_conferencia.php")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[1]/dd/input').clear()
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[1]/dd/input').send_keys("01/01/2022")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[2]/dd/input').clear
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[1]/dl[2]/dd/input').send_keys("{data_fim_formatada}")
    navegador.find_element('xpath', '//*[@id="conteudo"]/div[4]/form/div[3]/dl[4]/dd/button[1]').click()
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(60)

def baixar_relatorio_b2b(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_relatorios_b2b.php")
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
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=8&")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_cervejas(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=3")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_hortifruti(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=13")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_nao_alcoolicos(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=6")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_eventos(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=23")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_carnes_brancas(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=7")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_secos(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=20")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_condimentos(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=9")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_ovos_e_laticinios(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=14")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_destilados(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=4")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_descartaveis(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=22")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_congelados(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=24")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_limpeza(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=16")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_carvao(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=26")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_vinhos(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=5")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_produto_interno(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=19")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_pescados(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=15")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_confeitaria(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=10")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_material_escritorio(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=25")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_utensilios(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=30")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def baixar_relatorio_cmc_gas(navegador):
    navegador.get("https://superquadrabar.vucasolution.com.br/retaguarda/pg_financeiro_entradas_custodemateriaprima.php?data_inicio=01/01/2022&data_fim={data_fim_formatada}&id_custodemateriaprima=3&&id_custodemateriaprima=27")
    time.sleep(5)
    navegador.find_element('xpath', '//*[@id="btn-csv"]').click()
    time.sleep(5)

def login_zigpay_spq(navegador, url, cpf, senha):
    navegador.get(url)
    # Modifique os seletores de acordo com a estrutura da página
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[2]/div/div/input').send_keys(cpf)
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[3]/div/div/input').send_keys(senha)
    time.sleep(2)
    navegador.find_element('xpath', '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[4]/button/div').click()
    time.sleep(2)

def alterar_valor_campo_readonly(navegador, xpath, novo_valor):
    elemento = navegador.find_element("xpath", xpath)
    navegador.execute_script(f"arguments[0].value='{novo_valor}';", elemento)

def baixar_relatorio_spq_mane(navegador):
    navegador.get("https://multiloja.zigpay.com.br/productSold")
    time.sleep(5)
    # Obtenha o primeiro dia do mês atual e o dia atual
    hoje = datetime.date.today()
    primeiro_dia_mes = hoje.replace(day=1)

    # Crie um loop que passe por cada dia do intervalo
    data_atual = primeiro_dia_mes
    while data_atual <= hoje:
        # Formate a data no formato desejado (por exemplo, '01/01/2024')
        data_formatada = data_atual.strftime("%d/%m/%Y")

        # Chame a função para alterar as datas
        alterar_valor_campo_readonly(navegador, '//*[@id="zig-popup-anchor"]/div/div[2]/div[2]/div/div[1]/div[1]/div/input', data_formatada)
        alterar_valor_campo_readonly(navegador, '//*[@id="zig-popup-anchor"]/div/div[2]/div[2]/div/div[1]/div[2]/div/input', data_formatada)

        # Faça o download do relatório aqui, se necessário
        time.sleep(2)
        navegador.find_element('xpath', '//*[@id="more-header"]').click()
        navegador.find_element('xpath', '//*[@id="menu-list-grow"]/li').click()
        # Incremente a data para o próximo dia
        data_atual += datetime.timedelta(days=1)
        time.sleep(10)  # Ajuste o tempo de espera conforme necessário

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
    url_base = "https://superquadrabar.vucasolution.com.br/retaguarda/pg_relatorios_produtosdevenda.php?csv=1"
    parametros = f"&datahora_inicio={data_inicio.strftime('%d%%2F%m%%2F%Y+00%%3A00')}&datahora_fim={data_fim.strftime('%d%%2F%m%%2F%Y+23%%3A59')}&data_rapido={data_inicio.strftime('%d%%2F%m%%2F%Y+00%%3A00')}&unidades%5B%5D=826&unidades%5B%5D=1017&listagem=&turno=&tipo=recebido"
    navegador.get(url_base + parametros)
    # Aguarda um tempo para garantir que o download inicie.
    time.sleep(5)

# navegador = setup_browser()
# login_zigpay_spq(navegador, "https://multiloja.zigpay.com.br/login", "superquadra.mane", "B3550@=")

# baixar_relatorio_spq_mane(navegador)


# new function to close browser
def close_browser(navegador):
    navegador.quit()

navegador = setup_browser()
login_vuca(navegador, "https://superquadrabar.vucasolution.com.br/retaguarda/", "02966597119", "029")

local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")

# Acesso à biblioteca do SharePoint
sp_folder_superquadra = site.Folder('Documentos Compartilhados/Power BI/Superquadra')

# Acesso à biblioteca do SharePoint
sp_folder_superquadra_entradas = site.Folder('Documentos Compartilhados/Power BI/Superquadra/Entradas')

sp_folder_superquadra_produtos = site.Folder('Documentos Compartilhados/Power BI/Superquadra/Produtos')

# Exemplo de como usar a função modificada para o mês atual
hoje = datetime.now()
primeiro_dia_mes = datetime(2024, 2, 25)
data_atual = primeiro_dia_mes

while data_atual <= hoje:
    baixar_produtos(navegador, data_atual, data_atual)
    
    # Espera um momento para o download ser concluído
    time.sleep(2)  # Ajuste conforme necessário
    
    local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
    print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
    
    nome_arquivo_sharepoint = data_atual.strftime('%Y-%m-%d') + ".csv"
    upload_to_sharepoint_superquadra(sp_folder_superquadra_produtos, local_file_path, nome_arquivo_sharepoint)
    
    excluir_arquivo(local_file_path)
    
    data_atual += timedelta(days=1)

# call functions
baixar_relatorio_caixas_salao(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra, local_file_path, "Faturamento Loja - Superquadra.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_caixas_delivery(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra, local_file_path, "Faturamento Delivery - Superquadra.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_comandas(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra, local_file_path, "Comandas - Superquadra.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_b2b(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra, local_file_path, "B2B.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_carnes_vermelhas(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CARNES VERMELHAS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_cervejas(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "ALCOÓLICOS - CERVEJAS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_hortifruti(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "HORTIFRUTI.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_nao_alcoolicos(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "BEBIDAS NÃO ALCOÓLICAS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_eventos(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "EVENTOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_carnes_brancas(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CARNES BRANCAS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_secos(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "SECOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_condimentos(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CONDIMENTOS, MOLHOS E ÓLEOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_ovos_e_laticinios(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "OVOS E LATICÍNIOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_destilados(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "ALCOÓLICOS - DESTILADOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_descartaveis(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "DESCARTAVEIS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_congelados(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CONGELADOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_limpeza(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "LIMPEZA.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_carvao(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CARVÃO.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_vinhos(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "ALCOÓLICOS - VINHOS E ESPUMANTES.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_produto_interno(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "PRODUTO INTERNO.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_pescados(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "PESCADOS E FRUTOS DO MAR.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_confeitaria(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "CONFEITARIA.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_material_escritorio(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "MATERIAL DE ESCRITORIO.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_utensilios(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "UTENSÍLIOS.csv")
excluir_arquivo(local_file_path)

baixar_relatorio_cmc_gas(navegador)
print(ultimo_arquivo_baixado("C://Users/lucas/Downloads"))
local_file_path = caminho_ultimo_arquivo_baixado("C://Users/lucas/Downloads")
upload_to_sharepoint_superquadra(sp_folder_superquadra_entradas, local_file_path, "GÁS.csv")
excluir_arquivo(local_file_path)

close_browser(navegador)
