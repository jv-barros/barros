def login_zigpay_spq(navegador, url, cpf, senha):
    navegador.get(url)
    # Modifique os seletores de acordo com a estrutura da página
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[2]/div/div/input').send_keys(cpf)
    wait_and_find_element(navegador, By.XPATH, '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[3]/div/div/input').send_keys(senha)
    time.sleep(2)
    navegador.find_element('xpath', '//*[@id="zig-popup-anchor"]/div/div/div[2]/form/div[4]/button/div').click()
    time.sleep(2)


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



# navegador = setup_browser()
# login_zigpay_spq(navegador, "https://multiloja.zigpay.com.br/login", "superquadra.mane", "B3550@=")

# baixar_relatorio_spq_mane(navegador)