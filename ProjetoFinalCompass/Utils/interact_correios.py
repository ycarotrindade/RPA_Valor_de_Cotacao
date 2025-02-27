import re
from botcity.web import WebBot, By, element_as_select

# como guardar CONSTANTE? (.env? config.py? outra solução?)
URL_CORREIOS = "https://www2.correios.com.br/sistemas/precosPrazos/"


def interact_correios(
    bot: WebBot,
    service_type: str,
    cep_destiny: str,
    weight: str,
    dimensions: dict,
    cep_origin: str = "38182428",
    shipping_date: str = None,
    package_format: str = "caixa",
    package_type: str = "Outra Embalagem",
) -> tuple:
    """Coordena os subprocessos realizados no site dos correios.

    O parâmetro shipping_date é opcional. O campo já vem preenchido
    com a data atual. O Valor padrão None.

    O parâmetro cep_origin é opcional. Valor padrão "38182428".
    O valor é passado para a função _interact_cep.

    Os parâmetros package_format e package_type são opcionais. O valor padrão
    já foi definido no pdd do projeto.

    Args:
        bot (WebBot): instância da classe WebBot.
        service_type (str): string contendo tipo de serviço dos correios.
            Exemplo: 'PAC', 'SEDEX', 'SEDEX 10' etc.
        cep_destiny (str): cep de destino para envio de encomenda.
        O formato de oito dígitos em sequência.
        weight (str): o peso estimado da embalagem.
        dimensions (dict): um dicionário que deve possuir as chaves "height",
            "width" e "length".
        cep_origin (str): cep de origem para envio de encomenda. Valor padrão
            é "38182428".
        shipping_date (str): data no formato ddmmaaaa. Valor padrão None.
        package_format (str): uma string que indica o formato da embalagem.
            Pode ser uma dentre "caixa", "envelope" ou "rolo". O valor padrão
            é "caixa".
        package_type (str): uma string que indica se a embalagem do pacote é
            no padrão dos correios ("Embalagem dos Correios") ou outro padrão
            ("Outra Embalagem"). O valor padrão é "Outra Embalagem".

    Return:
        tuple: retorna uma tupla com os elementos deliver_time, total_price.
            deliver_time (str): prazo em dias úteis para entrega dos correios.
            total_price (str): valor total em R$ para a entrega dos correios.
    """

    bot.browse(url=URL_CORREIOS)
    # interagir com campo "Data de postagem"
    if shipping_date:
        bot.find_element("input#data", By.CSS_SELECTOR).clear()
        bot.find_element("input#data", By.CSS_SELECTOR).click()
        bot.paste(shipping_date)

    # interagir com campo "CEP de origem" e "CEP de destino"
    bot.find_element("//input[@name='cepOrigem']", By.XPATH).click()
    bot.paste(cep_origin)
    # bot.wait(1500)
    bot.find_element("//input[@name='cepDestino']", By.XPATH).click()
    bot.paste(cep_destiny)
    # bot.wait(200)

    correios_services = bot.find_element("//select[@name='servico']", By.XPATH)
    select_service = element_as_select(correios_services)
    select_service.select_by_visible_text(service_type)
    # bot.wait(300)

    # campo Formato
    bot.find_element(f"img.{package_format}", By.CSS_SELECTOR).click()
    # bot.wait(400)

    # campo embalagem
    correios_packages = bot.find_element(
        "//select[@name='embalagem1']",
        By.XPATH,
    )
    select_package = element_as_select(correios_packages)
    select_package.select_by_visible_text(package_type)
    # bot.wait(500)

    # # Campo Dimensões
    bot.find_element("//input[@name='Altura']", By.XPATH).click()
    bot.paste(dimensions["height"])
    bot.tab()
    bot.paste(dimensions["width"])
    bot.tab()
    bot.paste(dimensions["length"])
    # bot.wait(600)

    # campo Peso estimado (Kg)
    bot.find_element(
        "//select[@name='peso']",
        By.XPATH,
    ).send_keys(weight)
    # bot.wait(700)

    # Botão Calcular
    bot.find_element("input.btn2", By.CSS_SELECTOR).click()
    # bot.wait(2000)

    # Transfere controle para nova aba que site correios abre
    opened_tabs = bot.get_tabs()
    tab_correios_response = opened_tabs[1]

    # tab_correios_main = opened_tabs[0]
    # bot.wait(1000)
    bot.activate_tab(tab_correios_response)

    # Capta valores deliver_time e price
    deliver_time = bot.find_element(
        (
            "//tr[@class='destaque']/th[text()='Pra"
            "zo de entrega ']/following-sibling::td"
        ),
        By.XPATH,
    ).text
    total_price = bot.find_element(
        # o comentário abaixo é para desligar o formatador automático
        # fmt: off
        (
            "//tr[@class='destaque']/th[text()="
            "'Valor total ']/following-sibling::td"
        ),
        By.XPATH,
    ).text

    bot.stop_browser()

    # extrair informações do deliver_time
    deliver_time = re.search(r"\+ (\d+)", deliver_time).group(1)

    return deliver_time, total_price
