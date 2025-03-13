import pandas as pd
from pandas import Series


def check_variables(row: Series) -> tuple[dict, str, str, str]:
    """
    Verifica as variáveis usadas no site dos correios.

    A função recebe uma série de dados e verifica se as variáveis que
    interagirão com o site dos correios estão adequadas. São duas
    verificações: (1) verifica se as variáveis estão vazias;
    (2) verifica se os critérios para as dimensões da embalagem impostos
    pelo site dos correios foram atendidos.

    Args:
        row (Series): uma série de dados pandas.

    Raises:
        ValueError: se alguma das variáveis testadas estiver vazia ou se as
            dimensões da embalagem estiverem fora dos critérios dos correios.

    Return:
        tuple[dict, str, str, str]: uma tupla contendo:
            - um dicionário contendo as dimensões da embalagem.
            - uma string contendo o peso estimado.
            - uma string contendo o tipo de serviço postal.
            - uma string contendo o cep de destino.
    """

    PACKAGE_DIMENSIONS_KEYS = [
        "height",
        "width",
        "length",
    ]
    try:
        package_dimensions_values = row[
            "DIMENSÕES CAIXA (altura x largura x comprimento cm)"
        ]
        weight = row["PESO DO PRODUTO"]
        postal_service = row["TIPO DE SERVIÇO CORREIOS"]
        cep_destiny = str(row["CEP"])
        if (
            pd.isna(weight)
            or pd.isna(package_dimensions_values)
            or pd.isna(postal_service)
            or pd.isna(cep_destiny)
        ):
            raise ValueError("Erro ao realizar cotação Correios.")
        else:
            package_dimensions_values = package_dimensions_values.split(" x ")
            package_dimensions = dict(
                zip(
                    PACKAGE_DIMENSIONS_KEYS,
                    package_dimensions_values,
                )
            )

            if not are_package_dimensions_valid(**package_dimensions):
                raise ValueError("Erro ao realizar cotação Correios.")

            return package_dimensions, weight, postal_service, cep_destiny
    except Exception as err:
        raise err


def are_package_dimensions_valid(
    height: str,
    width: str,
    length: str,
) -> bool:
    """
    Analisa se as dimensões da embalagem passam nos critérios dos correios.

    A função aplica os critérios de dimensões de caixa exigidos pelo site dos
    correios.

    Args:
        height (str): uma string contendo a medida de altura da caixa.
        width (str): uma string contendo a medida de largura da caixa.
        length (str): uma string contendo a medida de comprimento da caixa.

    Return:
        bool: retorna verdadeiro se todas as restrições forem atendidas
            e retorna falso se alguma das restrições não forem atendidas.
    """

    height = float(height)
    width = float(width)
    length = float(length)
    sum_dimensions = height + width + length
    constrains = [
        0.4 <= height <= 100,
        8 <= width <= 100,
        13 <= length <= 100,
        21.4 <= sum_dimensions <= 200,
    ]
    return all(constrains)
