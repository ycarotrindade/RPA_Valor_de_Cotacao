from pandas import DataFrame
from botcity.web import WebBot
from Utils.check_correios_variables import check_variables
from Utils.interact_correios import interact_correios
from Utils.IntegratedLogger import IntegratedLogger


def interaction_df_correios(
    df_output: DataFrame,
    df_filtered: DataFrame,
    bot: WebBot,
    logger: IntegratedLogger,
) -> DataFrame:
    """
    Faz a interação entre a dataframe e o site dos correios.

    Para cada linha da df realiza extração de dados que são utilizados pelo
    bot para extrair cotação e prazo do site dos correios.

    Args:
        df_output (DataFrame): dataframe pandas com os dados de entrada.
        df_filtered (Dataframe): dataframe pandas que receberá dados de saída.
        bot (WebBot): instância do WebBot Botcity.
        logger (IntegratedLogger): instância do gerenciador de logs.

    Return: dataframe pandas com os dados de saída.
    """

    for index, row in df_filtered.iterrows():
        cnpj = row["CNPJ"]
        try:
            logger.info(
                f"Inicia teste {index} das variáveis usadas no site dos correios"
            )
            (
                package_dimensions,
                weight,
                postal_service,
                cep_destiny,
            ) = check_variables(row)
            logger.info(f"variáveis teste {index} aprovadas")
        except Exception as err:
            df_output.loc[df_output["CNPJ"] == cnpj, "STATUS"] = err
            logger.info("Variáveis não aprovadas")
            continue
        try:
            logger.info(f"Inicia iteração {index} no site dos correios")
            deliver_time, total_price = interact_correios(
                dimensions=package_dimensions,
                bot=bot,
                service_type=postal_service,
                cep_destiny=cep_destiny,
                weight=weight,
            )
            df_output.loc[df_output["CNPJ"] == cnpj, "PRAZO DE ENTREGA CORREIOS"] = (
                deliver_time
            )
            df_output.loc[df_output["CNPJ"] == cnpj, "VALOR COTAÇÃO CORREIOS"] = (
                total_price
            )

            logger.info(f"dados extraídos na iteração {index} no site dos correios.")
        except Exception as err:
            df_output.loc[df_output["CNPJ"] == cnpj, "STATUS"] = err
            logger.error(f"problema na iteração {index} no site dos correios")
            continue
    return df_output
