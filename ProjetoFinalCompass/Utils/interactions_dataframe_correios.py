from pandas import DataFrame
from botcity.web import WebBot
from Utils.check_correios_variables import check_variables
from Utils.interact_correios import interact_correios
from Utils.IntegratedLogger import IntegratedLogger

def interaction_df_correios(
    df_output: DataFrame,
    df_filtered:DataFrame,
    bot: WebBot,
    logger:IntegratedLogger
) -> tuple[DataFrame,DataFrame]:
    """
    Faz a interação entre a dataframe e o site dos correios.

    Para cada linha da df realiza extração de dados que são utilizados pelo
    bot para extrair cotação e prazo do site dos correios.

    Args:
        df (DataFrame): dataframe pandas.
        bot (WebBot): instância do WebBot Botcity.

    Return: uma dataframe pandas.
    """

    for index, row in df_filtered.iterrows():
        cnpj = row["CNPJ"] # depois do try dentro do for index, row in df_filtered.iterrows() aqui diz linha 27
        try:
            # log: inicia teste {index} das variáveis usadas no site correios.
            logger.info(f'Inicia teste {index} das variáveis usadas no site dos correios')
            (
                package_dimensions,
                weight,
                postal_service,
                cep_destiny,
            ) = check_variables(row)
            # log: "variáveis teste {index} aprovadas."
            logger.info(f'variáveis teste {index} aprovadas')
        except Exception as err:
            df_output.loc[df_output["CNPJ"] == cnpj, "STATUS"] = err # em todos os df_output.loc[df_output["CNPJ"] == cnpj, "STATUS"] = err  
            # tratamento de erro "variáveis teste {index} não aprovadas."
            logger.error('Teste de variáveis')
            continue
        try:
            # log: inicia iteração {index} no site dos correios.
            logger.info(f'Inicia iteração {index} no site dos correios')
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
            # log: dados extraídos na iteração {index} no site dos correios.
            logger.info(f'dados extraídos na iteração {index} no site dos correios.')
        except Exception as err:
            df_output.loc[df_output["CNPJ"] == cnpj, "STATUS"] = err
            # tratar erro: problema na iteração {index} no site dos correios
            logger.error(f'problema na iteração {index} no site dos correios')
            continue
    return df_filtered, df_output
