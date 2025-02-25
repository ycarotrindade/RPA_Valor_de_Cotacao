import os
import logging
import traceback
import time

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

from Utils.IntegratedLogger import *


def open_excel_file_to_dataframe(input_file_path):
    """ 
    Abre o arquivo excel em um DataFrame, faz modificações necessárias para o projeto e retona o DataFrame 
    
    Parâmetros:
        input_file_path (str): Caminho do arquivo Excel a ser aberto.
    
    Retorna:
        pd.DataFrame: DataFrame com os dados do arquivo Excel e as modificações realizadas.
    
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado no caminho especificado.
        Exception: Para qualquer outro erro que ocorra durante o processo.
    """
    try:
        logger.info("Iniciando abertura do arquivo Excel e retornando um DataFrame")

        # Verifica se o arquivo existe
        if not os.path.exists(input_file_path):
            logger.error("O arquivo não foi encontrado.")
            raise  FileNotFoundError(f"O arquivo {input_file_path} não foi encontrado.") 
        logger.info(f"O arquivo de Excel com os dados de entrada foi encontrado.")
        logger.debug(f"O arquivo foi encontrado na pasta indicada: {input_file_path}")
        
        # Abre o arquivo excel em DataFrame, na pasta 'Groupo 1' e preenche valores vazios como 'NA'
        df_input = pd.read_excel(input_file_path, "Grupo 1 ", na_values=["NA"])
        logger.info("DataFrame com base no arquivo Excel criado com sucesso")

        return df_input
    
    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def create_output_dataframe():
    """
    Cria um DataFrame vazio com colunas predefinidas.
    
    Retorna:
        pd.DataFrame: DataFrame com as colunas predefinidas.
    
    Raises:
        Exception: Para qualquer erro que ocorra durante o processo.
    """
    try:
        logger.info("Iniciando a criação do DataFrame para receber os dados de saída")
        
        # Definindo colunas
        columns = [
            "CNPJ", "RAZÃO SOCIAL", "NOME FANTASIA",
            "ENDEREÇO", "CEP", "DESCRIÇÃO MATRIZ FILIAL",
            "TELEFONE + DDD", "E-MAIL", "VALOR DO PEDIDO",
            "DIMENSÕES CAIXA", "PESO DO PRODUTO", "TIPO DE SERVIÇO JADLOG",
            "TIPO DE SERVIÇO CORREIOS", "VALOR COTAÇÃO JADLOG", "VALOR COTAÇÃO CORREIOS",
            "PRAZO DE ENTREGA CORREIOS", "STATUS"
        ]
        # Criando o DataFrame com as colunas definidas
        df_output = pd.DataFrame(columns=columns)
        logger.info("DataFrame para receber as saídas criado com sucesso.")

        # Alimenta o DataFrame de saída com as informações já existentes
        df_output.set_index("CNPJ").join(df_input.set_index("CNPJ"))
        
        return df_output
    
    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def save_df_output_to_excel(output_path, df_output):
    """
    Salva o DataFrame em um arquivo Excel no caminho especificado.
    
    Parâmetros:
        output_path (str): Caminho onde o arquivo Excel será salvo.
        cnpj (str): CNPJ que será usado no nome do arquivo.
        df_output (pd.DataFrame): DataFrame a ser salvo.
    
    Retorna:
        str: Caminho do arquivo Excel gerado.
    
    Raises:
        Exception: Para qualquer erro que ocorra durante o processo.
    """
    try:
        logger.info("Iniciando a criação da planilha Excel com os dados de saída")

        # Gerando nome do arquivo baseado na data e hora atual
        current_date = time.strftime("%Y-%m-%d_%Hh%Mm%Ss")
        file_name = f"cnpj_{current_date}.xlsx"
        logger.debug(f"Nome do arquivo criado: {file_name}")

        # Salvando DataFrame como arquivo Excel
        output_file_path = f"{output_path}/{file_name}"
        df_output.to_excel(output_file_path, index=False)
        logger.info(f"Sucesso, arquivo criado: {file_name}")
        logger.debug(f"Arquivo Excel criado com sucesso em: {output_path}")

        return output_file_path

    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def split_columns_box(df_to_split):
    try:
        logger.info("Iniciando as modificações necessarias do DataFrame")
        # Adiciona colunas para as dimensões dos produtos
        # ! deixar essa parte em um uma função separada para criar um df_output_quotation ?
        df_to_split[['ALTURA', 'LARGURA', 'COMPRIMENTO']] = df_to_split[
            'DIMENSÕES CAIXA (altura x largura x comprimento cm)'
            ].str.extract(r"(\d+)x(\d+)x(\d+)")
        logger.debug(f"Foram acrescentadas as colunas 'ALTURA', 'LARGURA', 'COMPRIMENTO' ao DataFrame.")
        logger.info("O arquivo foi modificado com sucesso. O DataFrame está pronto para uso.")
        df_input_split = df_to_split

        return df_input_split

    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise    


def clean_df_if_null(df_to_clean):
    """
    Remove linhas com células vazias de um DataFrame e registra os CNPJs e as colunas com células vazias.
    
    Args:
        df_to_clean (pd.DataFrame): DataFrame a ser limpo.
    
    Returns:
        pd.DataFrame: DataFrame limpo.
        list: Lista de dicionários com CNPJs e colunas com células vazias.
    
    Raises:
        Exception: Para qualquer erro que ocorra durante o processo.
    """
    try:
        logger.info("Iniciando o processo de limpeza das células vazias do DataFrame")
        # Cria uma lista para receber os CNPJs com alguma célula vazia
        empty_cells = []

        # Verifica se há células em branco em cada linha
        for _, row in df_to_clean.iterows():
            empty_rows = row[row.isnull()].index.tolist()
            # Registra as linhas vazias na lista como um dicionário
            if empty_rows:
                empty_cells.append({
                    "CNPJ": row["CNPJ"],
                    "NA": empty_rows
                    })

        # Registra os CNPJs com células vazias
        if empty_cells:
            logger.info(f"CNPJs com células vazias: {empty_cells}")
        else:
            logger.info("Nenhuma célula vazia encontrada.")

        # Remove as linhas que contêm células em branco
        df_clean = df_to_clean.dropna()
        logger.info("Células vazias/NA removidas do DataFrame")
        
        return df_clean, empty_cells

    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def write_if_null_output(df_output, empty_cells):
    """
    Atualiza a coluna 'Status' no DataFrame de saída com informações sobre células vazias.

    Parâmetros:
    df_output (pd.DataFrame): DataFrame que será atualizado.
    empty_cells (list): Lista de dicionários contendo CNPJs e colunas vazias.

    Retorna:
    pd.DataFrame: DataFrame atualizado com a coluna 'Status' preenchida.
    """
    try:
        logger.info("Iniciando o processo de registro de células vazias")

        # Itera sobre a lista de células vazias para preencher a coluna 'Status'
        for empty in empty_cells:
            # Verifica se, na linha, o índice CNPJ é igual
            index = df_output[df_output["CNPJ"] == empty["CNPJ"]].index

            # Verifica se a linha com o CNPJ existe no df_output e preenche
            if not index.empty:
                # Atualiza a coluna 'Status' na linha correspondente
                df_output.at[index, "Status"] = f"Os campos {empty['NA']} estão vazios"
                logger.info(f"Coluna 'Status' atualizada para o CNPJ {empty['CNPJ']}: {empty['NA']}")

        return df_output


    except Exception as erro:
        logging.error(f"Ocorreu um erro: {erro}")
        logging.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise    


def compare_quotation(df_output, output_file_path):
    """
    Compara os valores de cotação e destaca o menor valor em um arquivo Excel.

    Parâmetros:
    df_output (pd.DataFrame): DataFrame contendo os valores de cotação.
    output_file_path (str): Caminho do arquivo Excel de saída.
    """
    try:
        logger.info("Inciando o processo de comparação das cotações - Correios x JadLog")
        
        # Comparar os valores e criar uma nova coluna indicando o menor valor
        df_output['Menor_Valor'] = df_output[["VALOR COTAÇÃO CORREIOS", "VALOR COTAÇÃO JADLOG"]].idxmin(axis=1)
        logger.info("Comparação de valores concluída. Coluna 'Menor_Valor' criada.")
        
        # Carregar o arquivo Excel usando openpyxl
        workbook = load_workbook(output_file_path)
        worksheet = workbook.active
        logger.info("Arquivo Excel carregado com sucesso")
        
        # Definir a cor de preenchimento: verde
        green_fill = PatternFill(start_color='008000', 
                                end_color='008000', 
                                fill_type='solid')
        
        # Pintar a célula com o menor valor de verde
        for index, row in df_output.iterrows():
            # Verificar qual coluna é a de menor valor que deve ser pintada
            if row['Menor_Valor'] == "VALOR COTAÇÃO CORREIOS":
                column_index = df_output.columns.get_loc("VALOR COTAÇÃO CORREIOS") + 1
            else:
                column_index = df_output.columns.get_loc("VALOR COTAÇÃO JADLOG") + 1
        
            # Obter a célula correspondente
            cell = worksheet.cell(row=index+2, column=column_index)
            # Preenche a célula com a cor escolhida
            cell.fill = green_fill
            logger.debug(f"Célula na linha {index+2} preenchida de verde.")

        # Remover a coluna temporária
        df_output.drop('Menor_Valor', axis=1, inplace=True)
        logger.debug(f"Coluna temporária 'Menor_Valor' removida de {df_output}")
        
        # Salvar o arquivo Excel atualizado
        workbook.save(output_file_path)
        logging.info(f"Arquivo Excel atualizado e salvo em: {output_file_path}")

    except ValueError as erro:
        logging.error(f"Erro de valor: {erro}")
    except FileNotFoundError as erro:
        logging.error(f"Arquivo não encontrado: {erro}")
    except Exception as erro:
        logging.error(f"Ocorreu um erro: {erro}")
        logging.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise    