import os
import logging
import traceback
import time

import pandas as pd
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

from Utils.IntegratedLogger import *


# ? -> Verificar se a pasta Processar existe;
# TODO -> Se o DF encontrar valores NA, jogar em uma lista e remover as linhas para não causar conflito adiante
def open_excel_file_to_dataframe(input_file_path):
    """ 
    Abre o arquivo excel em um DataFrame, faz modificações necessárias para o projeto e retona o DataFrame 
    
    Args:
        input_file_path (str): Caminho do arquivo Excel a ser aberto.
    
    Returns:
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

        # Verificando as células em branco e realizando a limpeza
        df_clean, empty_cells = clean_df_if_null(df_input)
        df_input = df_clean

        logger.info("Iniciando as modificações necessarias do DataFrame")
        # Adiciona colunas para as dimensões dos produtos
        df_input[['ALTURA', 'LARGURA', 'COMPRIMENTO']] = df_input['DIMENSÕES CAIXA (altura x largura x comprimento cm)'
                                                                    ].str.extract(r"(\d+)x(\d+)x(\d+)")
        logger.debug(f"Foram acrescentadas as colunas 'ALTURA', 'LARGURA', 'COMPRIMENTO' ao DataFrame.")
        logger.info("O arquivo foi modificado com sucesso. O DataFrame está pronto para uso.")

        return df_input, empty_cells
    
    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def create_output_dataframe():
    """
    Cria um DataFrame vazio com colunas predefinidas.
    
    Returns:
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

        return df_output
    
    except Exception as erro:
        logger.error(f"Ocorreu um erro: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise


def save_df_output_to_excel(output_path, df_output):
    """
    Salva o DataFrame em um arquivo Excel no caminho especificado.
    
    Args:
        output_path (str): Caminho onde o arquivo Excel será salvo.
        cnpj (str): CNPJ que será usado no nome do arquivo.
        df_output (pd.DataFrame): DataFrame a ser salvo.
    
    Returns:
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

# ? -> Essa função quem deve fazer é quem manipula a API;
def write_output_worksheet(df_output, column_name, data):
    """
    Escreve dados em uma coluna específica de um DataFrame.
    
    Args:
        df_output (pd.DataFrame): DataFrame de saída onde os dados serão escritos.
        column_name (str): Nome da coluna onde os dados serão adicionados.
        data: Dados a serem adicionados na coluna especificada.
    
    Returns:
        pd.DataFrame: DataFrame atualizado com os novos dados na coluna especificada.
    
    Raises:
        ValueError: Se a coluna especificada não existir no DataFrame.
        Exception: Para qualquer outro erro que ocorra durante o processo.
    """
    try:
        logger.info(f"Iniciando a entrada de dados na coluna {column_name}")
        # Acessa o DataFrame do output passado
        # Verifica o nome da coluna <column_name>
        # Acrescenta o dado passado <data>
        # Verifica se a coluna existe
        # ? Outro tipo de verificação
        if column_name in df_output.columns:
            logger.debug(f"Coluna {column_name} encontrada no DataFrame. Adicionando dados...")
            df_output[column_name] = data
            logger.info(f"Dados adicionados com sucesso na coluna {column_name}")
            return df_output
        else:
            logger.error(f"A coluna '{column_name}' não existe no DataFrame")
            raise ValueError(f"A coluna '{column_name}' não existe no DataFrame")
        
    except ValueError as erro:
        logger.error(f"Erro de valor: {erro}")
        logger.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise
    except Exception as erro:
        logging.error(f"Ocorreu um erro: {erro}")
        logging.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
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

        #Verifica se há células em branco em cada linha
        for _, row in df_to_clean.iterows():
            empty_rows = row[row.isnull()].index.tolist()
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


# TODO Caso alguma coluna esteja vazia:
# Ir para "Planilha Saída" e preencher a coluna 'Status' com a mensagem: 'Os campos <nome> estão vazios' e seguir para o próximo caso;
def write_if_null(df_input, empty_columns):
    try:

        # Escreve na coluna 'Status'
        # recebendo a coluna vazia na mensagem padrão
        df_output["Status"] = f"Os campos {empty_columns} estão vazios"
        return df_output


    except Exception as erro:
        logging.error(f"Ocorreu um erro: {erro}")
        logging.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise    


# TODO Comparar qual cotação está mais barata e pintar a celula correspondente de verde
def compare_quotation(df_output, output_file_path):
    try:
        # Comparar os valores e criar uma nova coluna indicando o menor valor
        df_output['Menor_Valor'] = df_output[["VALOR COTAÇÃO CORREIOS", "VALOR COTAÇÃO JADLOG"]].idxmin(axis=1)
        
        # Carregar o arquivo Excel usando openpyxl
        workbook = load_workbook(output_file_path)
        worksheet = workbook.active
        
        # Definir o preenchimento verde
        green_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
        
        # Pintar a célula com o menor valor de verde
        for index, row in df_output.iterrows():
            if row['Menor_Valor'] == "VALOR COTAÇÃO CORREIOS":
                worksheet.cell(row=index+2, column=df_output.columns.get_loc("VALOR COTAÇÃO CORREIOS") + 1).fill = green_fill
            else:
                worksheet.cell(row=index+2, column=df_output.columns.get_loc("VALOR COTAÇÃO JADLOG") + 1).fill = green_fill
        
        # Remover a coluna temporária
        df_output.drop('Menor_Valor', axis=1, inplace=True)
        
        # Salvar o arquivo Excel atualizado
        workbook.save(output_file_path)
        
        logging.info(f"Arquivo Excel atualizado e salvo em: {output_file_path}")

    except Exception as erro:
        logging.error(f"Ocorreu um erro: {erro}")
        logging.debug(f"Detalhes do erro:\n{traceback.format_exc()}")
        # Para o processo para depuração manual
        raise    