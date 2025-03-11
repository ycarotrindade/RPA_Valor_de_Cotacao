import pandas as pd
import requests
import os
from Utils.IntegratedLogger import IntegratedLogger
from config import vars_map
from time import sleep

# Define o diretório para o arquivo de log
log_directory = vars_map['BASE_LOG_PATH']

# Cria o diretório se ele não existir
#os.makedirs(log_directory, exist_ok=True)

# Define o nome do arquivo de log
#log_file = os.path.join(log_directory, "arquivo_de_log.txt")

# Configuração de logger
# logger.basicConfig(
#     filename=log_file, level=logger.INFO, format="%(asctime)s - %(levelname)s: %(message)s"
# )

def read_excel_data(excel_file_path: str, logger:IntegratedLogger) -> list:
    """Lê dados de CNPJ de um arquivo Excel."""
    try:
        logger.info(f"Lendo dados do arquivo Excel")
        df = pd.read_excel(excel_file_path)
        cnpj_list = []
        for cnpj in df["CNPJ"]:
            cnpj = str(cnpj).strip()
            if len(cnpj) == 13:
                cnpj_list.append("0" + cnpj)
            else:
                cnpj_list.append(cnpj)
        logger.info(f"CNPJs lidos do Excel: {cnpj_list}")
        return cnpj_list
    except FileNotFoundError:
        logger.error("Processo read_excel_data")
        return
    except Exception as e:
        logger.error('Processo read_excel_data')
        return


def query_brasilapi(cnpj: str,logger:IntegratedLogger) -> tuple:  # Retorna uma tupla com dados e status
    """Consulta a API BrasilAPI para obter informações de CNPJ."""
    url = f"{vars_map['DEFAULT_BRASILAPI_URL']}{cnpj}"
    try:
        logger.info(f"Consultando API BrasilAPI para CNPJ: {cnpj}")
        headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
        response = requests.get(url=url,headers=headers,timeout=10)
        response.raise_for_status()
        return response.json(), "Sucesso"  # Retorna dados e status "sucesso"
    except requests.exceptions.RequestException as e:
        logger.info(f"Erro ao consultar CNPJ {cnpj}")
        return None, "falha"  # Retorna None e status "falha"

def create_companies_dataframe(companies_data: list,logger:IntegratedLogger) -> pd.DataFrame:
    """Cria um DataFrame Pandas com os dados das empresas."""
    try:
        logger.info("Criando DataFrame com dados da API.")
        data = []
        for item in companies_data:
            company_info = item['data']
            status = item['status']
            if company_info:
                company_info['status'] = status
                data.append(company_info)
        logger.debug(data.__repr__())
        df = pd.DataFrame(data,dtype=str)

        columns = [
            "cnpj",
            "razao_social",
            "nome_fantasia",
            "situacao_cadastral",
            "logradouro",
            "numero",
            "municipio",
            "cep",
            "descricao_identificador_matriz_filial",
            "ddd_telefone_1",
            "email",
            "status"
        ]
        df = df[columns].copy()
        df = df.dropna(subset=["cnpj"]).reset_index(drop=True)
        df = df.fillna("N/A")
        df = df.replace("", "N/A")
        status_map = {
            1: "Nula",
            2: "Ativa",
            3: "Suspensa",
            4: "Inapta",
            5: "Ativa Não Regular",
            8: "Baixada",
        }
        df["situacao_cadastral"] = df["situacao_cadastral"].map(status_map)
        df.columns = [
            "CNPJ",
            "RAZÃO SOCIAL",
            "NOME FANTASIA",
            "SITUAÇÃO CADASTRAL",
            "LOGRADOURO",
            "NÚMERO",
            "MUNICÍPIO",
            "CEP",
            "DESCRIÇÃO MATRIZ FILIAL",
            "TELEFONE + DDD",
            "E-MAIL",
            "STATUS"
        ]
        
        return df
    except Exception as e:
        logger.error(f"Processo create_companies_dataframe")
        return None


def save_dataframe_to_csv(df: pd.DataFrame, csv_file_path: str,logger:IntegratedLogger):
    """Salva o DataFrame em um arquivo CSV."""
    try:
        logger.info(f"Salvando dados processados em arquivo CSV: {csv_file_path}")
        df.to_csv(csv_file_path, index=False)
    except Exception as e:
        logger.error(f"Processo save_dataframe_to_csv")


def join_and_transform(excel_df, api_df):
    """
    Mescla dados de um arquivo CSV em um arquivo Excel, atualizando as colunas correspondentes.

    Args:
        excel_file_path (str): Caminho para o arquivo XLSX.
        csv_file_path (str): Caminho para o arquivo CSV.
        output_file_path (str): Caminho para salvar o arquivo XLSX atualizado.
    """
    try:
        excel_df = excel_df.astype('object')
        api_df = api_df.astype('object')
        # Carregar os arquivos com tipos de dados especificados para preservar zeros à esquerda

        # Verificar se a coluna CNPJ existe em ambos os DataFrames
        if "CNPJ" not in excel_df.columns or "CNPJ" not in api_df.columns:
            raise ValueError("A coluna CNPJ não foi encontrada em um ou ambos os arquivos.")

        # Criar a coluna ENDEREÇO no CSV e remover colunas auxiliares
        endereco_cols = ["LOGRADOURO", "NÚMERO", "MUNICÍPIO"]
        if all(col in api_df.columns for col in endereco_cols):
            api_df["ENDEREÇO"] = api_df[endereco_cols].agg(lambda x: ', '.join(x.dropna().astype(str)), axis=1)
            api_df.drop(columns=endereco_cols, inplace=True)

        # Mesclar os DataFrames
        merged_df = excel_df.update(api_df, on="CNPJ", how="left", suffixes=("", "_novo"))

        # Atualizar as colunas do Excel apenas onde houver novos valores
        for column in api_df.columns:
            if column != "CNPJ" and column in excel_df.columns:
                merged_df[column] = merged_df[column + "_novo"].combine_first(merged_df[column])
                merged_df.drop(columns=[column + "_novo"], inplace=True)
                merged_df[column].fillna("N/A", inplace=True)

        # Tratar a coluna SITUAÇÃO CADASTRAL separadamente
        if "SITUAÇÃO CADASTRAL" in merged_df.columns:
            merged_df["SITUAÇÃO CADASTRAL"].fillna("N/A", inplace=True)

        # Salvar o DataFrame atualizado
        return merged_df

    except FileNotFoundError as e:
        print(f"Erro: {e}")
    except ValueError as e:
        print(f"Erro: {e}")
    except Exception as e:
        print(f"Erro ao processar os dados: {e}")


def api_data_lookup(df_output:pd.DataFrame,logger:IntegratedLogger):
    """Função principal do programa."""
    logger.info("Iniciando busca de dados no site Brasil API.")
    
    excel_file_path = os.path.join(vars_map['DEFAULT_PROCESSAR_PATH'],'Planilha de Entrada Grupos.xlsx')
    # excel_final_file_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processar\cnpj_2025-02-24_12h16m30s.xlsx"
    # csv_file_path = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processados\dados_processados.csv"
    # output_excel = "C:/RPA/RPA_Valor_de_Cotacao/ProjetoFinalCompass/Processar/dados_completos.xlsx"

    
    cnpj_list = read_excel_data(excel_file_path,logger)
    if cnpj_list:
        companies_data = []
        missing_cnpjs = []
        for cnpj in cnpj_list:
            company_data, status = query_brasilapi(cnpj,logger)  # Recebe dados e status
            companies_data.append({'data': company_data, 'status': status})
            if status == 'falha':
                missing_cnpjs.append(cnpj)
        companies_df = create_companies_dataframe(companies_data,logger)
        if companies_df is not None:
            logger.info("Identificando CNPJs ausentes na API.")
            missing_cnpjs_api = [cnpj for cnpj in cnpj_list if cnpj not in companies_df["CNPJ"].tolist()]
            for cnpj in missing_cnpjs:
                df_output.loc[df_output['CNPJ'] == cnpj, 'STATUS'] = 'Sem retorno da API'
            logger.info(f"CNPJs não encontrados na API: {missing_cnpjs_api}")
            #save_dataframe_to_csv(companies_df, csv_file_path)

            # Chama a função join_and_transform após salvar o arquivo CSV
            #join_and_transform(excel_final_file_path, csv_file_path, output_excel)
    logger.info("Busca de dados finalizada.")
    return companies_df, df_output

if __name__ == "__main__":
    api_data_lookup()