import pandas as pd
import requests
import logging
import os

# Define o diretório para o arquivo de log
diretorio_log = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\log"

# Cria o diretório se ele não existir
os.makedirs(diretorio_log, exist_ok=True)

# Define o nome do arquivo de log
arquivo_log = os.path.join(diretorio_log, "arquivo_de_log.txt")

# Configuração de logging
logging.basicConfig(
    filename=arquivo_log, level=logging.INFO, format="%(asctime)s - %(levelname)s: %(message)s"
)


def ler_dados_excel(arquivo_excel: str) -> list:
    """Lê dados de CNPJ de um arquivo Excel.

    Args:
        arquivo_excel: O caminho para o arquivo Excel.

    Returns:
        Uma lista com os CNPJs lidos do arquivo.
    """
    try:
        logging.info(f"Lendo dados do arquivo Excel")
        df = pd.read_excel(arquivo_excel)
        lista_cnpj = []
        for cnpj in df["CNPJ"]:
            cnpj = str(cnpj).strip()  # Remover espaços em branco
            if len(cnpj) == 13:
                lista_cnpj.append("0" + cnpj)
            else:
                lista_cnpj.append(cnpj)
        logging.info(f"CNPJs lidos do Excel: {lista_cnpj}")
        return lista_cnpj
    except FileNotFoundError:
        logging.error(f"Arquivo Excel não encontrado: {arquivo_excel}")
        return []
    except Exception as e:
        logging.error(f"Erro ao ler arquivo Excel: {e}")
        return []

def consultar_api_brasilapi(cnpj: str) -> dict:
    """Consulta a API BrasilAPI para obter informações de CNPJ.

    Args:
        cnpj: O CNPJ a ser consultado.

    Returns:
        Um dicionário com os dados da empresa ou None em caso de erro.
    """
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    try:
        logging.info(f"Consultando API BrasilAPI para CNPJ: {cnpj}")
        response = requests.get(url, timeout=10)  # Adicionar timeout
        response.raise_for_status()  # Lança exceção para status HTTP diferentes de 200
        return response.json()
    except requests.exceptions.RequestException as e:
        logging.error(f"Erro ao consultar CNPJ {cnpj}: {e}")
        return None

def criar_dataframe_empresas(dados_empresas: list) -> pd.DataFrame:
    """Cria um DataFrame Pandas com os dados das empresas.

    Args:
        dados_empresas: Uma lista de dicionários com os dados das empresas.

    Returns:
        Um DataFrame com os dados das empresas.
    """
    try:
        logging.info("Criando DataFrame com dados da API.")
        df = pd.DataFrame(dados_empresas)
        colunas = [
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
        ]
        df = df[colunas].copy()
        df = df.dropna(subset=["cnpj"]).reset_index(drop=True)
        df = df.fillna("N/A")
        df = df.replace("", "N/A")
        mapa_situacao_cadastral = {
            1: "Nula",
            2: "Ativa",
            3: "Suspensa",
            4: "Inapta",
            5: "Ativa Não Regular",
            8: "Baixada",
        }
        df["situacao_cadastral"] = df["situacao_cadastral"].map(mapa_situacao_cadastral)
        df.columns = [
            "CNPJ",
            "Razão Social",
            "Nome Fantasia",
            "Situação Cadastral",
            "Logradouro",
            "Numero",
            "Município",
            "CEP",
            "Matriz ou Filial",
            "DDD e Telefone",
            "Email",
        ]
        return df
    except Exception as e:
        logging.error(f"Erro ao criar DataFrame: {e}")
        return None

def salvar_dados_csv(df: pd.DataFrame, arquivo_csv: str):
    """Salva o DataFrame em um arquivo CSV.

    Args:
        df: O DataFrame a ser salvo.
        arquivo_csv: O caminho para o arquivo CSV.
    """
    try:
        logging.info(f"Salvando dados processados em arquivo CSV: {arquivo_csv}")
        df.to_csv(arquivo_csv, index=False)
    except Exception as e:
        logging.error(f"Erro ao salvar arquivo CSV: {e}")

def consulta_api():
    """Função principal do programa."""

    logging.info("Iniciando busca de dados no site Brasil API.")

    arquivo_excel = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processar\Planilha de Entrada Grupos.xlsx"

    arquivo_csv = r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processados\dados_processados.csv"


    lista_cnpj = ler_dados_excel(arquivo_excel)

    if lista_cnpj:
        dados_empresas = []
        cnpj_ausentes = []

        for cnpj in lista_cnpj:
            dados_empresa = consultar_api_brasilapi(cnpj)
            if dados_empresa:
                dados_empresas.append(dados_empresa)
            else:
                cnpj_ausentes.append(cnpj)

        df_empresas = criar_dataframe_empresas(dados_empresas)

        if df_empresas is not None:
            logging.info("Identificando CNPJs ausentes na API.")
            cnpj_ausentes_api = [cnpj for cnpj in lista_cnpj if cnpj not in df_empresas["CNPJ"].tolist()]
            logging.info(f"CNPJs não encontrados na API: {cnpj_ausentes_api}")

            salvar_dados_csv(df_empresas, arquivo_csv)

    logging.info("Busca de dados finalizada.")

if __name__ == "__main__":
    consulta_api()