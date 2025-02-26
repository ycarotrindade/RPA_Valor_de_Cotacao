import json
import pandas as pd
import requests

endereco = r'C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processados\dados.json'

df = pd.read_excel(r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processar\Planilha de Entrada Grupos.xlsx") 

lista_cnpj = []
for cnpj in df['CNPJ']:
    if len(str(cnpj)) == 13:
        lista_cnpj.append("0" + str(cnpj))
    else:
        lista_cnpj.append(str(cnpj))
    

request_list = []

for cnpj in lista_cnpj:
    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
    response = requests.get(url)
    request_list.append(response.json())
    

df2 = pd.DataFrame(request_list)
colunas = ["cnpj", "razao_social", "nome_fantasia", "situacao_cadastral", "logradouro", "numero", "municipio", "cep", "descricao_identificador_matriz_filial", "ddd_telefone_1", "email"]

df3 = df2[colunas].copy()

df3 = df3.dropna(subset=['cnpj']).reset_index(drop=True)

df3 = df3.fillna("N/A")
df3 = df3.replace("", "N/A")

df3

mapa = {1: "Nula", 2: "Ativa", 3: "Suspensa", 4: "Inapta", 5: "Ativa Não Regular", 8: "Baixada"}

df3["situacao_cadastral"] = df3["situacao_cadastral"].map(mapa)

df3.columns = ["CNPJ", "Razão Social", "Nome Fantasia", "Situação Cadastral", "Logradouro", "Numero", "Município", "CEP", "Matriz ou Filial", "DDD e Telefone", "Email"]

df3

cnpj_ausentes = [cnpj for cnpj in lista_cnpj if cnpj not in df3["CNPJ"].tolist()]
print(cnpj_ausentes)

df3.to_csv(r"C:\RPA\RPA_Valor_de_Cotacao\ProjetoFinalCompass\Processados\dados_processados", index=False)

