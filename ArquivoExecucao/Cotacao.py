import requests
import pandas as pd
from datetime import datetime

# Fazendo a requisição da API
requisicao = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL")
requisicao_dic = requisicao.json()

# Pegando as cotações
cotacao_dolar = float(requisicao_dic['USDBRL']['bid'])
cotacao_euro = float(requisicao_dic['EURBRL']['bid'])
cotacao_btc = float(requisicao_dic['BTCBRL']['bid']) * 1000  # Convertendo BTC para mil

# Criando novas linhas com os dados atualizados
data_atual = datetime.now()
novas_linhas = pd.DataFrame([
    {"Moeda": "Dólar", "Cotação": cotacao_dolar, "Data da última atualização": data_atual},
    {"Moeda": "Euro", "Cotação": cotacao_euro, "Data da última atualização": data_atual},
    {"Moeda": "Bitcoin", "Cotação": cotacao_btc, "Data da última atualização": data_atual},
])

# Carregando ou criando o arquivo Excel
arquivo = "ArquivoExcel/Cotações.xlsx"
try:
    tabela = pd.read_excel(arquivo)

    # Verificação de tabela completamente vazia (evita FutureWarning)
    if tabela.empty or tabela.isna().all().all():
        tabela = novas_linhas
    else:
        tabela = pd.concat([tabela, novas_linhas], ignore_index=True)

except FileNotFoundError:
    # Se não existir ainda, começa com as novas linhas
    tabela = novas_linhas

# Salvando a planilha com as novas informações
tabela.to_excel(arquivo, index=False)
print(f"Cotações atualizadas e adicionadas com sucesso! {data_atual}")

