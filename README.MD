# 📈 Atualizador de Cotações – Dólar, Euro e Bitcoin

Este projeto realiza a coleta automática de cotações atualizadas para **Dólar (USD)**, **Euro (EUR)** e **Bitcoin (BTC)** em relação ao **Real Brasileiro (BRL)**, e salva essas informações em um arquivo Excel, mantendo um histórico de atualizações.

---

## 🚀 Funcionalidades

- Consulta automática das cotações via API da [AwesomeAPI](https://docs.awesomeapi.com.br/api-de-moedas).
- Armazena o histórico de cotações no arquivo `Cotações.xlsx`.
- Organiza as informações com colunas:
  - `Moeda`
  - `Cotação`
  - `Data da última atualização`

---

## 🛠️ Tecnologias utilizadas

- Python 3.13
- [Pandas](https://pandas.pydata.org/)
- [Requests](https://docs.python-requests.org/)
- [Poetry](https://python-poetry.org/) para gerenciamento de dependências

---

## 📦 Instalação

1️⃣ Clone o repositório:
    git clone https://github.com/Romulomrs/CotacaoDolar.git
    cd CotacaoDolar
2️⃣ Instale as dependências com **Poetry:**
    poetry install
3️⃣ Ative o ambiente virtual:
    poetry shell

⚙️ Como usar
1️⃣Execute o script principal:
    python src/atualiza_cotacoes.py
    (O arquivo Cotações.xlsx será criado/atualizado na pasta ArquivoExcel.
    A cada execução, novas linhas serão adicionadas automaticamente com as cotações atualizadas e a data/hora do momento.)

🗂️ Estrutura do Projeto
    ![Image](https://github.com/user-attachments/assets/7e129c53-fbc5-4de0-bae0-f2767e8b11b9)