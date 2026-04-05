# Automação de Extração e Consolidação de Dados - Sistema Empresa XFC

![Python](https://img.shields.io/badge/python-3.8+-blue.svg)
![Selenium](https://img.shields.io/badge/library-selenium-green.svg)
![Pandas](https://img.shields.io/badge/library-pandas-orange.svg)

## 📌 Descrição do Projeto
Este projeto automatiza o processo de extração de relatórios de produção do sistema **XFC**, realiza o tratamento dos dados brutos e consolida as informações em uma base de dados centralizada (Excel) utilizada por dashboards de **Power BI**.

O script elimina o trabalho manual de login, navegação, parametrização de datas, download e formatação de planilhas, reduzindo o tempo de execução e evitando erros humanos.

## 🚀 Funcionalidades
- **Web Scraping:** Login automático e navegação em sistema ERP via Selenium.
- **Data Cleaning:** Manipulação de arquivos Excel com Pandas (remoção de cabeçalhos inúteis, seleção de colunas e limpeza de nulos).
- **Consolidação:** Integração de novos dados em uma base histórica utilizando a biblioteca `openpyxl`.
- **Sincronização:** Backup e movimentação de arquivos entre redes locais e servidores de produção.

## 🛠️ Tecnologias Utilizadas
- [Python](https://www.python.org/)
- [Selenium WebDriver](https://www.selenium.dev/) (Automação Web)
- [Pandas](https://pandas.pydata.org/) (Tratamento de Dados)
- [Openpyxl](https://openpyxl.readthedocs.io/) (Manipulação de Excel)

## 📋 Pré-requisitos
Antes de executar o script, você precisará:
1. Python 3.8 ou superior instalado.
2. Microsoft Edge WebDriver compatível com a versão do seu navegador.
3. Bibliotecas necessárias instaladas:
   ```bash
   pip install pandas selenium openpyxl


📂 Estrutura de Pastas

├── data/               # Arquivos temporários e processados

├── main.py             # Script principal de automação

├── README.md           # Documentação do projeto

└── pyproject.toml      # Dependências do projeto
