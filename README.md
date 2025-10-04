# Script de Validação e Comparação de Planilhas Excel

Este projeto contém um script Python desenvolvido para automatizar a comparação e validação de dados entre duas planilhas Excel (.xlsx). O principal objetivo é identificar discrepâncias, registros exclusivos e correspondências exatas entre os dois arquivos, utilizando uma coluna-chave (como RE de funcionário) para cruzar as informações.

## Tecnologias Utilizadas
* **Python 3**
* **Pandas:** Biblioteca para manipulação e análise de dados.
* **NumPy:** Biblioteca para operações numéricas, utilizada aqui para a lógica de comparação.

## O Problema que ele Resolve

Em muitos ambientes corporativos, é comum ter informações de funcionários ou ativos espalhadas em diferentes relatórios (ex: um do RH e outro da TI). Validar manualmente se os dados (IMEI de celular, linha telefônica, etc.) batem em ambas as planilhas é um processo lento, repetitivo e sujeito a erros. Este script automatiza essa tarefa, gerando um relatório completo em segundos.

## Pré-requisitos

Antes de executar, você precisa ter o Python e as bibliotecas necessárias instaladas.

1.  **Instale o Python 3** do [site oficial](https://www.python.org/).
2.  **Instale as bibliotecas `pandas` e `openpyxl`** (necessária para o pandas ler arquivos .xlsx):
    ```bash
    pip install pandas openpyxl
    ```

## Como Usar o Script

A configuração é feita diretamente no início do arquivo de código. Siga os passos abaixo:

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/MatheusRenat0/scriptp-validacao.git](https://github.com/MatheusRenat0/scriptp-validacao.git)
    ```
2.  **Coloque suas planilhas na pasta:**
    Coloque os dois arquivos Excel que você deseja comparar na mesma pasta onde o script está localizado.

3.  **Configure as variáveis no script:**
    Abra o arquivo Python e edite a seção `# --- 1. CONFIGURAÇÃO ---`:
    * `caminho_planilha1`: Coloque o nome exato do seu primeiro arquivo (ex: `'relatorio_rh.xlsx'`).
    * `caminho_planilha2`: Coloque o nome exato do seu segundo arquivo (ex: `'relatorio_ti.xlsx'`).
    * `mapa_colunas_planilha1`: Para cada campo (`RE`, `Nome`, `IMEI`, etc.), informe o nome exato da coluna correspondente na sua primeira planilha.
    * `mapa_colunas_planilha2`: Faça o mesmo para a sua segunda planilha.
    * `coluna_chave`: Defina qual coluna será o identificador único para cruzar os dados (geralmente `RE` ou `ID do funcionário`).

4.  **Execute o script:**
    Abra um terminal na pasta do projeto e execute:
    ```bash
    python nome_do_seu_script.py
    ```

## Entendendo os Resultados

O script irá imprimir um relatório detalhado diretamente no terminal, dividido em quatro categorias:

1.  **Itens em COMUM COM DIVERGÊNCIAS:**
    Registros que existem em ambas as planilhas (mesmo RE), mas com alguma informação diferente (ex: o IMEI ou o nome do funcionário não batem).

2.  **Itens ENCONTRADOS APENAS na Planilha 1:**
    Registros (REs) que estão na primeira planilha, mas não foram encontrados na segunda.

3.  **Itens ENCONTRADOS APENAS na Planilha 2:**
    Registros (REs) que estão na segunda planilha, mas não foram encontrados na primeira.

4.  **Itens IDÊNTICOS em ambas as planilhas:**
    Registros que existem nas duas planilhas com todas as informações comparadas sendo exatamente iguais.
