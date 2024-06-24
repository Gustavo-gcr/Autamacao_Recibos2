# Automatizador Excel Word

## Visão Geral

O Automatizador Excel Word é uma aplicação desenvolvida em Python que automatiza a geração de contratos personalizados em formato Word (.docx) a partir de dados contidos em uma planilha Excel. Este sistema simplifica o processo de criação de documentos personalizados, permitindo que modelos de contrato sejam preenchidos com informações específicas de clientes, produtos ou serviços diretamente de uma planilha.

## Funcionalidades

### Geração de Contratos Personalizados:

- Utiliza um modelo de contrato (.docx) como base para preenchimento automático.
- Extrai dados relevantes de uma planilha Excel para preencher o modelo de contrato.
- Salva os contratos gerados com nomes personalizados.

### Processamento Automatizado:

- Automatiza a leitura de dados da planilha e a substituição de placeholders no documento Word.
- Converte os documentos gerados para PDF, se necessário.

## Requisitos do Sistema

- Python 3.x instalado
- Bibliotecas Python necessárias: os, sys, tkinter, openpyxl, python-docx, docx2pdf, datetime

## Como Usar

1. **Preencher a Planilha Base:**
   - Insira os dados necessários na planilha Excel conforme as instruções fornecidas.

2. **Baixar Python:**
   - [Python Downloads](https://www.python.org/downloads/)
   - Instale o Python, lembrando-se de marcar as caixinhas de seleção necessárias.

3. **Atualizar Bibliotecas:**
   - Execute o programa de atualização de bibliotecas.

4. **Executar Automação:**
   - Execute o script principal (`main.py`).

5. **Seleção dos Arquivos:**
   - Selecione os arquivos necessários: o arquivo Excel com os dados dos recibos, o modelo de contrato Word e a pasta de destino para os documentos gerados.

6. **Resultado:**
   - Os documentos Word serão preenchidos com os dados da planilha, salvos na pasta de destino e convertidos para PDF, se configurado.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e pull requests.

## Autores

- [Gustavo Coelho](https://github.com/Gustavo-gcr)
- [Iasmin Fernandes](https://github.com/IasminCQFernandes)
