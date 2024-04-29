# Consulta de CNPJs no Site do Simples Nacional

Este projeto tem como objetivo fornecer uma ferramenta simples e eficiente para a consulta de CNPJs (Cadastro Nacional da Pessoa Jurídica) diretamente no site do Simples Nacional (https://www8.receita.fazenda.gov.br/SimplesNacional/aplicacoes.aspx?id=21), facilitando o acesso às informações de empresas de forma rápida e direta.

## Funcionalidades

Permite consultar informações detalhadas de um CNPJ específico, como:
- Data consulta;
- CNPJ;
- Nome;
- Situação Simples Nacional;
- Situação SIMEI.

## Como Usar

- Clonar SSH do GitHub ou baixar Zip;
- Criar virtual ou utilizar Python global, desde que possua as bibliotecas do arquivo requirements.txt instaladas;
- Caso opte por usar o ambiente virtual, atualizar o endereço do python no arquivo "coletor.bat" (cd "C:\Users\...\consulta_cnpj\venv\Scripts")

### Criando ambiente virtual

- pip install virtualenv venv
- venv\scripts\activate
- pip install -r requirements.txt
