# Consulta de CNPJs no Site do Simples Nacional

Este projeto tem como objetivo fornecer uma ferramenta simples e eficiente para a consulta de CNPJs (Cadastro Nacional da Pessoa Jurídica) diretamente no site do Simples Nacional (https://www8.receita.fazenda.gov.br/SimplesNacional/aplicacoes.aspx?id=21), facilitando o acesso às informações de empresas de forma rápida e direta.

## Funcionalidades

Permite consultar informações detalhadas de um CNPJ específico, como:
- Data consulta;
- CNPJ;
- Nome;
- Situação Simples Nacional;
- Situação SIMEI.

## Baixar dependências

- Clonar SSH do GitHub ou baixar Zip;
- Criar virtual ou utilizar Python global (desde que possua as bibliotecas do arquivo requirements.txt instaladas).

## Criando ambiente virtual

No terminal, digite:
- cd C:\Users\...\consulta_cnpj (atualizar para o endereço real)
- pip install virtualenv venv
- venv\scripts\activate
- pip install -r requirements.txt

## Como usar

- Listar todos os números de CNPJ (C:\Users\...\01_consulta\consulta.xlsx);
- Restaurar dimensão do navegador para 50% da tela [win + up;  win + right] ou [win + up;  win + left];
- Edite o endereço presente no arquivo coletor.bat (C:\Users\...\consulta_cnpj\venv\Scripts") ou execute o arquivo "main.py";
- O tempo para inicializar o programa é de 10 segundos, sendo assim, deixe o navegador como janela ativa, para que o programa consiga percorrer todos os campos do html;
- Após finalizar o programa, será exibido um alerta >>> messagebox.showinfo("Atenção", "A coleta foi realizada com sucesso!").

## Processo executadas pelo programa

- 01: Ler a lista de CNPJs a serem consultados no arquivo "consulta.xlsx";
- 02: para cada CNPJ, consultar os dados no site do Simples Nacional;
- 03: salvar todos os dados coletados (sem tratamento) no arquivo "coleta_temp.xlsx";
- 04: tratar os dados do arquivo "coleta_temp.xlsx" e salvar um novo arquivo chamado no modelo "coleta_01012024_235959" (nome do arquivo + data + hora).


teste giovanni
