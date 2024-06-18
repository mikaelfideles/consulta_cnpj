# COLETOR CNPJ SIMPLES NACIONAL

## Introdução

O script "Coletor CNPJ Simples Nacional" automatiza a coleta de informações de CNPJ a partir de uma fonte específica, utilizando a biblioteca `pyautogui` para interação com a interface gráfica, `pandas` para manipulação de dados e `clipboard` para copiar informações da área de transferência.

## Bibliotecas Utilizadas

- **time**: Para gerenciar atrasos temporais durante a execução do script.
- **datetime**: Para registrar a data e hora do início da coleta de dados.
- **pyautogui**: Para simular eventos de teclado e mouse.
- **pandas**: Para manipulação de dados em planilhas Excel.
- **clipboard**: Para copiar e colar dados da área de transferência.
- **tkinter**: Para criar janelas e caixas de mensagem para interação com o usuário.

## Instalação das Bibliotecas

Para instalar as bibliotecas necessárias, certifique-se de que você tem o pip instalado. Em seguida, execute o comando abaixo no terminal:
pip install -r requirements.txt

O arquivo requirements.txt deve conter as seguintes linhas (exemplo):
- `clipboard==0.0.4`
- `et-xmlfile==1.1.0`
- `MouseInfo==0.1.3`
numpy==1.26.4
openpyxl==3.1.2
pandas==2.2.2
pillow==10.3.0
pip==24.0
PyAutoGUI==0.9.54
PyGetWindow==0.0.9
PyMsgBox==1.0.9
pyperclip==1.8.2
PyRect==0.2.0
PyScreeze==0.1.30
python-dateutil==2.9.0.post0
pytweening==1.2.0
pytz==2024.1
setuptools==69.5.1
six==1.16.0
tzdata==2024.1
wheel==0.43.0

## Funcionalidade

### Passo 1: Inicialização

1. **Mensagem de Boas-Vindas**:
   - Exibe uma mensagem inicial e aguarda 10 segundos.
   - Cria e esconde uma janela principal usando `tkinter`.

2. **Preparação da Planilha Excel**:
   - Lê os dados de CNPJ a serem consultados de um arquivo Excel.
   - Adiciona uma coluna 'COLETA' para armazenar os resultados.
   - Salva os dados temporariamente em um novo arquivo Excel.

3. **Registro do Início da Coleta**:
   - Armazena a data e hora de início da coleta de dados.

### Passo 2: Coleta de Dados

1. **Iteração sobre cada Linha do DataFrame**:
   - Atualiza a página web (tecla F5).
   - Navega para o campo de entrada de CNPJ (7 tabs).
   - Digita o CNPJ e executa a consulta (tab e enter).
   - Aguarda o carregamento da página.

2. **Cópia e Verificação dos Resultados**:
   - Copia o conteúdo da página para a área de transferência.
   - Verifica se a mensagem "Informe um CNPJ válido." está presente:
     - Se sim, formata uma mensagem padrão indicando CNPJ inválido.
   - Caso contrário, repete o processo se a quantidade de caracteres copiados for insuficiente.

3. **Salvar Resultados Temporários**:
   - Salva os resultados temporários em um arquivo Excel.

### Passo 3: Processamento Final dos Dados

1. **Leitura do Arquivo Temporário**:
   - Lê os dados coletados do arquivo Excel temporário.

2. **Divisão e Limpeza dos Dados**:
   - Divide as linhas de resultados em colunas específicas.
   - Remove textos desnecessários e formata os dados.

3. **Criação do DataFrame Final**:
   - Converte a lista de linhas divididas em um DataFrame.
   - Renomeia as colunas e seleciona as colunas desejadas.

4. **Salvar o Arquivo Final**:
   - Salva o DataFrame final em um arquivo Excel nomeado com a data e hora de início da coleta.

### Passo 4: Finalização

1. **Mensagem de Sucesso**:
   - Exibe uma mensagem de pop-up indicando que o programa foi executado com sucesso.
   - Fecha a janela `tkinter`.

## Uso

Para executar o script, é necessário ter as bibliotecas mencionadas instaladas e os arquivos Excel nas pastas especificadas. Execute o script e siga as instruções na tela. Os resultados serão salvos em um novo arquivo Excel na pasta `03_coleta`.

## Tratamento de Interrupção

Se o script coletor.py for interrompido no meio do processo, pode-se executar o tratamento dos dados pelo script tratamento.py.