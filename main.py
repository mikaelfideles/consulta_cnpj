import time
from datetime import datetime
import pyautogui as pg
import pandas as pd
import clipboard

# Ler o arquivo Excel com os dados das consultas
consulta = pd.read_excel(r'consulta_cnpj.xlsx', dtype=str)

# Preencher valores nulos com strings vazias
consulta = consulta.fillna('')

# Registrar o momento de início da coleta de dados
now = datetime.now()
inicio_coleta = now.strftime("%d %m %Y %Hh %Mmin %Ss")

time.sleep(10)
print('Iniciando coleta...')    

# Iterar através de cada linha do DataFrame 'consulta'
for i, linha in consulta.iterrows():

    pg.press('f5')

    time.sleep(2)

    # Exibir informações relevantes
    print('-----------------------------------------')
    now = datetime.now()
    dt_string = now.strftime("%d %m %Y %Hh %Mmin %Ss")
    print("date and time =", dt_string)
    print(linha['QTDE_CONSULTA'])
    print('num_doc: ' + linha['NUM_CNPJ'])
  
    presses = 7

    # Pressiona a tecla "Tab" várias vezes com um intervalo de 0.1 segundo
    for _ in range(presses):
        pg.press('tab')
        #time.sleep(0.1)        
    
    # Escrever o CNPJ na caixa de texto
    pg.write(linha['NUM_CNPJ'])
    time.sleep(0.1)

    pg.press('tab')
    time.sleep(0.1)
    pg.press('enter')

    time.sleep(5)

     # Selecionar e copiar a data da consulta
    pg.hotkey('ctrl', 'a')
    pg.hotkey('ctrl', 'c')
    
    # Atribuir a dados copiados copiada à coluna 'COLETA' da linha atual
    linha['COLETA'] = clipboard.paste()
    time.sleep(0.1)
    print(linha['COLETA'])
    time.sleep(0.1)
    #time.sleep(1)

    # Salvar o DataFrame 'consulta' em um arquivo Excel
    #consulta.to_excel(r'consulta/consulta_cnpj_iniciada_em ' + str(inicio_coleta) + '.xlsx', index=False)
    consulta.to_excel(r'coleta/coleta.xlsx', index=False)

# Exibir uma mensagem indicando que o programa foi finalizado
print('Coleta finalizada!')


# Ler o arquivo Excel
consulta = pd.read_excel(r'coleta\coleta.xlsx', dtype=str)

# Lista para armazenar as linhas divididas em colunas
linhas_divididas = []

# Iterar sobre as linhas da consulta
for i, linha in consulta.iterrows():
    # Dividir a linha em partes com base no caractere de quebra de linha '\n'
    partes = linha['COLETA'].split('\n')

    # Remover o texto "Data da consulta: " das linhas pertencentes à coluna "Data da consulta"
    partes[1] = partes[1].replace("Data da consulta: ", "")
    partes[3] = partes[3].replace("CNPJ: ", "")
    partes[6] = partes[6].replace("Nome Empresarial: ", "")
    partes[8] = partes[8].replace("Situação no Simples Nacional: ", "")
    partes[9] = partes[9].replace("Situação no SIMEI: ", "")

    # Adicionar as partes à lista de linhas divididas
    linhas_divididas.append(partes)

# Converter a lista de linhas divididas em um DataFrame do Pandas
df = pd.DataFrame(linhas_divididas)

# Renomear as colunas
nomes_colunas = {
    0: 'Consulta Optantes',
    1: 'Data da consulta',
    2: 'Identificação do Contribuinte',
    3: 'CNPJ',
    4: '4',
    5: '5',
    6: 'Nome Empresarial',
    7: '7',
    8: 'Situação no Simples Nacional',
    9: 'Situação no SIMEI',
    10: '10',
    11: '11'
}

df = df.rename(columns=nomes_colunas)

# Selecionar colunas específicas
colunas_selecionadas = ['Data da consulta',
                        'CNPJ',
                        'Nome Empresarial',
                        'Situação no Simples Nacional',
                        'Situação no SIMEI']
df = df[colunas_selecionadas]

# Exportar o DataFrame para um arquivo Excel
nome_do_arquivo_excel = 'coleta\coleta_dividida.xlsx'
df.to_excel(nome_do_arquivo_excel, index=False)

print(f'DataFrame exportado para {nome_do_arquivo_excel} com sucesso!')