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
print('Programa finalizado!')