import time
from datetime import datetime
import pyautogui as pg
import pandas as pd
import clipboard
import tkinter as tk
from tkinter import messagebox

# Criar uma janela Tkinter
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Ler o arquivo Excel com os dados das consultas, adicionar coluna 'COLETA' e substituir planilha
consulta = pd.read_excel(r'consulta.xlsx', dtype=str).fillna('')
consulta['COLETA'] = ''
consulta.to_excel(r'coleta_temp\coleta_temp.xlsx', index=False)
consulta = pd.read_excel(r'coleta_temp\coleta_temp.xlsx', dtype=str).fillna('')

# Registrar o momento de início da coleta de dados
now = datetime.now()
inicio_coleta = now.strftime("%d%m%Y_%H%M%S")

# Aguardar 10 segundos para iniciar coleta
time.sleep(10)
print('Iniciando coleta...')    

# Iterar através de cada linha do DataFrame 'consulta'
for i, linha in consulta.iterrows():

    # Atualizar página
    pg.press('f5')

    # Aguardar 2 segundos
    time.sleep(2)

    # Exibir informações relevantes
    print('-----------------------------------------')
    print(f'CONSULTA: {i + 1} de {len(consulta)}')
    print('CNPJ: ' + linha['NUM_CNPJ'])
  
    # Pressiona a tecla "tab" por n vezes
    presses = 6
    for _ in range(presses):
        pg.press('tab')    
    
    # Escrever o CNPJ na caixa de texto
    pg.write(linha['NUM_CNPJ'])
    time.sleep(0.1)

    # Pressionar "tab" e "enter" com intervalo de 0.1 segundo
    pg.press('tab')
    time.sleep(0.1)
    pg.press('enter')

    # Aguardar 2 segundos para carregamento da próxima página
    time.sleep(2)

     # Selecionar e copiar a data da consulta
    pg.hotkey('ctrl', 'a')
    pg.hotkey('ctrl', 'c')
    
    # Atribuir os dados copiados à coluna 'COLETA' da linha atual
    linha['COLETA'] = clipboard.paste()

    # Enquanto a quantidade de caracteres copiados for menor ou igual a 60, repetir o processo
    while len(linha['COLETA']) <= 100:
       
        # Aguardar 1 segundo antes de repetir o processo
        time.sleep(1)
        
        # Atualizar página
        pg.press('f5')

        # Aguardar 2 segundos
        time.sleep(2)

        # Exibir informações relevantes
        print('-----------------------------------------')
        print(linha['QTDE_CONSULTA'])
        print('num_doc: ' + linha['NUM_CNPJ'])
    
        # Pressiona a tecla "tab" por n vezes
        presses = 6
        for _ in range(presses):
            pg.press('tab')       
        
        # Escrever o CNPJ na caixa de texto
        pg.write(linha['NUM_CNPJ'])
        time.sleep(0.1)

        # Pressionar "tab" e "enter" com intervalo de 0.1 segundo
        pg.press('tab')
        time.sleep(0.1)
        pg.press('enter')

        # Aguardar 2 segundos para carregamento da próxima página
        time.sleep(2)

        # Selecionar e copiar a data da consulta
        pg.hotkey('ctrl', 'a')
        pg.hotkey('ctrl', 'c')
            
        # Atualizar os dados copiados na coluna 'COLETA'
        linha['COLETA'] = clipboard.paste()

    print(linha['COLETA'])
    time.sleep(0.1)

    nome_do_arquivo_excel = r'coleta_temp\coleta_temp.xlsx'
    consulta.to_excel(nome_do_arquivo_excel, index=False)

# Exibir uma mensagem indicando que o programa foi finalizado
print('Coleta finalizada!')

'---------------------------------------------------'

# Ler o arquivo Excel
consulta = pd.read_excel(r'coleta_temp\coleta_temp.xlsx', dtype=str)

# Lista para armazenar as linhas divididas em colunas
linhas_divididas = []

# Iterar sobre as linhas da consulta
for i, linha in consulta.iterrows():
    # Dividir a linha em partes com base no caractere de quebra de linha '\n'
    partes = linha['COLETA'].split('\n')

    # Remover o texto "Data da consulta: " das linhas pertencentes à coluna "Data da consulta"
    partes[1] = partes[1].replace("Data da consulta: ", "")
    partes[3] = partes[3].replace("CNPJ: ", "").replace(".", "").replace("/", "").replace("-", "")
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

nome_do_arquivo_excel = r'coleta\coleta_' + str(inicio_coleta) + '.xlsx'
df.to_excel(nome_do_arquivo_excel, index=False)

print(f'DataFrame exportado para {nome_do_arquivo_excel} com sucesso!')

# Exibir uma mensagem de pop-up indicando que o programa foi finalizado
messagebox.showinfo("Programa Finalizado", "O programa foi finalizado com sucesso!")

# Fechar a janela Tkinter
root.destroy()