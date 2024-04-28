import time
from datetime import datetime
import pyautogui as pg
import pandas as pd
import clipboard
import tkinter as tk
from tkinter import messagebox

# Registrar o momento de início da coleta de dados
now = datetime.now()
inicio_coleta = now.strftime("%d%m%Y_%H%M%S")

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

nome_do_arquivo_excel = 'coleta\coleta_' + str(inicio_coleta) + '.xlsx'
df.to_excel(nome_do_arquivo_excel, index=False)

print(f'DataFrame exportado para {nome_do_arquivo_excel} com sucesso!')

# Exibir uma mensagem de pop-up indicando que o programa foi finalizado
messagebox.showinfo("Programa Finalizado", "O programa foi finalizado com sucesso!")
