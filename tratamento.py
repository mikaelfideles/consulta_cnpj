# COLETOR CNPJ SIMPLES NACIONAL

# Bibliotecas
import time
from datetime import datetime
import pyautogui as pg
import pandas as pd
import clipboard
import tkinter as tk
from tkinter import messagebox

print('\nTRATAMENTO CNPJ SIMPLES NACIONAL')
print('\n////////////////////////////////////////////////////////////')
# time.sleep(10)

# Criar uma janela Tkinter
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Registrar o momento de início da coleta de dados
now = datetime.now()
inicio_coleta = now.strftime("%d%m%Y_%H%M%S")

# Ler o arquivo Excel
consulta = pd.read_excel(r'02_coleta_temp\coleta_temp.xlsx', dtype=str)

# Lista para armazenar as linhas divididas em colunas
linhas_divididas = []

# Função para verificar se uma linha está vazia
def is_not_empty(line):
    return any(column.strip() for column in line)

# Iterar sobre as linhas da consulta
for i, linha in consulta.iterrows():
    # Verificar se a linha não está vazia
    if pd.notna(linha['COLETA']) and is_not_empty(linha['COLETA'].split('\n')):
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
    0: 'CONSULTA_OPTANTES',
    1: 'DATA_CONSULTA',
    2: 'IDENTIFICACAO_CONTRIBUINTE',
    3: 'CNPJ',
    4: '4',
    5: '5',
    6: 'NOME',
    7: '7',
    8: 'SITUACAO_SIMPLES_NACIONAL',
    9: 'SITUACAO_SIMEI',
    10: '10',
    11: '11'
}

df = df.rename(columns=nomes_colunas)

# Selecionar colunas específicas
colunas_selecionadas = ['DATA_CONSULTA',
                        'CNPJ',
                        'NOME',
                        'SITUACAO_SIMPLES_NACIONAL',
                        'SITUACAO_SIMEI']

df = df[colunas_selecionadas]

nome_do_arquivo_excel = r'03_coleta\coleta_' + str(inicio_coleta) + '.xlsx'
df.to_excel(nome_do_arquivo_excel, index=False)

print(f'\nArquivo gerado com sucesso ...\{nome_do_arquivo_excel}\n')

# Exibir uma mensagem de pop-up indicando que o programa foi finalizado
messagebox.showinfo("MENSAGEM", 'Programa executado com sucesso!')

# Fechar a janela Tkinter
root.destroy()
