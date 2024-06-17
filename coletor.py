# COLETOR CNPJ SIMPLES NACIONAL

# Bibliotecas
import time
from datetime import datetime
import pyautogui as pg
import pandas as pd
import clipboard
import tkinter as tk
from tkinter import messagebox

print('\nCOLETOR CNPJ SIMPLES NACIONAL')
print('\n////////////////////////////////////////////////////////////')
time.sleep(10)

# Criar uma janela Tkinter
root = tk.Tk()
root.withdraw()  # Esconder a janela principal

# Ler o arquivo Excel com os dados das consultas, adicionar coluna 'COLETA' e substituir planilha
consulta = pd.read_excel(r'01_consulta\consulta.xlsx', dtype=str).fillna('')
consulta['COLETA'] = ''
consulta.to_excel(r'02_coleta_temp\coleta_temp.xlsx', index=False)
consulta = pd.read_excel(r'02_coleta_temp\coleta_temp.xlsx', dtype=str).fillna('')

# Registrar o momento de início da coleta de dados
now = datetime.now()
inicio_coleta = now.strftime("%d%m%Y_%H%M%S")

# Iterar através de cada linha do DataFrame 'consulta'
for i, linha in consulta.iterrows():

    # Atualizar página e aguardar 1 segundo
    pg.press('f5')
    time.sleep(1)

    # Print da etapa de consulta
    print(f'\nCONSULTA: {i + 1} de {len(consulta)}')
    print('CNPJ: ' + linha['NUM_CNPJ'])
  
    # Pressiona a tecla "tab" por n vezes
    presses = 7
    for _ in range(presses):
        pg.press('tab')    
    
    # Escrever o CNPJ na caixa de texto
    pg.write(linha['NUM_CNPJ'])
    time.sleep(0.1)

    # Pressionar "tab" e "enter" com intervalo de 0.1 segundo
    pg.press('tab')
    time.sleep(0.1)
    pg.press('enter')

    # Aguardar 3 segundos para carregamento da próxima página
    time.sleep(3)

    # Selecionar e copiar a data da consulta
    pg.hotkey('ctrl', 'a')
    pg.hotkey('ctrl', 'c')
    
    # Atribuir os dados copiados à coluna 'COLETA' da linha atual
    linha['COLETA'] = clipboard.paste()
    print(linha['COLETA'])
    print('////////////////////////////////////////////////////////////')

    # Verificar se a mensagem "Informe um CNPJ válido." está presente
    if "Informe um CNPJ válido." in linha['COLETA']:
        linha['COLETA'] = (
            "Consulta Optantes\n"
            "Data da consulta: 00/00/0000 00:00:00\n"
            "Identificação do Contribuinte - CNPJ Matriz\n"
            #"CNPJ: 00.000.000/0000-00\n"
            f"CNPJ: {linha['NUM_CNPJ']}\n"
            "A opção pelo Simples Nacional e/ou SIMEI abrange todos os estabelecimentos da empresa\n"
            "\n"
            "Nome Empresarial: <cnpj_invalido>\n"
            "Situação Atual\n"
            "Situação no Simples Nacional: <cnpj_invalido>\n"
            "Situação no SIMEI: <cnpj_invalido>\n"
        )
        continue  # Ir para a próxima iteração do loop

    # Enquanto a quantidade de caracteres copiados for menor ou igual a 100, repetir o processo
    while len(linha['COLETA']) <= 100:
       
        # Aguardar 1 segundo antes de repetir o processo
        time.sleep(1)
        
        # Atualizar página e aguardar 1 segundo
        pg.press('f5')
        time.sleep(1)

        # Print da etapa de consulta
        print(f'\nCONSULTA: {i + 1} de {len(consulta)}')
        print('CNPJ: ' + linha['NUM_CNPJ'])
    
        # Pressiona a tecla "tab" por n vezes
        presses = 7
        for _ in range(presses):
            pg.press('tab')       
        
        # Escrever o CNPJ na caixa de texto
        pg.write(linha['NUM_CNPJ'])
        time.sleep(0.1)

        # Pressionar "tab" e "enter" com intervalo de 0.1 segundo
        pg.press('tab')
        time.sleep(0.1)
        pg.press('enter')

        # Aguardar 3 segundos para carregamento da próxima página
        time.sleep(3)

        # Selecionar e copiar a data da consulta
        pg.hotkey('ctrl', 'a')
        pg.hotkey('ctrl', 'c')
            
        # Atualizar os dados copiados na coluna 'COLETA'
        linha['COLETA'] = clipboard.paste()
        print(linha['COLETA'])
        print('////////////////////////////////////////////////////////////')

        # Verificar novamente se a mensagem "Informe um CNPJ válido." está presente
        if "Informe um CNPJ válido." in linha['COLETA']:
            linha['COLETA'] = (
                "Consulta Optantes\n"
                "Data da consulta: 00/00/0000 00:00:00\n"
                "Identificação do Contribuinte - CNPJ Matriz\n"
                #"CNPJ: 00.000.000/0000-00\n"
                f"CNPJ: {linha['NUM_CNPJ']}\n"
                "A opção pelo Simples Nacional e/ou SIMEI abrange todos os estabelecimentos da empresa\n"
                "\n"
                "Nome Empresarial: <cnpj_invalido>\n"
                "Situação Atual\n"
                "Situação no Simples Nacional: <cnpj_invalido>\n"
                "Situação no SIMEI: <cnpj_invalido>\n"
            )
            break  # Sair do loop while e ir para a próxima iteração do loop for

    time.sleep(0.1)

    nome_do_arquivo_excel = r'02_coleta_temp\coleta_temp.xlsx'
    consulta.to_excel(nome_do_arquivo_excel, index=False)

#---------------------------------

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
