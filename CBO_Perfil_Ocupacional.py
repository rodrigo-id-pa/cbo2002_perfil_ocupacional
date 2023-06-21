### INCIANDO O SCRIPT ###

# bibliotecas padrão do python #
import importlib
import subprocess

# verificando se o pip está atualizado
def check_install_pip():
    try:
        import pip
        print('pip já está instalado.')
    except ImportError:
        print('pip não está instalado. Atualizando...')
        subprocess.check_call(['python.exe', '-m', 'pip', 'install', '--upgrade', 'pip'])
        print('pip atualizado com sucesso.')
check_install_pip()

# Lista de bibliotecas que você precisa verificar
bibliotecas = ['pandas', 'csv', 're', 'os']

# Verifica se cada biblioteca está instalada e, se necessário, instala
for biblioteca in bibliotecas:
    try:
        importlib.import_module(biblioteca)
        print(f'{biblioteca} já está instalada.')
    except ImportError:
        print(f'{biblioteca} não está instalada. Instalando...')
        subprocess.check_call(['pip', 'install', biblioteca])
        print(f'{biblioteca} instalada com sucesso.')

#importando as bibliotecas
import pandas as pd
import os
import csv
import re

# pegando ad do usuario e o pc
computerName = os.environ['USERNAME']
userName = os.environ['COMPUTERNAME']

# pegando os arquivos no Output
output_ = r'C:/Users/'+computerName+'/OneDrive - Firjan/'+userName+'/DIDAD/RPA 017 - CBO/Output/'

#listando os arquivos na pasta
arquivos_ = os.listdir(output_)

# montando o csv para leitura
csv_cbo = output_+arquivos_[3]

# criando uma lista
list_ = []

# abrindo o csv, criando os indices
with open(csv_cbo, 'r') as file:
    reader = csv.reader(file)
    for i, row in enumerate(reader):
        list_.append((i, row))
    print('exportado com sucesso.')

# percorrer cada indice e remover os ; por / dentro de parenteses
for i, row in list_:
    for j in range(len(row)):
        # Usando expressão regular para encontrar o padrão "(...)"
        match = re.search(r'\((.*?)\)', row[j])
        if match:
            # Substituindo os pontos e vírgulas por barras apenas dentro dos parênteses
            new_value = re.sub(r';', '/', match.group(1))
            # Atualizando o valor na lista original
            list_[i][1][j] = re.sub(r'\((.*?)\)', f'({new_value})', row[j])

# percorrendo a nova lista e removendo a linha com "coleta(bags;" pois ele não trocou ; por /
for item in list_:
    if any("coleta(bags;" in value for value in item[1]):
        list_.remove(item)
        new_index = len(list_)
        new_item = ['5;51;519;5192;519205;A;COLETAR MATERIAL RECICLÁVEL E REAPROVEITÁVEL;7;Fornecer recipientes para coleta de bags, conteineres, etc.']
        list_.append((new_index, new_item))
print(f'sucesso, total de linhas: {len(list_)}') 

# Removendo o índice da lista
df_list = [item[1] for item in list_]
df = pd.DataFrame(df_list)

# Juntar todas as colunas em uma única coluna
df['combined'] = df.apply(lambda row: ' '.join(map(str, row)), axis=1)

# Dropar as colunas de 0 até 9
df = df.drop(df.columns[0:10], axis=1)

# Substituir valores 'None' por espaços em branco na coluna 'combined'
df['combined'] = df['combined'].replace(r'\bNone\b', ' ', regex=True)

# Remover espaços em branco extras no final da string
df['combined'] = df['combined'].str.strip()

# Separar os dados por ponto e vírgula (;) e criar colunas separadas
df = df['combined'].str.split(';', expand=True)

# Definir a primeira coluna como o título
df.columns = df.iloc[0]

# Remover a primeira linha (título original)
df = df[1:]

# exportando o xlsx final
df.to_excel(f'{output_}cbo2002_perfil_ocupacional.xlsx', index=False)
print('excel exportado')

#verificando se existe valores nulos
df.info()