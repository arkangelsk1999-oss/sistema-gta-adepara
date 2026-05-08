import pandas as pd
import os
from database import init_db, importar_dataframe, registrar_arquivo, arquivo_ja_importado

init_db()

pasta = r'C:\Users\57216615\Desktop\Relatórios de GTAs'

arquivos = [
    ('2010.xlsx', 2010), ('2011.xlsx', 2011), ('2012.xlsx', 2012),
    ('2013.xlsx', 2013), ('2014.xlsx', 2014), ('2015.xlsx', 2015),
    ('2016.xlsx', 2016), ('2017.csv', 2017), ('2018.csv', 2018),
    ('2019.csv', 2019), ('2020.csv', 2020), ('2021.csv', 2021),
    ('2022.xlsx', 2022), ('2023.xlsx', 2023), ('2024.xlsx', 2024),
    ('2025.xlsx', 2025),
]

for nome, ano in arquivos:
    caminho = os.path.join(pasta, nome)
    if not os.path.exists(caminho):
        print(f'NAO ENCONTRADO: {nome}')
        continue
    if arquivo_ja_importado(nome):
        print(f'JA IMPORTADO: {nome}')
        continue
    print(f'Importando {nome}...')
    try:
        if nome.endswith('.csv'):
            df = pd.read_csv(caminho, encoding='latin-1', sep=';', dtype=str, low_memory=False)
        else:
            abas = pd.read_excel(caminho, sheet_name=None, dtype=str)
            df = pd.concat(abas.values(), ignore_index=True)
        linhas = importar_dataframe(df, ano, nome)
        registrar_arquivo(nome, ano, linhas)
        print(f'OK: {nome} — {linhas:,} linhas importadas')
    except Exception as e:
        print(f'ERRO em {nome}: {e}')

print('Concluido!')