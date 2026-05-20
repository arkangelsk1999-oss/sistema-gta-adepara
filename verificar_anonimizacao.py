import sqlite3, json

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

# Verifica casos onde CPF aparece com 11 dígitos mas nome do estabelecimento não está vazio
r = conn.execute("""
    SELECT 
        json_extract(dados_json, '$."CPF ou CNPJ do produtor de destino"') as cpf_dest,
        json_extract(dados_json, '$."nome da estabelecimento de destino"') as nome_estab,
        ano
    FROM gtas 
    WHERE ano BETWEEN 2012 AND 2023
    AND json_extract(dados_json, '$."nome da estabelecimento de destino"') IS NOT NULL
    AND json_extract(dados_json, '$."nome da estabelecimento de destino"') != ''
    AND length(replace(replace(replace(json_extract(dados_json, '$."CPF ou CNPJ do produtor de destino"'), '.', ''), '-', ''), '/', '')) = 11
    LIMIT 10
""").fetchall()

print(f"Registros com CPF + nome estabelecimento: {len(r)}")
for row in r:
    print(f"  ANO: {row[2]} | CPF: {row[0]} | ESTAB: {row[1]}")

conn.close()