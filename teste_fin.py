import sqlite3
conn = sqlite3.connect(r'C:\Users\57216615\sistema-gta-adepara\banco_gta.db')
rows = conn.execute(
    "SELECT DISTINCT json_extract(dados_json, '$.\"Finalidade do transporte\"') AS fin FROM gtas WHERE fin IS NOT NULL LIMIT 10"
).fetchall()
print(rows)
conn.close()