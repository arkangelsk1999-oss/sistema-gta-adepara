import sqlite3, json

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

anos = [row[0] for row in conn.execute("SELECT DISTINCT ano FROM gtas ORDER BY ano").fetchall()]

for ano in anos:
    row = conn.execute(f"SELECT dados_json FROM gtas WHERE ano={ano} LIMIT 1").fetchone()
    if row:
        dados = json.loads(row[0])
        colunas = list(dados.keys())
        print(f"\n=== ANO {ano} ({len(colunas)} colunas) ===")
        for i, col in enumerate(colunas, 1):
            print(f"  {i:3}. {col}")

conn.close()