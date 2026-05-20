import sqlite3

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

print("Exemplos orig_cpf com tamanho errado:")
r = conn.execute("SELECT orig_cpf, LENGTH(orig_cpf) FROM gtas WHERE LENGTH(orig_cpf) > 11 AND LENGTH(orig_cpf) < 14 LIMIT 10").fetchall()
for row in r:
    print(row)

print("\nExemplos dest_cpf com tamanho errado:")
r2 = conn.execute("SELECT dest_cpf, LENGTH(dest_cpf) FROM gtas WHERE LENGTH(dest_cpf) > 11 AND LENGTH(dest_cpf) < 14 LIMIT 10").fetchall()
for row in r2:
    print(row)

conn.close()