import sqlite3

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

r = conn.execute("SELECT COUNT(*) FROM gtas WHERE LENGTH(orig_cpf) > 11 AND LENGTH(orig_cpf) < 14").fetchone()
print('CPFs orig com tamanho errado:', r[0])

r2 = conn.execute("SELECT COUNT(*) FROM gtas WHERE LENGTH(dest_cpf) > 11 AND LENGTH(dest_cpf) < 14").fetchone()
print('CPFs dest com tamanho errado:', r2[0])

conn.close()