import sqlite3

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

print("Corrigindo orig_cpf...")
conn.execute("UPDATE gtas SET orig_cpf = '0' || orig_cpf WHERE LENGTH(orig_cpf) = 10")

print("Corrigindo dest_cpf...")
conn.execute("UPDATE gtas SET dest_cpf = '0' || dest_cpf WHERE LENGTH(dest_cpf) = 10")

conn.commit()
print("Concluído!")
conn.close()