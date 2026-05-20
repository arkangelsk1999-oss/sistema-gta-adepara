import sqlite3
import re

def norm_cpf(val):
    if not val: return ''
    s = str(val).strip()
    if re.match(r'^[\d,\.]+[Ee][+\-]?\d+$', s):
        try:
            s = str(int(float(s.replace(',','.'))))
        except:
            return s
    s = re.sub(r'[^\d]', '', s)
    if not s:
        return ''
    s_sem_zero = s.lstrip('0') or '0'
    if len(s_sem_zero) <= 11:
        return s_sem_zero.zfill(11)
    else:
        return s_sem_zero.zfill(14)

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None

print("Buscando CPFs para corrigir...")
rows = conn.execute("SELECT id, orig_cpf, dest_cpf FROM gtas WHERE LENGTH(orig_cpf) NOT IN (0,11,14) OR LENGTH(dest_cpf) NOT IN (0,11,14)").fetchall()
print(f"Encontrados {len(rows)} registros para corrigir...")

for row in rows:
    id_, orig, dest = row
    novo_orig = norm_cpf(orig)
    novo_dest = norm_cpf(dest)
    conn.execute("UPDATE gtas SET orig_cpf=?, dest_cpf=? WHERE id=?", (novo_orig, novo_dest, id_))

conn.commit()
print("Concluído!")
conn.close()