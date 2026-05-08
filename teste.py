from database import buscar_gtas
import json

r = buscar_gtas(nome='MERCURIO')
print('Anos:', sorted(r.keys()))

for ano in sorted(r.keys()):
    try:
        json.dumps(r[ano])
        print(f'{ano}: OK')
    except Exception as e:
        print(f'{ano}: ERRO - {e}')