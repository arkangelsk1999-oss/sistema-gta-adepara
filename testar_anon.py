import sqlite3, json, math

def _limpar_nan(dados):
    return {k: (None if isinstance(v, float) and math.isnan(v) else v) for k, v in dados.items()}

def _nome_estabelecimento_anonimizado(dados, ano):
    cpf_val = dados.get('CPF ou CNPJ do produtor de destino', '')
    digitos = ''.join(filter(str.isdigit, str(cpf_val)))
    print(f"  CPF bruto: {repr(cpf_val)}")
    print(f"  Dígitos: {digitos} (len={len(digitos)})")
    print(f"  Anonimizar: {len(digitos) == 11}")
    return len(digitos) == 11

conn = sqlite3.connect('banco_gta.db')
conn.row_factory = None
r = conn.execute("SELECT dados_json FROM gtas WHERE ano=2012 AND json_extract(dados_json,'$.\"nome da estabelecimento de destino\"')='FAZENDA PARAISO DO CAPIM GROSSO II' LIMIT 1").fetchone()
conn.close()

dados = _limpar_nan(json.loads(r[0]))
print("=== TESTE ===")
_nome_estabelecimento_anonimizado(dados, 2012)