import sqlite3, json
from pathlib import Path

DB_PATH = Path(__file__).parent / "banco_gta.db"

PADROES_EMISSOR = [
    'emitida_por_nome',
    'usuário emissor',
    'usuario emissor',
    'ususario emissor',
]

def detectar_emissor(dados):
    keys_lower = {k.lower().strip(): k for k in dados.keys()}
    for p in PADROES_EMISSOR:
        if p in keys_lower:
            val = dados[keys_lower[p]]
            if val and str(val).strip() and str(val).strip().upper() != 'NAN':
                return str(val).strip().upper()
    return 'SEM_EMISSOR'

def popular():
    conn = sqlite3.connect(str(DB_PATH))
    c = conn.cursor()

    total = c.execute("SELECT COUNT(*) FROM gtas WHERE emissor_nome IS NULL").fetchone()[0]
    print(f"Registros para atualizar: {total:,}")

    if total == 0:
        print("Nada a fazer! Todos já foram processados.")
        conn.close()
        return

    LOTE = 10000
    atualizados = 0

    while True:
        rows = c.execute(
            "SELECT id, dados_json FROM gtas WHERE emissor_nome IS NULL LIMIT ?",
            (LOTE,)
        ).fetchall()

        if not rows:
            break

        lote = []
        for row_id, dados_json in rows:
            try:
                dados = json.loads(dados_json)
                emissor = detectar_emissor(dados)
            except:
                emissor = 'SEM_EMISSOR'
            lote.append((emissor, row_id))

        c.executemany("UPDATE gtas SET emissor_nome = ? WHERE id = ?", lote)
        conn.commit()

        atualizados += len(rows)
        total_restante = c.execute("SELECT COUNT(*) FROM gtas WHERE emissor_nome IS NULL").fetchone()[0]
        pct = ((total - total_restante) / total) * 100
        print(f"Progresso: {atualizados:,} processados | Restam: {total_restante:,} ({pct:.1f}%)")

    print("Reconstruindo FTS5...")
    c.execute("INSERT INTO gtas_fts(gtas_fts) VALUES ('rebuild')")
    conn.commit()
    print("✅ Concluído!")
    conn.close()

popular()