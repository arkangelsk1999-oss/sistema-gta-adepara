import sqlite3, re
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).parent / "data" / "gtv.db"

# ─────────────────────────────────────────────
# CONEXÃO
# ─────────────────────────────────────────────

def get_conn_gtv():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


# ─────────────────────────────────────────────
# INICIALIZAÇÃO DO BANCO
# ─────────────────────────────────────────────

def init_db_gtv():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = get_conn_gtv()
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS gtv (
            gtv_numero TEXT,
            cultura TEXT,
            quantidade_carga TEXT,
            medida TEXT,
            valor TEXT,
            codigo_up_origem TEXT,
            procedencia TEXT,
            situacao_pedido TEXT,
            municipio_origem TEXT,
            destinatario_nome TEXT,
            municipio_destino TEXT,
            data_emissao TEXT,
            emitida_por_nome TEXT,
            data_impressao TEXT,
            veiculo TEXT,
            transito TEXT,
            quantidade_chegada TEXT
        )
    """)

    c.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS gtv_fts USING fts5(
            procedencia,
            destinatario_nome,
            emitida_por_nome,
            gtv_numero,
            municipio_origem,
            municipio_destino,
            cultura,
            content='gtv',
            content_rowid='rowid'
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS arquivos_importados_gtv (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_arquivo TEXT NOT NULL UNIQUE,
            linhas INTEGER,
            importado_em TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()


# ─────────────────────────────────────────────
# CONTROLE DE ARQUIVOS IMPORTADOS
# ─────────────────────────────────────────────

def arquivo_ja_importado_gtv(nome):
    conn = get_conn_gtv()
    r = conn.execute(
        "SELECT id FROM arquivos_importados_gtv WHERE nome_arquivo=?", (nome,)
    ).fetchone()
    conn.close()
    return r is not None


def registrar_arquivo_gtv(nome, linhas):
    conn = get_conn_gtv()
    conn.execute(
        "INSERT OR IGNORE INTO arquivos_importados_gtv (nome_arquivo, linhas, importado_em) VALUES (?,?,?)",
        (nome, linhas, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()


def listar_arquivos_gtv():
    conn = get_conn_gtv()
    rows = [dict(r) for r in conn.execute(
        "SELECT * FROM arquivos_importados_gtv ORDER BY importado_em DESC"
    ).fetchall()]
    conn.close()
    return rows


# ─────────────────────────────────────────────
# MAPEAMENTO DE COLUNAS
# ─────────────────────────────────────────────

MAPA_COLUNAS_GTV = {
    "Número da GTV":        "gtv_numero",
    "Cultura":              "cultura",
    "Quantidade da Carga":  "quantidade_carga",
    "Medida":               "medida",
    "Valor R$":             "valor",
    "Cod.UP de Origem":     "codigo_up_origem",
    "Procedência":          "procedencia",
    "situacao_pedido":      "situacao_pedido",
    "Município de Origem":  "municipio_origem",
    "Destino":              "destinatario_nome",
    "Município de Destino": "municipio_destino",
    "Data Emissão":         "data_emissao",
    "Emissor":              "emitida_por_nome",
    "Data Impressão":       "data_impressao",
    "Veiculo":              "veiculo",
    "Trânsito":             "transito",
    "Quantidade de Chegada":"quantidade_chegada",
}

COLUNAS_DB_GTV = [
    "gtv_numero", "cultura", "quantidade_carga", "medida", "valor",
    "codigo_up_origem", "procedencia", "situacao_pedido",
    "municipio_origem", "destinatario_nome", "municipio_destino",
    "data_emissao", "emitida_por_nome", "data_impressao",
    "veiculo", "transito", "quantidade_chegada",
]


# ─────────────────────────────────────────────
# IMPORTAÇÃO DE DATAFRAME
# ─────────────────────────────────────────────

def importar_dataframe_gtv(df, nome_arquivo):
    # Renomeia colunas usando o mapa
    novas_colunas = []
    for col in df.columns:
        col_strip = col.strip()
        if col_strip in MAPA_COLUNAS_GTV:
            novas_colunas.append(MAPA_COLUNAS_GTV[col_strip])
        else:
            novas_colunas.append("_ignorar_" + col_strip[:20])

    df.columns = novas_colunas

    # Remove colunas indesejadas
    cols_manter = [c for c in df.columns if not c.startswith("_ignorar_")]
    df = df[cols_manter]

    # Garante todas as colunas do banco
    for col in COLUNAS_DB_GTV:
        if col not in df.columns:
            df[col] = None

    df = df[COLUNAS_DB_GTV]
    df = df.replace('', None)
    df = df.where(df.notna(), None)

    conn = get_conn_gtv()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")

    total = len(df)
    batch_size = 5000
    for start in range(0, total, batch_size):
        batch = df.iloc[start:start + batch_size]
        batch.to_sql('gtv', conn, if_exists='append', index=False)

    # Rebuild FTS
    conn.execute("INSERT INTO gtv_fts(gtv_fts) VALUES('rebuild')")
    conn.commit()
    conn.close()

    registrar_arquivo_gtv(nome_arquivo, total)
    return total


# ─────────────────────────────────────────────
# BUSCA
# ─────────────────────────────────────────────

def _fts_query_gtv(termo):
    tokens = [t for t in termo.upper().split() if len(t) > 2]
    if not tokens:
        return None
    return ' AND '.join(f'"{t}"' for t in tokens)


def _nome_confere_gtv(pesquisado, campo):
    if not pesquisado or not campo:
        return False
    tokens = [t for t in pesquisado.upper().split() if len(t) > 2]
    campo_upper = str(campo).upper()
    matches = sum(1 for t in tokens if t in campo_upper)
    return matches >= min(2, len(tokens))


def buscar_gtv(nome='', emissor='', mes_ini=None, mes_fim=None):
    conn = get_conn_gtv()
    nome    = nome.strip().upper()
    emissor = emissor.strip().upper()

    params = []
    sql_where = []

    fts_parts = []
    if nome:
        fts_q = _fts_query_gtv(nome)
        if fts_q:
            fts_parts.append(f"procedencia:{fts_q} OR destinatario_nome:{fts_q}")
    if emissor:
        fts_q_e = _fts_query_gtv(emissor)
        if fts_q_e:
            fts_parts.append(f"emitida_por_nome:{fts_q_e}")

    if fts_parts:
        fts_full = ' OR '.join(fts_parts)
        sql_where.append("g.rowid IN (SELECT rowid FROM gtv_fts WHERE gtv_fts MATCH ?)")
        params.append(fts_full)

    if mes_ini:
        sql_where.append("g.data_emissao >= ?")
        params.append(mes_ini + '-01')

    if mes_fim:
        sql_where.append("g.data_emissao <= ?")
        params.append(mes_fim + '-31')

    if not sql_where:
        conn.close()
        return {}

    sql = f"SELECT g.* FROM gtv g WHERE {' AND '.join(sql_where)} ORDER BY g.data_emissao"
    rows = conn.execute(sql, params).fetchall()
    conn.close()

    resultado = {'gtv': {'origem': [], 'destino': [], 'colunas': COLUNAS_DB_GTV}}

    for row in rows:
        dados = dict(row)
        is_origem = nome and _nome_confere_gtv(nome, dados.get('procedencia', ''))
        is_destino = nome and _nome_confere_gtv(nome, dados.get('destinatario_nome', ''))

        if is_destino and not is_origem:
            resultado['gtv']['destino'].append(dados)
        else:
            resultado['gtv']['origem'].append(dados)

    return resultado


def buscar_gtv_lai():
    conn = get_conn_gtv()
    rows = conn.execute("SELECT * FROM gtv ORDER BY data_emissao").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def stats_gtv():
    try:
        conn = get_conn_gtv()
        total = conn.execute("SELECT COUNT(*) FROM gtv").fetchone()[0]
        datas = conn.execute("SELECT MIN(data_emissao), MAX(data_emissao) FROM gtv").fetchone()
        arquivos = conn.execute("SELECT COUNT(*) FROM arquivos_importados_gtv").fetchone()[0]
        conn.close()
        return {
            'total': total,
            'data_ini': datas[0],
            'data_fim': datas[1],
            'arquivos': arquivos,
        }
    except:
        return {'total': 0, 'data_ini': '', 'data_fim': '', 'arquivos': 0}
