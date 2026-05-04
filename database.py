import sqlite3, os, re, json, math
from pathlib import Path

DB_PATH = Path(__file__).parent / "banco_gta.db"

PADROES_ORIG_CPF  = ['cpf ou cnpj do produtor de origem', 'origem_identificacao']
PADROES_DEST_CPF  = ['cpf ou cnpj do produtor de destino', 'destinatario_identificacao']
PADROES_ORIG_NOME = ['nome do produtor de origem', 'origem_nome']
PADROES_DEST_NOME = ['nome do produtor de destino', 'destinatario_nome']

def detectar_col(colunas, padroes):
    cols_lower = {c.lower().strip(): c for c in colunas}
    for p in padroes:
        if p in cols_lower:
            return cols_lower[p]
    return None

def norm_cpf(val):
    if not val: return ''
    s = str(val).strip()
    if re.match(r'^[\d,\.]+[Ee][+\-]?\d+$', s):
        try: return str(int(float(s.replace(',','.'))))
        except: return s
    return re.sub(r'[^\d]', '', s)

def get_conn():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        cpf TEXT NOT NULL UNIQUE,
        email TEXT NOT NULL UNIQUE,
        senha_hash TEXT NOT NULL,
        orgao TEXT NOT NULL,
        nivel TEXT NOT NULL DEFAULT 'usuario',
        ativo INTEGER NOT NULL DEFAULT 1,
        criado_em TEXT NOT NULL,
        ultimo_acesso TEXT
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS gtas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ano INTEGER NOT NULL,
        arquivo TEXT NOT NULL,
        orig_cpf TEXT,
        orig_nome TEXT,
        dest_cpf TEXT,
        dest_nome TEXT,
        dados_json TEXT NOT NULL
    )''')

    c.execute('CREATE INDEX IF NOT EXISTS idx_orig_cpf  ON gtas(orig_cpf)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_dest_cpf  ON gtas(dest_cpf)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_ano       ON gtas(ano)')

    c.execute('''CREATE VIRTUAL TABLE IF NOT EXISTS gtas_fts
        USING fts5(
            orig_nome,
            dest_nome,
            content='gtas',
            content_rowid='id',
            tokenize='unicode61 remove_diacritics 1'
        )
    ''')

    c.execute('''CREATE TRIGGER IF NOT EXISTS gtas_ai
        AFTER INSERT ON gtas BEGIN
            INSERT INTO gtas_fts(rowid, orig_nome, dest_nome)
            VALUES (new.id, new.orig_nome, new.dest_nome);
        END
    ''')
    c.execute('''CREATE TRIGGER IF NOT EXISTS gtas_ad
        AFTER DELETE ON gtas BEGIN
            INSERT INTO gtas_fts(gtas_fts, rowid, orig_nome, dest_nome)
            VALUES ('delete', old.id, old.orig_nome, old.dest_nome);
        END
    ''')
    c.execute('''CREATE TRIGGER IF NOT EXISTS gtas_au
        AFTER UPDATE ON gtas BEGIN
            INSERT INTO gtas_fts(gtas_fts, rowid, orig_nome, dest_nome)
            VALUES ('delete', old.id, old.orig_nome, old.dest_nome);
            INSERT INTO gtas_fts(rowid, orig_nome, dest_nome)
            VALUES (new.id, new.orig_nome, new.dest_nome);
        END
    ''')

    c.execute('''CREATE TABLE IF NOT EXISTS arquivos_importados (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome_arquivo TEXT NOT NULL UNIQUE,
        ano INTEGER,
        linhas INTEGER,
        importado_em TEXT NOT NULL
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS auditoria (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data_hora TEXT NOT NULL,
        usuario_id INTEGER,
        usuario_nome TEXT,
        usuario_login TEXT,
        orgao TEXT,
        ip TEXT,
        localidade TEXT,
        cpf_cnpj_pesquisado TEXT,
        nome_pesquisado TEXT,
        total_resultados INTEGER
    )''')

    from werkzeug.security import generate_password_hash
    from datetime import datetime
    existe = c.execute("SELECT id FROM usuarios WHERE nivel='founder'").fetchone()
    if not existe:
        c.execute('''INSERT INTO usuarios 
            (nome, cpf, email, senha_hash, orgao, nivel, ativo, criado_em)
            VALUES (?,?,?,?,?,?,?,?)''', (
            'Gilliard Costa Rodrigues',
            '00000000000',
            'founder@arkangelsk.com',
            generate_password_hash('Arkangelsk@2025'),
            'Arkangelsk',
            'founder',
            1,
            datetime.now().isoformat()
        ))

    conn.commit()
    conn.close()


def migrar_fts():
    conn = get_conn()
    c = conn.cursor()
    print("Verificando se migração FTS já foi feita...")
    total_fts  = c.execute("SELECT COUNT(*) FROM gtas_fts").fetchone()[0]
    total_gtas = c.execute("SELECT COUNT(*) FROM gtas").fetchone()[0]

    if total_fts >= total_gtas:
        print(f"FTS já populado ({total_fts} registros). Nada a fazer.")
        conn.close()
        return

    print(f"Populando FTS5 com {total_gtas:,} registros. Aguarde...")
    c.execute("INSERT INTO gtas_fts(rowid, orig_nome, dest_nome) SELECT id, orig_nome, dest_nome FROM gtas")
    conn.commit()
    print("✅ Migração FTS5 concluída!")
    conn.close()


def arquivo_ja_importado(nome):
    conn = get_conn()
    r = conn.execute("SELECT id FROM arquivos_importados WHERE nome_arquivo=?", (nome,)).fetchone()
    conn.close()
    return r is not None

def registrar_arquivo(nome, ano, linhas):
    from datetime import datetime
    conn = get_conn()
    conn.execute(
        "INSERT OR IGNORE INTO arquivos_importados (nome_arquivo,ano,linhas,importado_em) VALUES (?,?,?,?)",
        (nome, ano, linhas, datetime.now().isoformat())
    )
    conn.commit()
    conn.close()

def importar_dataframe(df, ano, nome_arquivo):
    cols = list(df.columns)
    col_orig_cpf  = detectar_col(cols, PADROES_ORIG_CPF)
    col_dest_cpf  = detectar_col(cols, PADROES_DEST_CPF)
    col_orig_nome = detectar_col(cols, PADROES_ORIG_NOME)
    col_dest_nome = detectar_col(cols, PADROES_DEST_NOME)

    conn = get_conn()
    c = conn.cursor()
    lote = []

    for _, row in df.iterrows():
        orig_cpf  = norm_cpf(row.get(col_orig_cpf,  '')) if col_orig_cpf  else ''
        dest_cpf  = norm_cpf(row.get(col_dest_cpf,  '')) if col_dest_cpf  else ''
        orig_nome = str(row.get(col_orig_nome, '')).upper().strip() if col_orig_nome else ''
        dest_nome = str(row.get(col_dest_nome, '')).upper().strip() if col_dest_nome else ''
        dados     = json.dumps(row.to_dict(), ensure_ascii=False, default=str)
        lote.append((ano, nome_arquivo, orig_cpf, orig_nome, dest_cpf, dest_nome, dados))

        if len(lote) >= 5000:
            c.executemany(
                "INSERT INTO gtas (ano,arquivo,orig_cpf,orig_nome,dest_cpf,dest_nome,dados_json) VALUES (?,?,?,?,?,?,?)",
                lote
            )
            conn.commit()
            lote = []

    if lote:
        c.executemany(
            "INSERT INTO gtas (ano,arquivo,orig_cpf,orig_nome,dest_cpf,dest_nome,dados_json) VALUES (?,?,?,?,?,?,?)",
            lote
        )
        conn.commit()
    conn.close()
    return len(df)


def _limpar_nan(dados):
    return {k: (None if isinstance(v, float) and math.isnan(v) else v) for k, v in dados.items()}


def _fts_query(nome):
    tokens = [t for t in nome.upper().split() if len(t) > 2]
    if not tokens:
        return None
    return ' AND '.join(f'"{t}"' for t in tokens)


def buscar_gtas(nome='', cpf='', ano_ini=None, ano_fim=None):
    conn = get_conn()
    nome = nome.strip().upper()
    cpf  = norm_cpf(cpf)

    rows = []

    if nome and cpf:
        fts_q = _fts_query(nome)
        if fts_q:
            sql = """
                SELECT g.* FROM gtas g
                WHERE g.id IN (SELECT rowid FROM gtas_fts WHERE gtas_fts MATCH ?)
                  AND (g.orig_cpf = ? OR g.dest_cpf = ?)
            """
            params = [fts_q, cpf, cpf]
            if ano_ini: sql += " AND g.ano >= ?"; params.append(int(ano_ini))
            if ano_fim: sql += " AND g.ano <= ?"; params.append(int(ano_fim))
            sql += " ORDER BY g.ano"
            rows = conn.execute(sql, params).fetchall()

    elif nome:
        fts_q = _fts_query(nome)
        if not fts_q:
            conn.close()
            return {}
        sql = """
            SELECT g.* FROM gtas g
            WHERE g.id IN (SELECT rowid FROM gtas_fts WHERE gtas_fts MATCH ?)
        """
        params = [fts_q]
        if ano_ini: sql += " AND g.ano >= ?"; params.append(int(ano_ini))
        if ano_fim: sql += " AND g.ano <= ?"; params.append(int(ano_fim))
        sql += " ORDER BY g.ano"
        rows = conn.execute(sql, params).fetchall()

    elif cpf:
        sql = "SELECT * FROM gtas WHERE (orig_cpf = ? OR dest_cpf = ?)"
        params = [cpf, cpf]
        if ano_ini: sql += " AND ano >= ?"; params.append(int(ano_ini))
        if ano_fim: sql += " AND ano <= ?"; params.append(int(ano_fim))
        sql += " ORDER BY ano"
        rows = conn.execute(sql, params).fetchall()

    else:
        conn.close()
        return {}

    conn.close()

    resultado = {}
    for row in rows:
        ano   = row['ano']
        dados = _limpar_nan(json.loads(row['dados_json']))
        is_orig = (nome and nome in (row['orig_nome'] or '')) or (cpf and cpf == row['orig_cpf'])
        is_dest = (nome and nome in (row['dest_nome'] or '')) or (cpf and cpf == row['dest_cpf'])

        if ano not in resultado:
            resultado[ano] = {'origem': [], 'destino': [], 'colunas': list(dados.keys())}
        if is_orig:
            resultado[ano]['origem'].append(dados)
        if is_dest and not is_orig:
            resultado[ano]['destino'].append(dados)

    return resultado


def registrar_auditoria(usuario, ip, localidade, cpf_pesquisado, nome_pesquisado, total):
    from datetime import datetime
    conn = get_conn()
    conn.execute('''INSERT INTO auditoria 
        (data_hora,usuario_id,usuario_nome,usuario_login,orgao,ip,localidade,cpf_cnpj_pesquisado,nome_pesquisado,total_resultados)
        VALUES (?,?,?,?,?,?,?,?,?,?)''', (
        datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        usuario.get('id'),
        usuario.get('nome'),
        usuario.get('email'),
        usuario.get('orgao'),
        ip,
        localidade,
        cpf_pesquisado,
        nome_pesquisado,
        total
    ))
    conn.commit()
    conn.close()